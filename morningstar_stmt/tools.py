import os
import xlrd
import csv
import json
import re
import subprocess
import morningstar_stmt as ms
import time
import code
import glob
from lxml import html


def xls2csv(file):
    try:
        wb = xlrd.open_workbook(file)
        sh = wb.sheet_by_name('sheet1')
        fn = file.split('_', 2)
        suffix = 'balance.csv' if fn[-1] == 'Balance Sheet_Annual_As Originally Reported.xls' else 'cash.csv' if fn[-1] == 'Cash Flow_Annual_As Originally Reported.xls' else 'income.csv'
        your_csv_file = open(os.path.join('output', '{}_{}_{}'.format(fn[0], fn[1], suffix)), 'w')
        wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
        for rownum in range(sh.nrows):
            wr.writerow(sh.row_values(rownum))
        your_csv_file.close()
    except xlrd.biffh.XLRDError:
        print('failed_{}'.format(file))

def csv2json(file):
    data = {}
    with open(file) as csv_file:
        csv_reader = csv.DictReader(csv_file)
        for row in csv_reader:
            col = [c for c in row]
            key = row['Name'].strip()
            for dcol in col[1:]:
                value = int(row[dcol].replace(',', '')) if row[dcol] is not '' else None
                data[dcol] = data.get(dcol, {key : value})
                data[dcol].update({key : value})
    return data


def annual_stmt(ticker_list=ms.tickerlist.all_quote, from_annual=1970, save_file=True):
    b=ms.MorningStarStmtBrowser()
    done_cnt = 0
    all_cnt = len(ticker_list)
    ret_val = {}
    for item in ticker_list:
        market, ticker, _ = item.split(',', 2)
        output = '{}_{}_stmt.json'.format(market, ticker)
        if os.path.exists(output):
            #print('skip {}'.format(output))
            done_cnt += 1
            continue

        for file in glob.glob('*{}*.xls'.format(ticker)):
            os.remove(file)

        done_cnt += 1
        ret, _ = b.download_stmt(market, ticker)
        while ret < 0:
            print('retry {} {}'.format(market, ticker))
            ret, res = b.download_stmt(market, ticker)

        if ret != 0:
            print('download failed {}'.format(output))
            continue
        
        data = {}
        for xls_file in glob.glob('*{}*.xls'.format(ticker)):
            stmt_type = xls_file.split('_', 3)[2]
            stmt_type = 'balance' if stmt_type == 'Balance Sheet' else 'cash' if stmt_type == 'Cash Flow' else 'income'
            try:
                wb = xlrd.open_workbook(xls_file)
                sh = wb.sheet_by_name('sheet1')
            except:
                continue
            annuals = sh.row_values(0)
            for i in range(1, sh.nrows):
                row = sh.row_values(i)
                key = row[0].strip()
                for i in range(1, len(row)):
                    annual = annuals[i]
                    if annual != 'TTM' and int(annual) <= int(from_annual):
                        continue
                    value = int(row[i].replace(',', '')) if row[i] is not '' else None
                    annual_data = data.get(annual, {stmt_type: {key: value}})
                    type_data = annual_data.get(stmt_type, {key: value})
                    type_data.update({key: value})
                    annual_data.update({stmt_type: type_data})
                    data.update({annual: annual_data})

        for file in glob.glob('*{}*.xls'.format(ticker)):
            os.remove(file)

        if save_file:
            with open(output, 'w') as f:
                f.writelines(json.dumps(data, indent=4))
        else:
            ret_val[ticker] = data

    b.close()
    return ret_val


def key_ratio_annual_fixture():
    fixture = []
    pk = 0
    for fn in os.listdir('.'):
        if fn.endswith('.json') is False:
            continue

        data = {}
        market, code, _ = fn.split('_', 2)
        with open(fn, 'r') as f:
            v = json.load(f)
            for key, ratios in v['data'].items():
                if key == 'growth':
                    continue
                for annual, ratio_data in ratios.items():
                    annual = 'TTM' if annual == 'Latest Qtr' or annual == 'TTM' else annual.split('-')[0]
                    data[annual] = data.get(annual, {key : ratio_data})
                    data[annual].update({key : ratio_data})

        for annual, data in data.items():
            pk += 1
            fixture.append({
                'model' : 'tools.keyratioannual',
                'pk' : pk,
                'fields' : {
                    'code' : code,
                    'market' : market,
                    'annual' : annual,
                    'data' : data
                }
            })

    with open('keyratio_annual_fixture.json', 'w') as output:
        output.write(json.dumps(fixture, indent=4))


def make_django_annual_fixture():
    res = subprocess.check_output("find . -type f -iname '*.csv' | awk -F'[_|/]' '{print $2\"_\"$3}' | sort | uniq", shell=True)
    res = res.decode('utf-8')
    res = res.split('\n')
    fixture = []
    pk = 0
    for prefix in res:
        if prefix == '':
            continue
        data = {}
        print(prefix)
        market, code = prefix.split('_', 1)
        types = ['balance', 'income', 'cash']
        for type in types:
            fn = '{}_{}.json'.format(prefix, type)
            if os.path.exists(fn):
                with open(fn, 'r') as f:
                    v = json.load(f)
                for annual, stmt in v.items():
                    data[annual] = data.get(annual, {type : stmt})
                    data[annual].update({type : stmt})
        for annual, stmts in data.items():
            pk += 1
            fixture.append({
                'model' : 'tools.financialannualstmt_ms',
                'pk' : pk,
                'fields' : {
                    'code' : code,
                    'market' : market,
                    'annual' : annual,
                    'stmt' : stmts
                }
            })

    with open('fixture_annual.json', 'w') as output:
        output.write(json.dumps(fixture, indent=4))


def make_django_fixture():
    fixture = []
    pk = 0
    for f in sorted(os.listdir('.')):
        if f.endswith('.json'):
            market, code, type, _ = re.split('_|\.', f, 3)
            with open(f, 'r') as data_file:
                data = json.load(data_file)
            for annual, stmt in data.items():
                pk += 1
                fixture.append({
                    'model' : 'tools.financialstmt_ms',
                    'pk' : pk,
                    'fields' : {
                        'code' : code,
                        'market' : market,
                        'type' : type,
                        'annual' : annual,
                        'stmt' : stmt
                    }
                })
    
    with open('fixture.json', 'w') as output:
        output.write(json.dumps(fixture, indent=4))


def all_xls2csv():
    for f in os.listdir('.'):
        if f.endswith('.xls'):
            xls2csv(f)

def all_csv2json():
    for f in os.listdir('.'):
        if f.endswith('.csv'):
            data = csv2json(f)
            pre, ext = os.path.splitext(f)
            with open('{}.json'.format(pre), 'w') as output:
                output.write(json.dumps(data))


def all_quote_code():
    b=ms.MorningStarStmtBrowser()
    done = {}
    with open('all_quote_code', 'r') as f:
        for item in f.readlines():
            market, ticker, code = item.split(',', 2)
            done['{}{}'.format(market, ticker)] = code

    with open('all_quote_code', 'a+') as f:
        for item in ms.tickerlist.xase:
            market, ticker = item.split(',', 1)
            if '{}{}'.format(market, ticker) in done:
                continue

            b.browser.get('https://www.morningstar.com/stocks/{}/{}/quote'.format(market, ticker))
            tree=html.fromstring(b.browser.page_source)
            code = tree.xpath('//sal-components[@tab="stocks-quote"]/@share-class-id')
            if len(code) == 0:
                code = ''
            else:
                code = code[0]

            f.writelines('\'{},{},{}\'\n'.format(market, ticker, code))
            f.flush()


def all_key_ratio():
    b=ms.MorningStarStmtBrowser()
    all_cnt = len(ms.tickerlist.all_quote)
    done_cnt = 0
    for item in ms.tickerlist.all_quote:
        market, ticker, code = item.split(',', 2)
        output = '{}_{}_keyratio.json'.format(market, ticker)
        if code is '':
            #print('skip {} {} beacuse code is empty'.format(market, ticker))
            done_cnt += 1
            continue

        if os.path.exists(output):
            #print('skip {}'.format(output))
            done_cnt += 1
            continue

        print('downloading... {} ({}/{})'.format(output, done_cnt, all_cnt))
        done_cnt += 1
        b.browser.get('https://financials.morningstar.com/ratios/r.html?t={}&culture=en&platform=sal'.format(code))
        tree=html.fromstring(b.browser.page_source)
        wait_time = 0
        while len(tree) < 0 or len(tree.xpath('//table')) < 8:
            time.sleep(1)
            tree=html.fromstring(b.browser.page_source)
            wait_time += 1
            if wait_time >= 10:
                break

        if wait_time >= 10:
            print('download failed {}'.format(output))
            continue

        financials = key_ratio_table_to_json(tree, 'financials', '{} {}')
        profitability = key_ratio_table_to_json(tree, 'tab-profitability', '{} pr-margins {}')
        growth = key_ratio_growth_table_to_json(tree)
        cash = key_ratio_table_to_json(tree, 'tab-cashflow', '{} cf-cashflow {}')
        health = key_ratio_table_to_json(tree, 'tab-financial', '{} fh-balsheet {}')
        efficiency = key_ratio_table_to_json(tree, 'tab-efficiency', '{} ef-efficiency {}')
        data = {
            'ticker' : ticker,
            'market' : market,
            'data' : {
                'financials' : financials,
                'profitability' : profitability,
                'growth' : growth,
                'cash' : cash,
                'health' : health,
                'efficiency' : efficiency
            }
        }
        with open(output, 'w') as f:
            f.writelines(json.dumps(data, indent=4))


def key_ratio_growth_table_to_json(tree):
    tables = tree.xpath('//div[@id="tab-growth"]/table')
    if len(tables) == 0:
        print('No tab-growth table found')
        return 'ERROR'

    output = {}
    for table in tables:

        t_tx = table.xpath('.//tbody//th[@scope="row" and @align="left"]/text()')
        t_id = table.xpath('.//tbody//th[@scope="row" and @align="left"]/@id')
        tabs = [(t_tx[i], t_id[i]) for i in range(0, len(t_id))]
        c_tx = table.xpath('.//thead//th[@scope="col" and @align="right"]/text()')
        c_id = table.xpath('.//thead//th[@scope="col" and @align="right"]/@id')
        cols = [(c_tx[i], c_id[i]) for i in range(0, len(c_id))]
        r_tx = table.xpath('.//tbody//th[@scope="row" and @class="row_lbl"]/text()')
        r_id = table.xpath('.//tbody//th[@scope="row" and @class="row_lbl"]/@id')
        rows = [(r_tx[i], r_id[i]) for i in range(0, len(r_id))]
        for tab in tabs:
            t_tx, t_id = tab
            for col in cols:
                c_tx, c_id = col
                for row in rows:
                    r_tx, r_id = row
                    header_link = ' '.join([c_id, t_id, r_id]) if t_id is not '' else ' '.join([c_id, r_id])
                    values = table.xpath('.//tbody//td[@headers="{}"]/text()'.format(header_link))
                    if len(values) == 0:
                        continue

                    value = values[0]
                    output_of_t = output.get(t_tx, {})
                    output_of_t[c_tx] = output_of_t.get(c_tx, {r_tx.strip() : value})
                    output_of_t[c_tx].update({r_tx.strip() : value})
                    output.update({t_tx : output_of_t})
    
    return output
    
    pass


def key_ratio_table_to_json(tree, tab_id, header_link_format):

    tables = tree.xpath('//div[@id="{}"]/table'.format(tab_id))
    if len(tables) == 0:
        print('No {} table found'.format(tab_id))
        return 'ERROR'

    output = {}
    for table in tables:
        t_id = table.xpath('.//thead//th[@scope="col" and @align="left"]/@id')
        t_id = t_id[0] if len(t_id) > 0 else ''
        t_tx = table.xpath('.//thead//th[@scope="col" and @align="left"]/text()')
        t_tx = t_tx[0] if len(t_tx) > 0 else ''
        c_tx = table.xpath('.//thead//th[@scope="col" and @align="right"]/text()')
        c_id = table.xpath('.//thead//th[@scope="col" and @align="right"]/@id')
        cols = [(c_tx[i], c_id[i]) for i in range(0, len(c_id))]
        r_tx = table.xpath('.//tbody//th[@scope="row"]/text()')
        r_id = table.xpath('.//tbody//th[@scope="row"]/@id')
        rows = [(r_tx[i], r_id[i]) for i in range(0, len(r_id))]
        for col in cols:
            c_tx, c_id = col
            for row in rows:
                r_tx, r_id = row
                header_link = ' '.join([c_id, t_id, r_id]) if t_id is not '' else ' '.join([c_id, r_id])
                values = table.xpath('.//tbody//td[@headers="{}"]/text()'.format(header_link))
                value = values[0] if len(values) > 0 else ''
                #if t_tx is not '':
                #    output_of_t = output.get(t_tx, {})
                #    output_of_t[c_tx] = output_of_t.get(c_tx, {r_tx.strip() : value})
                #    output_of_t[c_tx].update({r_tx.strip() : value})
                #    output.update({t_tx : output_of_t})
                #else:
                output[c_tx] = output.get(c_tx, {r_tx.strip() : value})
                output[c_tx].update({r_tx.strip() : value})
    
    return output


def valuation_data_download(ticker_list=[]):
    b=ms.MorningStarStmtBrowser()
    done_cnt = 0
    l = ticker_list if len(ticker_list) > 0 else ms.tickerlist.all
    all_cnt = len(l)
    for item in l:
        market, ticker = item.split(',', 1)
        output = '{}_{}_valuation.json'.format(market, ticker)
        if os.path.exists(output):
            #print('skip {}'.format(output))
            done_cnt += 1
            continue

        done_cnt += 1
        b.browser.get('https://www.morningstar.com/stocks/{}/{}/valuation'.format(market, ticker))
        if 'Page Not Found | Morningstar' == b.browser.title or 'Error | Morningstar' == b.browser.title:
            print('Page not found {} ({}/{})'.format(output, done_cnt, all_cnt))
            continue

        tree=html.fromstring(b.browser.page_source)
        wait_time = 0
        while len(tree) < 0 or len(tree.xpath('//div[@ng-show="hasViewableData"]//table//tr')) < 5:
            time.sleep(1)
            tree=html.fromstring(b.browser.page_source)
            wait_time += 1
            if wait_time >= 10:
                break
            if len(tree.xpath('//div[@ng-show="hasViewableData"]//table//tr')) > 3:
                break

        if 'There is no Valuation data available.' in tree.xpath('//span[@class="ng-binding"]/text()'):
        #if len(tree.xpath('//table/tr')) < 5:
            print('No data {} ({}/{})'.format(output, done_cnt, all_cnt))
            continue

        if wait_time >= 10:
            print('download failed {}'.format(output))
            continue

        data = valuation_data_to_json(tree)

        if data != 'ERROR' or data == '':
            with open(output, 'w') as f:
                f.writelines(json.dumps(data, indent=4))
            print('download {} successful ({}/{})'.format(output, done_cnt, all_cnt))
        else:
            print('download {} failed ({}/{})'.format(output, done_cnt, all_cnt))
  

    
def valuation_data_to_json(tree):
    ret = {}
    tables = tree.xpath('.//table[@class="report-table ng-isolate-scope"]')
    if len(tables) == 0:
        return 'ERROR'

    table = tables[0]
    trs = table.xpath('.//tr')
    if len(trs) < 2:
        return 'ERROR'

    tdt = trs[0].xpath('.//td/text()')
    if len(tdt) == 0:
        return 'ERROR'
    
    annuals = tdt[1:]
    for tr in trs[2:]:
        tds = tr.xpath('.//td')
        if len(tds) < 2:
            continue
        tht = tds[0].xpath('.//span/text()')
        key = tht[0].strip() if len(tht) > 0 else ''
        if key == '':
            continue
        i = 0
        for td in tds[1:]:
            tdt = td.xpath('.//span/text()')
            value = tdt[0].strip().replace('\u2014', '') if len(tdt) > 0 else ''
            annual = annuals[i]
            ret[annual] = ret.get(annual, {key:value})
            ret[annual].update({key:value})
            i += 1

    return ret

def key_valuation_csv():
    lines = []
    pk = 0
    for fn in os.listdir('.'):
        if fn.endswith('.json') is False:
            continue

        data = {}
        market, code, _ = fn.split('_', 2)
        with open(fn, 'r') as f:
            v = json.load(f)
            for annual, data in v.items():
                pk += 1
                ps = float(data['Price/Sales'].replace(',', '')) if 'Price/Sales' in data and data['Price/Sales'] != '' else 0.0
                pe = float(data['Price/Earnings'].replace(',', '')) if 'Price/Earnings' in data and data['Price/Earnings'] != '' else 0.0
                pc = float(data['Price/Cash Flow'].replace(',', '')) if 'Price/Cash Flow' in data and data['Price/Cash Flow'] != '' else 0.0
                pb = float(data['Price/Book'].replace(',', '')) if 'Price/Book' in data and data['Price/Book'] != '' else 0.0
                lines.append('{},{},{},{:.2f},{:.2f},{:.2f},{:.2f}'.format(code, market, annual, ps, pe, pc, pb))

    with open('valuation.csv', 'w') as output:
        output.writelines('code,market,annual,ps,pe,pc,pb')
        for line in lines:
            output.writelines(line + '\n')


def key_valuation_fixture():
    fixture = []
    pk = 0
    for fn in os.listdir('.'):
        if fn.endswith('.json') is False:
            continue

        data = {}
        market, code, _ = fn.split('_', 2)
        with open(fn, 'r') as f:
            v = json.load(f)
            for annual, data in v.items():
                pk += 1
                fixture.append({
                    'model' : 'tools.valuation',
                    'pk' : pk,
                    'fields' : {
                        'code' : code,
                        'market' : market,
                        'annual' : annual,
                        'data' : data,
                        'ps' : float(data['Price/Sales'].replace(',', '')) if 'Price/Sales' in data and data['Price/Sales'] != '' else float(0.0),
                        'pe' : float(data['Price/Earnings'].replace(',', '')) if 'Price/Earnings' in data and data['Price/Earnings'] != '' else float(0.0),
                        'pc' : float(data['Price/Cash Flow'].replace(',', '')) if 'Price/Cash Flow' in data and data['Price/Cash Flow'] != '' else float(0.0),
                        'pb' : float(data['Price/Book'].replace(',', '')) if 'Price/Book' in data and data['Price/Book'] != '' else float(0.0),
                    }
                })

    with open('valuation_fixture.json', 'w') as output:
        output.write(json.dumps(fixture, indent=4))


#test code
#b=ms.MorningStarStmtBrowser()
#b.browser.get('https://financials.morningstar.com/ratios/r.html?t=0P00011HZW&culture=en&platform=sal')
#tree=html.fromstring(b.browser.page_source)
#while len(tree) < 0 or len(tree.xpath('//table')) < 8:
#    time.sleep(1)
#    tree=html.fromstring(b.browser.page_source)
#code.interact(local=locals())

#def main():
    #all_xls2csv()
    #csv2json('xase_AAMC_balance.csv')
    #all_csv2json()
    #make_django_fixture()
    #make_django_annual_fixture()
    #all_quote_code()
    #all_key_ratio()
    #key_ratio_annual_fixture()
    #valuation_data_download()
    #key_valuation_fixture()
    #key_valuation_csv()
    #annual_stmt()

#if __name__ == "__main__":
#    main()

