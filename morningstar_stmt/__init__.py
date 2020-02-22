import os
import time
import logging
import selenium
import morningstar_stmt.tickerlist
from selenium import webdriver
from typing import Tuple
from pathlib import Path

class WaitException(Exception):
    """Wait Failed"""
    pass

class MorningStarStmtBrowser(object):
    def __init__(self, download_dir='.', log_dir='.', log_level=logging.INFO):
        self.download_dir = download_dir
        self.__log_dir = log_dir
        self.__temp_dir = os.path.join(download_dir, '.temp')
        Path(self.__temp_dir).mkdir(parents=True, exist_ok=True)

        options = webdriver.ChromeOptions() 
        options.add_experimental_option("prefs", {
        "download.default_directory": self.__temp_dir,
        "download.prompt_for_download": False,
        "profile.default_content_setting_values.automatic_downloads": 1
        })
        self.browser = webdriver.Chrome(chrome_options=options)

        fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        fh = logging.FileHandler(os.path.join(log_dir, 'morningstar-stmt.log'))
        fh.setLevel(log_level)
        fh.setFormatter(fmt)
        ch = logging.StreamHandler()
        ch.setLevel(log_level)
        ch.setFormatter(fmt)
        self.logger = logging.getLogger('morningstar-stmt')
        self.logger.setLevel(log_level)
        self.logger.addHandler(fh)
        self.logger.addHandler(ch)

    
    def login(self, username, password):
        self.browser.get('https://www.morningstar.com/sign-in')
        self.browser.find_element_by_xpath('//input[@name="userName"]').send_keys(username)
        self.browser.find_element_by_xpath('//input[@name="password"]').send_keys(password)
        self.browser.find_element_by_xpath('//label[@class="mdc-checkbox mds-form__checkbox"]').click()
        self.browser.find_element_by_xpath('//button[@type="submit"]').click()
    

    def download_stmt(self, market, ticker) -> Tuple[int, str]:
        if market is "" or ticker is "":
            self.logger.warning('missing market={} or ticker={}'.format(market, ticker))
            return 1, 'Missing Market or Ticker'

        self.__clean_temp_dir()
        try:
            self.browser.get('https://www.morningstar.com/stocks/{}/{}/financials'.format(market, ticker))

            if 'Page Not Found | Morningstar' == self.browser.title:
                self.logger.info('market[{}] ticker[{}] financials page not found!'.format(market, ticker))
                return 1, 'Page Not Found'

            self.__wait_click('//a[@class="mds-link ng-binding" and contains(text(), "Income Statement")]')
            self.__wait_click('//div[@class="sal-financials-details__exportSection ng-scope"]//button[@class="sal-financials-details__export mds-button mds-button--small"]')
            self.__wait_click('//mds-button[@value="Balance Sheet"]')
            self.__wait_click('//div[@class="sal-financials-details__exportSection ng-scope"]//button[@class="sal-financials-details__export mds-button mds-button--small"]')
            self.__wait_click('//mds-button[@value="Cash Flow"]')
            self.__wait_click('//div[@class="sal-financials-details__exportSection ng-scope"]//button[@class="sal-financials-details__export mds-button mds-button--small"]')
        except selenium.common.exceptions.ElementClickInterceptedException:
            self.logger.error('market[{}] ticker[{}] download failed - click error'.format(market, ticker))
            self.browser.refresh()
            return -1, 'Download Failed - Click Error'
        except WaitException:
            return -1, 'Download Failed - Timeout'
        
        annual_balance = 'Balance Sheet_Annual_As Originally Reported.xls'
        annual_income = 'Income Statement_Annual_As Originally Reported.xls'
        annual_cash = 'Cash Flow_Annual_As Originally Reported.xls'

        balance_file = os.path.join(self.__temp_dir, annual_balance)
        income_file = os.path.join(self.__temp_dir, annual_income)
        cash_file = os.path.join(self.__temp_dir, annual_cash)
        self.logger.debug('market[{}] ticker[{}] wait for downloading finish'.format(market, ticker))
        while os.path.exists(balance_file) is False or os.path.exists(income_file) is False or os.path.exists(cash_file) is False:
            time.sleep(1)

        os.rename(balance_file, os.path.join(self.download_dir, '{}_{}'.format(ticker, annual_balance)))
        os.rename(income_file, os.path.join(self.download_dir, '{}_{}'.format(ticker, annual_income)))
        os.rename(cash_file, os.path.join(self.download_dir, '{}_{}'.format(ticker, annual_cash)))
        
        return 0, 'Successful'
    

    def __clean_temp_dir(self):
        for f in os.listdir(self.__temp_dir):
            try:
                os.remove(os.path.join(self.__temp_dir, f))
            except FileNotFoundError:
                continue


    def __wait_click(self, xpath, timeout=30, retry=False):
        self.logger.debug('wait click xpath<{}>'.format(xpath))
        wait_sec = 0
        ret = False
        while True:
            try:
                self.browser.find_element_by_xpath(xpath).click()
                return True
            except selenium.common.exceptions.NoSuchElementException:
                time.sleep(1)
                wait_sec += 1 
                if wait_sec % timeout == 0:
                    if retry:
                        self.logger.debug('retry wait click {}'.format(xpath))
                        self.browser.refresh()
                    else:
                        raise WaitException
                        break
                continue

        if ret is False:
            self.logger.error('wait click failed xpath<{}>'.format(xpath))
            raise WaitException

        return ret


class MorningstarAccount(object):
    username = ''
    password = ''
    def __init__(self, username, password):
        self.username = username
        self.password = password


def download_all_stmt(download_dir='.', log_dir='.', log_level=logging.DEBUG, account:MorningstarAccount=None):
    browser = MorningStarStmtBrowser(download_dir=download_dir, log_dir=log_dir, log_level=log_level)
    if account is not None:
        browser.login(account.username, account.password)
    
    done_list = {}
    if os.path.exists(os.path.join(browser.download_dir, 'done')):
        with open(os.path.join(browser.download_dir, 'done'), 'r', errors = 'ignore') as f:
            done_list = {x.strip() : 1 for x in f.readlines()}

    with open(os.path.join(browser.download_dir, 'done'), 'a+') as done_file:
        current = len(done_list)
        total = len(tickerlist.all)
        for market, ticker in tickerlist.all:
            current += 1
            if ticker in done_list:
                browser.logger.info('skip {} {}'.format(market, ticker))
                continue

            ret, res = browser.download_stmt(market, ticker)
            while ret < 0:
                browser.logger.warning('retry {} {}'.format(market, ticker))
                ret, res = browser.download_stmt(market, ticker)

            browser.logger.info('{} {} {} ({}/{})'.format(market, ticker, res, current, total))
            done_file.write("{}\n".format(ticker))
            done_file.flush()