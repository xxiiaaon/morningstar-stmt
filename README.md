morningstar-stmt
===============

#### Morningstar Financial Statement Downloader

morningstar-stmt is a simple Python module for downloading financial statement from [www.morningstar.com](http://www.morningstar.com/).

- Current Version: v0.0.2
- Version Released: 02/23/2020

Overview
--------
**morningstar-stmt** depends on [Selenium](https://selenium-python.readthedocs.io/installation.html) to simulate browser operations by following steps: _(goto) ticker's page_ -> _(click) statment tab_ -> _(click) export excel_ to download financial statments.

Features:

 - Support account login.
 - Full stock symbol list.
 - Download financial statmenst in xls format.
 - Only support Chrome webdriver.

Installation
-------------
- morningstar-stmt runs on Python >= 3.6.
- The package depends on [Selenium](https://selenium-python.readthedocs.io/installation.html) to work.

##### 1. Install morningstar-stmt using pip:

    $ pip install morningstar-stmt



##### 2. Install [Chrome webdriver](https://sites.google.com/a/chromium.org/chromedriver/downloads) and place it in **/usr/bin** or **/usr/local/bin**.


Usage Examples
--------------
### 1. Download financial statement for specific ticker

    import morningstar_stmt as ms
    browser = ms.MorningStarStmtBrowser()
    browser.download_stmt('xnas', 'AAPL')
    
##### Login Morningstar account before download

    browser.login('xxiiaaon', 'password')
    browser.download_stmt('xnas', 'AAPL')
    # market code xnas for NASDAQ and xnys for NYSE

##### Specify download directory and log level

    browser = ms.MorningStarStmtBrowser(download_dir='/User/xxiiaaon/stmt', log_level=logging.Info)
    
    
### 2. Download financial statements with ticker list

    import morningstar_stmt as ms
    ms.download_stmt(['xnas,AAPL', 'xnax,AAPL')])
    
##### Use predefined list

    import morningstar_stmt as ms
    from morningstar_stmt import tickerlist as tl
    ms.download_stmt(tl.all)
    # other list like tl.xnas for NASDAQ and tl.xnys for NYSE
    
##### Login when download with ticker list

    ms.download_all_stmt(ms.MorningstarAccount('xxiiaaon', 'password'), tickerlist.all)
