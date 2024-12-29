import requests
import smtplib
import copy
import pandas as pd

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from fake_useragent import UserAgent
from email.message import EmailMessage
from datetime import date

# Configure fake user agents
ua = UserAgent()
user_agent = ua.random
options = webdriver.ChromeOptions()
options.add_argument(f'--user-agent={user_agent}')

# Use headless browser
options.add_argument("--headless=new")

MARKET_CAP = 250
PE_TTM = 25
P_CF = 15
RELATIVE_PE = .7
ROA = .1
OPERATING_MARGIN = .25
NET_PROFIT_MARGIN = .1
ROIC = .1
CAGR = .1
TOTAL_DEBT_EQUITY = 2
CURRENT_RATIO = 1
SHORT_FLOAT = .05
PAYOUT_RATIO = .2

companies = {}
FINVIZ_URLS = ['https://finviz.com/screener.ashx?v=111&f=cap_microover,exch_nasd,fa_netmargin_o10,fa_roa_o10,sh_short_low&ft=4&r=','https://finviz.com/screener.ashx?v=111&f=cap_microover,exch_nyse,fa_netmargin_o10,fa_roa_o10,sh_short_low&ft=4&r=']

# Start driver
driver = webdriver.Chrome(options=options)

# Scrape NASDAQ & NYSE based on initial criteria
for url in FINVIZ_URLS:
    
    last = False
    r = 1

    # Iterate through urls until last result page 
    while not last:
        
        current_url = url + str(r)
        driver.get(current_url)
        driver.set_window_size(1920, 1080)
            
        html = driver.page_source
        soup = BeautifulSoup(html)

        # Iterate through table
        table = soup.find('table', class_='styled-table-new is-rounded is-tabular-nums w-full screener_table')
        for row in table.find_all('tr'):

            # Scrape ticker symbols
            for element in row.find_all('a', class_='tab-link'):
                if element.text not in companies:
                    ticker = element.text
                    companies[ticker] = {'valid': True}
                else:
                    last = True

            # Scrape market caps
            for i, element in enumerate(row.find_all('a')):
                if i == 6 and not last:
                    companies[ticker]['market_cap'] = element.text

        r += 20

# Exclude companies under $250M
for c in companies:
    if companies[c]['market_cap'][-1] == 'M':
        if float(companies[c]['market_cap'][:-1]) < MARKET_CAP:
            companies[c]['valid'] = False

driver.quit()

YAHOO_URL_ONE = 'https://finance.yahoo.com/quote/'
YAHOO_URL_TWO = '/key-statistics'

driver = webdriver.Chrome(options=options)
driver.set_window_size(1920, 1080)

for i, company in enumerate(companies):
    if companies[company]['valid']:
        ua = UserAgent()
        user_agent = ua.random
        headers = {"User-Agent": user_agent}
        driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent":user_agent})

        # Yahoo summary page
        summary_url = YAHOO_URL_ONE + company
        page = requests.get(summary_url, headers=headers)
        soup = BeautifulSoup(page.content)

        # Company Name
        try:
            company_name = ''
            company_name = soup.find_all('h1')[1]
            company_name = company_name.text if company_name else company_name
        except:
            pass

        # PE Ratio
        pe = soup.find('fin-streamer',attrs={"data-field": "trailingPE"})
        pe = pe.text if pe else pe
        try:
            pe = float(pe)
            companies[company]['valid'] = False if pe >= PE_TTM else companies[company]['valid']
        except:
            companies[company]['valid'] = False
        
        companies[company]['name'] = company_name
        companies[company]['pe'] = pe
        companies[company]['url'] = summary_url

        # Yahoo statistics page
        if companies[company]['valid']:
            
            stats_url = summary_url + YAHOO_URL_TWO
            driver.get(stats_url)
            
            # Operating Margin
            operating_margin = 0
            try:
                operating_margin = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nimbus-app"]/section/section/section/article/article/div/section[1]/div/section[2]/table/tbody/tr[2]/td[2]')))
                operating_margin = operating_margin.text[:-1] if operating_margin else operating_margin
                operating_margin = round(float(operating_margin) / 100, 2)
                companies[company]['valid'] = False if operating_margin <= OPERATING_MARGIN else companies[company]['valid']
            except:
                companies[company]['valid'] = False
            companies[company]['operating_margin'] = operating_margin
            
            # Total Debt Equity
            total_debt_equity = 3
            try:
                total_debt_equity = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nimbus-app"]/section/section/section/article/article/div/section[1]/div/section[5]/table/tbody/tr[4]/td[2]')))
                total_debt_equity = total_debt_equity.text[:-1] if total_debt_equity else total_debt_equity
                total_debt_equity = round(float(total_debt_equity) / 100, 2)
                companies[company]['valid'] = False if total_debt_equity >= TOTAL_DEBT_EQUITY else companies[company]['valid']
            except:
                companies[company]['valid'] = False
            companies[company]['total_debt_equity'] = total_debt_equity
            
            # Current Ratio
            current_ratio = 0
            try:
                current_ratio = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nimbus-app"]/section/section/section/article/article/div/section[1]/div/section[5]/table/tbody/tr[5]/td[2]')))
                current_ratio = current_ratio.text if current_ratio else current_ratio
                current_ratio = round(float(current_ratio), 2)
                companies[company]['valid'] = False if current_ratio <= CURRENT_RATIO else companies[company]['valid']
            except:
                companies[company]['valid'] = False
            companies[company]['current_ratio'] = current_ratio
            
            # Short % of Float
            short_float = 1
            try:
                short_float = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nimbus-app"]/section/section/section/article/article/div/section[2]/div/section[2]/table/tbody/tr[10]/td[2]')))
                short_float = short_float.text[:-1] if short_float else short_float
                short_float = round(float(short_float) / 100, 2)
                companies[company]['valid'] = False if short_float >= SHORT_FLOAT else companies[company]['valid']
            except:
                companies[company]['valid'] = False
            companies[company]['short_float'] = short_float
            
            # Payout Ratio
            payout_ratio = 1
            try:
                payout_ratio = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="nimbus-app"]/section/section/section/article/article/div/section[2]/div/section[3]/table/tbody/tr[6]/td[2]')))
                payout_ratio = payout_ratio.text[:-1] if payout_ratio else payout_ratio
                payout_ratio = round(float(payout_ratio) / 100, 2)
                companies[company]['valid'] = False if payout_ratio >= PAYOUT_RATIO else companies[company]['valid']
            except:
                companies[company]['valid'] = False
            companies[company]['payout_ratio'] = payout_ratio
    
    print(companies[company])
    print(f'{i+1} / {len(companies)} companies screened')

driver.quit()

company_copy = copy.deepcopy(companies)
valid_output = {} 

for c in company_copy:
    
    # Add valid companies to new dict
    if company_copy[c]['valid']:
        valid_output[c] = company_copy[c]

        # Convert floats to formatted strings
        for key in valid_output[c]:
            if key in ['operating_margin', 'total_debt_equity', 'short_float', 'payout_ratio']:
                if valid_output[c][key] not in ['-', '--']:
                    valid_output[c][key] = str(round(valid_output[c][key] * 100, 2)) + '%'

df = pd.DataFrame(data=valid_output).transpose()
excel_output = f'{date.today()} Stock Screener.xlsx'
df.to_excel(excel_output)

# Create a secure SSL context
PW = 'umtw vgqo wgdn egsg'
sender_email = 'liamstockscreener1@gmail.com'
receiver_emails = ['mags.liam@gmail.com', 'patrickmaguiremd@gmail.com']

# Send to list of receivers
for email in receiver_emails:
        msg = EmailMessage()
        msg['Subject'] = f'Stock Screener: {date.today()}'
        msg['From'] = 'Leeham'
        msg['To'] = email

        with open(excel_output, 'rb') as f:
                file_data = f.read()
                msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=excel_output)

        MyServer = smtplib.SMTP('smtp.gmail.com', 587)
        MyServer.starttls()
        MyServer.login(sender_email, PW)
        MyServer.send_message(msg)
        MyServer.quit()