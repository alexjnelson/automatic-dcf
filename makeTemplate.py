import datetime as dt
import json
import traceback
from time import sleep

import pandas as pd
import requests
import xlsxwriter as xls
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# TO DO:
# HOTFIX: fix bug causing summaries to be read wrong

# Make a "create config" script that lets the user customize the inputs to the config
# The user MUST pass a ticker, and has the options to:
# 1. allow a list of peers to be automatically generated
# 2. if the user chooses not to generate peers, they must pass a list of peers themselves
# 3. the user can set the RFR; it defaults to scraping Yahoo for the 30Y treasury yield
# 4. the user can set the MRP; it defaults to 5.5% (given by https://www.statista.com/statistics/664840/average-market-risk-premium-usa/)
# 5. the user can set the terminal growth; it defaults to the cagr of the gdp projection into 2050 by PWC (https://en.wikipedia.org/wiki/List_of_countries_by_past_and_projected_GDP_(nominal)#Long_term_GDP_estimates)
# create a script to run application, which checks that the file doesn't exist before overwriting

# Format the column widths in the DCF
# Clean up dcf implementation
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.89 Safari/537.36',
    'Cache-Control': 'no-cache'
}

options = webdriver.ChromeOptions()
options.add_argument('User-Agent="Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36"')
options.add_argument('Cache-Control="no-cache"')
options.add_argument('--headless')
options.add_argument("--log-level=3")

message_to_analyst = open('message.txt').read()


def get_statement(ticker, statement_name, driver: webdriver.Chrome):
    url = f'https://finance.yahoo.com/quote/{ticker}/{statement_name}?p={ticker}'

    driver.get(url)
    sleep(1)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f"//*[text() = 'Expand All']"))
    ).click()

    html = driver.execute_script('return document.body.innerHTML;')
    soup = BeautifulSoup(html, 'lxml')

    table = soup.find('section', {'data-test': 'qsp-financial'})

    items = []
    periods = []

    # determine how many line items there are, so they can be made into rows
    for item in table.find_all('span', 'Va(m)'):
        items.append(item.text)

    # determine how many periods there are, so they can be made into columns
    for data in (table_data := table.find_all('div', 'Ta(c)')):
        try:
            data['data-test']
        except:
            try:
                periods.append(dt.datetime.strptime(data.text, '%m/%d/%Y').strftime('%d %B %Y'))
            except ValueError:
                periods.append(data.text.upper().strip())

    # create a dataframe based on the determined dimensions of the table
    df = pd.DataFrame(columns=periods, index=items)

    # now iteratively read the data into the table
    for i, data in enumerate(table_data[len(periods):]):
        # data is read across the line item, ie period decreases until the next line item
        # the line wraps when all periods are filled out, so "i" divided by number of periods
        # determines the row number and "i" mod the number of periods determines the column number
        try:
            df.iloc[int(i / len(periods)), i % len(periods)] = float(data.text.strip().replace(',', ''))
        except ValueError:
            df.iloc[int(i / len(periods)), i % len(periods)] = data.text.strip()

    return df


# gets the width of every column in the dataframe for autofitting cells
def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]


# get the col letter from its int value
def colnum_string(n):
    n += 1
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def make_financials(ticker: str, book: xls.Workbook, driver: webdriver.Chrome):
    statements = ['financials', 'balance-sheet', 'cash-flow']  # "financials" means income statement
    names = ['Income Statement', 'Balance Sheet', 'Cash Flow']

    # keep a list of key figures that should be seperated by a border in the statement
    key_figures = {
        statements[0]: ['Gross Profit', 'Operating Income', 'Pretax Income', 'Net Income Common Stockholders'],

        statements[1]: ['Total Assets', 'Total Liabilities Net Minority Interest', 'Total Equity Gross Minority Interest', ],

        statements[2]: ['Investing Cash Flow', 'Financing Cash Flow', 'End Cash Position'],
    }
    # a list of figures that should have a double line under them
    bottom_lines = {
        statements[0]: ['Net Income Common Stockholders'],

        statements[1]: ['Common Stock Equity'],

        statements[2]: ['End Cash Position'],
    }
    # a list of figures that should have a seperator after them
    seperators = {
        statements[0]: ['Gross Profit', 'Operating Income', 'Net Income Common Stockholders', 'Diluted Average Shares', 'Total Expenses',
                        'Normalized Income', 'Net Interest Income', 'EBITDA', 'Reconciled Depreciation', 'Normalized EBITDA'],

        statements[1]: ['Total Assets', 'Cash And Cash Equivalents', 'Current Deferred Assets', 'Accumulated Depreciation',
                        'Other Intangible Assets', 'Non Current Deferred Assets', 'Total Liabilities Net Minority Interest', 'Accounts Payable',
                        'Current Debt', 'Long Term Debt', 'Non Current Deferred Taxes Liabilities',
                        'Total Non Current Liabilities Net Minority Interest', 'Other Non Current Liabilities',
                        'Total Equity Gross Minority Interest', 'Common Stock Equity',
                        'Gains Losses Not Affecting Retained Earnings', 'Other Equity Adjustments', 'Tangible Book Value', 'Net Debt'],

        statements[2]: ['Change in Other Working Capital', 'Purchase of Business', 'End Cash Position'],
    }

    ratio_headers = ['Liquidity', 'Efficiency', 'Capacity / Leverage', 'Profitability', 'Growth', 'Net Working Capital', 'Percent of Sales']
    # a list of figures that should have a seperator after them
    ratio_seperators = ['Cash & Securities / Assets', 'Cash Conversion Cycle', 'Gross Debt / EBITDA', 'Assets / Equity', 'Assets', 'Net Working Capital']

    header = book.add_format({
        'font_color': 'white',
        'bg_color': '#034638',
    })

    line_item = book.add_format({
        'align': 'left',
    })

    regular_data = book.add_format({
        'num_format': 43,
    })

    key_figure = book.add_format({
        'num_format': 43,
        'top': 1,
    })

    bottom_line = book.add_format({
        'num_format': 43,
        'top': 1,
        'bottom': 6,
    })

    # checks if the line item is a key figure and formats accordingly. returns 1 if a seperator was used, 0 if not
    def get_format(item, statement):
        if item in bottom_lines[statement]:
            return bottom_line
        if item in key_figures[statement]:
            return key_figure
        return regular_data

    dfs = {}
    # make financial statements
    for statement, name in zip(statements, names):
        dfs[statement] = get_statement(ticker, statement, driver)
        if statement == 'financials':
            tax_rate = dfs[statement].loc['Tax Rate for Calcs', dfs[statement].columns[0]]
            if tax_rate == 0:
                try:
                    # use most recent tax rate only
                    tax_rate = (dfs[statement].loc['Tax Provision', dfs[statement].columns[0]] / dfs[statement].loc['Pretax Income', dfs[statement].columns[0]])
                except KeyError:
                    pass

            tax_rate = max(tax_rate, 0)

        sheet = book.add_worksheet(name)
        sheet.freeze_panes(2, 1)

        # resize the columns
        for i, width in enumerate(get_col_widths(dfs[statement])):
            # add 3 to account for decimals which may not be present in original column, but must be present because of formatting
            sheet.set_column(i, i, width + 3)

        # write the headers row
        sheet.write(0, 0, '*In thousands, except per-share items')
        sheet.write(1, 0, name.upper(), header)  # name of statement
        for i, period in enumerate(dfs[statement]):
            sheet.write(1, i + 1, period, header)

        # write the line items column
        s = 0
        for i, item in enumerate(dfs[statement].index):
            # write data items with formatting before formatting the whole row, if necessary
            sheet.write_row(i + s + 2, 1, dfs[statement].loc[item], get_format(item, statement))
            # after formatting, write the line item so it overwrites the rest-of-row formatting
            sheet.write(i + s + 2, 0, item, line_item)
            # format the row based off the line item label. if a seperator should be placed after the
            # row, keep track of how many seperators were used so all other rows can be pushed down
            s += 1 if item in seperators[statement] else 0

    # ratios = pd.DataFrame(columns=['Data'])

    # ratios.loc['Current Ratio'] = dfs['balance-sheet']['Current Assets'] / dfs['balance-sheet']['Current Liabilities']
    # ratios.loc['Cash & Securities / Assets'] = dfs['balance-sheet']['Cash, Cash Equivalents & Short Term Investments'] / dfs['balance-sheet']['Total Assets']

    # ratios.loc['Days of Receivables'] = 365 * dfs['balance-sheet']['Accounts receivable'] / dfs['financials']['Total Revenue']

    return dfs, tax_rate


def to_datatype(s: str, parse_dates: bool = False):
    s = s.replace(',', '')
    if s.endswith('%'):
        try:
            return float(s.replace('%', '')) / 100
        except ValueError:
            pass
    elif s.endswith('T'):
        try:
            return float(s.replace('T', '')) * 1e12
        except ValueError:
            pass
    elif s.endswith('B'):
        try:
            return float(s.replace('B', '')) * 1e9
        except ValueError:
            pass
    elif s.endswith('M'):
        try:
            return float(s.replace('M', '')) * 1e6
        except ValueError:
            pass
    elif s.endswith('k'):
        try:
            return float(s.replace('k', '')) * 1e3
        except ValueError:
            pass
    try:
        return float(s)
    except ValueError:
        pass

    if parse_dates:
        try:
            return dt.datetime.strptime(s, '%M/%d/%Y')
        except ValueError:
            pass
        try:
            return dt.datetime.strptime(s, '%M-%d-%Y')
        except ValueError:
            pass

    return s


def get_peer(ticker, tax_rate, dfs):
    statements = ['financials', 'balance-sheet', 'cash-flow']
    labels = ['Peer', 'P/E Ratio', 'EV/Sales', 'EV/EBITDA', 'Market Cap', 'Total Debt', 'Cash and Equivalents',
              'Enterprise Value', 'Debt/Equity', 'Bond Rating (S&P)', 'Bond Spread (10Y)', 'Bond Spread (30Y)',
              'LTM Sales', 'LTM EBITDA', 'LTM Earnings', 'Share Price', 'Shares Outstanding', 'Equity Beta',
              'Unlevered Beta', 'Profit Margin', 'Operating Margin', 'Return on Assets', 'Return on Equity',
              'Revenue Growth (1Y)', 'Earnings Growth (1Y)', 'Key Notes']

    df = pd.DataFrame(columns=['Data'])
    for l in labels:
        df.loc[l] = ''

    url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics?p={ticker}'
    res = requests.get(url, headers=headers)
    sleep(1)
    soup = BeautifulSoup(res.text, features='lxml')

    try:
        df.loc['Peer'] = soup.find('h1').text
        df.loc['Share Price'] = to_datatype(soup.find('fin-streamer', {'data-test': 'qsp-price'}).text)

        df.loc['Equity Beta'] = to_datatype(soup.find(text='Beta (5Y Monthly)').parent.parent.next_sibling.text)
        if df.loc['Equity Beta', 'Data'] == 'N/A':  # handle invalid numbers
            df.loc['Equity Beta'] = 0.
        df.loc['Shares Outstanding'] = to_datatype(soup.find(text='Shares Outstanding').parent.parent.next_sibling.text) / 1000

        df.loc['Profit Margin'] = to_datatype(soup.find(text='Profit Margin').parent.parent.next_sibling.text)
        df.loc['Operating Margin'] = to_datatype(soup.find(text='Operating Margin').parent.parent.next_sibling.text)
        df.loc['Return on Assets'] = to_datatype(soup.find(text='Return on Assets').parent.parent.next_sibling.text)
        df.loc['Return on Equity'] = to_datatype(soup.find(text='Return on Equity').parent.parent.next_sibling.text)
        df.loc['Revenue Growth (1Y)'] = to_datatype(soup.find(text='Quarterly Revenue Growth').parent.parent.next_sibling.text)
        df.loc['Earnings Growth (1Y)'] = to_datatype(soup.find(text='Quarterly Earnings Growth').parent.parent.next_sibling.text)

        df.loc['LTM Sales'] = to_datatype(soup.find(text='Revenue').parent.parent.next_sibling.text) / 1000

        try:
            df.loc['LTM EBITDA'] = to_datatype(soup.find(text='EBITDA').parent.parent.next_sibling.text) / 1000
            if df.loc['LTM EBITDA', 'Data'] == 'N/A':
                raise TypeError
        # if ebitda is not a line item
        except (TypeError, AttributeError):
            # start with net income (this value should always be present on yahoo, everything else may not be)
            ebit_actual = dfs[statements[0]].loc['Net Income Common Stockholders', 'TTM']
            # add back taxes to net income
            try:
                ebit_actual += dfs[statements[0]].loc['Tax Provision', 'TTM']
            except (KeyError, TypeError):
                pass
            # try to add back interest expense and interest income
            try:
                ebit_actual += dfs[statements[0]].loc['Interest Expense', 'TTM']
            except (KeyError, TypeError):
                pass
            try:
                ebit_actual -= dfs[statements[0]].loc['Interest Income', 'TTM']
            except (KeyError, TypeError):
                pass
            # try to add back depreciation
            try:
                ebit_actual += dfs[statements[0]].loc['Reconciled Depreciation', 'TTM']
            except (KeyError, TypeError):
                pass
            df.loc['LTM EBITDA'] = ebit_actual

        df.loc['LTM Earnings'] = to_datatype(soup.find(text='Net Income Avi to Common').parent.parent.next_sibling.text) / 1000

        df.loc['Cash and Equivalents'] = to_datatype(soup.find(text='Total Cash').parent.parent.next_sibling.text) / 1000
        df.loc['Total Debt'] = to_datatype(soup.find(text='Total Debt').parent.parent.next_sibling.text) / 1000

        df.loc['Market Cap'] = df.loc['Share Price'] * df.loc['Shares Outstanding']
        df.loc['Enterprise Value'] = df.loc['Market Cap'] + df.loc['Total Debt'] - df.loc['Cash and Equivalents']

        df.loc['Debt/Equity'] = df.loc['Total Debt'] / df.loc['Market Cap']
        df.loc['Unlevered Beta'] = df.loc['Equity Beta'] / (1 + (1 - tax_rate) * df.loc['Debt/Equity'])

        df.loc['P/E Ratio'] = max(df.loc['Market Cap', 'Data'] / df.loc['LTM Earnings', 'Data'], 0.)
        df.loc['EV/Sales'] = max(df.loc['Enterprise Value', 'Data'] / df.loc['LTM Sales', 'Data'], 0.)
        df.loc['EV/EBITDA'] = max(df.loc['Enterprise Value', 'Data'] / df.loc['LTM EBITDA', 'Data'], 0.)
    except (TypeError, AttributeError):
        # print(traceback.format_exc())
        pass

    # get bond ratings and spread
    url = f'https://www.macroaxis.com/invest/bond/{ticker}'
    res = requests.get(url, headers=headers)
    sleep(1)
    soup = BeautifulSoup(res.text, features='lxml')

    try:
        data = []
        headings = []
        for header in soup.find(text='Issue Date').parent.parent.children:
            headings.append(header.text.strip())
        for row in soup.find(text='Issue Date').parent.parent.next_siblings:
            entry = {}
            for header, datum in zip(headings, row.children):
                entry[header] = to_datatype(datum.text.strip(), parse_dates=True)
            data.append(entry)

        debt_frame = pd.DataFrame(data).set_index('')
        debt_frame['Spread'] = debt_frame['Coupon'] - debt_frame['Ref Coupon']
        debt_frame['Length'] = debt_frame['Maturity'] - debt_frame['Issue Date']

        # get all bonds with an initial time to maturity of approximately 10 years (between 8 and 12) and calculate the average spread
        df.loc['Bond Spread (10Y)'] = debt_frame.loc[(debt_frame['Length'] >= dt.timedelta(days=2920)) & (debt_frame['Length'] <= dt.timedelta(days=4380))]['Spread'].mean() / 100
        # get all bonds with an initial time to maturity of approximately 30 years (between 28 and 32) and calculate the average spread
        df.loc['Bond Spread (30Y)'] = debt_frame.loc[(debt_frame['Length'] >= dt.timedelta(days=10220)) & (debt_frame['Length'] <= dt.timedelta(days=11680))]['Spread'].mean() / 100
    except AttributeError:
        pass

    try:
        df.loc['Bond Rating (S&P)'] = soup.find(text='Average S&P Rating').parent.next_sibling.text
    except AttributeError:
        pass

    return df


def get_summary(ticker):
    labels = ['Peer', 'Sector', 'Industry', 'Employees', 'Summary', 'Link']
    df = pd.DataFrame(columns=['Data'])
    for l in labels:
        df.loc[l] = ''

    url = f'https://finance.yahoo.com/quote/{ticker}/profile?p={ticker}'
    res = requests.get(url, headers=headers)
    sleep(1)
    soup = BeautifulSoup(res.text, features='lxml')

    try:
        df.loc['Peer'] = soup.find('h1').text
        df.loc['Summary'] = soup.find(text='Description').parent.parent.next_sibling.text
        df.loc['Employees'] = soup.find(text='Full Time Employees').parent.next_sibling.next_sibling.next_sibling.next_sibling.text
        df.loc['Sector'] = soup.find(text='Sector(s)').parent.next_sibling.next_sibling.next_sibling.next_sibling.text
        df.loc['Industry'] = soup.find(text='Industry').parent.next_sibling.next_sibling.next_sibling.next_sibling.text
    except (TypeError, AttributeError) as e:
        # print(traceback.format_exc())
        pass
    # even if getting the data fails, always include the link to the ticker
    df.loc['Link'] = f'https://finance.yahoo.com/quote/{ticker}'
    return df


# peers must be sent as a list-like
def make_peers(ticker, peers, tax_rate, dfs, book: xls.Workbook, peer_gen_depth=0, driver: webdriver.Chrome = None):
    # a list of columns that should be seperated from the next column
    column_splits = ['EV/EBITDA', 'Enterprise Value', 'Bond Spread (30Y)', 'LTM Earnings', 'Unlevered Beta']

    header = book.add_format({
        'bold': True,
        'top': 1,
        'bottom': 1,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
    })

    bold = book.add_format({
        'bold': True,
    })

    main_ticker = book.add_format({
        'num_format': 43,
        'bg_color': '#E7E6E6',
        'top': 1,
        'bottom': 1,
    })

    entry = book.add_format({
        'num_format': 43,
    })

    last_entry = book.add_format({
        'num_format': 43,
        'bottom': 1,
    })

    summary = book.add_format({

    })

    last_summary = book.add_format({
        'bottom': 1,
    })

    main_summary = book.add_format({
        'bg_color': '#E7E6E6',
        'top': 1,
        'bottom': 1,
    })

    # if "peer_gen_depth" > 0, generate a list of peers by reading the "People Also Watch" tab on Yahoo
    # for every level of depth, add peers of the given peer to the list.
    # e.g. peer_gen_depth = 1 means only the main  ticker's peers are added to the list,
    # peer_gen_depth = 2 means peers of the main ticker's peers are also added to the list, and so on
    def generate_peers(t: str, peers: list, driver: webdriver.Chrome, depth: int):
        # base case when the given depth was reached
        if depth == 0:
            return

        try:
            driver.get(f'https://finance.yahoo.com/quote/{t}?p={t}')
            sleep(5)
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"//section[@id = 'recommendations-by-symbol']//a"))
            )

            html = driver.execute_script('return document.body.innerHTML;')
            soup = BeautifulSoup(html, 'lxml')
            for p in soup.find('h2', text='People Also Watch').parent.find_all('a'):
                if p.text not in peers and p.text != t:
                    peers.append(p.text)
                    generate_peers(p.text, peers, driver, depth - 1)
        finally:
            return

    # generate the peer list and add it to the current peers
    # note that peers passed by the user to this "make_peers" function will not
    # have their peers checked, regardless of the depth
    generate_peers(ticker, peers, driver, peer_gen_depth)
    try:
        peers.remove(ticker)
    except ValueError:
        pass

    sheet = book.add_worksheet('Peers')
    summaries = book.add_worksheet('Peer Summaries')
    sheet.freeze_panes(4, 1)
    summaries.freeze_panes(1, 1)

    sheet.write(0, 0, '*Total share count and all monetary values are in thousands of USD where applicable, except per-share items')
    sheet.write(1, 0, 'Assumed tax rate:', bold)
    sheet.write(1, 1, tax_rate, bold)

    sheet.set_row(2, 52.8)

    def write_items(df, row, format, sheet):
        # keep track of seperators between columns
        s = 0
        for i, item in enumerate(df.index):
            try:
                sheet.write(row, i + s, item if format == header else df.loc[item, 'Data'], format)
            except Exception:
                # write an empty cell if the current data could not be written
                sheet.write(row, i + s, '', format)
            sheet.set_column(i + s, i + s, 16)
            if item in column_splits:
                s += 1
                sheet.write(row, i + s, None, format)
                sheet.set_column(i + s, i + s, 4)

        sheet.set_column(0, 0, 32)

    df = get_peer(ticker, tax_rate, dfs)
    # write the headers and the main ticker's info
    write_items(df, 2, header, sheet)
    write_items(df, 3, main_ticker, sheet)

    # write the summary headers and the main ticker's summary
    df = get_summary(ticker)
    write_items(df, 0, header, summaries)
    write_items(df, 1, main_summary, summaries)

    for p, peer in enumerate(peers):
        # write info about a peer
        df = get_peer(peer, tax_rate, dfs)
        write_items(df, p + 4, entry if p + 1 < len(peers) else last_entry, sheet)
        # write peer summary
        df = get_summary(peer)
        write_items(df, p + 2, summary if p + 1 < len(peers) else last_summary, summaries)


def make_dcf(dfs: dict, ticker: str, peers: list, tax_rate: float, rfr: float, mrp: float, terminal_growth: float, forecast_years: int, book: xls.Workbook):
    statements = ['financials', 'balance-sheet', 'cash-flow']  # "financials" means income statement

    # get values from financial statements required to make DCF
    revenue_actual = dfs[statements[0]].loc['Total Revenue', dfs[statements[0]].columns[1]]

    try:
        ebit_actual = dfs[statements[0]].loc['EBIT', dfs[statements[0]].columns[1]]
    # if EBIT is not a line item, construct it manually:
    except (KeyError, TypeError):
        # start with net income (this value should always be present on yahoo, everything else may not be)
        ebit_actual = dfs[statements[0]].loc['Net Income Common Stockholders', dfs[statements[0]].columns[1]]
        # try to add back taxes to net income
        try:
            ebit_actual += dfs[statements[0]].loc['Tax Provision', dfs[statements[0]].columns[1]]
        except (KeyError, TypeError):
            pass
        # try to add back interest expense and interest income (may not be a line item)
        try:
            ebit_actual += dfs[statements[0]].loc['Interest Expense', dfs[statements[0]].columns[1]]
        except (KeyError, TypeError):
            pass
        try:
            ebit_actual -= dfs[statements[0]].loc['Interest Income', dfs[statements[0]].columns[1]]
        except (KeyError, TypeError):
            pass

    revenue_ttm = dfs[statements[0]].loc['Total Revenue', dfs[statements[0]].columns[0]]

    try:
        ebit_ttm = dfs[statements[0]].loc['EBIT', dfs[statements[0]].columns[0]]
    # if EBIT is not a line item, construct it manually:
    except (KeyError, TypeError):
        # start with net income (this value should always be present on yahoo, everything else may not be)
        ebit_ttm = dfs[statements[0]].loc['Net Income Common Stockholders', dfs[statements[0]].columns[0]]
        # try to add back taxes to net income
        try:
            ebit_ttm += dfs[statements[0]].loc['Tax Provision', dfs[statements[0]].columns[0]]
        except (KeyError, TypeError):
            pass
        # try to add back interest expense and interest income (may not be a line item)
        try:
            ebit_ttm += dfs[statements[0]].loc['Interest Expense', dfs[statements[0]].columns[0]]
        except (KeyError, TypeError):
            pass
        try:
            ebit_ttm -= dfs[statements[0]].loc['Interest Income', dfs[statements[0]].columns[0]]
        except (KeyError, TypeError):
            pass

    ebit_margin = ebit_ttm / revenue_ttm

    # get percents nopat add-backs by taking the average of all recorded full fiscal years
    try:
        depr_amort_pct = (dfs[statements[0]].loc['Reconciled Depreciation'].drop(dfs[statements[0]].columns[0])
                          / dfs[statements[0]].loc['Total Revenue'].drop(dfs[statements[0]].columns[0])).mean()
    except KeyError:
        depr_amort_pct = 0
    try:
        capex_pct = -(dfs[statements[2]].loc['Capital Expenditure'].drop(dfs[statements[2]].columns[0])
                     / dfs[statements[0]].loc['Total Revenue'].drop(dfs[statements[0]].columns[0])).mean()
    except KeyError:
        capex_pct = 0
    try:
        nwc_pct = -(dfs[statements[2]].loc['Change in working capital'].drop(dfs[statements[2]].columns[0])
                   / dfs[statements[0]].loc['Total Revenue'].drop(dfs[statements[0]].columns[0])).mean()
    except KeyError:
        nwc_pct = 0

    try:
        # get the first two years' growth rate from the Analyst section of yahoo finance
        url = f'https://finance.yahoo.com/quote/{ticker}/analysis?p={ticker}'
        res = requests.get(url, headers=headers)
        soup = BeautifulSoup(res.text, features='lxml')
        growth_rate_1 = to_datatype([c for c in soup.find('td', text='Sales Growth (year/est)').next_siblings][2].text)
        growth_rate_2 = to_datatype([c for c in soup.find('td', text='Sales Growth (year/est)').next_siblings][3].text)
        # reverse the columns so pct_change gets the growth rate of the past revenues in the income statement
        rates = dfs[statements[0]][dfs[statements[0]].columns[::-1]].loc['Total Revenue'].pct_change().dropna()[:-1]
        # the growth rate to be used for the remaining years in the DCF forecast. not to be confused with the terminal growth rate.
        # uses the average of the past growth rates AND the two projected growth rates
        growth_rate_f = (rates.sum() + growth_rate_1 + growth_rate_2) / (len(rates) + 2)
    except Exception:
        rates = dfs[statements[0]][dfs[statements[0]].columns[:1:-1]].loc['Total Revenue'].pct_change().dropna().mean()
        growth_rate_1 = rates
        growth_rate_2 = rates
        growth_rate_f = rates

    header = book.add_format({
        'font_color': 'white',
        'bg_color': '#034638',
        'bold': True,
    })

    subheader = book.add_format({
        'bold': True,
    })

    line_item = book.add_format({
        'align': 'left',
    })

    regular_data = book.add_format({
        'num_format': 43,
    })

    percent = book.add_format({
        'num_format': 10,
    })

    supp_line_item = book.add_format({
        'align': 'left',
        'italic': True,
    })

    supp_data = book.add_format({
        'num_format': 43,
        'italic': True,
    })

    supp_percent = book.add_format({
        'num_format': 10,
        'italic': True,
    })

    key_figure = book.add_format({
        'num_format': 43,
        'top': 1,
    })

    bottom_line = book.add_format({
        'num_format': 43,
        'bottom': 6,
    })

    notes = book.add_format({
        'text_wrap': True,
        'valign': 'top',
    })

    sheet = book.add_worksheet('DCF')
    # make sure the header encompasses either the length of peer valuation table and the forecast table
    header_size = max(9, forecast_years + 2)

    # write the WACC header
    sheet.write(0, 0, '*in thousands, except per-share items. Next steps after generation: 1. Check that WACC inputs are correct (especially debt spread). 2. Update growth rates as desired. 3. Update all other DCF inputs as desired.')
    sheet.merge_range(1, 0, 1, header_size - 1, 'Cost of Capital', header)

    # write the WACC subheaders
    sheet.write(2, 0, 'Cost of Equity', subheader)
    sheet.write(2, 3, 'Cost of Debt', subheader)
    sheet.write(2, 6, 'WACC', subheader)

    # write the cost of equity line items and values
    sheet.write(3, 0, 'Risk Free Rate', line_item)
    sheet.write(3, 1, rfr, percent)
    sheet.write(4, 0, 'Market Risk Premium', line_item)
    sheet.write(4, 1, mrp, percent)
    # use competitor beta
    sheet.write(5, 0, 'Beta', line_item)
    sheet.write(5, 1, f"=AVERAGE(Peers!$W$5:$W${len(peers) + 4}) * (1 + (1-$B$14) * Peers!$K$4)", regular_data)
    sheet.write(6, 0, 'Cost of Equity', line_item)
    sheet.write(6, 1, "=B4+B5*B6", percent)

    # write the cost of debt line items and values
    sheet.write(3, 3, 'Risk Free Rate', line_item)
    sheet.write(3, 4, rfr, percent)
    sheet.write(4, 3, 'Spread', line_item)
    # use 30Y spread but if they are unavailable use 10 year spread
    sheet.write(4, 4, "=IF(ISBLANK(Peers!$N$4), Peers!$M$4, Peers!$N$4)", percent)
    # indicate which type of bond spread was used, or if no bond spreads could be found
    sheet.write(4, 5, '=IF(ISBLANK(Peers!$N$4), IF(ISBLANK(Peers!$M$4), "NEEDS TO BE UPDATED", "10Y spread"), "30Y spread")', line_item)
    sheet.write(5, 3, 'Cost of Debt', line_item)
    sheet.write(5, 4, "=E4+E5", percent)

    # write the WACC calculation lines
    sheet.write(3, 6, 'Weight of Equity', line_item)
    sheet.write(3, 7, "=Peers!$F$4/(Peers!$F$4 + Peers!$G$4)", percent)
    sheet.write(4, 6, 'Weight of Debt', line_item)
    sheet.write(4, 7, "=Peers!$G$4/(Peers!$F$4 + Peers!$G$4)", percent)
    sheet.write(5, 6, 'WACC', line_item)
    sheet.write(5, 7, "=B7 * H4 + E6 * H5", percent)

    # write the DCF inputs header
    sheet.merge_range(10, 0, 10, header_size - 1, 'DCF Model Inputs', header)

    # write the input lines and values
    sheet.write(11, 0, 'Cost of Capital', line_item)
    sheet.write(11, 1, "=H6", percent)
    sheet.write(12, 0, 'Terminal Growth Rate', line_item)
    sheet.write(12, 1, terminal_growth, percent)
    sheet.write(13, 0, 'Tax Rate', line_item)
    sheet.write(13, 1, tax_rate, percent)

    sheet.write(14, 0, 'Dep & Amort / Revenue', line_item)
    sheet.write(14, 1, depr_amort_pct, percent)
    sheet.write(15, 0, 'CAPEX / Revenue', line_item)
    sheet.write(15, 1, capex_pct, percent)
    sheet.write(16, 0, 'Change in Net Working Capital / Revenue', line_item)
    sheet.write(16, 1, nwc_pct, percent)

    sheet.write(17, 0, 'Current Debt Value', line_item)
    sheet.write(17, 1, "=Peers!$G$4", regular_data)
    sheet.write(18, 0, 'Current Cash Value', line_item)
    sheet.write(18, 1, "=Peers!$H$4", regular_data)
    sheet.write(19, 0, 'Shares Outstanding', line_item)
    sheet.write(19, 1, "=Peers!$U$4", regular_data)

    # write the DCF model header
    sheet.merge_range(23, 0, 23, header_size - 1, 'DCF Model', header)
    # write the period headers
    sheet.write(26, 1, f'{dt.datetime.today().year - 1} Actual', subheader)
    for i in range(forecast_years):
        sheet.write(26, i + 2, dt.datetime.today().year + i, subheader)

    # write the DCF model line items
    sheet.write(24, 0, 'Revenue Growth', supp_line_item)
    sheet.write(25, 0, 'EBIT Margin', supp_line_item)
    sheet.write(27, 0, 'Revenue', line_item)
    sheet.write(28, 0, 'EBIT', line_item)
    sheet.write(29, 0, 'Tax', line_item)
    sheet.write(30, 0, 'NOPAT', line_item)
    sheet.write(31, 0, 'Depreciation & Amortization', line_item)
    sheet.write(32, 0, 'Capital Expenditures', line_item)
    sheet.write(33, 0, 'Increase in Working Capital', line_item)
    sheet.write(34, 0, 'Free Cash Flow', line_item)
    sheet.write(35, 0, 'Discounted FCF', line_item)
    sheet.write(36, 0, 'Terminal Value', line_item)

    # write the historic DCF values
    sheet.write(27, 1, revenue_actual, key_figure)
    sheet.write(28, 1, ebit_actual, regular_data)
    sheet.write(34, 1, "", key_figure)

    # write the forecast figures
    for i in range(forecast_years):
        growth = growth_rate_1 if i == 0 else (growth_rate_2 if i == 1 else growth_rate_f)
        c = colnum_string(i + 2)
        p = colnum_string(i + 1)

        sheet.write(24, i + 2, growth, supp_percent)
        sheet.write(25, i + 2, ebit_margin, supp_percent)
        sheet.write(27, i + 2, f"={p}28 * ({c}25 + 1)", key_figure)
        sheet.write(28, i + 2, f"={c}28 * {c}26", regular_data)
        sheet.write(29, i + 2, f"={c}29 * $B$14", regular_data)
        sheet.write(30, i + 2, f"={c}29 - {c}30", regular_data)
        sheet.write(31, i + 2, f"={c}28 * $B$15", regular_data)
        sheet.write(32, i + 2, f"={c}28 * $B$16", regular_data)
        sheet.write(33, i + 2, f"={c}28 * $B$17", regular_data)
        sheet.write(34, i + 2, f"={c}31 + {c}32 - {c}33 - {c}34", key_figure)
        sheet.write(35, i + 2, f"={c}35 / ((1+$B$12) ^ {i + 1})", regular_data)

    # write the terminal value
    sheet.write(
        36,
        forecast_years + 1,
        f"={colnum_string(forecast_years + 1)}36 / (1+$B$12) / ($B$12 - $B$13)",
        regular_data
    )

    # write the DCF outputs header
    sheet.merge_range(40, 0, 40, header_size - 1, 'DCF Model Outputs', header)

    # write the DCF outputs
    sheet.write(41, 0, 'Enterprise Value', line_item)
    sheet.write(41, 1, f"=SUM($C$36:${colnum_string(forecast_years + 1)}$37)", regular_data)
    sheet.write(42, 0, 'Equity Value', line_item)
    sheet.write(42, 1, "=$B$42 - $B$18 + $B$19", regular_data)
    sheet.write(43, 0, 'Equity Value per Share', line_item)
    sheet.write(43, 1, "=$B$43 / $B$20", bottom_line)

    # write the peer valuations header and subheaders
    sheet.merge_range(47, 0, 47, header_size - 1, 'Peer Implied Valuations', header)
    sheet.write(48, 1, f'{ticker}', subheader)
    sheet.write(48, 2, 'Peer Average', subheader)
    sheet.write(48, 3, 'Peer Minimum', subheader)
    sheet.write(48, 4, 'Peer Maximum', subheader)
    sheet.write(48, 6, 'Implied Valuation', subheader)
    sheet.write(48, 7, 'Implied Range Minimum', subheader)
    sheet.write(48, 8, 'Implied Range Maximum', subheader)
    # write the peer implied valuations
    sheet.write(49, 0, 'Price/Earnings', line_item)
    sheet.write(49, 1, f"=Peers!$B$4")
    sheet.write(49, 2, f'=AVERAGEIF(Peers!$B$5:$B${len(peers) + 4}, ">0", Peers!$B$5:$B${len(peers) + 4})', regular_data)
    sheet.write(49, 3, f"=_xlfn.MINIFS(Peers!$B$5:$B${len(peers) + 4}, Peers!$B$5:$B${len(peers) + 4}, \">0\")", regular_data)
    sheet.write(49, 4, f"=MAX(Peers!$B$5:$B${len(peers) + 4})", regular_data)
    sheet.write(49, 6, f'=IF(C51 > 0, C50 * Peers!$R$4 / $B$20, "")', regular_data)
    sheet.write(49, 7, f'=IF(C51 > 0, D50 * Peers!$R$4 / $B$20, "")', regular_data)
    sheet.write(49, 8, f'=IF(C51 > 0, E50 * Peers!$R$4 / $B$20, "")', regular_data)

    sheet.write(50, 0, 'EV/Sales', line_item)
    sheet.write(50, 1, f"=Peers!$C$4")
    sheet.write(50, 2, f'=AVERAGEIF(Peers!$C$5:$C${len(peers) + 4}, ">0", Peers!$C$5:$C${len(peers) + 4})', regular_data)
    sheet.write(50, 3, f"=_xlfn.MINIFS(Peers!$C$5:$C${len(peers) + 4}, Peers!$C$5:$C${len(peers) + 4}, \">0\")", regular_data)
    sheet.write(50, 4, f"=MAX(Peers!$C$5:$C${len(peers) + 4})", regular_data)
    sheet.write(50, 6, f'=IF(C51 > 0, ($B$19 - $B$18 + C51 * Peers!$P$4) / $B$20, "")', regular_data)
    sheet.write(50, 7, f'=IF(C51 > 0, ($B$19 - $B$18 + D51 * Peers!$P$4) / $B$20, "")', regular_data)
    sheet.write(50, 8, f'=IF(C51 > 0, ($B$19 - $B$18 + E51 * Peers!$P$4) / $B$20, "")', regular_data)

    sheet.write(51, 0, "EV/EBITDA", line_item)
    sheet.write(51, 1, f"=Peers!$D$4")
    sheet.write(51, 2, f'=AVERAGEIF(Peers!$D$5:$D${len(peers) + 4}, ">0", Peers!$D$5:$D${len(peers) + 4})', regular_data)
    sheet.write(51, 3, f"=_xlfn.MINIFS(Peers!$D$5:$D${len(peers) + 4}, Peers!$D$5:$D${len(peers) + 4}, \">0\")", regular_data)
    sheet.write(51, 4, f"=MAX(Peers!$D$5:$D${len(peers) + 4})", regular_data)
    sheet.write(51, 6, f'=IF(C52 > 0, ($B$19 - $B$18 + C52 * Peers!$Q$4) / $B$20, "")', regular_data)
    sheet.write(51, 7, f'=IF(C52 > 0, ($B$19 - $B$18 + D52 * Peers!$Q$4) / $B$20, "")', regular_data)
    sheet.write(51, 8, f'=IF(C52 > 0, ($B$19 - $B$18 + E52 * Peers!$Q$4) / $B$20, "")', regular_data)

    # add a section for analysis notes
    sheet.merge_range(55, 0, 55, header_size - 1, 'Analyst Notes', header)
    sheet.merge_range(56, 0, 106, header_size - 1, '', notes)
    sheet.write(56, 0, message_to_analyst, notes)


def make_template(ticker, peers, rfr, mrp, terminal_growth, forecast_years, peer_gen_depth=0, outfile=None):
    if outfile == None:
        outfile = ticker.replace('.', '-') + '.xlsx'

    driver = webdriver.Chrome(options=options)
    try:
        writer = xls.Workbook(outfile)
        dfs, tax_rate = make_financials(ticker, writer, driver)
        make_peers(ticker, peers, tax_rate, dfs, writer, peer_gen_depth, driver)
        make_dcf(dfs, ticker, peers, tax_rate, rfr, mrp, terminal_growth, forecast_years, writer)
    finally:
        driver.close()
        writer.close()


if __name__ == '__main__':
    config = json.loads(open('config.json').read())

    ticker = config['ticker']
    peers = config['peers']
    # the current yield on 30Y US Treasury bonds
    rfr = config['rfr']
    # market risk premium as determined by statista:
    # https://www.statista.com/statistics/664840/average-market-risk-premium-usa/
    mrp = config['mrp']
    # default terminal growth is the CAGR of the American GDP projection by PWC from 2016 to 2050
    # =(1+(34102-18562)/18562)^(1/(2050-2016))-1
    terminal_growth = config['terminal_growth']
    forecast_years = config['forecast_years']

    make_template(ticker, peers, rfr, mrp, terminal_growth, forecast_years, 2)
