import argparse
import requests
from bs4 import BeautifulSoup
from time import sleep

from src.makeTemplate import make_template, headers


if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    parser.add_argument('ticker', type=str, help='The stock ticker to create a DCF for')
    parser.add_argument('--generate_peers', '-gp', help='Set this flag to automatically create a list of peers. If this is not set, you must pass a list of peers using the -p flag', action='store_true')
    parser.add_argument('--peers', '-p', help='Set a list of peers to compare the given ticker to. Must be set if --generate_peers is not set.', type=str, nargs='+')
    parser.add_argument('--risk_free_rate', '-rfr', help='Set the risk-free rate for cost of capital calculations. Defaults to the current 10Y American Treasury yield.', type=float)
    parser.add_argument('--market_risk_premium', '-mrp', help='Set the market risk premium for cost of capital calculations as a decimal. Defaults to the most recent American average MRP as given by Statista ~0.055.', type=float, default=0.055)
    parser.add_argument('--terminal_growth', '-tg', help='Set the terminal growth rate in the DCF model as a decimal. Defaults to PricewaterhouseCoopers 50Y projected American GDP annual growth rate ~0.0181.', type=float, default=0.018050372)
    parser.add_argument('--forecast_years', '-fy', help='Set how many years to make projections for in the DCF mode as a decimal. Defaults to 5 years.', type=int, default=5)
    parser.add_argument('--min_tax_rate', '-tr', help='Set the minimum tax rate for a company as a decimal. Tax rate is typically calculated by the program, but the calculated value will not be used if it is below the minimum. Defaults to 0.2.', type=float, default=0.2)
    parser.add_argument('--output', '-o', help='Set the filename to save the DCF in. Must be an .xlsx file.', type=str)

    args = parser.parse_args().__dict__

    if args['generate_peers'] is None and args['peers'] is None:
        raise ValueError('Please either set the --generate_peers flag or pass a list of peers using the --peers flag (pass the -h flag for help).')
    if args['output'] is not None and not args['output'].endswith('.xlsx'):
        raise ValueError('Output file must use a .xlsx extension.')

    if args['risk_free_rate'] is None:
        url = f'https://finance.yahoo.com/quote/%5ETNX'
        res = requests.get(url, headers=headers)
        soup = BeautifulSoup(res.text, features='lxml')
        sleep(1)
        args['risk_free_rate'] = float(soup.find('fin-streamer', {'data-test': 'qsp-price'}).text) / 100

    make_template(
        args['ticker'],
        [] if not args['peers'] else args['peers'],
        args['risk_free_rate'],
        args['market_risk_premium'],
        args['terminal_growth'],
        args['forecast_years'],
        args['min_tax_rate'],
        0 if not args['generate_peers'] else 2,
        None if not args['output'] else args['output']
    )
