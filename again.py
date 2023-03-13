"""
author: Ruchir Garg
date:   7 March 2023

run:
python3 again.py

"""

import pandas as pd
import math

# years can be added arbitrarily here and report will adjust automatically
years = [2022, 2023, 2024]

NAME = 'Company'
TICKER = 'Ticker'
BROKER = 'Broker'
YEAR = 'year'
EPS_CHG = 'EPS chg'
EPS = 'new EPS'
VS_CONS = 'vs Cons'
SALES_CHG = 'sales chg'
EBIT_CHG = 'EBIT chg'
PE = 'PE'
COMMENT = 'comments'

# top level data dictionary, ticker is key

dconsolidated = {}


# helper function to add data to dconsolidated

def process(compname, ticker, broker, cols):
  if ticker not in dconsolidated:
    dconsolidated[ticker] = []

  dconsolidated[ticker].append([compname, broker, cols])


# called after all data has been read

def print_consolidated(outfilename):
  header_keys1 = [
    EPS_CHG,
    VS_CONS,
    PE
  ]
  header_keys2 = [
    EPS,
    SALES_CHG,
    EBIT_CHG
  ]

  header_keys3 = [
    TICKER,
    NAME,
    BROKER
  ]

  percentage_keys = [EPS_CHG, VS_CONS, SALES_CHG, EBIT_CHG]

  # "together" join data for multiple years

  def together(cols, key):
    vals = [cols.get(year, {}).get(key, None) for year in years]
    return list(map(lambda x: str(round(x * 100 if key in percentage_keys else x, 2)) \
                     if x and isinstance(x, float) and not math.isnan(x) else '', vals))

  # makes the headers

  def make_header():
    h1 = header_keys3
    for key in header_keys1:
      h1 += [f'{key} {year}' for year in years]
    h1.append(COMMENT)
    for key in header_keys2:
      h1 += [f'{key} {year}' for year in years]

    return h1

  h1 = make_header()
 
  tickers = list(dconsolidated.keys())
  tickers.sort()

  # use last_ticker to not repeat ticker and company name between multiple brokers
  last_ticker = ''
  data = []
  for ticker in tickers:
    for item in dconsolidated[ticker]:
      cols = item[2]

      # if multiple comments, aggregate them
      comments = [cols.get(year, {}).get(COMMENT, None) for year in years]
      comment = ''.join([c for c in comments if c is not None and not pd.isnull(c)])

      print_cols = [
        ticker if ticker != last_ticker else '', 
        item[0] if ticker != last_ticker else '',
        item[1]
      ]
      for key in header_keys1:
        print_cols += together(cols, key)
      print_cols.append(comment)
      for key in header_keys2:
        print_cols += together(cols, key)

      last_ticker = ticker
      data.append(print_cols)

  df = pd.DataFrame(data, columns = h1)
  print(df)
  df.to_excel(outfilename)

# [COMPNAME,TICKER,YEAR,EPS_CHG,EPS,VS_CONS,SALES_CHG,EBIT_CHG,PE,COMMENT
configs = {
  'GS':         [1,2,3,4,5,8,10,None,7,9],
  'JPMORGAN':   [0,1,2,3,4,6,None,9,None,10],
  'ML':         [0,1,2,3,4,6,11,None,9,10],
  'DANSKE':     [1,2,3,4,5,7,11,12,9,10],
  'INTERMONTE': [0,1,5,8,7,10,None,None,12,None]
}

# parse data frame that has been read from file

def parse_df(broker, df):

  if broker not in configs:
    raise RuntimeError(f'broker {broker} not in configs')
  crec = configs[broker]

  i = 0
  nlen = len(df)
  ticker = ''
  lastname = ''
  cols = {}

  while i < nlen:
    row = df.iloc[i].tolist()
    name = row[crec[0]]
    if not pd.isnull(name):
      if lastname:
        process(lastname, ticker, broker, cols)
      # next company
      ticker = row[crec[1]]
      cols = {}
      lastname = name

    s_year = row[crec[2]]
    if not pd.isnull(s_year):
      year = int(s_year[:-1]) if isinstance(s_year, str) else int(s_year)
         
      if year in years:
        comment = '' if crec[9] is None else row[crec[9]]
        if pd.isnull(comment):
          comment = ''

        cols[year] = {
          NAME:      name,
          TICKER:    ticker,
          EPS_CHG:   row[crec[3]] if crec[3] else None,
          EPS:       row[crec[4]] if crec[4] else None,
          VS_CONS:   row[crec[5]] if crec[5] else None,
          SALES_CHG: row[crec[6]] if crec[6] else None,
          EBIT_CHG:  row[crec[7]] if crec[7] else None,
          PE:        row[crec[8]] if crec[8] else None,
          COMMENT:   comment
        }
    i += 1

  process(lastname, ticker, broker, cols)
# end of parse_df

# read goldman file

def read_gs(filename):
  df = pd.read_excel(filename, skiprows=4, header=None)
  parse_df('GS', df)

# read jpm file

def read_jpmorgan(filename):
  df = pd.read_excel(filename, skiprows=5, header=None)
  parse_df('JPMORGAN', df)

# read ml file

def read_ml(filename, sheetname):
  df = pd.read_excel(filename, sheet_name = sheetname, skiprows=3, header=None)
  parse_df('ML', df)

# read danske file

def read_danske(filename, sheetname):
  df = pd.read_excel(filename, sheet_name = sheetname, skiprows=2, header=None)
  parse_df('DANSKE', df)

# read intermonte file

def read_intermonte(filename, sheetname):
  df = pd.read_excel(filename, sheet_name = sheetname, skiprows=4, header=None)
  parse_df('INTERMONTE', df)


if __name__ == '__main__':
  read_gs(r'EPS_CHANGE_20221028_GS.xlsx')
  read_jpmorgan(r'EPS_CHANGE_20221028_JPMCAZ.xlsx')
  read_ml(r'EPS_CHANGES_20221028_ML.xls', 'Est Changes')
  read_danske(r'EPS_CHANGES_20221028_DANSKE.xlsm', 'DailyChanges')
  read_intermonte(r'EPS_CHANGES_20221028_INTERMONTE.xlsx', 'Foglio1')

  # the output file name can be passed on command line, 
  # but i have omitted to do that for this test

  outfilename = 'consol.xlsx'
  print_consolidated(outfilename)

  print(f'created {outfilename} with {len(dconsolidated)} tickers')
