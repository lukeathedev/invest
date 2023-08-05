import os, sys
from time import sleep
import xlwings as xw
import yahooquery as yq

# args: x, y, w

TBL_ROW   = 2
TBL_COL   = 3
TBL_WIDTH = 13

if len(sys.argv) < 4:
  print("You can supply the table position and width as")
  print("'x y w' when calling the program.")

if len(sys.argv) >= 2:
  TBL_ROW   = int(sys.argv[1])
if len(sys.argv) >= 3:
  TBL_COL   = int(sys.argv[2])
if len(sys.argv) >= 4:
  TBL_WIDTH = int(sys.argv[3])

path = os.path.dirname(os.path.realpath(__file__))
fn   = os.path.join(path, "stocks.xlsm")

wb  = xw.Book(fn)
wks = xw.sheets

# Sheet 0 (1)
ws = wks[0]

# Get stocks
i = TBL_ROW; stocks = []
while True:
  cell = ws.cells(i, TBL_COL).value
  if cell is None:
    break

  stocks.append(cell)
  i += 1

# TODO: multithreading / concurrent fetching

for i, stock in enumerate(stocks):
  print(f"Fetching data for {stock}...")

  r_st = (i + TBL_ROW, TBL_COL + 1)
  r_ed = (i + TBL_ROW, TBL_COL + TBL_WIDTH)

  # Reset cells
  ws.range(r_st, r_ed).value = ""

  # Get stock information
  ticker = yq.Ticker(stock)
  info1 = ticker.price[stock]
  info2 = ticker.summary_detail[stock]
  info3 = ticker.summary_profile[stock]

  if type(info1) is str or type(info2) is str or type(info3) is str:
    print(f"{info1}")
    continue

  # Price / Equity
  bals = ticker.balance_sheet(trailing=True).iloc[-1] # Latest report
  price_equity = info2["marketCap"] / bals["StockholdersEquity"]

  print(f"Reference date: {info1['regularMarketTime']}")

  try:
    fte = info3["fullTimeEmployees"]
  except: 
    fte = "N/A"

  # Add to worksheet
  ws.range(r_st, r_ed).value = [
    info1["longName"],                    # Nome
    info3["sector"],                      # Segmento
    fte,                                  # Funcionários
    info2["currency"],                    # Moeda
    info1["regularMarketPrice"],          # Preco
    info2["forwardPE"],                   # P/E
    price_equity,                         # Price / Equity
    info2["fiftyTwoWeekLow"],             # Min. YTD
    info2["fiftyTwoWeekHigh"],            # Max. YTD
    info2["trailingAnnualDividendYield"], # DY (%)
    info2["marketCap"],                   # Market Cap
    info1['regularMarketTime']            # Data
  ]
  
  # Set formatting
  # ws.range(r_st, (i + TBL_ROW, TBL_COL)).number_format = [
  #   "",                # Nome
  #   "",                # Segmento
  #   "",                # Funcionários
  #   "",                # Moeda
  #   "0.00",            # Preço
  #   "0.00",            # P/E
  #   "0.00",            # Price / Equity
  #   "0.00",            # Min. YTD
  #   "0.00",            # Máx. YTD
  #   "0.00%",           # DY (TTM)
  #   "0",               # Market Cap
  #   "dd/mm/yyyy hh:mm" # Data
  # ]

  i += 1