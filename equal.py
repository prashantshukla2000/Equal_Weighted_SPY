import numpy as np
import pandas as pd
import yfinance as yf
import math

# Read the list of S&P 500 stocks
# Change the source repository as per your file location.
stocks = pd.read_csv(r"C:\Users\MAA\Desktop\pATTERSON\My projects\Equal_weight\algorithmic-trading-python\starter_files\sp_500_stocks.csv")

# Define columns for final DataFrame
my_columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns=my_columns)

# Fetch stock data for each ticker
for stock in stocks['Ticker']:
    ticker = yf.Ticker(stock)
    info = ticker.info
    new_row = pd.DataFrame([{
        'Ticker': stock,
        'Stock Price': info.get('currentPrice', 'N/A'),
        'Market Capitalization': info.get('marketCap', 'N/A'),
        'Number of Shares to Buy': 'N/A'
    }])
    final_dataframe = pd.concat([final_dataframe, new_row], ignore_index=True)

# Input for portfolio size from user
while True:
    portfolio_size = input('Enter the value of your Portfolio: ')
    try:
        portfolio_size = float(portfolio_size)
        break  # exit loop if valid
    except ValueError:
        print("That's not a number. Please try again.\n")


# Clean Stock Price column to numeric
final_dataframe['Stock Price'] = pd.to_numeric(final_dataframe['Stock Price'], errors='coerce')


# Calculate position size per stock
position_size = portfolio_size / len(final_dataframe.index)

# Calculate number of shares to buy for each stock
for i in range(len(final_dataframe.index)):
    stock_price = final_dataframe.loc[i, 'Stock Price']
    if pd.notnull(stock_price) and stock_price > 0:
        final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / stock_price)
    else:
        final_dataframe.loc[i, 'Number of Shares to Buy'] = 0


# Export the final DataFrame to Excel with formatting
excel_file_path = r"C:\Users\MAA\Desktop\pATTERSON\My projects\Equal_weight\algorithmic-trading-python\starter_files\Recommended trades.xlsx"

with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
    # Write DataFrame to Excel sheet
    final_dataframe.to_excel(writer, sheet_name='Recommended Trades', index=False)

    # Access workbook and worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Recommended Trades']

    # Define cell formats
    background_colour = '#0a0a23'
    font_colour = '#ffffff'

    string_format = workbook.add_format({
        'font_color': font_colour,
        'bg_color': background_colour,
        'border': 1
    })

    dollar_format = workbook.add_format({
        'num_format': '$0.00',
        'font_color': font_colour,
        'bg_color': background_colour,
        'border': 1
    })

    integer_format = workbook.add_format({
        'num_format': '0',
        'font_color': font_colour,
        'bg_color': background_colour,
        'border': 1
    })

# Define format mapping for each column
    column_formats = {
        'A': ['Ticker', string_format],
        'B': ['Stock Price', dollar_format],
        'C': ['Market Capitalization', dollar_format],
        'D': ['Number of Shares to Buy', integer_format]
    }

# Apply formatting and set column widths
    for column in column_formats.keys():
        worksheet.set_column(f'{column}:{column}', 20, column_formats[column][1])
        worksheet.write(f'{column}1', column_formats[column][0], column_formats[column][1])


#  Message

print(f"\n Recommended trades saved successfully at:\n{excel_file_path}")
