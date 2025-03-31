import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Load and preprocess data
def load_data(filepath):
    df = pd.read_excel(filepath, parse_dates=['Date'], index_col='Date')
    df = df[['Close']].dropna()
    return df

# Compute Moving Averages
def moving_average(data, window, method='EMA'):
    if method == 'SMA':
        return data.rolling(window=window).mean()
    else:
        return data.ewm(span=window, adjust=False).mean()

# Compute MACD and Signal Line
def compute_macd(df, method='EMA'):
    df['Short_MA'] = moving_average(df['Close'], 12, method)
    df['Long_MA'] = moving_average(df['Close'], 26, method)
    df['MACD'] = df['Short_MA'] - df['Long_MA']
    df['Signal_Line'] = moving_average(df['MACD'], 9, method)
    df['Histogram'] = df['MACD'] - df['Signal_Line']
    return df

# Identify Buy/Sell signals
def generate_signals(df, threshold=0):
    df['Signal'] = 0
    df.loc[(df['MACD'] > df['Signal_Line']) & (df['Histogram'].abs() > threshold), 'Signal'] = 1
    df.loc[(df['MACD'] < df['Signal_Line']) & (df['Histogram'].abs() > threshold), 'Signal'] = -1
    return df

# Execute Trades
def execute_trades(df, initial_capital=100000, commission=0.00125):
    capital = initial_capital
    holdings = 0
    trade_log = []
    
    for date, row in df.iterrows():
        if row['Signal'] == 1 and capital > 0:
            holdings = capital / row['Close']
            capital -= capital * commission
            trade_log.append((date, 'BUY', row['Close'], capital, holdings))
        elif row['Signal'] == -1 and holdings > 0:
            capital = holdings * row['Close']
            capital -= capital * commission
            holdings = 0
            trade_log.append((date, 'SELL', row['Close'], capital, holdings))
    
    return capital, trade_log

# Compare with Buy-Hold Strategy
def buy_hold_strategy(df, initial_capital=100000, commission=0.00125):
    buy_price = df.iloc[0]['Close']
    sell_price = df.iloc[-1]['Close']
    capital = (initial_capital / buy_price) * sell_price
    capital -= capital * commission * 2
    return capital

# Run the system
def main(filepath, method='EMA', threshold=0):
    df = load_data(filepath)
    df = compute_macd(df, method)
    df = generate_signals(df, threshold)
    macd_final_capital, trade_log = execute_trades(df)
    buy_hold_final_capital = buy_hold_strategy(df)
    
    print("MACD Strategy Final Capital:", macd_final_capital)
    print("Buy-Hold Strategy Final Capital:", buy_hold_final_capital)
    
    # Plot results
    plt.figure(figsize=(12,6))
    plt.plot(df.index, df['MACD'], label='MACD')
    plt.plot(df.index, df['Signal_Line'], label='Signal Line')
    plt.legend()
    plt.title('MACD vs Signal Line')
    plt.show()
    
    return trade_log

# Example Usage
trade_log = main('SPY_2016_2021.xlsx', method='EMA', threshold=0.01)