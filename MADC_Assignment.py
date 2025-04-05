import queue
import pandas as pd

#read the csv file
def read_excel():
    while True:
        file_name = input("Enter the file name (include the file type e.g. file.xlsx): ")
        try:
            data = pd.read_excel(file_name, parse_dates=['Date'], index_col='Date')#'SPY_2016_2021.xlsx'
            data = data[['Close']].dropna()
            print('File read!')
            return data
        except FileNotFoundError:
            print(f"Error: The file '{file_name}' was not found.")
        except ValueError as e:
            print(f"Error: There was a problem with the file format. {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

#Let the user choose calculation type
def cal_type():
    calculation = ''
    while calculation not in ['SMA', 'EMA']:
        calculation = input('Select calculation for moving average (SMA/EMA):').strip().upper()
    return calculation
    
def moving_average(data, window, method='EMA'):
    if method == 'SMA':
        return data.rolling(window=window).mean()
    else:
        return data.ewm(span=window, adjust=False).mean()

def compute_macd(df, method='EMA'):
    df['Short_MA'] = moving_average(df['Close'], 12, method)
    df['Long_MA'] = moving_average(df['Close'], 26, method)
    df['MACD'] = df['Short_MA'] - df['Long_MA']
    df['Signal_Line'] = moving_average(df['MACD'], 9, method)
    df['Histogram'] = df['MACD'] - df['Signal_Line']
    df.fillna(0, inplace=True) # for SMA calculation
    return df

def generate_signals(df, threshold=0.01):
    df['Signal'] = 0  
    # Buy when MACD crosses above Signal Line and difference is significant
    df.loc[(df['MACD'] > df['Signal_Line']) & (df['MACD'].shift(1) <= df['Signal_Line'].shift(1)) & 
           (abs(df['MACD'] - df['Signal_Line']) > threshold), 'Signal'] = 1

    # Sell when MACD crosses below Signal Line and difference is significant
    df.loc[(df['MACD'] < df['Signal_Line']) & (df['MACD'].shift(1) >= df['Signal_Line'].shift(1)) & 
           (abs(df['MACD'] - df['Signal_Line']) > threshold), 'Signal'] = -1

def macd_calculation(new_df):
    holdings = 0
    buy_list = []  
    total_commission = 0
    threshold = 0.1 #to counter weak trend reversal
    trade_log = []
    commission=0.00125
    capital = 100000.00

    for i, row in new_df.iterrows():
        macd_signal_diff = abs(row['MACD'] - row['Signal_Line'])
        
        # Ignore weak signals
        if macd_signal_diff < threshold:
            continue
        
        # BUY Condition
        if row['Signal'] == 1 and holdings == 0:
            if capital > 0:
                buy_price = row['Close']
                quantity = capital / buy_price  # Buy as many shares as possible
                commission_fee = quantity * buy_price * commission  # Calculate commission
                
                capital -= commission_fee  # Deduct commission
                total_commission += commission_fee
                holdings += quantity  # Increase holdings
                buy_list.append((buy_price, quantity))  # Store buy transaction
                
                trade_log.append({
                    'Date': row.name, 'Trade': 'BUY', 'Price': buy_price,
                    'Holdings': holdings, 'Capital': capital,
                    'Profit': 0, 'Commission_Lost': commission_fee
                })
        
        # SELL Condition
        elif row['Signal'] == -1 and holdings > 0:
            sell_price = row['Close']
            if sell_price < trade_log[-1]['Price']:
                continue
            else:
                sell_value = holdings * sell_price
                commission_fee = sell_value * commission  # Commission on sale
            
                # Calculate profit based on all buy transactions
                total_cost = sum(buy_price * qty for buy_price, qty in buy_list)
                profit = sell_value - total_cost - commission_fee
                
                capital = sell_value - commission_fee  # Update capital
                holdings = 0  # Reset holdings
                buy_list.clear()  # Clear buy records after selling
                total_commission += commission_fee
                
                trade_log.append({
                    'Date': row.name, 'Trade': 'SELL', 'Price': sell_price,
                    'Holdings': holdings, 'Capital': capital,
                    'Profit': profit, 'Commission_Lost': commission_fee
                })

    if trade_log[-1]['Trade'] == 'BUY' and holdings > 0:
        last_date = new_df.index[-1]
        last_price = new_df.loc[last_date, 'Close']
        
        sell_value = holdings * last_price
        commission_fee = sell_value * commission
        total_cost = sum(buy_price * qty for buy_price, qty in buy_list)

        profit = sell_value - total_cost - commission_fee
        capital = sell_value - commission_fee
        holdings = 0
        total_commission += commission_fee
        
        trade_log.append({
            'Date': last_date,
            'Trade': 'SELL',
            'Price': last_price,
            'Holdings': holdings,
            'Capital': capital,
            'Profit': profit,
            'Commission_Lost': commission_fee
        })
    trade_df = pd.DataFrame(trade_log)
    return trade_df

def buy_hold_sell(new_df):
    buy_hold_sell = []
    commission=0.00125
    capital_2 = 100000.00
    total_commission_2 = 0
    buy_list=[]
    holdings = 0

    first = new_df.iloc[0]
    last = new_df.iloc[-1]

    buy_price = first['Close']
    quantity = capital_2 / buy_price  # Buy as many shares as possible
    commission_fee_2 = quantity * buy_price * commission  # Calculate commission

    capital_2 -= commission_fee_2  # Deduct commission
    total_commission_2 += commission_fee_2
    holdings += quantity  # Increase holdings
    buy_list.append((buy_price, quantity))  # Store buy transaction

    buy_hold_sell.append({
        'Date': first.name, 'Trade': 'BUY', 'Price': buy_price,
        'Holdings': holdings, 'Capital': capital_2,
        'Profit': 0, 'Commission_Lost': commission_fee_2
    })

    sell_price = last['Close']
    sell_value = holdings * sell_price
    commission_fee_3 = sell_value * commission  # Commission on sale

    # Calculate profit based on all buy transactions
    total_cost = sum(buy_price * qty for buy_price, qty in buy_list)
    profit = sell_value - total_cost - commission_fee_3

    capital_2 = sell_value - commission_fee_3  # Update capital
    holdings = 0  # Reset holdings
    buy_list.clear()  # Clear buy records after selling
    total_commission_2 += commission_fee_3

    buy_hold_sell.append({
        'Date': last.name, 'Trade': 'SELL', 'Price': sell_price,
        'Holdings': holdings, 'Capital': capital_2,
        'Profit': profit, 'Commission_Lost': commission_fee_3
    })

    buy_hold_sell_df = pd.DataFrame(buy_hold_sell)
    return buy_hold_sell_df

def print_menu():
    print("""
    ***************************************
    *                                     *
    *    MACD Trading Console Program     *
    *                                     *
    ***************************************
    *    1. Start MACD Trading            *
    *    0. Exit                          *
    ***************************************
    """) 

def main():
    while True:
        print_menu()
        # Take user input
        user_choice = input("Please select an option (1 to Start, 0 to Exit): ").strip()

        if user_choice == "1":
            df = read_excel()
            calculation = cal_type()
            new_df = compute_macd(df, calculation)
            print(f"Selected Calculation: {calculation}")
            print('Generating report...')
            generate_signals(new_df)
            trade_df = macd_calculation(new_df)
            buy_hold_sell_df = buy_hold_sell(new_df)
            # Display trade dataframe in a cleaner way
            print("\n===== MACD Trade Log =====")
            print(trade_df[['Date', 'Trade', 'Price', 'Holdings', 'Capital', 'Profit', 'Commission_Lost']].to_string(index=False))  

            # Display Buy-Hold-Sell comparison
            print("\n===== Buy-Hold-Sell Summary =====")
            print(buy_hold_sell_df.to_string(index=False))

            #save stock series
            with pd.ExcelWriter('stock_series.xlsx') as writer:  
                trade_df.to_excel(writer, sheet_name='macd')
                buy_hold_sell_df.to_excel(writer, sheet_name='buy_hold_sell')
            print('Saved stock series to excel file')

            # Summary statistics
            total_trades = int(len(trade_df) / 2)
            avg_return = trade_df['Profit'].sum() / total_trades if total_trades > 0 else 0
            bhs = buy_hold_sell_df.iloc[1]['Capital']
            macd = trade_df.iloc[-1]['Capital']

            # Pretty report
            print("\n===== Strategy Comparison Report =====")
            print(f"Total number of completed MACD trades: {total_trades}")
            print(f"Average return per completed trade: ${avg_return:.2f}")
            print(f"Final Capital using Buy-Hold-Sell strategy: ${bhs:,.2f}")
            print(f"Final Capital using MACD strategy: ${macd:,.2f}")

            if macd > bhs:
                print(f"✅ MACD strategy outperformed Buy-Hold-Sell by: ${macd - bhs:,.2f}")
            elif bhs > macd:
                print(f"📉 Buy-Hold-Sell strategy outperformed MACD by: ${bhs - macd:,.2f}")
            else:
                print("⚖️ Both strategies resulted in the same capital.")
        elif user_choice == "0":
            print("Exiting program...")
            break
        else:
            print("Invalid input. Please select either 1 or 0.")

if __name__ == "__main__":
    main()
    
