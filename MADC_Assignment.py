import queue
import pandas as pd

#read the csv file
def read_excel():
    while True:
        file_name = input("Enter the file name : ")
        try:
            data = pd.read_excel(file_name, parse_dates=['Date'], index_col='Date')#'SPY_2016_2021.xlsx'
            data = data[['Close']].dropna()
            print(data.head())
            return data
        except FileNotFoundError:
            print(f"Error: The file '{file_name}' was not found.")
        except ValueError as e:
            print(f"Error: There was a problem with the file format. {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

#declare constant variables
SMA = 12 
LMA = 26 
Commission = 0.0125
Initial_capital = 100,000.00

#Let the user choose calculation type
def cal_type():
    calculation = ''
    while calculation not in ['SMA', 'EMA']:
        calculation = input('Select calculation for moving average (SMA/EMA):').strip()
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
    return df
'''
Step 1: find all trend reversal and put in table
Step 2: Find the buy (even if the first few is sell signal ignore)
Step 3: Find the sell (only can sell if there is a previous buy) 
- cannot buy multiple times (assume use all captial for one buy and sell pair)
- Requirements to sell: price at least 10% higher than bought price
- If the cycle end and still have shares, sell it on last day prices no matter the price
- advance: find highest and lowest peak index, save it
Step 4: store all that shit inside an excel
Step 5: Compare with buy sell strat
'''

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
            read_excel()
            calculation = cal_type()  # Store the result of cal_type()
            calculation.upper()
            print(f"Selected Calculation: {calculation}")
        elif user_choice == "0":
            print("Exiting program...")
            break
        else:
            print("Invalid input. Please select either 1 or 0.")

if __name__ == "__main__":
    main()
    
