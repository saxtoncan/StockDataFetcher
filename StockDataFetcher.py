import yfinance as yf
import pandas as pd
import time
import tkinter as tk
from tkinter import messagebox, simpledialog
from tkinter import ttk
from tkcalendar import DateEntry
from threading import Thread
import os

# Global flag to stop fetching process
cancel_flag = False


# Function to fetch data and save to Excel
def fetch_and_save_data(tickers, start_date, end_date, interval, filename, progress_var, progress_label):
    global cancel_flag
    total_tickers = len(tickers)

    # Create DataFrames to store individual stock data
    individual_data = {}

    # Create DataFrames to store combined data
    combined_data = {stat: pd.DataFrame() for stat in ['Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']}

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for i, ticker in enumerate(tickers):
            if cancel_flag:
                messagebox.showinfo("Cancelled", "Data fetching process was cancelled.")
                break
            try:
                print(f"Fetching data for {ticker} with interval {interval}")
                stock_data = yf.download(ticker, start=start_date, end=end_date, interval=interval)
                stock_data.index = stock_data.index.strftime('%m/%d/%Y')  # Format the date column

                # Write individual stock data to sheets
                stock_data[['Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']].to_excel(writer, sheet_name=ticker)

                # Store individual stock data
                individual_data[ticker] = stock_data

                # Combine individual data for each statistic
                for stat in ['Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']:
                    combined_data[stat][ticker] = stock_data[stat]

                time.sleep(2)  # Delay to avoid hitting rate limits
            except Exception as e:
                print(f"Error fetching data for {ticker}: {e}")

            # Update progress
            progress = int((i + 1) / total_tickers * 100)
            progress_var.set(progress)
            progress_label.config(text=f'Progress: {progress}%')

        # Remove individual sheets containing only one type of data
        for ticker in tickers:
            if f"{ticker}" in writer.sheets and f"{ticker}_Adjusted_Prices" in writer.sheets:
                writer.sheets.pop(f"{ticker}")
                writer.sheets.pop(f"{ticker}_Adjusted_Prices")

        # Write combined data for each statistic to separate sheets, in the desired order
        for stat in ['Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']:
            combined_data[stat].to_excel(writer, sheet_name=stat)

    if not cancel_flag:
        messagebox.showinfo("Success", f"Data saved successfully to {filename}")

    # Reset cancel flag
    cancel_flag = False


# Function to handle the fetch button click
def on_fetch_click():
    global cancel_flag
    cancel_flag = False
    tickers = ticker_entry.get().split(',')
    start_date = start_date_entry.get_date().strftime('%Y-%m-%d')
    end_date = end_date_entry.get_date().strftime('%Y-%m-%d')
    interval = interval_var.get()
    filename = simpledialog.askstring("Input", "Enter the filename for the Excel file:",
                                      initialvalue="historical_prices.xlsx")
    if filename:
        # Run data fetching in a separate thread to keep the GUI responsive
        thread = Thread(target=fetch_and_save_data,
                        args=(tickers, start_date, end_date, interval, filename, progress_var, progress_label))
        thread.start()


# Function to handle the cancel button click
def on_cancel_click():
    global cancel_flag
    cancel_flag = True


# Create the main window
root = tk.Tk()
root.title("Stock Data Fetcher")

# Create and place the widgets
tk.Label(root, text="Tickers (comma separated):").grid(row=0, column=0, padx=10, pady=10)
ticker_entry = tk.Entry(root, width=50)
ticker_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Start Date:").grid(row=1, column=0, padx=10, pady=10)
start_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
start_date_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="End Date:").grid(row=2, column=0, padx=10, pady=10)
end_date_entry = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
end_date_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Interval:").grid(row=3, column=0, padx=10, pady=10)
interval_var = tk.StringVar(value='1d')  # Default interval is set to '1d'
interval_menu = ttk.Combobox(root, textvariable=interval_var, values=['1d', '1wk', '1mo'])
interval_menu.grid(row=3, column=1, padx=10, pady=10)

fetch_button = tk.Button(root, text="Fetch Data", command=on_fetch_click)
fetch_button.grid(row=4, column=0, columnspan=2, pady=10)

cancel_button = tk.Button(root, text="Cancel", command=on_cancel_click)
cancel_button.grid(row=5, column=0, columnspan=2, pady=10)

progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky='ew')

progress_label = tk.Label(root, text="Progress: 0%")
progress_label.grid(row=7, column=0, columnspan=2)

# Run the application
root.mainloop()
