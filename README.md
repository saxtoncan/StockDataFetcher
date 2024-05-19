# Stock Data Fetcher
A simple and easy-to-use Python code for retrieving stock data from Yahoo Finance and organizing it into a spreadsheet. This Data Fetcher tool was inspired by the functionality of the Multiple Stock Quote Downloader spreadsheet available at https://investexcel.net/multiple-stock-quote-downloader-for-excel/. While the spreadsheet works well on most Windows systems, it does not function on Mac. Therefore, I developed an alternative that performs the same tasks across different operating systems. This code includes a user-friendly GUI designed to be accessible for beginners. Additionally, it offers the feature of searching for stock tickers by company name, which was not available in the original spreadsheet method.

# How-To-Setup

1. [Install Python for MacOS/Linux](https://www.python.org/downloads/). For Windows it's best to download newest version from Mocrosoft Store.

2. (Optional) You can install the necessary libraries listed in the `requirements.txt` file ahead of time, or you can wait and let the errors guide you on what you need to install. If you choose to use this method, place the `requirements.txt` file in the same folder as the `.py` file in the next step, and then run `pip3 install -r requirements.txt` in the terminal in the directory specified in step four. This command should install all the required libraries.

3. To use the code, first download the `.py` (Python) file found in the file list of this repository. Move this file into a folder on your computer. You can rename it if you wish, but keep it simple.

4. Now open a terminal in that directory. Instructions on how to do so are provided in the "How to Open a Terminal" section.

5. Once the directory has been established in the terminal, use `python3 whatyounamedit.py` to run the program.

6. If any library has not been installed, an error will occur, indicating which module is not found. Simply install the missing module as instructed by the error message.

7. Once the code runs successfully, a GUI will pop up. Basic instructions for how to use it can be found in the "How-To-Use" section. 

# Library information:
1. See requirements.txt file for full list.
2. To install a library run "pip3 install module" in the terminal, replacing module with the required library.

# How to Open a Terminal 

## Linux
1. **Using a File Manager**:
   - Open your file manager and navigate to the desired directory.
   - Right-click inside the folder.
   - Select "Open in Terminal" from the context menu.

2. **Using the Terminal**:
   - Open a terminal.
   - Use the `cd` command to change to the desired directory. For example:
     ```bash
     cd /path/to/your/directory
     ```

## macOS
1. **Using Finder**:
   - Open Finder and navigate to the desired directory.
   - Right-click (or Control-click) inside the folder.
   - Select "New Terminal at Folder" (you might need to enable this service first in System Preferences under "Keyboard" > "Shortcuts" > "Services").

2. **Using the Terminal**:
   - Open a terminal.
   - Use the `cd` command to change to the desired directory. For example:
     ```bash
     cd /path/to/your/directory
     ```

## Windows
1. **Using File Explorer**:
   - Open File Explorer and navigate to the desired directory.
   - Click on the address bar, type `cmd`, and press `Enter` to open Command Prompt in that directory.
   - Alternatively, type `powershell` and press `Enter` to open PowerShell in that directory.

2. **Using Command Prompt or PowerShell**:
   - Open Command Prompt or PowerShell.
   - Use the `cd` command to change to the desired directory. For example:
     ```cmd
     cd C:\path\to\your\directory
     ```
# How-To-Use
1. Run the code (refer to step #5 in the How-To-Setup section).
2. List your tickers. You can do this by either listing the ticker symbols separated by commas, or by using the search function to type in a company name. This will open a window with options for the stock you intend to select; choosing the stock will add it to the list.
3. Choose the start and end date by typing the date or using the calendar. Then, select the time interval you wish to use and click "Fetch Data".
4. Name the Excel file and click "OK".

# Results
1. To view the results, navigate to the area where your code is located. There will be an Excel sheet named as you specified. Each stock will have its own sheet of data containing the open, high, low, close, adjusted closing, and volume. In addition to the individual stock sheets, there are summary sheets with all the stocks combined for each of those categories.

# Thanks
Please let me know if you have any suggestions or upgrades you would like implemented, and I will look into them.
