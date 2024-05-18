# Stock Data Fetcher
A simple and easy-to-use Python code for retrieving stock data from Yahoo Finance and organizing it into a spreadsheet. This Data Fetcher tool was inspired by the functionality of the Multiple Stock Quote Downloader spreadsheet available at https://investexcel.net/multiple-stock-quote-downloader-for-excel/. While the spreadsheet works well on most Windows systems, it does not function on Mac. Therefore, I developed an alternative that performs the same tasks across different operating systems. This code includes a user-friendly GUI designed to be accessible for beginners. Additionally, it offers the feature of searching for stock tickers by company name, which was not available in the original spreadsheet method.

# How-To-Setup
1. Install Python by following this guide: [https://github.com/PackeTsar/Install-Python](https://github.com/PackeTsar/Install-Python).

2. (Optional) You can install the necessary libraries listed in the `requirements.txt` file ahead of time, or you can wait and let the errors tell you what you need to install.

3. To use the code, first download the `.py` (Python) file found in the file list of this repository. Move this file into a folder somewhere on your computer. You can rename it if you wish; keep it simple.

4. Now open a terminal in that directory. Instructions on how to do so are listed in the "How to Open a Terminal" section.

5. Once the directory has been established in the terminal, use `python3 whatyounamedit.py` to run the program.

6. If any library has not been installed, an error will occur, and it will tell you which module is not found. Simply download what the message says is missing.

7. Once the code runs successfully, a GUI will pop up. Basic instructions for how to use this, can be found in the "How-To-Use" section. 


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
1. Run the code (#5 in the How-To-Setup) section.
2. List your tickers. You can do this by simply listing the ticker symbols with a comma separating them, or alternatively you can use the search function by typing in a company name. This will pull up a window with options of which stock you mean to select, choosing the stock will add it to the list.
3. Choose start and end date by typing the date or using the calendar. Choose the time interval you wish to use. Click fetch data.
4. Name the excel file and click "OK". 

# Results
1. To view results, navigate to the area in which your code is found. There will be an excel sheet listed in which you named. 
