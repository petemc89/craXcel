# craXcel
Remove pesky Workbook and Worksheet protection from Microsoft Excel files.

### Why have I made this?
Excel Workbook and Worksheet protection is great. It serves its purpose to prevent the defacing (accidental or otherwise) of precious files of both business and personal users. 

However, sometimes full access to the files are needed and the password simply isn't known, whether it's been lost, forgotten, or not provided. 

Comimg from an IT background, typical cases I've personally encountered are:
- I've forgotten my own passwords (d'oh!)
- The previous owner of a spreadsheet has left the business and access is required by their successor
- Suppliers/vendors/partners provide complex spreadsheets as templates that are password protected, and in an effort to backward engineer them to better integrate a solution they need to be unlocked
  - With this one I'd recommend simply asking for the passwords first!
  
Although craXcel has been built to deal with these cases, the reason I created it was to put what I've learnt so far with Python (and programming in general) to the test. 
 
I appreciate it may not be the cleanest of code, nor the most efficient or bug-free, but it's a small program I'm proud to call my own.

### What craXcel is...
craXcel ("crack-excel") is a tool that helps you remove Workbook and Worksheet protection from Microsoft Excel files. It achieves this by accessing the underlying XML structure of the file and removing the areas associated with protection.

### What craXcel isn't...
craXcel does not ever "crack" or otherwise know the passwords that protect the Workbook and Worksheets, it simply circumvents them. It also cannot break file encryption (i.e. when a password is needed just to open the file).

## Installation

### Prerequisites								
1. Download and install Python (https://www.python.org/downloads/)								

### Step-by-Step								
1. Clone or download the repository from GitHub to a local folder of your choosing								
2. Open a terminal of your choice (i.e. cmd, powershell, bash)								
3. In the terminal navigate to the folder with craxcel (from step 1)								
    - If you're not familiar with how to do this, search online for: "[name of terminal] change directory"								
4. In the terminal, enter the command: <b>pip install -r requirements.txt</b>								
    - If you have trouble, try opening the terminal as an administrator								
5. You are now good to go! Refer to the Usage section below for instructions on how to use craXcel								

### If You Get Stuck…								
- Guide to Downloading and Installing Python (https://realpython.com/installing-python/)							
- Cloning a GitHub Repository (https://help.github.com/en/github/creating-cloning-and-archiving-repositories/cloning-a-repository)								

## Usage

### Basic								
1. Copy the file you wish to unlock to the folder with craxcel								
2. Open a terminal of your choice (i.e. cmd, powershell, bash)								
3. In the terminal, navigate to the folder with craxcel								
4. In the terminal, enter the command: python craxcel.py yourfilename.xlsx								
    - The terminal doesn't necessarily have to be in the same folder as craxcel, and nor does your file								
    - If craxcel is not located in the current directory, simply give the terminal the full path, i.e.:								
      - <b>python 'c:/users/me/downloads/craxcel/craxcel.py' yourfile.xlsx</b>								
    - The same applies in the case of your file, simply give the terminal the full path, i.e.:								
      - <b>python craxcel.py 'c:/users/me/documents/yourfile.xlsx'</b>								
    - And you can, of course, combine these for both if required, i.e.:								
      - <b>python 'c:/users/me/downloads/craxcel/craxcel.py' 'c:/users/me/documents/yourfile.xlsx'</b>								

### Advanced								
- craXcel has several options that can be passed in for more advanced uses, i.e.:								
  - Selecting to only remove Workbook protection (leaving Worksheet protection intact)								
    - <b>python craxcel.py yourfile.xlsx -wb</b>							
  - Selecting to only remove Worksheet protection (leaving Workbook protection intact)								
    - <b>python craxcel.py yourfile.xlsx -ws</b>							
  - Running without creating a backup file								
    - <b>python craxcel.py yourfile.xlsx --no_backup</b>								
- For a full list of options, enter the command: <b>python craxcel.py --help</b>							

### Supported Formats
Right now, supported formats are limited to the more modern .xlsx and .xlsm formats. Other formats may work, but have not been explicity tested.

### Contribution
If you have a feature you would like to see you can either raise an issue in the GitHub repository, or branch off and give it a go yourself!
