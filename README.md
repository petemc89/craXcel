# craXcel-cli (v2.0)
Python command line application to unlock Microsoft Office password protected files.

---

![craxcel-cli-basic](https://user-images.githubusercontent.com/50495755/95125116-60ed3780-074c-11eb-8547-0e28cb4f28c5.png)

![craxcel-cli-list](https://user-images.githubusercontent.com/50495755/95125877-7adb4a00-074d-11eb-9a1c-d6a7406717af.png)

---

# What is craXcel

craXcel ("crack-cel") is a tool that makes removing various password protections from Microsoft Office files seemless. It works by directly amending the underlying XML files that make up modern Microsoft Office files.

Please note that craXcel cannot unlock encrypted files.

---

# Supported applications

As of V2.0:

- Microsoft Excel (workbook, worksheet, vba)
  - .xlsx
  - .xlsm
- Microsoft Word (modify, format, vba)
  - .docx
  - .docm
- Microsoft Powerpoint (modify, vba)
  - .pptx
  - .pptm

Others may work, but have not been tested.

## Important note on unlocking the VBA project of macro files

Upon unlocking the VBA Project of a Macro Enabled file, that file will state it has encountered issues and needs to recover... __DO NOT PANIC__, this is normal.

The steps to follow to complete the unlock is as follows:

1. Open the unlocked file and click 'Enable Content' on the warning:

![image](https://user-images.githubusercontent.com/50495755/94193731-9e2e0b80-fea8-11ea-818f-45ac9ac7b80e.png)

2. Click 'OK' on the following pop-up:

![image](https://user-images.githubusercontent.com/50495755/94193790-b56cf900-fea8-11ea-8f73-2b27378b1e3d.png)

3. Open Visual Basic from the Developer toolbar:

![image](https://user-images.githubusercontent.com/50495755/94193894-d59cb800-fea8-11ea-9cc6-6a88008a853e.png)

4. Open VBAProject Propeties under Tools:

![image](https://user-images.githubusercontent.com/50495755/94193982-f5cc7700-fea8-11ea-8dad-9d0ccb3cf921.png)

5. Navigate to the Protection tab and enter a new password (a one character password is fine, as we will be removing it again straight away). Click 'OK'.

![image](https://user-images.githubusercontent.com/50495755/94194050-0ed52800-fea9-11ea-9cf9-315a1a0fc7fc.png)

6. Head back in to VBAProject Properties > Protection tab, and de-select the 'Lock project for viewing' checkbox and clear any passwords in the boxes below. Click 'OK'.

7. The modules will now be unlocked and you can save the document without having to repeat these steps.

![image](https://user-images.githubusercontent.com/50495755/94194188-40e68a00-fea9-11ea-9f1d-77ea49010a4b.png)

__Note to developers:__ If you're willing to take on the challenge of automating these steps (preferably without user input mimicking...) you are welcome to contribute!

---

## Installation

### Prerequisites								
1. Download and install Python (v3+) (https://www.python.org/downloads/)								

### Step-by-Step								
1. Clone or download the repository from GitHub to a local folder of your choosing								
1. Open a terminal of your choice (i.e. cmd, powershell, bash)								
1. In the terminal navigate to the folder with craxcel (from step 1)								
    - If you're not familiar with how to do this, search online for: "[name of terminal] change directory"								
1. In the terminal, enter the command: __pip install -r requirements.txt__							
    - If you have trouble, try opening the terminal as an administrator								
1. You are now good to go! Refer to the Usage section below for instructions on how to use craXcel								

### If You Get Stuck…								
- Guide to Downloading and Installing Python (https://realpython.com/installing-python/)							
- Cloning a GitHub Repository (https://help.github.com/en/github/creating-cloning-and-archiving-repositories/cloning-a-repository)								

## Usage

### Basic			
1. Open a terminal of your choice (i.e. cmd, powershell, bash)								
1. In the terminal, navigate to the folder with craxcel								
1. In the terminal, enter the command: python craxcel.py yourfilename.xlsx								
    - The terminal doesn't necessarily have to be in the same folder as craxcel, and nor does your file								
    - If craxcel is not located in the current directory, simply give the terminal the full path, i.e.:								
      - __python 'c:/users/me/downloads/craxcel/craxcel.py' yourfile.xlsx__						
    - The same applies in the case of your file, simply give the terminal the full path, i.e.:								
      - __python craxcel.py 'c:/users/me/documents/yourfile.xlsx'__								
    - And you can, of course, combine these for both if required, i.e.:								
      - __python 'c:/users/me/downloads/craxcel/craxcel.py' 'c:/users/me/documents/yourfile.xlsx'__
1. The unlocked file will be saved in the created 'unlocked' folder where the app is installed

### List Mode
craXcel also has the ability to unlock multiple files at a time!

1. Create a .txt file with a line for each filepath (see __file-list-example.txt__ for an example)
1. Instead of entering the filename of an individual Microsoft Office application, enter the .txt filename
1. Finish the command by entering '--list', i.e.
   - __python craxcel.py 'c:/users/me/documents/list-of-files.txt' --list__

### Options								
- craXcel has several options that can be passed in for more advanced uses, i.e.:			
  - Unlock the VBA Project (macro file) of a macro enabled file
    - __python craxcel.py yourfile.xlsm --vba__
  - Selecting to only remove Workbook protection (leaving Worksheet protection intact)								
    - __python craxcel.py yourfile.xlsx -wb__						
  - Selecting to only remove Worksheet protection (leaving Workbook protection intact)								
    - __python craxcel.py yourfile.xlsx -ws__						
  - Run without deleting the temporary XML files							
    - __python craxcel.py yourfile.xlsx --debug__							
- For a full list of options, enter the command: __python craxcel.py --help__						

### Contribution
If you have a feature you would like to see you can either raise an issue in the GitHub repository, or branch off and give it a go yourself!
