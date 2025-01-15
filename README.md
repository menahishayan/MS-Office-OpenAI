# OpenAI/Gemini Integration for Microsoft Office 365 Apps

This project attempts to create an open source implementation of OpenAI and Gemini to work with MS Office 365 Apps such as Excel, Word, and Powerpoint.

The key<sup>(no pun intended)</sup> differentiator between this project and an Out-of-box solution from the Marketplace is that you don't have to pay middle-man subscription fees to an OpenAI wrapper. The only thing you need to pay for is direct access to OpenAI/Gemini Keys.

> Note: this has currently only been tested with Excel 2019 and above, but feel free to test and raise a PR/Issue for other versions.

## Installation

### API Setup
- Generate your API key from OpenAI/Gemini. Be sure to add a card for OpenAI keys to work.
- `cp core/.env.sample core/.env`
- Paste your API keys in **core/.env** as `OPENAI_API_KEY=` or `GEMINI_API_KEY=`

### MacOS/Linux
```sh
mv . ~/Documents # Move the root folder to a convenient location
cd core
python3 -m venv . # Set up a Python virtual environment
source bin/activate # Activate environment
```
```sh
pip install -r requirements.txt # Install dependencies
deactivate # Exit environment
```
```sh
cp ../ShellExecutor.scpt ~/Library/Application Scripts/com.microsoft.Excel/ # copy shell executor to Excel Scripts
```
- Open the Excel file that you're looking to integrate your AI into.
- Enable the Developer tab: 
  - Excel > Preferences > Ribbon & Toolbar 
  - On the Ribbon page scroll to the bottom of the Main Tabs list
  - Check the Developer box.
- Developer tab > Visual Basic 
- File > Import File > (Select PyInterface.bas) > Open
- Save and Exit the Visual Basic Editor
- All set. You can now use the API in your Excel cells via `=AI("~/Documents/MS-Office-AI/core", "Your chat prompt")`