# BadgerWifiTools

### What is it?
BadgerWiFiTools is a wxPython GUI application built by Nick Turner for Wi-Fi Engineers.

### What does it do?
The tool contains several automation scripts for performing common Ekahau Wi-Fi reporting tasks.
* Summarise Simulated Design Project details
* Rename APs within Ekahau
* Export APs to XLSX
* Export custom asset maps from Ekahau
* Perform image insertion into DOCX files

### How do I install and use it?
The tool is not compiled, this means you can look inside the code and see how it works.

To install the application, you will need to have Python 3.9 or later installed.
To check what you currently have installed, if any:


## Windows users
Open command prompt (CMD)
`python --version`
if you don't have Python installed, this prompt will offer to take you to the Windows App Store to download and install Python.


## macOS users
Open terminal
`python3 --version`
if you don't have Python installed you can download Python from the official website: https://www.python.org/downloads/ or use Homebrew to install it.


## Regardless of your operating system
You will also need Git installed.
Try the command
`git` from your terminal window.
If you don't have Git installed, you can download it from the official website: https://git-scm.com/downloads
On Windows all the default options are fine.


## Installation
Clone the repository to your local machine using the following command:
Decide the folder where you want to clone the badgerwifi-tools repository to, then navigate to that folder in your terminal window and run the following command:
`git clone https://github.com/nickjvturner/badgerwifi-tools`

Navigate to badgerwifi-tools directory:
`cd badgerwifi-tools`

Install all the missing dependencies with the following command:
`pip install .`

## Usage
You can now run the application from your terminal window by executing main.py:
macOS: `python3 main.py`
Windows: `python main.py`


## How do I use the application?
This should be somewhat self-explanatory, if not please reach out to me and tell me what you are not finding intuitive. At some point I will create a user guide.

## What are the project profiles and where are they stored?
The project profiles are `.py` files that define project specific conventions, requirements stored in the `profiles` directory. Each project profile is a JSON file that contains the project details and the settings for each of the automation scripts.