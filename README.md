##Disclaimer

This is a reconstruction of a project I developed for Tellworth Investment. The project involves parsing data from JSON files and populating them into an Excel file. The project also requires the use of sensitive information, which are not included in this repository for security reasons.

Originally the file handled path via Json structure, but then it was deployed in 4 costumised version installed in 4 different terminals.

The architecture of the project is designed to run on Windows only. It uses Python to automate the data parsing and Excel manipulation, and VBA to customize the behavior of Excel. In addition, the project supports scheduling via the Windows Task Scheduler.

Although some parts of the project are not included due to confidentiality reasons, this repository provides a detailed description of the project's architecture.

## Get Started with Python-Excel-VBA-Automation Project
The Python-Excel-VBA-Automation project is a collection of VBA macros and Python scripts that automate tasks related to financial data analysis and reporting.


## Installation
To install the required packages to run the VBA_CARD project, use the following command:

```
pip install -r requirements.txt
```
## Usage
Navigate to the src directory in the command line.

Run the main Python script using the following command:

```
python run_automated_process.py
```

This will run the automated process, which will:

Open the links.json file located in the data/jsons/ directory.
Iterate through the file paths and open the corresponding Excel workbooks.
Overwrite and Execute the VBA macros contained in the workbooks.
Close the workbooks.
