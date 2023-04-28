## Get Started with Python-Excel-VBA-Automation Project
The Python-Excel-VBA-Automation project is a collection of VBA macros and Python scripts that automate tasks related to financial data analysis and reporting.

## Project Structure
The VBA_CARD project is structured as follows:

```
VBA_CARD/
├── archive/
├── data/
│   ├── jsons/
│   │   └── links.json
│   ├── macros/
│   └── test_2/
│       └── DB/
├── src/
│   └── run_automated_process.py
├── test.py
├── README.md
├── requirements.txt
└── run_test.py
```

**The archive directory** contains files that are no longer needed for the project but are being kept for reference.
**The data directory** contains data files used by the VBA project.
**The jsons subdirectory** contains a JSON file named links.json that contains links to external data sources.
**The macros subdirectory** contains VBA macro files used in the project.
**The test_2 subdirectory** contains Excel spreadsheets.
**The src directory** contains the main engine, a file named run_automated_process.py that runs an automated process.
**The README.md file** contains an overview of the VBA_CARD repository and its contents.
**The requirements.txt file** lists the required packages and their versions needed to run the VBA project.
**The test.py file** runs a test suite for the VBA project.
**The run_test.py file** runs a script to test the Python code.

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
