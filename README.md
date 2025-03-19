# SRUM Dump

## Overview
SRUM Dump is a Python application designed to analyze the System Resource Usage Monitoring (SRUM) database on Windows systems. It extracts data from the SRUM database and presents it in an easily readable Excel format (XLSX). The application features a user-friendly interface and supports different database interaction methods.

## Features
- All new GUI featuring a walk through wizard
- TEMPLATE from older version is now a .json text file
- Editible configuration file supports dirty words, friendly table names and column names, User name resolution, WIFI resolution, dynamic column additions and more 
- Analyzes any SRUDB.dat database.
- Suppliments it with data from SOFTWARE registry hive. That data is placed in the config for editing.
- Attempts live acquisition using VSS then attempts to repairs the extracted SRUDB.dat file 
- Generates an XLSX file or CSV with the extracted data.
- Supports interchangeable database interaction methods pyesedb or dissect

# Known issues
- Extracting a SRUM from a live Windows 11 box is problematic. Works on older Versions.
- On WIndows 11 I suggest dumping the file with some other tool.

## Project Structure
```

├── srum-dump
│   ├── srum_dump.py        # Main entry point of the application
│   ├── ui_simple.py        # Simple GUI implementation
│   ├── ui_tk.py            # Tkinter GUI implementation
│   ├── db_ese.py           # ESE database interface
│   ├── db_dissect.py       # Dissect database interface
│   └── helpers.py          # Utility functions and default configs
├── requirements.txt         # Project dependencies
└── README.md                # Project documentation
```

## Installation
To set up the project, clone the repository and install the required dependencies. You can do this using pip:

```bash
pip install -r requirements.txt
```

If you use Python3.12 then you can grab pyesedb from here:
https://github.com/log2timeline/l2tbinaries 

Otherwise you must compile it.

## Dependencies
The project requires the following Python libraries:
- simplegui
- pyesedb
- openpyxl
- dissect
- tkinter (included with Python standard library)

## Contributing
Contributions are welcome! Please feel free to submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.