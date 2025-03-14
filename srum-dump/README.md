# SRUM Dump

## Overview
SRUM Dump is a Python application designed to analyze the System Resource Usage Monitoring (SRUM) database on Windows systems. It extracts data from the SRUM database and presents it in an easily readable Excel format (XLSX). The application features a user-friendly interface and supports different database interaction methods.

## Features
- Extracts data from the SRUM database.
- Generates an XLSX file with the extracted data.
- Provides a choice between different user interface implementations (simple GUI or tkinter).
- Supports interchangeable database interaction methods.

## Project Structure
```
srum-dump
├── src
│   ├── srum_dump.py        # Main entry point of the application
│   ├── ui_simple.py        # Simple GUI implementation
│   ├── ui_tk.py            # Tkinter GUI implementation
│   ├── db_ese.py           # ESE database interface
│   ├── db_dissect.py       # Dissect database interface
│   └── utils
│       └── helpers.py      # Utility functions
├── requirements.txt         # Project dependencies
└── README.md                # Project documentation
```

## Installation
To set up the project, clone the repository and install the required dependencies. You can do this using pip:

```bash
pip install -r requirements.txt
```

## Usage
To run the application, execute the main script:

```bash
python src/srum_dump.py
```

Follow the on-screen instructions to select the SRUM database file and specify the output XLSX file location.

## Dependencies
The project requires the following Python libraries:
- simplegui
- pyesedb
- openpyxl
- tkinter (included with Python standard library)

## Contributing
Contributions are welcome! Please feel free to submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.