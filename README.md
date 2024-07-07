
# SR ID Merger Script

## Description

This Python script is designed to merge new SR IDs from a daily received Excel file into an existing "archive" file. It checks multiple specified sheets for new SR IDs and appends only the new entries to the existing file, ensuring that any notes or existing data in the original file are preserved.

## Features

- Merges new SR IDs from a new Excel file into an existing archive file.
- Checks multiple specified sheets for new SR IDs.
- Appends only new entries, preserving existing data and notes.
- User-friendly file selection dialog using `tkinter`.

## Requirements

- Python 3.x
- pandas
- openpyxl
- tkinter

## Installation

1. Clone the repository or download the script.
2. Install the required Python packages:

```bash
pip install pandas openpyxl
```

## Usage

1. Run the script:

```bash
python sr_id_merger.py
```

2. A file selection dialog will appear. Select the existing final_results file.
3. Another file selection dialog will appear. Select the new final_results file.
4. The script will process the files and append new SR IDs to the specified sheets.

## Configuration

- Modify the `sheets_to_check` list to specify the sheets you want to check for new SR IDs:

```python
sheets_to_check = ["Aggregated Data", "Summary of Actions", "Last Drop"]
```

## Author

Panos Bonotis

- [LinkedIn](https://www.linkedin.com/in/panagiotis-bonotis-351a7996/)

## License

This project is licensed under the MIT License.
