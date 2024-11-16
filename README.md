# Create calender file (ics) from excel file:
This Python script reads event data from an Excel file and generates an .ics calendar file that can be imported into calendar applications like Google Calendar, Outlook, and Apple Calendar. It uses `openpyxl` to handle Excel files.

&nbsp;
&nbsp;

## Requirements
- `openpyxl` (for reading Excel files)
- `ics` (for generating `.ics` files)

&nbsp;
&nbsp;


## Usage
Run the script:
```bash
python main.py
```

The script will generate an events.ics file in the same directory.

&nbsp;
&nbsp;

## Installation
Install the required packages by running:

```bash
pip install -r requirements.txt
```

&nbsp;
&nbsp;

## Sample Excel file

Title	Start | Date |	Start Time |	End Date |	End Time |	Description |	Location |	Repeat Frequency |	Repeat Until Date
--- | --- | --- | --- | --- | --- | --- | --- | ---
Meeting |	12/1/2024 |	10:00 |	12/1/2024 |	11:00 |	Project meeting |	Office Room 101	| None |
Workshop |	12/2/2024 |	14:00 |	12/2/2024 |	16:00 |	Data workshop |	Conference Hall	| None | 

&nbsp;
&nbsp;
