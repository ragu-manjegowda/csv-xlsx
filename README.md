# csv-xlsx
Python script to convert csv to xlsx

## Background
This is a small python code that I wrote to convert the all CSV's (Comma Separated Values) into sheets in a Spreadsheet.
It has cell formatting for first column and conditional formatting for second column.

## Usage

### To convert
All CSV files should be placed in a folder within a folder named `csvFiles`.

file structure would look something like this,

```
$ tree .
.
├── LICENSE
├── README.md
├── csvFiles
│   └── convertMyCSVTOXLSX
│       ├── set1.csv
│       ├── set2.csv
│       ├── set3.csv
│       ├── set4.csv
└── update.py
```

Then just run this command

```
$ python update.py
```

Final output will be Microsoft Excel sheet with the name named after folder name
> For this example, file name would be `convertMyCSVTOXLSX.xlsx`

# Dependencies

```
$ pip install openpyxl
$ pip install python-csv
```


