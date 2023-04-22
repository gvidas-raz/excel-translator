# Excel translator
This python script takes any excel file and translates text from a given column
to another column. It uses the DeepL API to do all translations. Currently is
set to translate from Spanish to English.

Given the excel file it backs up the original and then proceeds with the
translation.

It takes the text from source cell and goes down the column until it hits an
empty cell.

It writes the translation starting with the given destination cell and moves down
the column in sync the source column.

## Prerequisites

The script requares these python modules to be installed:
`openpyxl`
`requests`
`absl-py`
`dotenv`

All of these can be installed with:
```
pip install <module_name>
```

It also requires an `.env` file which contains an `API_KEY` which is the API
key for your DeepL account.

## Usage
```
python ./translate --file <excel_filename> --source <source_cell> --dest <dest_cell>
```

`--file or -f` - the file path to the excel file
`--source or -s` - the source text cell to start going down the column from e.g.: 'A1'
`--dest or -s` - the destination text cell to start going down the column from
 and write the translation to e.g.: 'B1'
