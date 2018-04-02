# SpreadScript: Use a spreadsheet as a function
This project provides a way to use spreadsheets from the command line or from
Python programs. In this way, spreadsheets can be used in automated data
analysis processes.

The inputs and outputs are defined by two tables in a new sheet named
"Interface". SpreadScript will read the input variables from column `B` and the
values from column `C`. Likewise, the output variables are read from column `E`
and their values from column `F`. In both cases, the variables are read from
row `4` onward until an empty cell is encountered.

This method should work with any format that is supported by
[LibreOffice Calc](https://en.wikipedia.org/wiki/LibreOffice_Calc). It has been
tested using file formats ODS, XLS and XLSX.


## Installation
Prerequisites:

    apt install python3-uno

Via [PyPI](https://pypi.python.org/pypi/spreadscript):

    pip3 install spreadscript

From source:

    git clone https://github.com/jfjlaros/spreadscript.git
    cd spreadscript
    pip3 install .


## Usage
Suppose we have the following table with heights of family members.

![Example table.](data/example_table.png)

Since Janie and Johnny have not reached full height, we might want to export
their heights as input variables. Suppose we are interested in the average
height and the tallest person in this family. This information goes to the
"Interface" sheet.

![Example interface.](data/example_interface.png)

In this sheet we put the input variables in column `B` and references to the
values in column `C`. The value of `C4` is `=$Sheet1.C5` and that of `C5` is
`=$Sheet1.C6`.

Likewise, the output variables are put in column `E` and references to the
values in column `F`. The value of `F4` is `=$Sheet1.C7` and that of `F5` is
`=$Sheet1.C9`.

### Command line interface
With the command line interface, the input and output table can be read.

    $ spreadscript read_input data/test.ods
    {"height_janie": 1.41, "height_johnny": 1.52}

    $ spreadscript read_output data/test.ods
    {"tallest": 1.76, "average": 1.5775}

To manipulate the input, use the `process` subcommand:

    $ spreadscript process data/test.ods  '{"height_johnny": 1.56}'
    {"tallest": 1.76, "average": 1.5875}


### Library
First import the `SpreadScript` class and load a spreadsheet.

```python
>>> from spreadscript import SpreadScript
>>> 
>>> spreadsheet = SpreadScript('data/test.ods')
```

The input and output variables can be read with the `read_input` and
`read_output` methods respectively.

```python
>>> spreadsheet.read_input()
{'height_johnny': 1.52, 'height_janie': 1.41}
>>> 
>>> spreadsheet.read_output()
{'average': 1.5775, 'tallest': 1.76}
```

The `write_input` method is used to update any variables. 

```python
>>> spreadsheet.write_input({'height_johnny': 1.56})
>>> spreadsheet.read_output()
{'average': 1.5875, 'tallest': 1.76}
```
