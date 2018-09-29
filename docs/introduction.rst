Introduction
============

This project provides a command line interface and an API for spreadsheets.

The inputs and outputs are defined by two tables in a new sheet named
"Interface". SpreadScript will read the input variables from column
``B`` and the values from column ``C``. Likewise, the output variables
are read from column ``E`` and their values from column ``F``. In both
cases, the variables are read from row ``4`` onward until an empty cell
is encountered.

This method should work with any format that is supported by `LibreOffice
Calc`_. It has been tested using file formats ODS, XLS and XLSX.


.. _LibreOffice Calc: https://en.wikipedia.org/wiki/LibreOffice_Calc
