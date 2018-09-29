Library
=======

First import the ``SpreadScript`` class and load a spreadsheet.

.. code:: python

    >>> from spreadscript import SpreadScript
    >>> 
    >>> spreadsheet = SpreadScript('data/test.ods')

The input and output variables can be read with the ``read_input`` and
``read_output`` methods respectively.

.. code:: python

    >>> spreadsheet.read_input()
    {'height_johnny': 1.52, 'height_janie': 1.41}
    >>> 
    >>> spreadsheet.read_output()
    {'average': 1.5775, 'tallest': 1.76}

The ``write_input`` method is used to update any variables.

.. code:: python

    >>> spreadsheet.write_input({'height_johnny': 1.56})
    >>> spreadsheet.read_output()
    {'average': 1.5875, 'tallest': 1.76}
