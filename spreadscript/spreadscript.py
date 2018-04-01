import os
import subprocess

import uno
import unohelper
from com.sun.star.connection import NoConnectException


class SpreadScript(object):
    def __init__(self, file_name=None):
        """Initialise the class.

        :arg str file_name: File name.
        """
        self._desktop = None
        self._start_soffice()
        self._connect_soffice()
        if file_name:
            self.open(file_name)

    def __del__(self):
        """Close the soffice instance."""
        if self._desktop:
            self.close()

    def _start_soffice(self):
        """Start soffice in the background."""
        process_id = os.fork()
        if not process_id:
            subprocess.call(
                'soffice --accept="socket,host=localhost,port=2002;urp;" ' + 
                '--norestore --nologo --nodefault --headless', shell=True)
            exit()

    def _connect_soffice(self):
        """Connect to a running soffice instance."""
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            'com.sun.star.bridge.UnoUrlResolver', local_context)
        while True:
            try:
                component_context = resolver.resolve(
                    'uno:socket,host=localhost,port=2002;urp;' +
                    'StarOffice.ComponentContext')
            except NoConnectException:
                pass
            else:
                break
        smgr = component_context.ServiceManager
        self._desktop = smgr.createInstanceWithContext(
            'com.sun.star.frame.Desktop', component_context)

    def _get_cell_text(self, column, row):
        return self._interface.getCellByPosition(column, row).getString()

    def _get_cell_value(self, column, row):
        return self._interface.getCellByPosition(column, row).getValue()

    def _set_cell_value(self, column, row, value):
        self._interface.getCellByPosition(column, row).setValue(value)

    def _read_table(self, column):
        """Read names and values from a table.

        :arg int column: Upper-left coordinate of the table content.

        :returns dict: Table content.
        """
        inputs = {}

        row = 3
        while True:
            name = self._get_cell_text(column, row)
            value = self._get_cell_value(column + 1, row)
            if not name:
                break
            inputs[name] = value
            row += 1

        return inputs

    def _write_table(self, column, data):
        """Write values to a table.

        :arg int column: Upper-left coordinate of the table content.
        :arg dict data: Data to be written.
        """
        row = 3
        while True:
            name = self._get_cell_text(column, row)
            if not name:
                break
            if name in data:
                self._set_cell_value(column + 1, row, data[name])
            row += 1

    def open(self, file_name):
        """Open a spreadsheet.

        :arg str file_name: File name.
        """
        doc_url = unohelper.systemPathToFileUrl(file_name)
        self._desktop.loadComponentFromURL(doc_url, '_blank', 0, ())
        current_component = self._desktop.getCurrentComponent()
        if 'Interface' not in current_component.Sheets:
            raise ValueError('no sheet named "Interface" found')
        self._interface = current_component.Sheets.Interface

    def close(self):
        """Close the soffice instance."""
        self._desktop.terminate()

    def read_input(self):
        return self._read_table(1)

    def write_input(self, data):
        self._write_table(1, data)

    def read_output(self):
        return self._read_table(4)
