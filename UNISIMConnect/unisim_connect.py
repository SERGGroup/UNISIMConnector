import win32com.client


class UNISIMConnector:

    def __init__(self, filepath=None, close_on_completion=True):

        self.app = win32com.client.Dispatch('UnisimDesign.Application')
        self.__doc = None
        self.solver = None
        self.close_unisim = close_on_completion

        self.open(filepath)

    def __enter__(self):

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):

        if self.close_unisim:

            try:

                self.doc.Visible = False
                self.__doc.Close()

            except:

                pass

            self.app.Visible = False
            self.app.Quit()

    def get_spreadsheet(self, spreadsheet_name):

        try:

            spreadsheet = UNISIMSpreadsheet(self, spreadsheet_name)

        except:

            return None

        return spreadsheet

    def reset_iterators(self, sub_element = None):

        if self.doc is not None:

            if sub_element is not None:

                flowsheet = sub_element

            else:

                flowsheet = self.doc.Flowsheet

            for operation_name in flowsheet.Operations.Names:

                operation = self.doc.Flowsheet.Operations.Item(operation_name)

                if operation.TypeName == "adjust":

                    operation.SolveAdjust()

                elif operation.TypeName == "Standard Sub-Flowsheet":

                    self.reset_iterators(sub_element=operation.OwnedFlowsheet)

    def show(self):

        try:

            self.app.Visible = True
            self.doc.Visible = True

        except:

            pass

    def open(self, filepath):

        if self.is_ready:

            self.__doc.Close()

        if filepath is None:

            self.doc = self.app.SimulationCases.Add()

        else:

            self.doc = self.app.SimulationCases.Open(filepath)

    @property
    def doc(self):

        return self.__doc

    @doc.setter
    def doc(self, doc):

        try:

            self.solver = doc.Solver

        except:

            self.__doc = None
            self.solver = None

        else:

            self.__doc = doc

    @property
    def is_ready(self):

        return self.__doc is not None

    def wait_solution(self):

        try:

            while self.solver.IsSolving:

                pass

        except:

            pass


class UNISIMSpreadsheet:

    def __init__(self, parent:UNISIMConnector, spreadsheet_name):

        self.parent = parent
        self.spreadsheet = parent.doc.Flowsheet.Operations.Item(spreadsheet_name)

    def set_cell_from_list(self, input_list: list):

        for value_dict in input_list:

            self.set_cell_value(value_dict["cell name"], value_dict["value"])

    def get_value_from_list(self, input_list: list):

        return_dict = dict()

        for cell_name in input_list:

            return_dict.update({cell_name: self.get_cell_value(cell_name)})

    def set_cell_value(self, cell_name, value):

        try:

            cell = self.spreadsheet.Cell(cell_name)
            cell.CellValue = value

        except:

            pass

    def get_cell_value(self, cell_name):

        try:

            cell = self.spreadsheet.Cell(cell_name)
            value = cell.CellValue

        except:

            return None

        return value