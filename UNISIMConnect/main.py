from unisim_connect import UNISIMConnector
from tkinter import filedialog


if __name__ == '__main__':

    file_path = filedialog.askopenfilename()

    if file_path != "":

        with UNISIMConnector(file_path) as unisim:

            spreadsheet = unisim.get_spreadsheet("Calculation")
            result = list()

            for i in range(1,10):

                spreadsheet.set_cell_value("F30", i)
                unisim.wait_solution()
                result.append(spreadsheet.get_cell_value("E30"))

            print(result)