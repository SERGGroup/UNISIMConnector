# %% ------- IMPORT MODULES                               ------------------------------------------------------------ #
from UNISIMConnect.constants import UNISIM_EXAMPLE_FOLDER, os
from UNISIMConnect import UNISIMConnector
import time


# %% ------- CALCULATE                                    ------------------------------------------------------------ #
file_path = os.path.join(UNISIM_EXAMPLE_FOLDER, "simple_example.usc")
with UNISIMConnector(file_path) as unisim:

    unisim.show()
    spreadsheet = unisim.get_spreadsheet("Calculation")
    result = list()

    for i in range(1, 10):

        time.sleep(2)
        spreadsheet.set_cell_value("B2", i)
        unisim.wait_solution()
        result.append(spreadsheet.get_cell_value("B4"))

        print(spreadsheet.get_cell_value("B4"))

    print(result)
