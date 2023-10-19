from calculations.zekeryia_project import DataElement, UNISIMConnector
from xml_generation import create_rstudio_xml
from tkinter import filedialog
from pyDOE2 import *
import tkinter as tk
import os

cells = {

    "other": {

        "col": {

            "values": "F"

        },
        "row": {

            "m_ratio": "4",
            "m_ORC_1": "5",
            "m_ORC_2": "6",

            "W_net": "17",
            "eta": "18",
            "error_check": "20"

        }

    }

}


class ModDataElement(DataElement):

    def __init__(self, unisim: UNISIMConnector, set_values: list, check_values: list, input_cells=None):

        super().__init__(unisim, set_values, check_values, input_cells=input_cells, spreadsheet_name="INPUT-OUTPUT")


def get_cell_name(input_dict: dict):

    type = input_dict["type"]
    col = input_dict["value"]
    row = input_dict["point"]

    return cells[type]["col"][col] + cells[type]["row"][row]


def convert_excel_to_xml_zekeryia():

    root = tk.Tk()
    root.withdraw()
    unisim_path = filedialog.askopenfilename()

    input_dict = {"names": ["m_ratio", "m_ORC_1", "m_ORC_2"],
                  "units": ["[-]", "[-]", "[-]"],
                  "min": [0.3, 0.25, 0.25],
                  "max": [0.6, 0.45, 0.45],
                  "values": list()}

    output_dict = {"names": ["W_net", "eta", "error_check"],
                   "units": ["[kW]", "[-]", "[-]"],
                    "values": list()}

    with UNISIMConnector(filepath=unisim_path) as unisim:

        unisim.show()
        result_dict = return_values(unisim, input_dict, output_dict, n_runs=1000)

    create_rstudio_xml(os.path.dirname(unisim_path), result_dict["input"], result_dict["output"])
    return result_dict


def return_values(unisim: UNISIMConnector, input_dict: dict, output_dict: dict, n_runs=10):

    n_param = len(input_dict["names"])
    LHS_design = lhs(n_param, n_runs)

    input_list = generate_list(input_dict["names"], is_input=True)
    output_list = generate_list(output_dict["names"])

    for curr_dict in [input_dict, output_dict]:

        for i in range(len(curr_dict["names"])):

            curr_dict["values"].append(list())

    data_element = ModDataElement(unisim, input_list, list(), input_cells=cells)

    for i in range(n_runs):

        input_array = LHS_design[i]
        mod_input_array = list()

        for j in range(n_param):

            min_value = input_dict["min"][j]
            max_value = input_dict["max"][j]

            mod_input_array.append((max_value - min_value)*input_array[j] + min_value)
            input_list[j]["value"] = mod_input_array[-1]

        data_element.set_values = input_list
        data_element.calculate()
        result = data_element.get_values(output_list)

        result_correct = data_element.result_dict["other"]["error_check"]["values"] == 3
        iteration_counter = 0

        while not result_correct and iteration_counter < 10:

            unisim.reset_iterators()
            data_element.calculate()

            result = data_element.get_values(output_list)

            result_correct = data_element.result_dict["other"]["error_check"]["values"] == 3
            iteration_counter += 1

        if result_correct:

            for curr_dict in [input_dict, output_dict]:

                if curr_dict == input_dict:
                    append_list = mod_input_array

                else:
                    append_list = result

                for j in range(len(append_list)):

                    curr_dict["values"][j].append(append_list[j])

        print("point: " + str(i + 1) + " (" + str(len(input_dict["values"][0])) + " correct)")

    return {"input": input_dict, "output": output_dict}


def generate_list(param_names, is_input=False):

    new_list = list()

    for param in param_names:

        input_dict = {"type": "other",
                      "point": param,
                      "value": "values"}

        if is_input:

            new_list.append({"dict": input_dict,
                             "value": 0})

        else:

            new_list.append(input_dict)

    return new_list


if __name__ == "__main__":

    a = convert_excel_to_xml_zekeryia()
    print("process completed")
