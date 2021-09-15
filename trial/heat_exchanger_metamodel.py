from trial.zekeryia_project import DataElement, UNISIMConnector
from xml_generation import create_rstudio_xml
from tkinter import filedialog
from pyDOE2 import *
import tkinter as tk
import os

cells = {

    "points": {

        "col": {

            "P": "E", "T": "F", "m": "G",
            "h": "H", "s": "I", "Q": "J", "rho": "K"

        },
        "row": {

            "brine_in": "5", "brine_out": "6",
            "ORC_in": "7", "ORC_out": "8"

        }

    },
    "other": {

        "col": {

         "values": "B"

        },
        "row": {

         "Q_HE": "3",
         "eta": "4",
         "UA": "5",
         "m_ratio": "10",
         "DP_brine": "12",
         "DP_ORC": "13"

        }

    }

}


def get_cell_name(input_dict: dict):

    type = input_dict["type"]
    col = input_dict["value"]
    row = input_dict["point"]

    return cells[type]["col"][col] + cells[type]["row"][row]


def convert_excel_to_xml_zekeryia():

    root = tk.Tk()
    root.withdraw()
    unisim_path = filedialog.askopenfilename()

    input_dict = {

        "names": ["T_brine", "m_brine", "P_ORC", "T_ORC", "m_ratio"],
        "units": ["[C]", "[kg/s]", "[kPa]", "[-]", "[-]"],
        "min": [60, 100, 800, 20, 0.4],
        "max": [140, 500, 1500, 55, 1.4],
        "values": list()

    }
    output_dict = {

        "names": ["Q_HE", "eta", "UA", "DP_brine", "DP_ORC"],
        "units": ["[kJ/h]", "[-]", "[kJ/(h C)]", "[kPa]", "[kPa]"],
        "values": list()

    }

    with UNISIMConnector(filepath=unisim_path) as unisim:

        unisim.show()
        result_dict = return_values(unisim, input_dict, output_dict, n_runs=5000)

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

    data_element = DataElement(unisim, input_list, list(), input_cells=cells)

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

        if not data_element.result_dict["points"]["brine_out"]["T"] == 0:

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

        if param in cells["other"]["row"].keys():

            input_dict = {"type": "other",
                          "point": param,
                          "value": "values"}

        elif "vap_indicator" in param:

            if "ORC" in param:

                input_dict = {"type": "other",
                              "point": "vap_ind_ORC",
                              "value": "values"}

            else:

                input_dict = {"type": "other",
                              "point": "vap_ind_brine",
                              "value": "values"}

        else:

            input_dict = {"type": "points"}

            for value in cells["points"]["col"].keys():

                if value + "_" in param:

                    input_dict.update({"value": value})
                    point_param = param.strip(value + "_")
                    break

            if point_param in cells["points"]["row"].keys():

                input_dict.update({"point": point_param})

            elif "ORC" in point_param:

                input_dict.update({"point": "ORC_in"})

            else:

                input_dict.update({"point": "brine_in"})

        if is_input:

            new_list.append({"dict": input_dict,
                             "value": 0})

        else:

            new_list.append(input_dict)

    return new_list


if __name__ == "__main__":

    a = convert_excel_to_xml_zekeryia()
    print("process completed")
