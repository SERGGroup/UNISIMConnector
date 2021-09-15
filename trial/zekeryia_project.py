from UNISIMConnect import UNISIMConnector
from scipy.stats import norm

cells = {

    "points": {

        "col": {

            "P": "K", "T": "L", "h": "M", "s": "N", "Q": "O", "rho": "P"

        },
        "row": {

            # brine points
            "1": "4", "2": "5", "3": "6",
            "4": "7", "5": "8", "6": "9",
            "7": "10", "8": "11", "9": "12",

            # ORC-1 points
            "10": "14", "10.1": "15", "11": "16",
            "12": "17", "13": "18", "14": "19",
            "15": "20", "12.sat": "29", "11.iso": "32",

            # ORC-2 points
            "16": "22", "16.1": "23", "17": "24",
            "18": "25", "19": "26", "20": "27",
            "18.sat": "30", "17.iso": "33",

            "amb": "35"

        }

    },

    "turbines": {

        "col": {

            "Power": "E", "P ratio": "F", "eta": "G", "phi": "H"

        },
        "row": {

            "TURB-1": "4",
            "TURB-2": "5",
            "PUMP-1": "6",
            "PUMP-2": "7"

        }

    },

    "heat exchangers": {

        "col": {

            "Power": "E", "UA": "F"

        },
        "row": {

            "VAP-1": "13",
            "VAP-2": "16",
            "TPH-1": "14",
            "PH-1": "15",
            "PH-2": "17"

        }

    },

    "stodola curves": {

        "col": {

            "Yid": "E", "n stages": "F"

        },
        "row": {

            "TURB-1": "30",
            "TURB-2": "31"

        }

    },

    "heat losses": {

        "col": {

            "Heat Exchanged": "E", "Relative Heat": "F"

        },
        "row": {

            "ORC-1": "23", "ORC-2": "24"

        }

    },

    "other": {

        "col": {

            "values": "B"

        },
        "row": {

            "m brine": "7",
            "m ORC-1": "8",
            "m ORC-2": "9",
            "m ratio": "11",

            "W gross": "16",
            "W net": "17",
            "Q in": "18",
            "eta": "19"

        }

    }

}


def evaluate_probability(x, dx, mean, sd):
    distribution = norm(loc=mean, scale=sd)
    max_probability = abs(distribution.cdf(mean + dx) - distribution.cdf(mean - dx))
    return abs(distribution.cdf(x + dx) - distribution.cdf(x - dx)) / max_probability


class DataElement:

    def __init__(self, unisim: UNISIMConnector, set_values: list, check_values: list, input_cells=None, spreadsheet_name="Calculation"):

        self.unisim = unisim
        self.sheet = unisim.get_spreadsheet(spreadsheet_name)
        self.set_values = set_values
        self.check_values = check_values

        self.likelihood = 1
        self.result_dict = dict()

        if input_cells is None:
            input_cells = cells

        self.cells = input_cells

    def calculate(self):

        for set_value in self.set_values:
            self.__set_value(set_value["dict"], set_value["value"])

        self.initialize_iteration()
        self.__collect_results()

    def get_values(self, input_list: list):

        return_list = list()

        for element in input_list:
            return_list.append(self.__get_value(element))

        return return_list

    def initialize_iteration(self):

        # CALLED IN "CALCULATE" BEFORE SOLUTION WAIT
        # TO BE OVEWRITTEN IF NEEDED

        pass

    def __collect_results(self):

        self.unisim.wait_solution()
        self.result_dict = dict()

        for type in self.cells.keys():

            sub_result = dict()
            sub_dict = self.cells[type]

            for row in sub_dict["row"].keys():

                sub_result.update({row: dict()})

                for col in sub_dict["col"].keys():

                    cell_name = sub_dict["col"][col] + sub_dict["row"][row]
                    result = self.sheet.get_cell_value(cell_name)

                    try:

                        result = float(result)

                    except:

                        result = 0.

                    sub_result[row].update({col: result})

            self.result_dict.update({type: sub_result})

    def __calculate_likelihood(self, ):

        self.likelihood = 1
        # TODO

    def __set_value(self, input_dict: dict, value):

        cell_name = self.get_cell_name(input_dict)
        self.sheet.set_cell_value(cell_name, value)

    def __get_value(self, input_dict: dict):

        type = input_dict["type"]
        col = input_dict["value"]
        row = input_dict["point"]

        return self.result_dict[type][row][col]

    def get_cell_name(self, input_dict: dict):

        type = input_dict["type"]
        col = input_dict["value"]
        row = input_dict["point"]

        return self.cells[type]["col"][col] + self.cells[type]["row"][row]


if __name__ == "__main__":

    with UNISIMConnector(filepath="C:\\Users\\Utente\\Documents\\Università\\Dottorato\\Progetto Zekeriya\\Plant V3.1 "
                                  "(Baesian Stodola).usc") as unisim:

        for i in range(1, 10):

            data_element = DataElement(unisim, [{"dict": {"type": "stodola curves",
                                                          "value": "n stages",
                                                          "point": "TURB-1"},
                                                 "value": i},

                                                {"dict": {"type": "stodola curves",
                                                          "value": "n stages",
                                                          "point": "TURB-2"},
                                                 "value": i}], list())

            data_element.calculate()

            print(data_element.get_values([{"type": "stodola curves",
                                            "value": "Yid",
                                            "point": "TURB-1"},

                                           {"type": "stodola curves",
                                            "value": "Yid",
                                            "point": "TURB-2"}]))
