import xml.etree.ElementTree as ETree
import matplotlib.pyplot as plt
from sklearn.svm import SVC
import numpy, os


def create_rstudio_xml(folder_path, input_dict, output_dict):

    data = generate_xml(input_dict, output_dict)
    file_path = os.path.join(folder_path,"data.xml")

    xml_file = open(file_path, "wb")
    xml_file.write(ETree.tostring(data))
    xml_file.close()


def generate_xml(input_dict, output_dict) -> ETree.Element:

    root = ETree.Element("data")
    root.set("n_exp", str(len(input_dict["values"][0])))

    inputs = ETree.SubElement(root, "inputs")
    inputs.set("n_input", str(len(input_dict["values"])))

    outputs = ETree.SubElement(root, "outputs")
    outputs.set("n_output", str(len(output_dict["values"])))

    for dict in [input_dict, output_dict]:

        if dict == input_dict:

            parent = inputs
            name = "input"

        else:

            parent = outputs
            name = "output"

        for i in range(len(dict["values"])):

            sub_element = ETree.SubElement(parent, name)

            sub_element.set("index", str(i + 1))

            sub_element.set("name", dict["names"][i])
            sub_element.set("measure_unit", dict["units"][i])

            sub_element.set("max", str(numpy.max(dict["values"][i])))
            sub_element.set("min", str(numpy.min(dict["values"][i])))
            sub_element.set("mean", str(numpy.nanmean(dict["values"][i])))

            value_str = str(dict["values"][i][0])
            for value in dict["values"][i][1:]:

                value_str += ";" + str(value)

            sub_element.set("values", value_str)

    return root


def generate_SVM(return_dict, x_index, y_index):

    input_values_list = return_dict["input"]["values"]
    output_values_list = return_dict["output"]["values"]

    input_names_list = return_dict["input"]["names"]
    input_units_list = return_dict["input"]["units"]

    fig1, ax = plt.subplots()

    x = numpy.array(input_values_list).transpose()

    model = SVC(kernel='linear', C=1E10)
    model.fit(x, output_values_list[-1])

    ax.scatter(input_values_list[x_index], input_values_list[y_index], c=output_values_list[-1])
    ax.set_xlabel(input_names_list[x_index] + " " + input_units_list[x_index])
    ax.set_ylabel(input_names_list[y_index] + " " + input_units_list[y_index])

    counter_dict = dict()

    for value in output_values_list[-1]:

        value = '{:0>4}'.format(int(value))

        if value in counter_dict.keys():

            counter_dict[value] += 1

        else:

            counter_dict[value] = 1

    print(counter_dict)

    plt.show()