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