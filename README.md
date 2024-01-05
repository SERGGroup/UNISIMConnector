# UNISIM connector

__UNISIM connector__ is a tools developed by the [SERG research group](https://www.dief.unifi.it/vp-177-serg-group-english-version.html) 
of the [University of Florence](https://www.unifi.it/changelang-eng.html) for modifying UNISIM files from python.

The beta version can be downloaded using __PIP__:

```
pip install UNISIM_connector
```
Once the installation has been completed the user can import the tool and initialize the connector itself passing 
the path of the UNISIM file that has to be opened (in the example below the file is selected by the user trough a 
dialog box. If possible use a with statement for the initialization.

```python
from UNISIMConnect import UNISIMConnector
from tkinter import filedialog
import tkinter as tk

root = tk.Tk()
root.withdraw()
unisim_path = filedialog.askopenfilename()

with UNISIMConnector(unisim_path) as unisim:

    # insert your code here

```
Finally, you can ask the program to modify values in the spreadsheets inside the UNISIM file and wait until a solution 
has been reached
```python
with UNISIMConnector(unisim_path) as unisim:
  
    spreadsheet = unisim.get_spreadsheet("CALCULATION")
    
    spreadsheet.set_cell_value("A5", 15)
    unisim.wait_solution()
    spreadsheet.get_cell_value("A6")

```

If you need to keep UNISIM open once the calculation has been completed you can set the option "close_on_completion=False".
```python
with UNISIMConnector(unisim_path, close_on_completion=False) as unisim:
  
    spreadsheet = unisim.get_spreadsheet("CALCULATION")
    
    spreadsheet.set_cell_value("A5", 15)
    unisim.wait_solution()
    spreadsheet.get_cell_value("A6")

```
__-------------------------- !!! THIS IS A BETA VERSION !!! --------------------------__ 

please report any bug or problems in the installation to _pietro.ungar@unifi.it_<br/>
for further information visit: https://tinyurl.com/SERG-3ETool
