import win32gui, win32con
import win32com.client
import pyautogui
import time


class UNISIMConnector:

    def __init__(self, filepath=None, close_on_completion=True):

        self.app = win32com.client.Dispatch('UnisimDesign.Application')
        self.__doc = None
        self.solver = None
        self.close_unisim = close_on_completion

        self.solution_time = 0.

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

    def reset_iterators(self, sub_element=None):

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

    @staticmethod
    def __get_window_titles():

        titles = []

        def enum_window_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title:  # Ignore windows without titles
                    titles.append(title)

        win32gui.EnumWindows(enum_window_callback, None)
        return titles

    @staticmethod
    def __find_button(hwnd_parent, text):
        """ Find a button with a specific text by hwnd_parent """
        def callback(hwnd, _):
            if win32gui.GetWindowText(hwnd) == text:
                buttons.append(hwnd)

        buttons = []
        win32gui.EnumChildWindows(hwnd_parent, callback, None)
        return buttons[0] if buttons else None

    @staticmethod
    def __click_button(hwnd_button):
        """ Send a click message to a button """
        # Store the original position of the mouse
        original_pos = pyautogui.position()

        # Identify the button position
        rect = win32gui.GetWindowRect(hwnd_button)
        x = (rect[0] + rect[2]) // 2
        y = (rect[1] + rect[3]) // 2

        # Move the mouse to the target position and click
        pyautogui.moveTo(x, y)
        pyautogui.click()

        # Wait a moment for the click action to be processed
        time.sleep(0.1)

        # Move the mouse back to the original position
        pyautogui.moveTo(original_pos)

    def __close_unisim_popups(self):

        titles = self.__get_window_titles()
        for title in titles:

            if "UniSim" in title:
                hwnd_main = win32gui.FindWindow(None, title)

                if hwnd_main:
                    hwnd_button = self.__find_button(hwnd_main, "OK")

                    if hwnd_button:
                        self.__click_button(hwnd_button)

    def __check_consistency_windows(self):

        titles = self.__get_window_titles()
        for title in titles:

            if "UniSim" in title:

                hwnd_main = win32gui.FindWindow(None, title)

                def callback(hwnd, _):
                    if win32gui.GetWindowText(hwnd) == 'Consistency Error':
                        consistency_windows.append(hwnd)

                consistency_windows = []
                win32gui.EnumChildWindows(hwnd_main, callback, None)

                if len(consistency_windows) > 0:

                    consistent_window = consistency_windows[0]
                    win32gui.PostMessage(consistent_window, win32con.WM_CLOSE, 0, 0)

                    return hwnd_main

                return None

    def __check_for_activate_button(self):

        titles = self.__get_window_titles()
        activate_button = []
        for title in titles:

            if "Document Already Loaded" in title:
                hwnd_main = win32gui.FindWindow(None, title)

                def callback(hwnd, _):
                    text = win32gui.GetWindowText(hwnd)

                    if "Activate" in text:
                        activate_button.append(hwnd)

                win32gui.EnumChildWindows(hwnd_main, callback, None)

        if len(activate_button) > 0:

            self.__click_button(activate_button[0])

    def __restart_solver(self, main_window):

        self.__check_for_activate_button()

        # Place the UNISIM window in foreground
        win32gui.SetForegroundWindow(main_window)
        time.sleep(0.1)

        # Press "F8" to restart the solver
        pyautogui.press('f8')
        time.sleep(0.1)

    def wait_solution(self, timeout: float = None, check_pop_ups: float = None, check_consistency_error: int = 1):

        """

            Waits until solution is available. timeout can be given in seconds.
            Check_pop_ups = t_pop, activates an autoclicker that automatically closes Unisim warnings, t_pop is the
            time in seconds that the software waits before checking if there's some windows to close (the operation
            slows down the calculation).
            If Check_consistency = N, the code will try to close the UNISIM consistency warning and
            restart the calculations N times.

        """

        time_start = time.time()
        time_start_pop = time_start
        timed_out = False

        try:

            while True:

                if self.solver.IsSolving:

                    if timeout is not None:
                        if (time.time() - time_start) > timeout:
                            timed_out = True
                            break

                    if check_pop_ups is not None:
                        if (time.time() - time_start_pop) > check_pop_ups:
                            time_start_pop = time.time()
                            self.__close_unisim_popups()

                elif check_consistency_error > 0:

                    main_window = self.__check_consistency_windows()
                    consistency_found = main_window is not None

                    if consistency_found:
                        check_consistency_error -= 1
                        self.__restart_solver(main_window)

                    else:
                        break

                else:
                    break

            self.solution_time = time.time() - time_start

        except:

            pass

        return timed_out

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