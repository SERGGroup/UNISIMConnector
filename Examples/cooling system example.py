# %% 1 - IMPORT MODULES
from UNISIMConnect.constants import UNISIM_EXAMPLE_FOLDER, os
from bayes_opt import BayesianOptimization
from UNISIMConnect import UNISIMConnector
import numpy as np


unisim_path = os.path.join(UNISIM_EXAMPLE_FOLDER, "cooling system example.usc")


# %% 2 - EVALUATE RESULTS
optimize_mean = False
weights = [1, 1, 1]

range_x = (0.01, 0.99)
range_y = (0.01, 0.99)
pbounds = {'hot': range_x, 'cold': range_y}

t_GCs = [35, 40, 45, 50]
t_eva_range = np.linspace(-10, 0, 5)
tot_results = np.zeros((len(t_GCs), len(t_eva_range), 17))
tot_results[:] = np.nan

with UNISIMConnector(unisim_path, close_on_completion=False) as unisim:

    sprsht = unisim.get_spreadsheet("OVR-IO")

    def safe_divide(value, dividend):
        try:
            return value / dividend
        except TypeError:
            return np.nan

    def find_GC_UA(n_it=15):

        x_ranges = [

            np.array([0.00001, 0.999999]),
            np.array([0.00001, 0.999999])

        ]

        x_means = [0.5, 0.5]

        cols = ["B", "C"]

        for k in range(n_it):

            for n, col in enumerate(cols):
                x_means[n] = np.mean(x_ranges[n])
                sprsht.set_cell_value("{}31".format(col), x_means[n])

            unisim.wait_solution(check_pop_ups=True)

            for n, col in enumerate(cols):

                error = sprsht.get_cell_value("{}33".format(col))
                dt_app = sprsht.get_cell_value("{}32".format(col))

                if dt_app > 0:

                    if error < 0:

                        x_ranges[n][1] = x_means[n]

                    else:

                        x_ranges[n][0] = x_means[n]

                else:

                    x_ranges[n][1] = x_means[n]

    def retrieve_results(n_res, check_flow_rates=False):

        max_res = 17
        result = np.zeros(max_res)
        result[:] = np.nan

        q_ihx = safe_divide(sprsht.get_cell_value("F15"), 3600)
        UA_dc = safe_divide(sprsht.get_cell_value("E7"), 1)

        m_dot_hot = safe_divide(sprsht.get_cell_value("B19"), 3600)
        m_dot_cold = safe_divide(sprsht.get_cell_value("C19"), 3600)
        m_dot_ihx = safe_divide(sprsht.get_cell_value("D19"), 3600)

        is_not_nan = (not np.isnan(q_ihx)) and (not np.isnan(m_dot_cold)) and (not np.isnan(UA_dc)) and (
            not np.isnan(m_dot_hot)) and (not np.isnan(m_dot_ihx))

        if is_not_nan and (UA_dc > 0) and (q_ihx > 0):

            if (not check_flow_rates) or (np.max([m_dot_cold, m_dot_hot, 2 * m_dot_ihx]) < 100):

                result[0] = m_dot_hot                       # m_pump_hot     [kg/s]
                result[1] = m_dot_cold                      # m_pump_cols    [kg/s]
                result[2] = m_dot_ihx                       # m_pump_ihx     [kg/s]

                result[3] = sprsht.get_cell_value("E6")     # UA_IHX         [kW/K]
                result[4] = sprsht.get_cell_value("E7")     # UA_DC          [kW/K]

                result[5] = sprsht.get_cell_value("B28")    # T DC out       [°C]
                result[6] = sprsht.get_cell_value("C12")    # T hot in       [°C]
                result[7] = sprsht.get_cell_value("F7")     # T hot out      [°C]
                result[8] = sprsht.get_cell_value("C14")    # T cold in      [°C]
                result[9] = sprsht.get_cell_value("F6")     # T cold out     [°C]
                result[10] = sprsht.get_cell_value("B27")   # T ihx out      [°C]

                result[11] = safe_divide(sprsht.get_cell_value("F11"), 3600)    # Q gas cooler  [kW]
                result[12] = safe_divide(sprsht.get_cell_value("F12"), 3600)    # Q air cooler  [kW]
                result[13] = safe_divide(sprsht.get_cell_value("F13"), 3600)    # W comp        [kW]
                result[14] = safe_divide(sprsht.get_cell_value("F14"), 3600)    # Q eva         [kW]
                result[15] = safe_divide(sprsht.get_cell_value("F15"), 3600)    # Q ihx         [kW]
                result[16] = safe_divide(sprsht.get_cell_value("F16"), 3600)    # Q dry cooler  [kW]

        return result[:min(n_res, max_res)]

    def evaluate_balance(t_percs):

        t_hot_perc, t_cold_perc = t_percs

        try:

            sprsht.set_cell_value("B31", 0.2)
            sprsht.set_cell_value("C31", 0.2)
            unisim.wait_solution(check_pop_ups=True)

            sprsht.set_cell_value("B12", t_hot_perc)
            sprsht.set_cell_value("B14", t_cold_perc)
            unisim.wait_solution(check_pop_ups=True)
            results = retrieve_results(15, check_flow_rates=False)

            if np.isnan(results[0]):
                return np.inf

            find_GC_UA()

            results = retrieve_results(15, check_flow_rates=False)

            if np.isnan(results[0]):
                return np.inf

        except:
            return np.inf

        sub_results = results[0:3] * weights

        curr_std = np.std(sub_results)

        if optimize_mean:

            curr_std += np.mean(sub_results)

        return curr_std

    def bayesian_balance(hot, cold):

        return 1 / evaluate_balance([hot, cold])


    for j, t_GC in enumerate(t_GCs):

        sprsht.set_cell_value("B5", t_GC)
        unisim.wait_solution(check_pop_ups=True)

        for i, t_eva in enumerate(t_eva_range):

            sprsht.set_cell_value("B3", t_eva)
            unisim.wait_solution(check_pop_ups=True)

            optimizer = BayesianOptimization(
                f=bayesian_balance,
                pbounds=pbounds,
                random_state=1,
            )

            optimizer.maximize(
                init_points=10,
                n_iter=25,
            )

            bayesian_balance(optimizer.max["params"]["hot"], optimizer.max["params"]["cold"])
            tot_results[j, i, :] = retrieve_results(17, check_flow_rates=True)
