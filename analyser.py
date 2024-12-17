import pandas as pd
import math
import numpy as np
import matplotlib.pyplot as plt
from os import path, mkdir, read

class Analyzer ():
    def __init__(self, filename: str):
        self.filename = filename
        self.year_data = {}
        self.all_tmp_observations = None
        self.load_df(filename)
        self.bootstrap()
    
    def load_df(self, filename: str):
        fname = path.basename(filename).split(".xlsx")[0]
        pkl_path = f"{fname}_pkls"
        self.complete_df = {}
        self.complete_df = self.read_excel(filename)
        # if path.isdir(pkl_path):
        #     self.complete_df = self.read_pickles(pkl_path) 
        # else:
        #     mkdir(pkl_path)
        #     self.complete_df = self.read_excel(filename)
        #     self.cache_df(pkl_path)

    def cache_df(self, basepath: str):
        for k in self.complete_df:
            self.complete_df[k].to_pickle(f"{basepath}/{k}.pkl")


    def read_pickles(self, basepath: str):
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        complete_df = {}
        for month in months:
            temp_key = f"{month} Temp"
            rh_key = f"{month} RH"
            complete_df[temp_key] = pd.read_pickle(f"{basepath}/{month} Temp.pkl")
            complete_df[rh_key] = pd.read_pickle(f"{basepath}/{month} RH.pkl")
        return complete_df

    def read_excel(self, filepath: str) -> pd.DataFrame:
        return pd.read_excel(filepath, sheet_name=None, engine='openpyxl')
    
    def bootstrap(self):
        # loop through months in a year
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        for month in months:
            end_row = 508 if month == "Feb" else 511
            temp = self.complete_df[month + ' Temp'].iloc[0:end_row, 2:26]
            rh = self.complete_df[month + ' RH'].iloc[0:end_row, 2:26]
            self.year_data[month] = {
                'temp': temp,
                'rh': rh
            }
        return self
    
    def calculate_coincindent_rhs(self):
        for month in self.year_data:
            temp_df = self.year_data[month]['temp']
            rh_df = self.year_data[month]['rh']
            coincident_rhs = self._calculate_concindent_rh(temp_df, rh_df)
            self.year_data[month]['coincident_rhs'] = coincident_rhs
        return self
    
    def plot_coincindent_rh(self, pd: str):
        # Get the last row of the DataFrame
        last_row = pd.iloc[-1]
        
        # Plot the column names and the last row
        plt.figure(figsize=(10, 6))
        plt.scatter(last_row.index, last_row.values)
        plt.xlabel('Temp')
        plt.ylabel('Average RH Coincidences')
        plt.title('Temp vs Average RH Coincidences')
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

    def _verified_data(self, the_data):
        try:
            if the_data is not None:
                if isinstance(the_data, (int, float)):
                    if the_data >= 1:
                        return True
        except:
            pass
        return False
    
    def _calculate_concindent_rh(self, temp_df: pd.DataFrame, rh_df: pd.DataFrame):
        averages = []
        max_rows = 0
        coincident_df = pd.DataFrame()

        num_rows = temp_df.shape[0]  # Get the number of rows in the DataFrame

        for i in range(num_rows):  # Loop through the rows
            for j in range(temp_df.shape[1]):  # Loop through the columns
                temp_value = temp_df.iloc[i, j]
                if self._verified_data(temp_value):
                    rh_value = rh_df.iloc[i, j]
                    coincident_df = self._average_rh(coincident_df, temp_value, rh_value)

        for col in coincident_df.columns:
            column_data = coincident_df[col].dropna()
            if not column_data.empty:
                avg = column_data.mean()
                averages.append(avg)
                if len(column_data) > max_rows:
                    max_rows = len(column_data)

        result_df = pd.DataFrame(coincident_df, index=range(max_rows + 1), columns=coincident_df.columns)

        for j, avg in enumerate(averages):
            result_df.iloc[max_rows, j] = round(avg, 2)

        return result_df
    
    def _average_rh(self, coincident_df, temp_val, rh_val):
        bool_present = False
        if temp_val in coincident_df.columns:
            col = coincident_df[temp_val]
            next_row = col.first_valid_index()
            if next_row is None:
                next_row = 0
            else:
                next_row += 1

            coincident_df.at[next_row, temp_val] = rh_val
            bool_present = True
        else:
            new_data = {temp_val: [rh_val]}
            new_df = pd.DataFrame(new_data)
            coincident_df = pd.concat([coincident_df, new_df], axis=1)

        return coincident_df
    
    # compute averages
    def compute_averages(self, data):
        temp_averages = []
        rh_averages = []
        temp = data['temp']
        rh = data['rh']

        for col in temp.columns:
            temp_values = []
            rh_values = []

            for row in temp.index:
                temp_val = temp.at[row, col]
                rh_val = rh.at[row, col]
                if self._verified_data(temp_val):
                    temp_values.append(temp_val)
                    rh_values.append(rh_val)

            if temp_values:
                temp_avg = round(float(sum(temp_values)) / len(temp_values), 1)
                rh_avg = round(float(sum(rh_values)) / len(rh_values), 1)
            else:
                temp_avg = ""
                rh_avg = ""

            temp_averages.append(temp_avg)
            rh_averages.append(rh_avg)

        filtered_temp_averages = list(filter(lambda x: x != "", temp_averages))
        filtered_rh_averages = list(filter(lambda x: x != "", rh_averages))

        max_temp_avg = float(max(filtered_temp_averages))
        min_temp_avg = float(min(filtered_temp_averages))
        avg_of_avgs_temp = round(float(sum(filtered_temp_averages)) / len(filtered_temp_averages), 1)

        max_rh_avg = float(max(filtered_rh_averages))
        min_rh_avg = float(min(filtered_rh_averages))
        avg_of_avgs_rh = round(float(sum(filtered_rh_averages)) / len(filtered_rh_averages), 1)

        temp_std_dev = self._calculate_standard_deviation(temp, temp_averages, avg_of_avgs_temp)
        rh_std_dev = self._calculate_standard_deviation(rh, rh_averages, avg_of_avgs_rh)

        hourly_temp_values = []
        hourly_rh_values = []
        for j in temp.columns:
            if temp_averages[j-1] != "":
                tasterix = (2 * math.pi * (j - 1)) / 24
                hourly_temp = round(self._calculate_hourly_temp(avg_of_avgs_temp, max_temp_avg, min_temp_avg, tasterix), 1)
                hourly_temp_values.append(hourly_temp)
            else:
                hourly_temp_values.append("")

            if rh_averages[j-1] != "":
                tasterix = (2 * math.pi * (j - 1)) / 24
                hourly_rh = round(self._calculate_hourly_rh(avg_of_avgs_rh, max_rh_avg, min_rh_avg, tasterix), 1)
            hourly_rh_values.append(hourly_rh)
        else:
            hourly_rh_values.append("")


        if len(hourly_rh_values) == 25:
            hourly_rh_values[24] = avg_of_avgs_rh
        else:
            hourly_rh_values.append(avg_of_avgs_rh)
        
        if len(hourly_temp_values) == 25:
            hourly_temp_values[24] = avg_of_avgs_temp
        else:
            hourly_temp_values.append(avg_of_avgs_temp)

        statics = {
            "Max Average Temp": max_temp_avg,
            "Min Average Temp": min_temp_avg,
            "Temp Range": round(max_temp_avg - min_temp_avg, 1),
            "Temp STD": temp_std_dev,
            "Max Average RH": max_rh_avg,
            "Min Average RH": min_rh_avg,
            "RH STD": rh_std_dev,
            "RH Range": round(max_rh_avg - min_rh_avg, 1),
        }
        averages = {
            "Temperature": hourly_temp_values,
            "Relative Humidity": hourly_rh_values
        }
        return statics, averages

    def _calculate_standard_deviation(self, df: pd.DataFrame, averages, avg_of_avgs):
        valid_hours = 0
        std_dev = 0

        for col in df.columns:
            for row in df.index:
                cell_value = df.at[row, col]
                if self._verified_data(cell_value):
                    valid_hours += 1
                    std_dev += (cell_value - avg_of_avgs) ** 2

        return round(math.sqrt(std_dev / valid_hours), 1)

    def _calculate_hourly_temp(self, the_av_temp, the_max_temp, the_min_temp, tasterix):
        try:
            result = the_av_temp + (the_max_temp - the_min_temp) * (
                0.4535 * math.cos(tasterix - 3.7522) +
                0.1207 * math.cos(2 * tasterix - 0.3895) +
                0.0146 * math.cos(3 * tasterix - 0.8927) +
                0.0212 * math.cos(4 * tasterix - 0.2674)
            )
            return result
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def _calculate_hourly_rh(self, the_av_rh, the_max_rh, the_min_rh, tasterix):
        try:
            result = the_av_rh + (the_max_rh - the_min_rh) * (
                0.4602 * math.cos(tasterix - 0.6038) +
                0.1255 * math.cos(2 * tasterix - 3.5427) +
                0.0212 * math.cos(3 * tasterix - 4.2635) +
                0.0255 * math.cos(4 * tasterix - 0.3833)
            )
            return result
        except Exception as e:
            print(f"An error occurred: {e}")
            return None

    def plot_averages_stats(self, month: str, results):
        df = pd.DataFrame.from_dict(results, orient='index', columns=['Value'])
        # Plot the table
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, rowLabels=df.index, cellLoc='center', loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.2)
        plt.subplots_adjust(left=0.4, right=0.5, top=0.9, bottom=0.1)
        ax.set_title(f'{month} Averages and Statistics', fontsize=14, fontweight='bold')
        plt.show()

    def plot_averages(self, month:str, averages):
        # Create a DataFrame from the averages dictionary
        df = pd.DataFrame(averages)

        # Plot the table
        fig, ax = plt.subplots(figsize=(12, 6))
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, rowLabels=df.index, cellLoc='center', loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.2)

        # Set the title
        ax.set_title(f'{month} Hourly Averages for Temperature and Relative Humidity', fontsize=14, fontweight='bold')

        # Adjust the subplot margins
        plt.subplots_adjust(left=0.2, right=0.8, top=0.8, bottom=0.1)

        plt.show()

    def freq_of_data_in_range(self, df, data_type):
        int_data_point_count = 0
        probabilities = pd.DataFrame(0, index=[0, 1], columns=range(1, 51))

        for col in df.columns:
            for row in df.index:
                val = df.at[row, col]

                if self._verified_data(val):
                    int_data_point_count += 1
                    if data_type == 'temp':
                        k = self._check_temp_range_of_data(val)
                    elif data_type == 'rh':
                        k = self._check_rh_range_of_data(val)
                    else:
                        raise ValueError("Invalid data type. Use 'temp' or 'rh'.")

                    if pd.notna(probabilities.at[0, k]):
                        probabilities.at[0, k] += 1
                    else:
                        probabilities.at[0, k] = 1

        # Ensure the entire second row is of float type
        probabilities.loc[1] = probabilities.loc[1].astype(float)

        for k in range(1, probabilities.shape[1] + 1):
            if pd.notna(probabilities.at[0, k]):
                probabilities.at[1, k] = round(float(probabilities.at[0, k]) / int_data_point_count, 6)

        return probabilities

    def _check_temp_range_of_data(self, the_temp_data):
        if '.' in str(the_temp_data):
            return int(str(the_temp_data).split('.')[0])
        else:
            return int(the_temp_data)

    def _check_rh_range_of_data(self, the_rh_data):
        rh_div_5 = the_rh_data / 5
        if '.' in str(rh_div_5):
            if int(str(rh_div_5).split('.')[1]) > 0:
                return int(str(rh_div_5).split('.')[0]) + 1
            else:
                return int(rh_div_5) + 1
        else:
            return int(rh_div_5) + 1
        
    def compute_fourier(self, temp_df, rh_df, txt_new_fourier_tav, txt_new_fourier_trange, txt_new_fourier_rhav, txt_new_fourier_rhrange):

        sng_calc_temp_av = 0
        sng_calc_rh_av = 0
        sng_max_av = 0
        sng_min_av = 10000
        sng_tav = 0
        sng_std = 0
        int_count_temp = 0
        int_count_rh = 0
        int_valid_hours = 0
        int_av_of_av_count = 0

        fourier_th = pd.DataFrame(index=[0, 1], columns=temp_df.columns)
        fourier_rh = pd.DataFrame(index=[0, 1], columns=temp_df.columns)

        for j in temp_df.columns:  # loop through DataFrame columns
            sng_calc_temp_av = 0
            int_count_temp = 0
            sng_calc_rh_av = 0
            int_count_rh = 0

            for i in temp_df.index:  # loop through DataFrame rows
                temp_val = temp_df.iloc[i, j]
                rh_val = rh_df.iloc[i, j]

                if self._verified_data(temp_val):
                    int_valid_hours = j
                    sng_calc_temp_av += temp_val
                    int_count_temp += 1

                    sng_calc_rh_av += rh_val
                    int_count_rh += 1

            if sng_calc_temp_av != 0:
                fourier_th.loc[0, j - 2] = round(sng_calc_temp_av / int_count_temp, 1)
                fourier_rh.loc[0, j - 2] = round(sng_calc_rh_av / int_count_rh, 1)
            else:
                fourier_th.loc[0, j - 2] = ""
                fourier_rh.loc[0, j - 2] = ""

        sng_calc_temp_av = 0
        sng_min_av = 10000
        sng_max_av = 0
        int_av_of_av_count = 0

        for j in temp_df.columns:
            if fourier_th.loc[0, j] != "":
                if fourier_th.loc[0, j] > sng_max_av:
                    sng_max_av = fourier_th.loc[0, j]

                if fourier_th.loc[0, j] < sng_min_av:
                    sng_min_av = fourier_th.loc[0, j]

                int_av_of_av_count += 1
                sng_calc_temp_av += fourier_th.loc[0, j]

        sng_calc_temp_av = round(sng_calc_temp_av / int_av_of_av_count, 1)

        for j in temp_df.columns:
            if pd.notna(fourier_th.loc[0, j]):
                sng_tasterix = (2 * math.pi * (j - 1)) / 24
                fourier_th.loc[1, j] = round(self._calculate_fourier_hourly_temp(float(txt_new_fourier_tav), float(txt_new_fourier_trange), sng_tasterix), 1)

        int_valid_hours = 0
        sng_std = 0
        for j in temp_df.columns:
            for i in temp_df.index:
                temp_val = temp_df.iloc[i, j]
                if self._verified_data(temp_val):
                    int_valid_hours += 1
                    sng_std += (temp_val - float(txt_new_fourier_tav)) ** 2

        sng_calc_rh_av = 0
        sng_min_av = 10000
        sng_max_av = 0
        int_av_of_av_count = 0

        for j in temp_df.columns:
            if fourier_rh.loc[0, j] != "":
                if fourier_rh.loc[0, j] > sng_max_av:
                    sng_max_av = fourier_rh.loc[0, j]

                if fourier_rh.loc[0, j] < sng_min_av:
                    sng_min_av = fourier_rh.loc[0, j]

                int_av_of_av_count += 1
                sng_calc_rh_av += fourier_rh.loc[0, j]

        sng_calc_rh_av = round(sng_calc_rh_av / int_av_of_av_count, 2)

        for j in temp_df.columns:
            if pd.notna(fourier_rh.loc[0, j]):
                sng_tasterix = (2 * math.pi * (j - 1)) / 24
                fourier_rh.loc[1, j] = round(self._calculate_fourier_hourly_rh(float(txt_new_fourier_rhav), float(txt_new_fourier_rhrange), sng_tasterix), 1)

        return fourier_th, fourier_rh
    
    def _calculate_fourier_hourly_temp(self, the_av_temp, the_range, tasterix):
        result = the_av_temp + the_range * (
            0.4535 * math.cos(tasterix - 3.7522) +
            0.1207 * math.cos(2 * tasterix - 0.3895) +
            0.0146 * math.cos(3 * tasterix - 0.8927) +
            0.0212 * math.cos(4 * tasterix - 0.2674)
        )
        return result

    def _calculate_fourier_hourly_rh(self, the_av_rh, the_range, tasterix):
        result = the_av_rh + the_range * (
            0.4606 * math.cos(tasterix - 0.6038) +
            0.1255 * math.cos(2 * tasterix - 3.5247) +
            0.0212 * math.cos(3 * tasterix - 4.2635) +
            0.0255 * math.cos(4 * tasterix - 0.3833)
        )
        return result
    
    def compute_heating_degree_days(self, base_temp_model_hdd, tav_hdd, sigma_m_model_hdd, num_of_days_in_month_hdd):
        sng_heating_hb = (float(base_temp_model_hdd) - float(tav_hdd)) / (float(sigma_m_model_hdd) * math.sqrt(float(num_of_days_in_month_hdd)))
        avg_monthly_hdd = round((float(sigma_m_model_hdd) * math.pow(float(num_of_days_in_month_hdd), 3 / 2)) * (0.072196 + (sng_heating_hb / 2) + (1 / 9.6) * math.log(math.cosh(4.8 * sng_heating_hb))), 2)
        return avg_monthly_hdd


    def compute_cooling_degree_days(self, tav_cdd, base_temp_model_cdd, sigma_m_model_cdd, num_of_days_in_month_cdd):
        sng_cooling_hb = (float(tav_cdd) - float(base_temp_model_cdd)) / (float(sigma_m_model_cdd) * math.sqrt(float(num_of_days_in_month_cdd)))
        avg_monthly_cdd = round((float(sigma_m_model_cdd) * math.pow(float(num_of_days_in_month_cdd), 3 / 2)) * (0.072196 + (sng_cooling_hb / 2) + (1 / 9.6) * math.log(math.cosh(4.8 * sng_cooling_hb))), 2)
        return avg_monthly_cdd
    
    def freq_of_rh_in_temp_range(self, temp_df, rh_df, the_end_row):
        int_probability_count = 0
        the_table = pd.DataFrame(0, index=temp_df.columns, columns=range(1, 51))
        the_probability_table = pd.DataFrame(0, index=temp_df.columns, columns=range(1, 51))
        for i in temp_df.index:  # Loop through DataFrame rows
            for j in temp_df.columns:  # Loop through DataFrame columns
                temp_val = temp_df.iloc[i, j]
                if self._verified_data(temp_val):
                    int_probability_count += 1
                    k = self._check_temp_range_of_data(temp_val)
                    rh_val = rh_df.iloc[i, j]
                    l = self._check_rh_range_of_data(rh_val)
                    if pd.notna(the_table.iloc[l, k]):
                        the_table.iloc[l, k] += 1
                    else:
                        the_table.iloc[l, k] = 1

        for i in range(1, the_table.shape[0]):
            for j in range(1, the_table.shape[1]):
                if pd.notna(the_table.iloc[i, j]) and the_table.iloc[i, j] != 0:
                    dec_probability = the_table.iloc[i, j] / int_probability_count
                    the_probability_table.iloc[i, j] = self._rounding_process(dec_probability)
                else:
                    the_table.iloc[i, j] = 0
                    the_probability_table.iloc[i, j] = 0

        return the_table, the_probability_table

    def _rounding_process(self, val):
        val_str = str(val)
        int_decimal_position = val_str.find(".")

        # Return six places after decimal point
        str_check = val_str[:int_decimal_position + 7]

        int_last_digit = int(str_check[-1])
        str_check = val_str[:int_decimal_position + 6]

        if int_last_digit > 4:
            str_output = str(float(str_check) + 0.000001)
            return str_output
        else:
            return val_str[:int_decimal_position + 6]

    def plot_generic_table(self, title: str, data: pd.DataFrame):
        df = pd.DataFrame.from_dict(data)
        # Plot the table
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.axis('tight')
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, rowLabels=df.index, cellLoc='center', loc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.2)
        plt.subplots_adjust(left=0.4, right=0.5, top=0.9, bottom=0.1)
        ax.set_title(f'{title}', fontsize=14, fontweight='bold')
        plt.show()