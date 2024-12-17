from analyser import Analyzer
from os import path

def __main__():
    # create an instance of the Analyzer class
    analyzer = Analyzer(path.dirname(__file__) + "/data/calabar.xlsx")
    
    month = "Feb"
    data = analyzer.year_data[month]
    temp_df = data["temp"]
    rh_df = data["rh"]
    # analyzer.calculate_coincindent_rhs()
    # analyzer.plot_coincindent_rh(data['coincident_rhs'])
    # stats, avgs = analyzer.compute_averages(data)
    # analyzer.plot_averages(month, avgs)
    # analyzer.plot_averages_stats(month, stats)
    # temp_range = analyzer.freq_of_data_in_range(temp, rh)
    # rh_range = analyzer.freq_of_data_in_range(data["rh"], "rh")
    # print(rh_range)
    # print(temp_range)


    # end_row = 483  # could also be 5800 for a year
    # basic_table, the_probability_table = analyzer.freq_of_rh_in_temp_range(temp_df, rh_df, end_row)
    # analyzer.plot_generic_table("Basic Table", basic_table)
    # analyzer.plot_generic_table("Probability Table", the_probability_table)

    # txt_new_fourier_tav = "20.0"
    # txt_new_fourier_trange = "10.0"
    # txt_new_fourier_rhav = "50.0"
    # txt_new_fourier_rhrange = "20.0"
    # fourier_th, fourier_rh = analyzer.compute_fourier(data["temp"], data["rh"], txt_new_fourier_tav, txt_new_fourier_trange, txt_new_fourier_rhav, txt_new_fourier_rhrange)
    # print(fourier_th)
    # print(fourier_rh)

    # base_temp_model_hdd = "30.0"
    # tav_hdd = "25.0"
    # sigma_m_model_hdd = "2.0"
    # num_of_days_in_month_hdd = "30"

    # tav_cdd = "25.0"
    # base_temp_model_cdd = "18.0"
    # sigma_m_model_cdd = "2.0"
    # num_of_days_in_month_cdd = "30"

    # avg_monthly_hdd = analyzer.compute_heating_degree_days(base_temp_model_hdd, tav_hdd, sigma_m_model_hdd, num_of_days_in_month_hdd)
    # avg_monthly_cdd = analyzer.compute_cooling_degree_days(tav_cdd, base_temp_model_cdd, sigma_m_model_cdd, num_of_days_in_month_cdd)

    # print(f"Average Monthly HDD: {avg_monthly_hdd}")
    # print(f"Average Monthly CDD: {avg_monthly_cdd}")





# initialize the program
if __name__ == "__main__":
    __main__()