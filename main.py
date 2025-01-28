from analyser import Analyzer
from os import path, mkdir

def __main__():
    cities = ['owerri', 'ibadan',  "ikeja", "oshogbo","yola","minna","kano", "jos","ilorin", "abuja", "abeokuta", "port-harcourt", "uyo", "benin", "calabar", "enugu"]
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for city in cities:
        print(f"\n Processing for {city} \n")
        analyzer = Analyzer(path.dirname(__file__) + f"/data/{city}.xlsx")

        for month in months:
            print(f"\n Processing for {month} \n")


            mkdir(f"results/{city}") if path.exists(f"results/{city}") == False else None 
            mkdir(f"results/{city}/{month}") if path.exists(f"results/{city}/{month}") == False else None 
            data = analyzer.year_data[month]

            rh_range = analyzer.freq_of_data_in_range(data["rh"], "rh")
            analyzer.plot_generic_table(rh_range, "Frequency of Temp in RH Range", city, '')

            temp_range = analyzer.freq_of_data_in_range(data["temp"], "temp")
            analyzer.plot_generic_table(temp_range, "Frequency of RH in Temp Range", city, '')
            
            temp_df = data["temp"]
            rh_df = data["rh"]


            analyzer.calculate_coincindent_rhs()
            analyzer.plot_coincindent_rh(city, month, data['coincident_rhs'])
            stats, avgs = analyzer.compute_averages(data)
            analyzer.plot_averages(city, month, avgs)
            analyzer.plot_averages_stats(city, month, stats)


            end_row = 512  # could also be 5800 for a year
            basic_table, the_probability_table = analyzer.freq_of_rh_in_temp_range(temp_df, rh_df, end_row, city, month)
            analyzer.plot_generic_table(basic_table, "Basic Table",city, month)
            analyzer.plot_generic_table(the_probability_table, "Probability Table", city, month)

            txt_new_fourier_tav = "20.0"
            txt_new_fourier_trange = "10.0"
            txt_new_fourier_rhav = "50.0"
            txt_new_fourier_rhrange = "20.0"
            fourier_th, fourier_rh = analyzer.compute_fourier(data["temp"], data["rh"], txt_new_fourier_tav, txt_new_fourier_trange, txt_new_fourier_rhav, txt_new_fourier_rhrange)

            analyzer.plot_generic_table(fourier_rh, "Fourier Analysis - RH", city, month)
            analyzer.plot_generic_table(fourier_th, "Fourier Analysis - Temp", city, month)

            base_temp_model_hdd = "30.0"
            tav_hdd = "25.0"
            sigma_m_model_hdd = "2.0"
            num_of_days_in_month_hdd = "30"

            tav_cdd = "25.0"
            base_temp_model_cdd = "18.0"
            sigma_m_model_cdd = "2.0"
            num_of_days_in_month_cdd = "30"

            avg_monthly_hdd = analyzer.compute_heating_degree_days(base_temp_model_hdd, tav_hdd, sigma_m_model_hdd, num_of_days_in_month_hdd)

            analyzer.plot_generic_table(avg_monthly_hdd, "Average Heating Degree Days", city, month)
            avg_monthly_cdd = analyzer.compute_cooling_degree_days(tav_cdd, base_temp_model_cdd, sigma_m_model_cdd, num_of_days_in_month_cdd)
            analyzer.plot_generic_table(avg_monthly_cdd, "Average Cooling Degree Days", city, month)





# initialize the program
if __name__ == "__main__":
    __main__()