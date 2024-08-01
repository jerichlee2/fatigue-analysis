#class for Funkter: a basic data->FELA mapping 
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import pandas as pd
import rainflow
import os

class Funkter:
    def __init__(self):
        self.hardbank = np.array([['Platform', '950 962 M&L', '950 962 AU2020', '966 972', '980 982'],
                            ['Hardbank', 70, 45, 90, 90],
                            ['Truck Loading', 0, 25, 0, 0],
                            ['Bulldoze/Backdrag', 9, 9, 8, 8],
                            ['Non-Damaging', 21, 21, 2, 2],
                            ['Sum', 100, 100, 100, 100]])
        
        self.truck_loading =  np.array([['Platform', '950 962 M&L', '950 962 AU2020', '966 972', '980 982'],
                            ['Hardbank', 10, 10, 0, 0],
                            ['Truck Loading–2" Rock', 35, 35, 45, 45],
                            ['Truck Loading–Pea Gravel', 35, 35, 45, 45],
                            ['Bulldoze/Backdrag', 9, 9, 8, 8],
                            ['Non-Damaging', 11, 11, 2, 2],
                            ['Sum', 100, 100, 100, 100]])


    def data_ingestion(self, sensor_data):
        pass

    def work_profiles(self, sensor_data, machine):
        self.hardbank = pd.DataFrame(data=self.hardbank[1:,1:], index=self.hardbank[1:,0], columns=self.hardbank[0,1:])
        self.truck_loading = pd.DataFrame(data=self.truck_loading[1:,1:], index=self.truck_loading[1:,0], columns=self.truck_loading[0,1:])
#
    def data_histogram(self, data, min, max, increment):
        # Define histogram range boundaries and increment
        min_range = min
        max_range = max
        increment = increment

        # Create bins for the histogram
        bins = list(range(min_range, max_range + increment, increment))

        # Compute histogram using pandas
        hist, edges = pd.cut(data, bins=bins, include_lowest=True, right=False, retbins=True)
        hist_counts = hist.value_counts()
        total_samples = len(data)

        # Calculate percentage
        hist_percent = (hist_counts / total_samples) * 100

        # Create DataFrame
        histogram_df = pd.DataFrame({
            'kPa': edges[:-1],
            'percent': hist_percent
        })
        return histogram_df
    

    def composite_histogram(self, data, min_val, max_val, increment, name):
        results_dir = "Histograms/"
        if not os.path.isdir(results_dir):
            os.makedirs(results_dir) 
        plt.figure(figsize=(10, 6))
        xaxis = np.arange(0, 50000, 2000).tolist()
        if len(data[0]) < 25:
            for i in range(len(data)):
                data[i].append(0)
                data[i].append(0)

        species = tuple(xaxis)
        penguin_means = {
            'Hardbank Composite': tuple(data[0]),
            'Truck Loading Composite': tuple(data[1]),
        }

        x = np.arange(len(species))  # the label locations
        width = 0.5  # the width of the bars
        multiplier = 0

        fig, ax = plt.subplots(layout='constrained')

        for attribute, measurement in penguin_means.items():
            offset = width * multiplier
            rects = ax.bar(x + offset, measurement, width, label=attribute)
            # ax.bar_label(rects, padding=3)
            multiplier += 1

        # Add some text for labels, title and custom x-axis tick labels, etc.
        ax.set_ylabel('Percentage')
        ax.set_title(name.capitalize())
        ax.set_xticks(x + width, species)
        ax.set_xticklabels(labels=species, fontsize=5, va='bottom', ha='left')
        ax.legend(loc='upper left', ncols=2)
        ax.set_ylim(0, max(max(data[0]), max(data[1]))+max(max(data[0]), max(data[1]))/10)
        # Save the histogram as a PNG file
        ax.figure.savefig(os.path.join(results_dir, f"{name}.png"))


    def rainflow(self, data, num_bins, lower_bound, upper_bound, cutoff):
        # Convert data from kPa to MPaq
        data = [x / 1000 for x in data]
        ranges = rainflow.count_cycles(data)
        # cutoff_value = (max(data) - min(data)) * cutoff / 100
        # ranges = [x for x in ranges if abs(x[0] - x[1]) > cutoff_value]
        ranges = [x for x in ranges]
        hist, bins = np.histogram(ranges, bins=num_bins, range=(lower_bound, upper_bound))
        df = pd.DataFrame({'Range (MPa)': bins[:-1], 'Count': hist})
        # df.to_csv('rainflow_histogram.csv', index=False)
        # print("Rainflow histogram successfully saved to 'rainflow_histogram.csv'")

        return df
    

    
    def load_severity(self, data, lsindex, cutoff):
        loadseverity = 0
        for i in range(0, len(data)):
            loadseverity += (data.iloc[i].iloc[0])**(lsindex)*(data.iloc[i].iloc[1])
 
        return loadseverity
    
    def combined_load_severity(self, profile, loadseverities):
        pass

    def test_cycles(self, loadseverity, targetlife, fela, lsindex):
        return (loadseverity*targetlife)/(fela**lsindex)
    



