import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from fatigue import Funkter
from openpyxl import Workbook
from spreadsheet import Spreadsheet
import sys
import os
sys.path.insert(0, os.environ['DATK_INSTALL_PATH'] + '\\datk\\bin\\bin.win64')
# os.add_dll_directory
if hasattr(os, "add_dll_directory"):
  os.add_dll_directory(os.environ['DATK_INSTALL_PATH'] + '\\datk\\bin\\bin.win64')
import numpy
import D2D_Analysis
import graphviz



g = graphviz.Digraph('g', filename='btree.gv',
                     node_attr={'shape': 'record', 'height': '.1'})

# Get the list of all files and directories


#gui will allow user to open folder
# file = 'C:\\Users\\leej85\\Desktop\\CAT_Internship_Jerich_Lee_2024\\Projects\\Structural_Analysis\\Cylinder_Analysis\\python\\python-examples\\950L_TB_OMLA_2014\\700_TruckLoading_2inchRock_04_10_15.thd'

def truncate(s, n):
    return s[:-n]

def combined(func, side, pos):
   return func+"_"+side+"_"+pos

def combined_tlt(func, pos):
   return func+"_"+pos

def get_last_folder(path):
   return os.path.basename(os.path.normpath(path))

def process_files(path):

   file_list = os.listdir(path)

   ending = '.thd'

# Use list comprehension to filter strings with the specified ending
   cleaned_file_list = [s for s in file_list if s.endswith(ending)]
   #there are no R and L for TLT...
   # cylinder_func = ['LFT', 'TLT', 'STR']
   cylinder_func = ['LFT', 'STR', 'TLT']
   cylinder_pos = ['HE', 'RE']

   #counters
   num_stat = 0
   num_lift = 0
   num_tilt = 0
   num_steer = 0

   wb = Workbook()
   excel_instance = Spreadsheet(wb)
   excel_instance.construct(path)
   excel_instance.constant_sheet('constant_files\Volvo_L150.xlsx', "Cylinder")
   excel_instance.constant_sheet('constant_files\Work_Profiles.xlsx', "Work Profiles")
   for i in range(len(cleaned_file_list)):
      file = path + "\\" + cleaned_file_list[i]
      for func in cylinder_func:
         if func == 'TLT':
            sides = ['L']
         else:
            sides = ['L', 'R']
         for side in sides:
            for pos in cylinder_pos:
               if func == 'TLT':
                  combine = combined_tlt(func, pos)
               else:
                  combine = combined(func, side, pos)
               current_file = truncate(cleaned_file_list[i], 13)
               data_path = numpy.empty([7], dtype = "S33")
               data_path[0] = 'DYNAMIC DATA'
               data_path[1] = current_file
               data_path[2] = current_file
               data_path[3] = 'TIME'
               data_path[4] = 'TIME VECTORS'
               data_path[5] = combine
               data_path[6] = 'ORIGINAL'

               # Send the parameters to wave
               D2D_Analysis.set_wave_data(data_path, 'data_path')
               D2D_Analysis.set_wave_data(file, 'file')

               # get current working directory path
               cwd = os.path.dirname(os.path.realpath(__file__))

               # call the DATK functions
               D2D_Analysis.wave_command('cd, "' + cwd + '"')
               D2D_Analysis.wave_command('fo_tag = FOT_ADD(file, !GDF)')
               D2D_Analysis.wave_command('FO_OPEN_FILE, fo_tag')
               D2D_Analysis.wave_command('tag = do_create( fo_tag, data_path)')
               D2D_Analysis.wave_command('do_read, tag, data, indep, /ALL')
               D2D_Analysis.wave_command('FO_CLOSE_FILE, fo_tag')

               # now get the output data
               data  = D2D_Analysis.get_wave_data('data')
               indep = D2D_Analysis.get_wave_data('indep')

               funkter_instance = Funkter()
               sensor_data = data  # Example sensor data 

               # we need these methods, just turning them off to test other methods
               d2d_statistics = D2D_Analysis.D2D_Analysis('D2D_STATISTICS')
               results_stat = d2d_statistics(sensor_data, indep, 1)
               excel_instance.statistics(results_stat, num_stat, current_file, combine) 

               analysis = D2D_Analysis.D2D_Analysis('d2d_fela')
               results_fela = analysis(sensor_data, indep, 3.0, 10000, 10000000, 100000)
               excel_instance.fela(results_fela, num_stat)
               
      
               df_histogram = funkter_instance.data_histogram(sensor_data, 0, 50000, 2000)

               if func == 'LFT':
                  excel_instance.pressure_histograms(df_histogram, num_lift, current_file, combine, func)
                  num_lift += 1
               elif func == 'TLT':
                  excel_instance.pressure_histograms(df_histogram, num_tilt, current_file, combine, func)
                  num_tilt += 1
               else:
                  excel_instance.pressure_histograms(df_histogram, num_steer, current_file, combine, func)
                  num_steer += 1

               num_stat += 1 

               df_rainflow = funkter_instance.rainflow(sensor_data, 100, 0, 54, 5)
               
               load_severity = funkter_instance.load_severity(df_rainflow, 3, 5)
               excel_instance.load_severity(load_severity, current_file, combine, func, num_stat, 3)

               load_severity = funkter_instance.load_severity(df_rainflow, 6, 5)
               excel_instance.load_severity(load_severity, current_file, combine, func, num_stat, 6 )

   excel_instance.average_pressure_loads() 
   funkter_instance = Funkter()
   composite_lift_head = excel_instance.composite_lift_head()
   # print(funkter_instance.composite_histogram(composite_lift_head, 0, 50000, 2000, "test"))
   funkter_instance.composite_histogram(composite_lift_head, 0, 50000, 2000, "lift_head")
   excel_instance.histogram_chart("Histograms/lift_head.png", 1)

   composite_lift_rod = excel_instance.composite_lift_rod()
   funkter_instance.composite_histogram(composite_lift_rod, 0, 50000, 2000, "lift_rod")
   excel_instance.histogram_chart("Histograms/lift_rod.png", 30)

   composite_steer_head = excel_instance.composite_steer_head()
   funkter_instance.composite_histogram(composite_steer_head, 0, 50000, 2000, "steer_head")
   excel_instance.histogram_chart("Histograms/steer_head.png", 60)

   composite_steer_rod = excel_instance.composite_steer_rod()
   funkter_instance.composite_histogram(composite_steer_rod, 0, 50000, 2000, "steer_rod")
   excel_instance.histogram_chart("Histograms/steer_rod.png", 90)

   composite_tilt_head = excel_instance.composite_tilt_head()
   funkter_instance.composite_histogram(composite_tilt_head, 0, 50000, 2000, "tilt_head")
   excel_instance.histogram_chart("Histograms/tilt_head.png", 120)

   composite_tilt_rod = excel_instance.composite_tilt_rod()
   funkter_instance.composite_histogram(composite_tilt_rod, 0, 50000, 2000, "tilt_rod")
   excel_instance.histogram_chart("Histograms/tilt_rod.png", 150)

   excel_instance.composite_load_severity()

   # g.view()

   # Save and open the Excel file
   excel_file_path = os.path.join(path, f"MWL_Load_Severity_{get_last_folder(path)}.xlsx")
   wb.save(excel_file_path)
   os.startfile(excel_file_path)

def select_folder():
   folder_path = filedialog.askdirectory()
   if folder_path:
       try:
           process_files(folder_path)
           messagebox.showinfo("Success", "Files processed successfully!")
       except Exception as e:
           messagebox.showerror("Error", str(e))
app = tk.Tk()
app.title("Funkter")

# Load and resize logo image
logo_path = "constant_files/logo.png"  # Replace with your logo's path
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((250, 250))  # Adjust the size as needed
logo_photo = ImageTk.PhotoImage(logo_image)
# Create a label for the logo
logo_label = tk.Label(app, image=logo_photo)
logo_label.pack(pady=10)

frame = tk.Frame(app)
frame.pack(pady=20)
select_folder_button = tk.Button(frame, text="Select Folder", command=select_folder)
select_folder_button.pack()
app.mainloop()


