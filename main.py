import pandas
import os
import re
import tkinter as tk
import time

# Note: the following sequence shows up throughout the code:
#       tk_widget.config(text="updated")
#       tk_widget.update()
# This is because widget updates are delayed by default until the callback function is finished. The tk_widget.update()
# syntax forces the GUI to update the widgets immediately, for real-time response to user changes.

# Note: double backslashes are used throughout the code for single backslashes, due to escape sequences

# Color strings for the GUI window
bg_color_str = 'black'
text_color_str = 'black'
label_color_str = 'gray'
error_text_color_str = 'red'
error_label_color_str = 'red'

# Other default formatting values for the GUI window
global_padding_x = 5
global_padding_y = 5
cell_entry_width = 5

# Default entry values
sn_start_default = 33700
sn_end_default = 33726
search_dir_default = "C:\\Machines"
output_dir_default = "C:\\Desktop"
sheet_name_default = "Summary"
keyword_default = "cal"

cell_entry_list = []
default_cal_test_cell_list = ['A34', 'B4', 'E2', 'E3', 'B20', 'B21', 'B22', 'B23',
                               'C20', 'C21', 'C22', 'C23', 'F20', 'G20', 'H20', 'I20']
# default_light_cal_cell_list = ['D13', 'K14', 'M14', 'Q14', 'K17']
cell_entry_count = len(default_cal_test_cell_list)
# cell_entry_count = len(default_light_cal_cell_list)

# RegEx for validating user input in the cell entry widgets. The '^' symbol signifies the start of the match pattern,
# the '$' sign signifies the end of the match pattern.
valid_cell_regex = re.compile('^[a-zA-Z]\d+$')
# RegEx for validating serial number input. Serial number must be a five-digit number only
valid_sn_regex = re.compile('^[0-9][0-9][0-9][0-9]$')


# function that converts uppercase letters to a column number
def convert_column_letter_to_num(char):
    return ord(char.lower()) - 97


def check_settings_inputs():
    valid_input = True
    return valid_input


def get_cell_list(convert_columns=True):
    cell_list = []
    for entry in cell_entry_list:
        cell_input = entry.get()
        if cell_input:
            if valid_cell_regex.match(cell_input):
                if convert_columns:
                    cell_list.append([int(cell_input[1:]), convert_column_letter_to_num(cell_input[0])])
                else:
                    cell_list.append([int(cell_input[1]-2), cell_input[0]])
    return cell_list


# This function is called when the "Submit Search" button is pressed, and contains the main functionality for this
# program. It takes the user input from the entry fields and processes the selected serial numbers using nested
# for loops. List comprehensions are used throughout to filter unneeded directories and files based on the keywords
# provided by the user.
def begin_search():

    extracted_data_list = []

    # Get the user-inputted sn's, directory, and keyword from the entry widgets in the 'Search Settings' frame
    first_mach_sn = int(sn_first_entry.get())
    last_mach_sn = int(sn_last_entry.get())
    cm_mach_dir = search_dir_entry.get()
    keyword = file_keyword_entry.get().lower()
    sheet_name = sheet_name_entry.get()

    cell_list = get_cell_list()
    start_time = time.perf_counter()
    # This for loop steps through the specified directory in sequential order by serial number
    # noinspection PyInterpreter
    for current_mach_sn in range(first_mach_sn, last_mach_sn + 1):
        # Generate a list of directories and files in the sub-folder using os.listdir()  and
        # filter out all non-directory objects in dir_list using a list comprehension with a conditional calling
        # os.path.isdir() to check if the path leads to a directory, and if the that path includes our
        # user-supplied keyword.
        matching_file_list = []
        # Update the progress label each time the program moves on to a new serial number
        update_str = "Processing: {} of {}: SN{}".format(current_mach_sn - first_mach_sn + 1,
                                                         last_mach_sn - first_mach_sn + 1,
                                                         current_mach_sn)
        prog_label.config(text=update_str)
        prog_label.update()
        active_dir = "{}\\SN{}".format(cm_mach_dir, current_mach_sn)

        # This try/except block attempts to
        try:
            matching_dir_list = ["{}\\{}".format(active_dir, name) for name in os.listdir(active_dir)
                                 if (os.path.isdir("{}\\{}".format(active_dir, name))
                                     and keyword.lower() in "{}\\{}".format(active_dir, name).lower())]
        except FileNotFoundError:
            data_entry = [current_mach_sn, active_dir]
            for cell in cell_list:
                data_entry.append('FNF')
            extracted_data_list.append(data_entry)
            continue

        for directory in matching_dir_list:
            for file in os.listdir(directory):
                if '.xlsm' in file.lower() and keyword.lower() in file.lower():
                    matching_file_list.append("{}\\{}".format(directory, file))
            # active_xl_df = pandas.read_excel(directory, sheet_name=sheet_name)

        # Loop through the matching file list and extract the cell data
        for file_str in matching_file_list:

            active_xl_df = pandas.read_excel(file_str, dtype=object, sheet_name=sheet_name, index_col=None, header=None)
            start_df_time = time.perf_counter()
            data_entry = [current_mach_sn, file_str]
            for cell in cell_list:
                data_entry.append(active_xl_df.at[cell[0] - 1, cell[1]])
            extracted_data_list.append(data_entry)
            print("data frame proc time = {}".format(time.perf_counter() - start_df_time))

    # check the extracted data list
    for entry in extracted_data_list:
        print(entry)
    column_list = ['SN', 'Directory']
    for cell in cell_list:
        column_list.append(cell)

    pandas.DataFrame(extracted_data_list).to_excel(output_dir_entry.get() + "\\output.xlsx")

    update_str = "Finished: {} directories processed in {}s".format(last_mach_sn - first_mach_sn + 1,
                                                                   time.perf_counter() - start_time)
    prog_label.config(text=update_str)
    prog_label.update()
    return True


# Define the tkinter GUI window and add some minimal formatting
gui = tk.Tk()
gui.title("Machine Surfer")
gui.configure(bg='black')
gui.minsize(1000, 250)

# Define all of the tkinter widgets in the GUI
main_frame = tk.Frame(gui, bg=label_color_str)
picture_frame = tk.Frame(gui, bg=label_color_str)
settings_frame = tk.Frame(main_frame, bg=label_color_str)
settings_title_label = tk.Label(settings_frame, text='Search Settings', bg=label_color_str)
search_dir_label = tk.Label(settings_frame, text='Search Directory:', bg=label_color_str, fg=text_color_str)
search_dir_entry = tk.Entry(settings_frame, fg=text_color_str)
search_dir_entry.insert(0, search_dir_default)
output_dir_label = tk.Label(settings_frame, text='Output Directory', bg=label_color_str, fg=text_color_str)
output_dir_entry = tk.Entry(settings_frame, fg=text_color_str)
output_dir_entry.insert(0, output_dir_default)
sn_first_label = tk.Label(settings_frame, text='First SN:', bg=label_color_str, fg=text_color_str)
sn_first_entry = tk.Entry(settings_frame, fg=text_color_str)
sn_first_entry.insert(0, str(sn_start_default))
sn_last_label = tk.Label(settings_frame, text='Last SN:', bg=label_color_str, fg=text_color_str)
sn_last_entry = tk.Entry(settings_frame, fg=text_color_str)
sn_last_entry.insert(0, str(sn_end_default))
file_keyword_label = tk.Label(settings_frame, text='Filename Keyword:', bg=label_color_str, fg=text_color_str)
file_keyword_entry = tk.Entry(settings_frame, fg=text_color_str)
file_keyword_entry.insert(0, str(keyword_default))
sheet_name_label = tk.Label(settings_frame, text='Sheet Name:', bg=label_color_str, fg=text_color_str)
sheet_name_entry = tk.Entry(settings_frame, fg=text_color_str)
sheet_name_entry.insert(0, sheet_name_default)
start_btn = tk.Button(settings_frame, text='Start Search', command=begin_search, bg=label_color_str, fg=text_color_str)
prog_label = tk.Label(settings_frame, bg=label_color_str, text='Progress', anchor='w')

cell_selection_frame = tk.Frame(main_frame, bg=label_color_str)
for i in range(cell_entry_count):
    cell_entry_list.append(
        tk.Entry(cell_selection_frame, width=cell_entry_width, fg=text_color_str))
    cell_entry_list[i].insert(0, default_cal_test_cell_list[i])
"""check_cells_btn = tk.Button(cell_selection_frame, bg=label_color_str, text='Check Cells', anchor='w')
check_cells_btn.grid(row=1, column=0, sticky='w', padx=global_padding_x, pady=global_padding_y)"""

# Build the GUI using the tkinter grid layout manager
# tkinter widgets contained within the search settings frame
main_frame.grid(row=0, column=0, padx=global_padding_x, pady=global_padding_y)
picture_frame.grid(row=0, column=0, padx=global_padding_x, pady=global_padding_y)
settings_frame.grid(row=0, column=0, padx=global_padding_x, pady=global_padding_y)
settings_title_label.grid(row=0, column=0, columnspan=4, sticky='ew', padx=global_padding_x, pady=global_padding_y)
search_dir_label.grid(row=1, column=0, columnspan=1, padx=global_padding_x, pady=global_padding_y)
search_dir_entry.grid(row=1, column=1, columnspan=3, sticky='ew', padx=global_padding_x, pady=global_padding_y)
output_dir_label.grid(row=2, column=0, columnspan=1, padx=global_padding_x, pady=global_padding_y)
output_dir_entry.grid(row=2, column=1, columnspan=3, sticky='ew', padx=global_padding_x, pady=global_padding_y)
sn_first_label.grid(row=3, column=0, padx=global_padding_x, pady=global_padding_y)
sn_first_entry.grid(row=3, column=1, padx=global_padding_x, pady=global_padding_y)
sn_last_label.grid(row=3, column=2, padx=global_padding_x, pady=global_padding_y)
sn_last_entry.grid(row=3, column=3, padx=global_padding_x, pady=global_padding_y)
sheet_name_label.grid(row=4, column=0, padx=global_padding_x, pady=global_padding_y)
sheet_name_entry.grid(row=4, column=1, padx=global_padding_x, pady=global_padding_y)
file_keyword_label.grid(row=4, column=2, padx=global_padding_x, pady=global_padding_y)
file_keyword_entry.grid(row=4, column=3, padx=global_padding_x, pady=global_padding_y)
prog_label.grid(row=5, column=0, columnspan=2, sticky='ew', padx=global_padding_x, pady=global_padding_y)
settings_title_label = tk.Label(settings_frame, text='Search Settings', bg=label_color_str)
start_btn.grid(row=6, column=0, columnspan=2, sticky='ew', padx=global_padding_x, pady=global_padding_y)

# tkinter widgets contained within the cell selection frame
cell_selection_frame.grid(row=0, column=1, padx=global_padding_x, pady=global_padding_y)
for i in range(cell_entry_count):
    cell_entry_list[i].grid(row=0, column=i, padx=global_padding_x, pady=global_padding_y)

# Initiate the tkinter GUI
gui.mainloop()
