import pandas as pd
import numpy as np
from sys import exit


##################
# helper functions
##################

# function for obtaining the detailed job report from the user
# arguments: none
# returns the detailed job report as an array
# includes error checking
def obtain_detailed_job_report():
    djr_array = []  # initialize detailed job report array
    user_error = True
    while user_error:  # obtaining the Detailed Job Report (.xlsx) spreadsheet by name from directory
        choice = input(
            "Enter (1) to enter the Detailed Job Report by name or (2) to enter the Detailed Job Report by path: ")
        if choice == "1":
            file_found = False
            while not file_found:
                djr_name = input("Enter the file name of the Detailed Job Report: ")
                djr_name_split = djr_name.split(".")

                try:
                    contains_file_type = False
                    if djr_name_split[len(djr_name_split) - 1] == "xlsx" and len(djr_name_split) > 1:
                        contains_file_type = True

                    if not contains_file_type:
                        djr_name = djr_name + ".xlsx"

                    djr_dataframe = pd.read_excel(djr_name, sheet_name='Sheet1')
                    djr_array = djr_dataframe.to_numpy()  # convert data frame to an array

                    file_found = True
                    print("File found.")
                except:
                    print("Error: no such file found in current directory")
                    print("Please try again")
            user_error = False

        elif choice == "2":  # obtaining the Detailed Job Report (.xlsx) spreadsheet by path
            path = ""

            file_found = False
            while not file_found:
                try:
                    path = input("Paste the path to the Detailed Job Report here: ")

                    djr_dataframe = pd.read_excel(path)
                    djr_array = djr_dataframe.to_numpy()  # convert from dataframe to numpy array
                    print("File found.")
                    file_found = True
                except:
                    print("Error: no such file found")
                    print("Please try again")
            user_error = False

        else:
            print("Invalid option, please try again")

    return djr_array


# function for obtaining which machine for analysis from the user
# arguments: the detailed job report array
# returns the machine selected by the user
# includes error checking
def obtain_machine_to_analyze(detailed_job_report):
    list_of_machines = []  # initialize list of all possible machines within the detailed job report
    for row in range(ROWS):
        append_element_in_array(list_of_machines, detailed_job_report[row][MACHINE_COL_NUM])

    print_list_of_machines(list_of_machines)

    while True:
        machine_selection = input("Enter the number associated to the machine you would like to analyze: ")

        try:
            if 1 <= int(machine_selection) <= len(list_of_machines):
                return list_of_machines[int(machine_selection) - 1]
            else:
                print("Invalid input. Please try again.")
        except:
            print("Error. Please try again.")


# function for printing all machines found within a detailed job report
# prints each machine name (as found in the detailed job report) along with an associated number
# arguments: the list of all machines found within the detailed job report
# returns nothing
def print_list_of_machines(machines_list):
    # find length of longest machine name
    longest_machine_name_length = 0
    for item in machines_list:
        if len(item) > longest_machine_name_length:
            longest_machine_name_length = len(item)

    print("\nList of machines in the detailed job report:")

    for dashes in range(longest_machine_name_length + 11):
        print("-", end="")
    print()

    for machine_num in range(len(machines_list)):
        print("| (" + str(machine_num + 1) + ")", end="")
        if machine_num < 9:
            print(" ", end="")
        print(" - ", end="")
        for spaces in range(int((longest_machine_name_length - len(machines_list[machine_num])) / 2)):
            print(" ", end="")
        print(machines_list[machine_num], end="")
        for spaces in range(int((longest_machine_name_length - len(machines_list[machine_num])) / 2)):
            print(" ", end="")
        if len(machines_list[machine_num]) % 2 != 0:
            print(" ", end="")
        print(" |")

    for dashes in range(longest_machine_name_length + 11):
        print("-", end="")
    print()


# function to print all possible instructions and receive input from the user
# function also performs error checking to ensure user's input is valid
# arguments: none
# returns nothing
def obtain_instruction():
    print("\nWhat would you like to do?")
    print("----------------------------------------")
    print("| 1 - calculate open down time         |")
    print("| 2 - calculate total feeds            |")
    print("| 3 - calculate average setup time     |")
    print("| 4 - break down feeds per day         |")
    print("| 5 - analyze orders by type           |")
    print("| 6 - calculate average run speed      |")
    print("| 7 - sort top three best/worst orders |")
    print("| 8 - return to home page              |")
    print("| 9 - exit                             |")
    print("----------------------------------------")

    user_error = True
    user_input = ""
    while user_error:
        user_input = input("Enter the number associated to your command: ")
        if user_input == "1" or user_input == "2" or user_input == "3" or user_input == "4" or user_input == "5" or user_input == "6" or user_input == "7" or user_input == "8" or user_input == "9":
            user_error = False
        else:
            print("Error: please try again.")

    return int(user_input)


# function to obtain user's sub command after user has entered their main command
# arguments: an option variable that dictates which main command the user entered earlier
# returns the user's sub command option as a string
def obtain_sub_instruction(option):
    error = True
    choice = " "

    while error:
        if option == 1:
            choice = input("Enter (1) to calculate overall ODT for all shifts/crews, (2) to calculate ODT by shift, (3) to calculate ODT by crew or (4) to produce Pareto chart: ")
        elif option == 2:
            choice = input("Enter (1) to calculate total feeds by shift or (2) to total feeds by crew: ")
        elif option == 3:
            choice = input("Enter (1) to calculate general average setup time, (2) to calculate average setup time by shift or (3) to calculate average setup time by crew: ")
        elif option == 4:
            choice = input("Enter (1) to calculate overall daily average feeds for all shifts/crews, (2) to display feeds per day by shift or (3) to display feeds per day by crew: ")
        elif option == 5:
            choice = input("Enter (1) to calculate the average order size, (2) to analyze jobs by the number of colors or (3) to analyze jobs by the number of ups: ")
        elif option == 6:
            choice = input("Enter (1) to calculate the average run speed by shift or (2) to calculate the average run speed by crew: ")
        elif option == 7:
            choice = input("Enter (1) to find the top three best orders or (2) to find the top three worst orders: ")
        else:
            choice = input("Enter (1) to sort by efficiency, (2) to sort by ODT, (3) to sort by setup time or (4) to sort by total time spent on each order: ")

        if option == 1 or option == 8:
            if choice == "1" or choice == "2" or choice == "3" or choice == "4":
                error = False
            else:
                print("Please try again.")
        elif option == 3 or option == 4 or option == 5:
            if choice == "1" or choice == "2" or choice == "3":
                error = False
            else:
                print("Please try again.")
        else:
            if choice == "1" or choice == "2":
                error = False
            else:
                print("Please try again.")

    return choice


# function for obtaining a "yes" or "no" command from the user
# called in any instance where a yes/no answer is required from the user
# arguments: an option integer variable to dictate what phrase we are asking the user
# returns true or false depending on the user's decision + includes error checking
def yes_or_no(option):
    choice = ""

    while True:
        if option == 1: # asking the user if they would like to use AI
            choice = input("Warning: gaps found in Employee Name column. Use AI to compute (y/n)? ")
        elif option == 2: # asking if the user would like to write to Excel
            choice = input("Would you like to write this data to Excel (y/n)? ")
        elif option == 3: # asking if the user would like to see additional information
            choice = input("Would you like to see a breakdown of all rows with potential machine errors (y/n)? ")
        elif option == 4: # asking if the user would like to analyze a specific charge code
            choice = input("Would you like to calculate the ODT by operator for a specific charge code (y/n)? ")
        elif option == 5: # asking if the user would like to remove holidays
            choice = input("Would you like to remove holidays and days off from the table (y/n)? ")
        elif option == 6: # asking if the user would like to calculate the average feeds per day
            choice = input("Would you like to calculate the average feeds per day (y/n)? ")

        if choice == "y" or choice == "Y":
            return True
        elif choice == "n" or choice == "N":
            return False
        else:
            print("Error. Please try again.")


# function to obtain a date as an input from the user and return that same date
# function performs error checking to ensure date is valid and within the detailed job report
# arguments: the detailed job report array and the total number of rows in the detailed job array
# returns the date entered by the user as a string
def obtain_date_string(detailed_job_report):
    result = ""

    error = True
    while error:
        try:
            day, month, year = "", "", ""
            date = input()
            date_split = date.split("/")

            if int(date_split[1]) < 10:
                if (date_split[1])[0] == "0":
                    month = date_split[1]
                else:
                    month = "0" + date_split[1]
            else:
                month = date_split[1]

            if int(date_split[2]) < 10:
                if (date_split[2])[0] == "0":
                    day = date_split[2]
                else:
                    day = "0" + date_split[2]
            else:
                day = date_split[2]

            year = date_split[0]

            result = year + "-" + month + "-" + day

            for item in range(ROWS):
                temp = str(detailed_job_report[item][WORK_DATE_COL_NUM]).split(" ")
                date = temp[0]
                if result == date:
                    error = False
                    break

            if error:
                print("Please enter a valid date found in the detailed job report: ", end="")
        except:
            print("Error, please try again: ", end="")

    return result


# function to obtain a end date as an input from the user and return that same date
# function performs error checking to ensure date is valid and in the future relative to the start date
# arguments: the detailed job report array and the start date as a string
# returns the end date (entered by the user) as a string
def obtain_second_date_string(detailed_job_report_array, first_date_string):
    second_date_string = ""
    error = True
    while error:
        print("Enter the end date (YYYY/MM/DD): ", end="")
        second_date_string = obtain_date_string(detailed_job_report_array)

        if int(second_date_string[5:7]) < int(first_date_string[5:7]) or int(second_date_string[8:10]) < int(first_date_string[8:10]):
            print("Error: end date precedes start date. Please try again.")
        else:
            error = False

    return second_date_string


# function to increment a particular column of a particular row onto a particular variable
# error-checks if the elapsed hours of that row are negative or excessive
# arguments: the detailed job report, the row we are analyzing, the variable we wish to increment on,
# the list of rows with negative elapsed hours, the list of rows with excessive elapsed hours and an option
# 1 = increment on elapsed hours and 2 = increment on total FG quantity
# returns the updated incremented variable
def variable_incrementer(detailed_job_report, row, variable_to_increment, negative_num_rows, excessive_num_rows, option):
    if detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
        append_element_in_array(negative_num_rows, row)
    elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
        append_element_in_array(excessive_num_rows, row)
    else:
        if option == 1:
            variable_to_increment += detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
        else:
            variable_to_increment += detailed_job_report[row][GROSS_FG_QTY_COL_NUM]

    return variable_to_increment


# function to check for gaps in the Employee Name column for a given date frame
# arguments: the detailed job report, the start date as an int and the end date as an int
# returns true or false depending on whether gaps were found or not
def check_for_gaps_in_data(detailed_job_report, start_date_num, end_date_num):
    for row in range(ROWS):
        if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(djr_array[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
            if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                return True

    return False


# algorithm to continue incrementing a counter variable
# (either total machine hours, ODT, total setup time or total feeds)
# when a missing name occurs in the employee name column
# arguments: the detailed job report array, the value of the counter variable so far, crew name from employee column,
# the row being parsed, the array of rows with empty names still undetermined by the algorithm,
# the list of rows with negative elapsed hours
# and an int value (either 1, 2, 3 or 4) to determine which variable to increment
# returns new value of the counter variable
def name_filling_algorithm(detailed_job_report, counter, crew, row, empty_name_rows, negative_num_rows, excessive_num_rows, option):
    if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
        if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
            if option == 1 and detailed_job_report[row][DOWNTIME_COL_NUM] != "Closed Downtime": # increment on total machine hours
                assumed_name = assume_name(detailed_job_report, empty_name_rows, row)
                counter = update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, False)
            elif option == 2 and detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime": # increment on ODT
                assumed_name = assume_name(detailed_job_report, empty_name_rows, row)
                counter = update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, False)
            elif option == 3 and detailed_job_report[row][DOWNTIME_COL_NUM] == "Setup": # increment on total setup time
                assumed_name = assume_name(detailed_job_report, empty_name_rows, row)
                counter = update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, False)
            elif option == 4: # increment on total feeds
                assumed_name = assume_name(detailed_job_report, empty_name_rows, row)
                counter = update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, True)

    return counter


# function to assume name to a particular empty Employee Name cell
# or to include the row as part of rows the AI could not assign a name to
# arguments: the detailed job report as an array,
# the list of empty names the AI could not attribute a name to
# and the row of the detailed job report being analyzed
# returns the assumed name
def assume_name(detailed_job_report, empty_name_rows, row):
    shift_num = detailed_job_report[row][SHIFT_COL_NUM]
    assumed_name = "nan"

    keep_going = True # check rows prior
    iterator = -1
    while keep_going:
        if row + iterator > 0:
            if detailed_job_report[row + iterator][MACHINE_COL_NUM] == MACHINE:
                if detailed_job_report[row + iterator][SHIFT_COL_NUM] == shift_num and str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM]) != "nan" and int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) - int(str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[8:10]) < 2:
                    assumed_name = str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM])
                    keep_going = False
                elif detailed_job_report[row + iterator][SHIFT_COL_NUM] == shift_num and str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM]) == "nan" and int(str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[8:10]) - int(str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[8:10]) < 2:
                    iterator = iterator - 1
                else:
                    keep_going = False
            else:
                iterator = iterator - 1
        else:
            keep_going = False

    if assumed_name == "nan": # check rows after
        keep_going = True
        iterator = 1
        while keep_going:
            if row + iterator < ROWS:
                if detailed_job_report[row + iterator][MACHINE_COL_NUM] == MACHINE:
                    if detailed_job_report[row + iterator][SHIFT_COL_NUM] == shift_num and str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM]) != "nan" and int(str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[8:10]) - int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) < 2:
                        assumed_name = str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM])
                        keep_going = False
                    elif detailed_job_report[row + iterator][SHIFT_COL_NUM] == shift_num and str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM]) == "nan"  and int(str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[8:10]) - int(str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[8:10]) < 2:
                        iterator = iterator + 1
                    else:
                        keep_going = False
                        already_included = False
                        for item in empty_name_rows:
                            if row == item:
                                already_included = True

                        if not already_included:
                            empty_name_rows.append(row)
                else:
                    iterator = iterator + 1
            else:
                keep_going = False

    return assumed_name


# function for adding a value in the detailed job report to the counter if the assumed name from the name-assuming
# algorithm matches the crew being analyzed and the elapsed hours is above zero and below the threshold
# otherwise, the row is appended to its respective list of erroneous rows
# arguments: the detailed job report, the row being analyzed, the name assumed by the name-assuming algorithm,
# the crew being analyzed, the counter variable, the list of rows with negative elapsed hours,
# the list of rows with excessive elapsed hours and whether we are incrementing on the total feeds or not
# returns the new updated counter value
def update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, is_total_feeds):
    if assumed_name == crew:
        if is_total_feeds:
            counter = variable_incrementer(detailed_job_report, row, counter, negative_num_rows, excessive_num_rows, 2)
        else:
            counter = variable_incrementer(detailed_job_report, row, counter, negative_num_rows, excessive_num_rows, 1)

    return counter


# sorting algorithm
# arguments: takes a single one-dimension array
# and sorts its values from smallest to largest
# returns nothing
def sorting_algorithm(array):
    for i in range(len(array) - 1, 0, -1):
        for j in range(0, i):
            if array[j] > array[i]:
                temp = array[j]
                array[j] = array[i]
                array[i] = temp


# function to create array of unique crew members
# arguments: the detailed job report array, the start date as an integer,
# the end date as an integer and the empty list of crew members
# returns nothing
def generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list, rows_with_no_name, use_algo):
    first_row = 0
    last_row = 0
    for row in range(ROWS): # determine which row to start from
        if start_date_num == int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]):
            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                first_row = row
                break
    for row in range(first_row, ROWS): # determine which row to end at
        if end_date_num == int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]):
            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                last_row = row

    for row in range(ROWS): # must parse the entire detailed job report still in case dates within the date frame are dispersed throughout the spreadsheet
        if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) != "nan":
                    append_element_in_array(crews_list, str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]))

    # add crew members if first or last row contain empty cells for "Employee Name" column and user wishes to use AI
    if use_algo:
        if detailed_job_report[first_row][MACHINE_COL_NUM] == MACHINE:
            if str(detailed_job_report[first_row][EMPLOYEE_NAME_COL_NUM]) == "nan": # if first row has blank employee name
                assumed_name = assume_name(detailed_job_report, rows_with_no_name, first_row)
                if assumed_name != "nan":
                    append_element_in_array(crews_list, assumed_name)
            if str(detailed_job_report[last_row][EMPLOYEE_NAME_COL_NUM]) == "nan": # if last row has blank employee name
                assumed_name = assume_name(detailed_job_report, rows_with_no_name, last_row)
                if assumed_name != "nan":
                    append_element_in_array(crews_list, assumed_name)


# function to append an element in an array
# only appends the element if it is unique (if it is not already found in the array)
# arguments: the array itself and the element to be appended
# returns nothing
# called to append rows with negative elapsed hours, rows with excessive elapsed hours,
# crew names to a list of crews and unique orders to a list of all unique orders
def append_element_in_array(array, element):
    already_included = False
    for item in array:
        if item == element:
            already_included = True
            break

    if not already_included:
        array.append(element)


# function for creating the top three orders array
# arguments: the array we are building on, the dictionary containing the data we need,
# an option which represents if we want the top three best or worst orders
# and a sub option that represents the data the user wishes to see
# returns nothing but modifies the array passed as an argument
def build_top_three_orders_array(array, dictionary, option, sub_option):
    for row in range(1, 4):
        key_to_pop = 0
        for key in dictionary.keys():
            if (option == 1 and sub_option == 1) or (option == 2 and sub_option != 1):
                if dictionary[key] > array[row][1]:
                    array[row][0], array[row][1] = key, dictionary[key]
                    key_to_pop = key
            elif (option == 1 and sub_option != 1) or (option == 2 and sub_option == 1):
                if dictionary[key] < array[row][1]:
                    array[row][0], array[row][1] = key, dictionary[key]
                    key_to_pop = key
        dictionary.pop(key_to_pop)


# function to obtain Excel spreadsheet name and write data from user's previous command to said spreadsheet
# arguments: array containing the data to write to spreadsheet and its length
# returns nothing
def write_to_excel(array, length_of_array):
    spreadsheet_name = ""

    error = True
    choice = ""
    while error:
        choice = input("Enter (1) if you would like to write the data to a pre-existing spreadsheet or (2) if you would like to create a new spreadsheet: ")
        if choice == "1" or choice == "2":
            error = False
        else:
            print("Please try again.")

    if choice == "1": # user wants to write to pre-existing spreadsheet
        file_found = False
        while not file_found:
            spreadsheet_name = input("Enter the file name of the destination spreadsheet: ")
            spreadsheet_name_split = spreadsheet_name.split(".")

            try:
                contains_file_type = False
                if spreadsheet_name_split[len(spreadsheet_name_split) - 1] == "xlsx":
                    contains_file_type = True

                if not contains_file_type:
                    spreadsheet_name = spreadsheet_name + ".xlsx"

                # if spreadsheet can be read, then spreadsheet_name is valid
                spreadsheet_dataframe = pd.read_excel(spreadsheet_name)
                spreadsheet_array = spreadsheet_dataframe.to_numpy()

                file_found = True
                print("File found.")
            except:
                print("Error: no such file found in current directory")
                print("Please try again")
    else: # user wants to create new spreadsheet
        spreadsheet_name = validate_spreadsheet_name()
        spreadsheet_name_split = spreadsheet_name.split(".")

        contains_file_type = False
        for item in range(len(spreadsheet_name_split)):
            if spreadsheet_name_split[item] == "xlsx":
                contains_file_type = True

        if not contains_file_type:
            spreadsheet_name = spreadsheet_name + ".xlsx"

    # clear any rows with irrelevant data
    rows_to_delete = [] # list to remember which rows to delete
    counter = 0
    for row in range(length_of_array):
        if array[row][0] == 0 and array[row][1] == 0:
            rows_to_delete.append(row)
    for row in rows_to_delete:
        array = np.delete(array, row - counter, axis=0)
        # size of array decreases with every deletion
        counter = counter + 1
        length_of_array = length_of_array - 1

    charge_code_dataframe = pd.DataFrame(array[0:length_of_array]) # convert array containing data into a dataframe
    charge_code_dataframe.to_excel(spreadsheet_name, index=False) # write dataframe to Excel spreadsheet
    print("Data successfully copied to spreadsheet.")


# function for ensuring the name of a newly created spreadsheet follows Microsoft's naming conventions
# arguments: none
# returns the spreadsheet name once the user has entered an acceptable file name
def validate_spreadsheet_name():
    unacceptable_characters = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    while True:
        valid_name = True
        name = input("Enter the name for the new spreadsheet: ")

        if name and name.strip() and name[0] != " ": # check for exceptions in input
            print_error_phrase = False
            for item in unacceptable_characters:
                for char in name:
                    if char == item:
                        if not print_error_phrase:
                            print("Error. File name cannot contain the following character(s): ", end="")
                            print_error_phrase = True
                        print(item + " ", end="")
                        valid_name = False
                        break

            if valid_name:
                return name
            else:
                print()
        else:
            print("Error. Please try again.")


# function for printing the resulting table from any command
# arguments: the table to print as an array
# returns nothing
def print_table(array_to_print):
    num_rows, num_cols = np.shape(array_to_print)

    column_headers = []
    for col in range(num_cols):
        column_headers.append(array_to_print[0][col])

    # print table header
    print_dashes(array_to_print, column_headers, num_rows, num_cols)
    print("| ", end="")
    for col in range(num_cols):
        longest_element_length = find_length_of_longest_element(array_to_print, col, num_rows)

        for space in range(abs(int((longest_element_length - len(str(column_headers[col]))) / 2))):
            print(" ", end="")
        print(column_headers[col], end="")
        for space in range(abs(int((longest_element_length - len(str(column_headers[col]))) / 2))):
            print(" ", end="")
        if abs(longest_element_length - len(str(column_headers[col]))) % 2 != 0:
            print(" ", end="")
        print(" | ", end="")

        if col == len(column_headers) - 1:
            print()
    print_dashes(array_to_print, column_headers, num_rows, num_cols)

    # print rest of table
    for row in range(num_rows - 1):
        print("| ", end="")
        for col in range(num_cols):
            longest_element_length = find_length_of_longest_element(array_to_print, col, num_rows)

            if longest_element_length > len(str(column_headers[col])):
                print_table_element(array_to_print[row + 1][col], longest_element_length)
            else:
                print_table_element(array_to_print[row + 1][col], len(str(column_headers[col])))
            print(" | ", end="")

            if col == num_cols - 1:
                print()
    print_dashes(array_to_print, column_headers, num_rows, num_cols)


# function for finding the length of the longest element in a given column of a table
# arguments: the table to print as an array, the column in question and the total number of rows in the table
# returns the length of the longest element found in the column
def find_length_of_longest_element(array_to_print, col, num_rows):
    longest_element_length = 0
    for row in range(num_rows):
        if len(str(array_to_print[row][col])) > longest_element_length:
            longest_element_length = len(str(array_to_print[row][col]))

    return longest_element_length


# function for printing dashes for the resulting table of any command
# arguments: the table to print as an array, the list of all column headers,
# the total number of rows in the table and the total number of columns in the table
# returns nothing
def print_dashes(array_to_print, column_headers, num_rows, num_cols):
    for col in range(num_cols):
        longest_element_length = find_length_of_longest_element(array_to_print, col, num_rows)

        if longest_element_length > len(str(column_headers[col])):
            for dash in range(longest_element_length + 3):
                print("-", end="")
        else:
            for dash in range(len(str(column_headers[col])) + 3):
                print("-", end="")

        if col == len(column_headers) - 1:
            print("-")


# function to print an element in a cell of a table according to a specific space to fill
# arguments: the element to print, either as a string or an integer
# and the number of total spaces to fill as an int
# returns nothing
def print_table_element(element, length_to_fill):
    for space in range(int((length_to_fill - len(str(element))) / 2)):
        print(" ", end="")
    print(element, end="")
    for space in range(int((length_to_fill - len(str(element))) / 2)):
        print(" ", end="")

    if len(str(element)) % 2 != 0 and length_to_fill % 2 == 0:
        print(" ", end="")
    elif len(str(element)) % 2 == 0 and length_to_fill % 2 != 0:
        print(" ", end="")


# function to print all incorrect numbers in a given list
# arguments: the number of numbers in an array, the array itself
# and an option to indicate what sort of wrong numbers we are printing
# 1 = negative elapsed hours and 2 = excessive elapsed hours
# returns nothing
def print_wrong_nums_list(num_numbers, wrong_nums_list, option):
    if option == 1:
        print("Row(s) with negative elapsed hours:\n")
    else:
        print("Row(s) with excessive elapsed hours (greater than " + str(EXCESSIVE_THRESHOLD) + "):\n")

    counter = 0
    for row in range(num_numbers):
        print(wrong_nums_list[row] + 2, end=" ")
        counter = counter + 1

        if counter > 15:
            print("")
            counter = 0

    if counter != 0:
        print()
    print("\nThese row(s) were omitted from calculations")
    print("---------------------------------------------")


# function for printing both the list of rows with negative elapsed hours
# and the list of rows with excessive elapsed hours
# arguments: the list of rows with negative elapsed hours and
# the list of rows with excessive elapsed hours
# used in all situations except when user requests information by crew
# returns nothing
def print_incorrect_hours(negative_num_rows, excessive_num_rows):
    print("\n")

    if len(negative_num_rows) > 0:
        sorting_algorithm(negative_num_rows)
        print_wrong_nums_list(len(negative_num_rows), negative_num_rows, 1)

    if len(excessive_num_rows) > 0:
        sorting_algorithm(excessive_num_rows)
        print_wrong_nums_list(len(excessive_num_rows), excessive_num_rows, 2)


# function for printing all the rows in which AI was unable to attribute a name
# arguments: the completed list of rows with no name
# returns nothing
def print_rows_with_no_name(rows_with_no_name):
    if len(rows_with_no_name) > 0:
        print("AI was unable to attribute names to the following row(s) due to insufficient information:\n")
        counter = 0
        for item in rows_with_no_name:
            counter = counter + 1

            if counter < 19:
                print(str(item + 2) + " ", end="")
            else:
                print(str(item + 2) + " ")
                counter = 0

        if counter != 0:
            print("")
        print("\nThese row(s) were omitted from calculations")


# function for displaying both the list of rows where no crew name could be attributed by the AI,
# the list of rows with negative elapsed hours and the list of excessive elapsed hours
# arguments: whether the algorithm was used, the list of rows with no names,
# the list of rows with negative elapsed hours and the list of rows with excessive elapsed hours
def print_additional_info(algorithm_used, rows_with_no_name, negative_num_rows, excessive_num_rows):
    print("\n")

    # print list of rows where no name was attributed
    if algorithm_used:
        sorting_algorithm(rows_with_no_name)
        print_rows_with_no_name(rows_with_no_name)
        if len(negative_num_rows) <= 0 and len(excessive_num_rows) <= 0 and len(rows_with_no_name) > 0:
            print("---------------------------------------------")

    # print list of rows with negative elapsed hours
    if len(negative_num_rows) > 0:
        sorting_algorithm(negative_num_rows)
        if len(rows_with_no_name) > 0:
            print("---------------------------------------------")
        print_wrong_nums_list(len(negative_num_rows), negative_num_rows, 1)

    # print list of rows with excessive elapsed hours
    if len(excessive_num_rows) > 0:
        sorting_algorithm(excessive_num_rows)
        if len(rows_with_no_name) > 0 and len(negative_num_rows) == 0:
            print("---------------------------------------------")
        print_wrong_nums_list(len(excessive_num_rows), excessive_num_rows, 2)


# function for converting an integer corresponding to a date
# to its string equivalent
# arguments: the date as an integer
# returns the date as a string
def convert_date_int_to_string(date_num):
    day = str(date_num)[6:8]

    if int(str(date_num)[4:6]) == 1:
        month = "Jan"
    elif int(str(date_num)[4:6]) == 2:
        month = "Feb"
    elif int(str(date_num)[4:6]) == 3:
        month = "Mar"
    elif int(str(date_num)[4:6]) == 4:
        month = "Apr"
    elif int(str(date_num)[4:6]) == 5:
        month = "May"
    elif int(str(date_num)[4:6]) == 6:
        month = "Jun"
    elif int(str(date_num)[4:6]) == 7:
        month = "Jul"
    elif int(str(date_num)[4:6]) == 8:
        month = "Aug"
    elif int(str(date_num)[4:6]) == 9:
        month = "Sep"
    elif int(str(date_num)[4:6]) == 10:
        month = "Oct"
    elif int(str(date_num)[4:6]) == 11:
        month = "Nov"
    else:
        month = "Dec"

    year = str(date_num)[0:4]

    return day + "-" + month + "-" + year


# function for calculating the average feeds by shift
# arguments: the feeds per day by shift array and the length of said array
# returns nothing
def calculate_average_feeds_by_shift(feeds_per_day_array, len_feeds_per_day_array):
    resulting_average_table = [[0 for i in range(3)] for j in range(4)]

    resulting_average_table[0][0], resulting_average_table[0][1], resulting_average_table[0][2] = "Shift", "Average Feeds per Day", "Opportunity (%)"
    resulting_average_table[1][0], resulting_average_table[2][0], resulting_average_table[3][0] = 1, 2, 3

    # calculate average feeds per day for each shift
    for index in range(3):
        total_for_shift = 0
        denominator = 0
        for row in range(len_feeds_per_day_array - 1):
            if feeds_per_day_array[row + 1][index + 1] != 0:
                total_for_shift = total_for_shift + feeds_per_day_array[row + 1][index + 1]
                denominator = denominator + 1

        if denominator != 0:
            resulting_average_table[index + 1][1] = total_for_shift / denominator
        else:
            resulting_average_table[index + 1][1] = "N/A"

    # calculate opportunity for each shift
    highest_average = 0  # first calculate the highest average to set as the datum
    for index in range(3):
        if resulting_average_table[index + 1][1] != "N/A":
            if resulting_average_table[index + 1][1] > highest_average:
                highest_average = resulting_average_table[index + 1][1]

    for index in range(3):
        if resulting_average_table[index + 1][1] != "N/A":
            resulting_average_table[index + 1][2] = ((highest_average - resulting_average_table[index + 1][1]) / highest_average) * 100
        else:
            resulting_average_table[index + 1][2] = "N/A"

    print_table(resulting_average_table)

    if yes_or_no(2):
        write_to_excel(resulting_average_table, len(resulting_average_table))


# function for calculating the average feeds by crew
# arguments: the array of feeds per day by crew, the length of the array of feeds per day by crew,
# and the list of all crew members considered
# returns nothing
def calculate_average_feeds_by_crew(feeds_per_day_array, len_feeds_per_day_array, list_of_crews):
    resulting_average_table = [[0 for i in range(3)] for j in range(len(list_of_crews) + 1)]

    # print header and crew names first
    resulting_average_table[0][0], resulting_average_table[0][1], resulting_average_table[0][2] = "Crew", "Average Feeds per Day", "Opportunity (%)"
    for index in range(len(list_of_crews)):
        resulting_average_table[index + 1][0] = list_of_crews[index]

    # calculate average feeds per day for each crew
    for index in range(len(list_of_crews)):
        total_for_crew = 0
        denominator = 0
        for row in range(len_feeds_per_day_array - 1):
            if feeds_per_day_array[row + 1][index + 1] != 0:
                total_for_crew = total_for_crew + feeds_per_day_array[row + 1][index + 1]
                denominator = denominator + 1

        if denominator != 0:
            resulting_average_table[index + 1][1] = total_for_crew / denominator
        else:
            resulting_average_table[index + 1][1] = "N/A"

    # calculate opportunity for each crew
    highest_average = 0 # first calculate the highest average to set as the datum
    for index in range(len(list_of_crews)):
        if resulting_average_table[index + 1][1] != "N/A":
            if resulting_average_table[index + 1][1] > highest_average:
                highest_average = resulting_average_table[index + 1][1]

    for index in range(len(list_of_crews)):
        if resulting_average_table[index + 1][1] != "N/A":
            resulting_average_table[index + 1][2] = ((highest_average - resulting_average_table[index + 1][1]) / highest_average) * 100
        else:
            resulting_average_table[index + 1][2] = "N/A"

    print_table(resulting_average_table)

    if yes_or_no(2):
        write_to_excel(resulting_average_table, len(resulting_average_table))


# function for calculating the total ODT in hours by crew based on a selected charge code
# arguments: the pareto chart array, the detailed job report, the start and end dates as integers,
# the list of negative num rows and the list of excessive num rows
# returns nothing
def calculate_ODT_by_crew(charge_code_array, detailed_job_report, start_date_num, end_date_num):
    # obtain which charge code the user would like to analyze
    charge_code = ""
    charge_code_error = True
    while charge_code_error:
        charge_code = input("Enter the name of the charge code you would like to consider: ")

        for code in range(len(charge_code_array) - 1):
            if charge_code_array[code + 1][0] == charge_code:
                charge_code_error = False

        if charge_code_error:
            print("Error. Could not find charge code. Please try again.")

    # check if there are gaps in data
    gap_name = check_for_gaps_in_data(detailed_job_report, start_date_num, end_date_num)

    # ask user if they want to use name-filling algorithm
    use_algo = False
    if gap_name:
        use_algo = yes_or_no(1)
    rows_with_no_name = []

    # generate list of crews
    crews_list = []
    generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list, rows_with_no_name, use_algo)

    if len(crews_list) != 0:
        crew_ODT_by_charge_code_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
        crew_ODT_by_charge_code_array[0][0], crew_ODT_by_charge_code_array[0][1] = "Crew", (charge_code + " (hours elapsed)")

        for index in range(len(crews_list)): # write all crew names to resulting table
            crew_ODT_by_charge_code_array[index + 1][0] = crews_list[index]

        for index in range(len(crews_list)): # compute ODT elapsed hours for each crew member based on charge code
            elapsed_hours_for_crew = 0
            for row in range(len(detailed_job_report)):
                if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE and detailed_job_report[row][CHARGE_CODE_COL_NUM] == charge_code:
                    if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if crews_list[index] == detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]:
                            if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                elapsed_hours_for_crew += detailed_job_report[row][ELAPSED_HOURS_COL_NUM]

                        elif use_algo and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                            if assume_name(detailed_job_report, rows_with_no_name, row) == crews_list[index]:
                                if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                    elapsed_hours_for_crew += detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                            if assume_name(detailed_job_report, rows_with_no_name, row) == "nan":
                                append_element_in_array(rows_with_no_name, row)

            crew_ODT_by_charge_code_array[index + 1][1] = elapsed_hours_for_crew

        print_table(crew_ODT_by_charge_code_array)

        if use_algo and len(rows_with_no_name) > 0:
            if yes_or_no(3):
                print("\n")
                print_rows_with_no_name(rows_with_no_name)
                print("---------------------------------------------")

        if yes_or_no(2):
            write_to_excel(crew_ODT_by_charge_code_array, len(crew_ODT_by_charge_code_array))
    else:
        print("There is nothing to show here")


# function to keep adding unique orders to the unique orders list and to keep incrementing on the total quantity
# used when computing the average order quantity for a given date frame
# arguments: the detailed job report, the row being parsed, the total quantity so far
# and the list of unique orders so far
# returns the new total quantity
def incrementing_algo(detailed_job_report, row, total_quantity, unique_orders_list):
    unique = True
    for item in unique_orders_list:
        unique = True
        if detailed_job_report[row][ORDER_NUM_COL_NUM] == item:
            unique = False
            break
    if unique:
        unique_orders_list.append(detailed_job_report[row][ORDER_NUM_COL_NUM])
        total_quantity = total_quantity + detailed_job_report[row][ORDER_QTY_COL_NUM]

    return total_quantity


# function to display ODT by either shift or crew
# arguments: the detailed job report array, which type of ODT the user wants,
# the start date and the end date
# returns nothing
def display_ODT(detailed_job_report, user_option, start_date, end_date):
    negative_num_rows = []
    excessive_num_rows = []
    start_date_num = int(start_date[0:4] + start_date[5:7] + start_date[8:10])
    end_date_num = int(end_date[0:4] + end_date[5:7] + end_date[8:10])

    if user_option == "1": # user wants overall ODT for all shifts/crews
        total_machine_hours = 0
        total_ODT = 0

        for row in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                    if detailed_job_report[row][DOWNTIME_COL_NUM] != "Closed Downtime":
                        total_machine_hours = variable_incrementer(detailed_job_report, row, total_machine_hours, negative_num_rows, excessive_num_rows, 1)
                    if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime":
                        total_ODT = variable_incrementer(detailed_job_report, row, total_ODT, negative_num_rows, excessive_num_rows, 1)

        print("Total machine hours: " + str(total_machine_hours))
        print("Total ODT hours: " + str(total_ODT))
        if total_ODT != 0:
            print("ODT %: " + str(total_ODT / total_machine_hours * 100))
        else:
            print("ODT%: N/A")

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            print()
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

    elif user_option == "2": # user wants ODT by shift
        # create & initialize array to hold data in case user wants to write data to an Excel spreadsheet
        ODT_by_shift_array = [[0 for x in range(2)] for y in range(4)]
        ODT_by_shift_array[0][0], ODT_by_shift_array[0][1] = "Shift", "ODT (%)"

        for shift in range(3):
            total_machine_hours = 0
            total_ODT = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                        if detailed_job_report[row][SHIFT_COL_NUM] == shift + 1:
                            if detailed_job_report[row][DOWNTIME_COL_NUM] != "Closed Downtime":
                                total_machine_hours = variable_incrementer(detailed_job_report, row, total_machine_hours, negative_num_rows, excessive_num_rows, 1)
                            if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime":
                                total_ODT = variable_incrementer(detailed_job_report, row, total_ODT, negative_num_rows, excessive_num_rows, 1)

            if total_machine_hours != 0:
                result = (total_ODT/total_machine_hours) * 100
                ODT_by_shift_array[shift + 1][0], ODT_by_shift_array[shift + 1][1] = shift + 1, result
            else:
                ODT_by_shift_array[shift + 1][0], ODT_by_shift_array[shift + 1][1] = shift + 1, "N/A"

        print_table(ODT_by_shift_array)

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

        if yes_or_no(2):
            write_to_excel(ODT_by_shift_array, len(ODT_by_shift_array))

    elif user_option == "3": # user wants ODT by crew
        # check if there are gaps in data
        gap_name = check_for_gaps_in_data(detailed_job_report, start_date_num, end_date_num)

        # ask user if they want to use name-filling algorithm
        use_algo = False
        if gap_name:
            use_algo = yes_or_no(1)
        rows_with_no_name = []

        # create list of crews
        crews_list = []
        generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list, rows_with_no_name, use_algo)

        if len(crews_list) != 0:
            # create & initialize array to hold data in case user wants to write data to an Excel spreadsheet
            ODT_by_crew_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
            ODT_by_crew_array[0][0], ODT_by_crew_array[0][1] = "Crew", "ODT (%)"

            # longest_name_len = print_crew_header(ODT_by_crew_array, crews_list, 1)

            counter = 1
            for crew in crews_list:
                total_machine_hours = 0
                total_ODT = 0
                for row in range(ROWS):
                    if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                            if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew:
                                if detailed_job_report[row][DOWNTIME_COL_NUM] != "Closed Downtime":
                                    total_machine_hours = variable_incrementer(detailed_job_report, row, total_machine_hours, negative_num_rows, excessive_num_rows, 1)
                                if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime":
                                    total_ODT = variable_incrementer(detailed_job_report, row, total_ODT, negative_num_rows, excessive_num_rows, 1)

                            elif use_algo:
                                total_machine_hours = name_filling_algorithm(detailed_job_report, total_machine_hours, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 1)
                                total_ODT = name_filling_algorithm(detailed_job_report, total_ODT, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 2)

                if total_machine_hours > 0:
                    ODT_by_crew_array[counter][0], ODT_by_crew_array[counter][1] = crew, (total_ODT/total_machine_hours) * 100
                else:
                    ODT_by_crew_array[counter][0], ODT_by_crew_array[counter][1] = crew, "N/A"
                counter = counter + 1

            # print_rest_of_table(ODT_by_crew_array, longest_name_len)
            print_table(ODT_by_crew_array)

            if len(rows_with_no_name) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    print_additional_info(use_algo, rows_with_no_name, negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(ODT_by_crew_array, len(ODT_by_crew_array))
        else:
            print("There is nothing to show here")

    else: # user wants Pareto chart
        # create list of charge codes
        charge_code_list = []
        for row_iterator in range(ROWS):
            if detailed_job_report[row_iterator][MACHINE_COL_NUM] == MACHINE:
                already_included = False
                for charge_code in charge_code_list:
                    if detailed_job_report[row_iterator][CHARGE_CODE_COL_NUM] == charge_code:
                        already_included = True
                        break

                if not already_included and detailed_job_report[row_iterator][CHARGE_CODE_COL_NUM] != "RUN" and detailed_job_report[row_iterator][CHARGE_CODE_COL_NUM] != "SET UP":
                    charge_code_list.append(detailed_job_report[row_iterator][CHARGE_CODE_COL_NUM])

        if len(charge_code_list) != 0:
            # temporary dictionary to contain each charge code and their corresponding ODT
            charge_code_dict = {}
            for charge_code in charge_code_list:
                charge_code_dict[charge_code] = 0
            for row_num in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row_num][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_num][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_num][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row_num][MACHINE_COL_NUM] == MACHINE:
                        charge_code_found = False
                        for charge_code in charge_code_list:
                            if detailed_job_report[row_num][CHARGE_CODE_COL_NUM] == charge_code:
                                charge_code_found = True
                                break
                        if charge_code_found:
                            charge_code_dict[detailed_job_report[row_num][CHARGE_CODE_COL_NUM]] = variable_incrementer(detailed_job_report, row_num, charge_code_dict[detailed_job_report[row_num][CHARGE_CODE_COL_NUM]], negative_num_rows, excessive_num_rows, 1)
            num_rows = 0
            for key in charge_code_dict.keys():
                if charge_code_dict[key] > 0:
                    num_rows = num_rows + 1

            charge_code_array = [[0 for x in range(2)] for y in range(num_rows + 1)]
            charge_code_array[0][0], charge_code_array[0][1] = "Charge Code", "ODT (hours)"

            # sort charge_code_array from largest ODT to smallest
            counter = 1
            while len(charge_code_dict) > 0:
                largest_ODT_key = list(charge_code_dict.keys())[0]
                for key in charge_code_dict.keys():
                    if charge_code_dict[key] > charge_code_dict[largest_ODT_key]:
                        largest_ODT_key = key

                if charge_code_dict[largest_ODT_key] > 0:
                    charge_code_array[counter][0], charge_code_array[counter][1] = largest_ODT_key, charge_code_dict[largest_ODT_key]

                charge_code_dict.pop(largest_ODT_key)
                counter = counter + 1

            print_table(charge_code_array)

            if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    print_incorrect_hours(negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(charge_code_array, len(charge_code_array))

            if yes_or_no(4):
                calculate_ODT_by_crew(charge_code_array, detailed_job_report, start_date_num, end_date_num)
        else:
            print("There is nothing to show here")


# function to display total feeds either by shift or by crew
# arguments: the detailed job report as an array, whether the user wants
# total feeds by shift or by crew, the start date as an integer and
# the end date as an integer
# returns nothing
def display_total_feeds(detailed_job_report, user_choice, start_date_num, end_date_num):
    negative_num_rows = []
    excessive_num_rows = []

    if user_choice == "1": # user wants total feeds by shift
        total_feeds_by_shift_array = [[0 for x in range(2)] for y in range(4)]
        total_feeds_by_shift_array[0][0], total_feeds_by_shift_array[0][1] = "Shift", "Total Feeds"

        shift_counter = 1
        for shift in range(3):
            total_feeds = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                        if detailed_job_report[row][SHIFT_COL_NUM] == shift + 1:
                            total_feeds = variable_incrementer(detailed_job_report, row, total_feeds, negative_num_rows, excessive_num_rows, 2)

            if total_feeds > 0:
                total_feeds_by_shift_array[shift_counter][0], total_feeds_by_shift_array[shift_counter][1] = shift + 1, total_feeds
            shift_counter = shift_counter + 1

        print_table(total_feeds_by_shift_array)

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

        if yes_or_no(2):
            write_to_excel(total_feeds_by_shift_array, len(total_feeds_by_shift_array))

    else: # user wants total feeds by crew
        # check for gaps in Employee Name column
        gap_name = check_for_gaps_in_data(detailed_job_report, start_date_num, end_date_num)

        # ask if user would like to use AI to compute
        use_algo = False
        if gap_name:
            use_algo = yes_or_no(1)
        empty_name_rows = []

        crews_list = []
        generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list, empty_name_rows, use_algo)

        if len(crews_list) != 0:
            total_feeds_by_crew_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
            total_feeds_by_crew_array[0][0], total_feeds_by_crew_array[0][1] = "Crew", "Total Feeds"

            counter = 1
            for crew in crews_list:
                total_feeds = 0
                for row in range(ROWS):
                    if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                            if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew:
                                total_feeds = variable_incrementer(detailed_job_report, row, total_feeds, negative_num_rows, excessive_num_rows, 2)

                            elif use_algo:
                                total_feeds = name_filling_algorithm(detailed_job_report, total_feeds, crew, row, empty_name_rows, negative_num_rows, excessive_num_rows, 4)

                total_feeds_by_crew_array[counter][0], total_feeds_by_crew_array[counter][1] = crew, total_feeds
                counter = counter + 1

            print_table(total_feeds_by_crew_array)

            if len(empty_name_rows) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    print_additional_info(use_algo, empty_name_rows, negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(total_feeds_by_crew_array, len(total_feeds_by_crew_array))
        else:
            print("There is nothing to show here")


# function to display average setup time either by shift or by crew
# arguments: the detailed job report as an array, whether the user wants
# average setup time by shift or by crew, the start date as an integer and
# the end date as an integer
# returns nothing
def display_average_setup_time(detailed_job_report, user_choice, start_date_num, end_date_num):
    negative_num_rows = []  # list to track rows with negative elapsed hours
    excessive_num_rows = [] # list to track rows with excessive elapsed hours

    if user_choice == "1":  # user wants general average setup time
        total_elapsed_hours = 0  # counter for total elapsed hours
        unique_orders_list = []

        for elapsed_hours in range(ROWS):
            if detailed_job_report[elapsed_hours][MACHINE_COL_NUM] == MACHINE:
                if start_date_num <= int(str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[elapsed_hours][DOWNTIME_COL_NUM] == "Setup":
                        total_elapsed_hours = variable_incrementer(detailed_job_report, elapsed_hours, total_elapsed_hours, negative_num_rows, excessive_num_rows, 1)
                        if 0 <= detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                            append_element_in_array(unique_orders_list, detailed_job_report[elapsed_hours][ORDER_NUM_COL_NUM])

        print("\nTotal Elapsed Hours: " + str(total_elapsed_hours))
        print("Total Unique Orders: " + str(len(unique_orders_list)))

        # final calculation
        if len(unique_orders_list) != 0:
            average_setup_time = total_elapsed_hours / len(unique_orders_list)
            print("Average Setup Time: " + str(average_setup_time * 60) + " minutes")  # display average setup time in minutes
        else:
            print("Average Setup Time: N/A")

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            print()
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

    elif user_choice == "2": # user wants average setup time by shift
        average_setup_time_by_shift_array = [[0 for x in range(3)] for y in range(4)]
        average_setup_time_by_shift_array[0][0], average_setup_time_by_shift_array[0][1], average_setup_time_by_shift_array[0][2] = "Shift", "Average Setup Time (minutes)", "Total Number of Setups"

        counter = 1
        for shift in range(3):
            total_elapsed_hours = 0  # counter for total elapsed hours
            unique_orders_list = []  # list to track all unique orders

            for elapsed_hours in range(ROWS):
                if detailed_job_report[elapsed_hours][DOWNTIME_COL_NUM] == "Setup":
                    if start_date_num <= int(str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if detailed_job_report[elapsed_hours][MACHINE_COL_NUM] == MACHINE:
                            if detailed_job_report[elapsed_hours][SHIFT_COL_NUM] == shift + 1:
                                total_elapsed_hours = variable_incrementer(detailed_job_report, elapsed_hours, total_elapsed_hours, negative_num_rows, excessive_num_rows, 1)
                                if 0 <= detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                    append_element_in_array(unique_orders_list, detailed_job_report[elapsed_hours][ORDER_NUM_COL_NUM])

            # final calculation
            if len(unique_orders_list) != 0:
                # print(str(total_elapsed_hours) + " " + str(total_unique_orders))
                average_setup_time = total_elapsed_hours / len(unique_orders_list) * 60
                average_setup_time_by_shift_array[counter][0], average_setup_time_by_shift_array[counter][1], average_setup_time_by_shift_array[counter][2] = shift + 1, average_setup_time, len(unique_orders_list)
            else:
                average_setup_time_by_shift_array[counter][0], average_setup_time_by_shift_array[counter][1], average_setup_time_by_shift_array[counter][2] = shift + 1, "N/A", "N/A"
            counter = counter + 1

        print_table(average_setup_time_by_shift_array)

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

        if yes_or_no(2):
            write_to_excel(average_setup_time_by_shift_array, len(average_setup_time_by_shift_array))

    else:  # user wants average setup time by crew
        # check if there are gaps in data
        gap_name = check_for_gaps_in_data(detailed_job_report, start_date_num, end_date_num)

        # ask user if they want to use name-filling algorithm
        use_algo = False
        if gap_name:
            use_algo = yes_or_no(1)
        rows_with_no_name = []

        crews_list = []
        generate_crews_list(djr_array, start_date_num, end_date_num, crews_list, rows_with_no_name, use_algo)

        if len(crews_list) != 0:
            average_setup_time_by_crew_array = [[0 for x in range(3)] for y in range(len(crews_list) + 1)]
            average_setup_time_by_crew_array[0][0], average_setup_time_by_crew_array[0][1], average_setup_time_by_crew_array[0][2] = "Crew", "Average Setup Time (minutes)", "Total Number of Setups"

            counter = 1
            for crew in crews_list:
                total_elapsed_hours = 0  # counter for total elapsed hours
                unique_orders_list = []  # list to track all unique orders

                for elapsed_hours in range(ROWS):
                    if detailed_job_report[elapsed_hours][DOWNTIME_COL_NUM] == "Setup":
                        if start_date_num <= int(str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                            if detailed_job_report[elapsed_hours][MACHINE_COL_NUM] == MACHINE:
                                if detailed_job_report[elapsed_hours][EMPLOYEE_NAME_COL_NUM] == crew:
                                    total_elapsed_hours = variable_incrementer(detailed_job_report, elapsed_hours, total_elapsed_hours, negative_num_rows, excessive_num_rows, 1)
                                    if 0 <= detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                        append_element_in_array(unique_orders_list, detailed_job_report[elapsed_hours][ORDER_NUM_COL_NUM]) # append order if unique
                                elif use_algo:
                                    # increment elapsed hours
                                    total_elapsed_hours = name_filling_algorithm(detailed_job_report, total_elapsed_hours, crew, elapsed_hours, rows_with_no_name, negative_num_rows, excessive_num_rows, 3)

                                    # append order if unique
                                    if str(detailed_job_report[elapsed_hours][EMPLOYEE_NAME_COL_NUM]) == "nan":
                                        if 0 <= detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                            if assume_name(djr_array, rows_with_no_name, elapsed_hours) == crew:
                                                append_element_in_array(unique_orders_list, detailed_job_report[elapsed_hours][ORDER_NUM_COL_NUM])

                # final calculation
                if len(unique_orders_list) != 0:
                    # print(str(total_elapsed_hours) + " " + str(total_unique_orders))
                    average_setup_time = total_elapsed_hours / len(unique_orders_list) * 60
                    average_setup_time_by_crew_array[counter][0], average_setup_time_by_crew_array[counter][1], average_setup_time_by_crew_array[counter][2] = crew, average_setup_time, len(unique_orders_list)
                else:
                    average_setup_time_by_crew_array[counter][0], average_setup_time_by_crew_array[counter][1], average_setup_time_by_crew_array[counter][2] = crew, "N/A", "N/A"
                counter = counter + 1

            print_table(average_setup_time_by_crew_array)

            if len(rows_with_no_name) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    print_additional_info(use_algo, rows_with_no_name, negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(average_setup_time_by_crew_array, len(average_setup_time_by_crew_array))
        else:
            print("There is nothing to show here")


# function to display chart for feeds per day either by shift or by crew
# arguments: the detailed job report as an array, whether the user wants
# feeds per day by shift or by crew, the start date as an integer and
# the end date as an integer
# returns nothing
def display_daily_feeds(detailed_job_report, user_choice, start_date_num, end_date_num):
    negative_num_rows = []
    excessive_num_rows = []

    if user_choice == "1": # user wants average daily feeds for all three shifts/crews
        total_feeds = 0
        unique_days = []
        for row in range(ROWS):
            if str(start_date_num) <= str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10] <= str(end_date_num):
                if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                    total_feeds = variable_incrementer(detailed_job_report, row, total_feeds, negative_num_rows, excessive_num_rows, 2)
                    if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                        append_element_in_array(unique_days, detailed_job_report[row][WORK_DATE_COL_NUM])

        print("Total feeds: " + str(total_feeds))
        print("Total work days: " + str(len(unique_days)))
        if len(unique_days) != 0:
            print("Average daily feeds for all shifts and crews: " + str(total_feeds / len(unique_days)))
        else:
            print("Average daily feeds for all shifts and crews unavailable for the given date frame.")

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            print()
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

    elif user_choice == "2": # user wants to display feeds per day by shift
        # create header for resulting table
        resulting_table = np.array([[0 for x in range(4)] for y in range(int(end_date_num - start_date_num + 2))], dtype='object')
        resulting_table[0][0], resulting_table[0][1], resulting_table[0][2], resulting_table[0][3] = "Work Date", "Shift 1", "Shift 2", "Shift 3"

        # generate rest of table
        for row_table in range(len(resulting_table) - 1):
            resulting_table[row_table + 1][0] = convert_date_int_to_string(start_date_num + row_table)
        for row_djr in range(ROWS):
            if detailed_job_report[row_djr][MACHINE_COL_NUM] == MACHINE:
                for row_table in range(len(resulting_table) - 1):
                    if str(start_date_num + row_table) == str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[8:10]:
                        for shift in range(3):
                            if detailed_job_report[row_djr][SHIFT_COL_NUM] == shift + 1:
                                resulting_table[row_table + 1][shift + 1] = variable_incrementer(detailed_job_report, row_djr, resulting_table[row_table + 1][shift + 1], negative_num_rows, excessive_num_rows, 2)

        if yes_or_no(5):
            # delete unnecessary rows
            location = 1
            for row_table in range(len(resulting_table) - 1):
                all_zero = True
                for col in range(3):
                    if resulting_table[row_table + 1][col + 1] > 0:
                        all_zero = False
                if not all_zero:
                    for col in range(4):
                        resulting_table[location][col] = resulting_table[row_table + 1][col]

                    location = location + 1

            for row in range(len(resulting_table) - location):
                resulting_table = np.delete(resulting_table, location, axis=0)
            for row in range(location - 1):
                for col in range(1, 4):
                    resulting_table[row + 1][col] = int(resulting_table[row + 1][col])

        print_table(resulting_table)

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

        if yes_or_no(2):
            write_to_excel(resulting_table, len(resulting_table))

        if yes_or_no(6):
            calculate_average_feeds_by_shift(resulting_table, len(resulting_table))

    else: # user wants to display feeds per day by crew
        # check if there are gaps in data
        gap_name = check_for_gaps_in_data(detailed_job_report, start_date_num, end_date_num)

        # ask user if they want to use name-filling algorithm
        use_algo = False
        if gap_name:
            use_algo = yes_or_no(1)
        empty_name_rows = []

        crews_list = []
        generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list, empty_name_rows, use_algo)

        if len(crews_list) != 0:
            # create header for resulting table
            resulting_table = np.array([[0 for x in range(len(crews_list) + 1)] for y in range(int(end_date_num - start_date_num + 2))], dtype='object')
            resulting_table[0][0] = "Work Date"
            for col in range(len(crews_list)):
                resulting_table[0][col + 1] = crews_list[col]

            # generate rest of table
            for row_table in range(len(resulting_table) - 1):
                resulting_table[row_table + 1][0] = convert_date_int_to_string(start_date_num + row_table)
            for row_djr in range(ROWS):
                if detailed_job_report[row_djr][MACHINE_COL_NUM] == MACHINE:
                    for row_table in range(len(resulting_table) - 1):
                        if str(start_date_num + row_table) == str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[8:10]:
                            if str(detailed_job_report[row_djr][EMPLOYEE_NAME_COL_NUM]) != "nan":
                                for crew_counter in range(len(crews_list)):
                                    if detailed_job_report[row_djr][EMPLOYEE_NAME_COL_NUM] == crews_list[crew_counter]:
                                        resulting_table[row_table + 1][crew_counter + 1] = variable_incrementer(detailed_job_report, row_djr, resulting_table[row_table + 1][crew_counter + 1], negative_num_rows, excessive_num_rows, 2)
                            if str(detailed_job_report[row_djr][EMPLOYEE_NAME_COL_NUM]) == "nan" and use_algo:
                                for crew_index in range(len(crews_list)):
                                    feeds_calculated_by_AI = 0
                                    feeds_calculated_by_AI = name_filling_algorithm(detailed_job_report, feeds_calculated_by_AI, crews_list[crew_index], row_djr, empty_name_rows, negative_num_rows, excessive_num_rows, 4)
                                    resulting_table[row_table + 1][crew_index + 1] = resulting_table[row_table + 1][crew_index + 1] + feeds_calculated_by_AI

            if yes_or_no(5):
                # delete unnecessary rows
                location = 1
                for row_table in range(len(resulting_table) - 1):
                    all_zero = True
                    for col in range(len(crews_list)):
                        if resulting_table[row_table + 1][col + 1] > 0:
                            all_zero = False
                    if not all_zero:
                        for col in range(len(crews_list) + 1):
                            resulting_table[location][col] = resulting_table[row_table + 1][col]

                        location = location + 1

            for row in range(len(resulting_table) - location):
                resulting_table = np.delete(resulting_table, location, axis=0)
            for row in range(location - 1):
                for col in range(1, len(crews_list) + 1):
                    resulting_table[row + 1][col] = int(resulting_table[row + 1][col])

            print_table(resulting_table)

            if len(empty_name_rows) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    print_additional_info(use_algo, empty_name_rows, negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(resulting_table, len(resulting_table))

            if yes_or_no(6):
                calculate_average_feeds_by_crew(resulting_table, len(resulting_table), crews_list)

        else:
            print("There is nothing to show here")


# function to display the average order size, or the total number of jobs by # of colors or ups
# arguments: the detailed job report as an array, whether the user wants average order size, the total number of jobs
# by the number of colors or the total number of jobs by the number of ups,
# the start date as an integer and the end date as an integer
# returns nothing
def display_order_type(detailed_job_report, option, start_date_num, end_date_num):
    if option == 1: # if user chooses option 1, find average order size
        unique_orders = []
        total_quantity = 0

        # sum all unique orders + total feed quantity
        for row in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                    total_quantity = incrementing_algo(detailed_job_report, row, total_quantity, unique_orders)

        print("\nTotal Order Quantity: " + str(total_quantity))
        print("Total Unique Orders: " + str(len(unique_orders)))
        print("Average Order Size: ", end="")
        if len(unique_orders) > 0:
            print(total_quantity / len(unique_orders))
        else:
            print("N/A")

    else: # any other option means user wants total orders by the # of colors/ups
        # first find the largest number of items in the detailed job report (either the # of colors or ups)
        largest_num_items = 0
        for row in range(ROWS):
            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                if option == 2: # find the largest number of colors used
                    if str(detailed_job_report[row][NUM_COLORS_COL_NUM]) != "nan":
                        if detailed_job_report[row][NUM_COLORS_COL_NUM] > largest_num_items:
                            largest_num_items = int(detailed_job_report[row][NUM_COLORS_COL_NUM])
                else: # find the largest number of ups used
                    if str(detailed_job_report[row][NUM_UPS_COL_NUM]) != "nan":
                        if detailed_job_report[row][NUM_UPS_COL_NUM] > largest_num_items:
                            largest_num_items = int(detailed_job_report[row][NUM_UPS_COL_NUM])

        # create resulting array for printing/writing
        resulting_array = np.array([[0 for x in range(largest_num_items + 2)] for y in range(2)], dtype='object')
        if option == 2:
            resulting_array[0][0] = "Number of Colors"
        else:
            resulting_array[0][0] = "Number of Ups"
        resulting_array[1][0] = "Total Orders"

        counter = 1
        for num_items in range(largest_num_items + 1):
            unique_orders = []  # array to keep track of unique orders
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                        if option == 2: # user wants jobs by number of colors
                            if detailed_job_report[row][NUM_COLORS_COL_NUM] == num_items:
                                append_element_in_array(unique_orders, detailed_job_report[row][ORDER_NUM_COL_NUM])
                        else: # user wants jobs by number of ups
                            if detailed_job_report[row][NUM_UPS_COL_NUM] == num_items:
                                append_element_in_array(unique_orders, detailed_job_report[row][ORDER_NUM_COL_NUM])

            if len(unique_orders) > 0:
                resulting_array[0][counter] = num_items
                resulting_array[1][counter] = len(unique_orders)
                counter = counter + 1

        # remove unnecessary columns
        for col in range(largest_num_items - counter + 2):
            resulting_array = np.delete(resulting_array, counter, axis=1)
        # reconvert all numbers back to integers
        for row in range(2):
            for col in range(counter - 1):
                resulting_array[row][col + 1] = int(resulting_array[row][col + 1])

        print_table(resulting_array)

        if yes_or_no(2):
            write_to_excel(resulting_array, 2)


# function for printing the average run speed
# arguments: the detailed job report, an option to determine what we need to print
# the start date as an integer and the end date as an integer
# 1 = average run speed by shift and 2 = average run speed by crew
# returns nothing
def display_average_run_speed(detailed_job_report, option, start_date_num, end_date_num):
    negative_num_rows = []
    excessive_num_rows = []

    if option == "1": # user wants average run speed by shift
        # create & initialize array to hold data in case user wants to write data to an Excel spreadsheet
        average_run_speed_by_shift_array = [[0 for x in range(2)] for y in range(4)]
        average_run_speed_by_shift_array[0][0], average_run_speed_by_shift_array[0][1] = "Shift", "Average Run Speed (feeds/hour)"

        for shift in range(3):
            total_feeds = 0
            total_run_hours = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][CHARGE_CODE_COL_NUM] == "RUN": # calculate total run elapsed hours for specific shift
                        if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                            total_feeds = variable_incrementer(detailed_job_report, row, total_feeds, negative_num_rows, excessive_num_rows, 2)
                            total_run_hours = variable_incrementer(detailed_job_report, row, total_run_hours, negative_num_rows, excessive_num_rows, 1)

            if total_run_hours != 0:
                average_run_speed = (total_feeds / total_run_hours)
                average_run_speed_by_shift_array[shift + 1][0], average_run_speed_by_shift_array[shift + 1][1] = shift + 1, average_run_speed
            else:
                average_run_speed_by_shift_array[shift + 1][0], average_run_speed_by_shift_array[shift + 1][1] = shift + 1, "N/A"

        print_table(average_run_speed_by_shift_array)

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

        if yes_or_no(2):
            write_to_excel(average_run_speed_by_shift_array, len(average_run_speed_by_shift_array))

    else: # user wants average run speed by crew
        # check if there are gaps in data
        gap_name = check_for_gaps_in_data(detailed_job_report, start_date_num, end_date_num)

        # ask user if they want to use name-filling algorithm
        use_algo = False
        if gap_name:
            use_algo = yes_or_no(1)
        rows_with_no_name = []

        # create list of crews
        crews_list = []
        generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list, rows_with_no_name, use_algo)

        if len(crews_list) != 0:
            average_run_speed_by_crew_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
            average_run_speed_by_crew_array[0][0], average_run_speed_by_crew_array[0][1] = "Crew", "Average Run Speed (feeds/hour)"

            crew_counter = 1
            for crew in crews_list:
                total_feeds = 0
                total_run_hours = 0
                for row in range(ROWS):
                    if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew and detailed_job_report[row][CHARGE_CODE_COL_NUM] == "RUN":
                            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                                total_feeds = variable_incrementer(detailed_job_report, row, total_feeds, negative_num_rows, excessive_num_rows, 2)
                                total_run_hours = variable_incrementer(detailed_job_report, row, total_run_hours, negative_num_rows, excessive_num_rows, 1)
                        elif use_algo and detailed_job_report[row][CHARGE_CODE_COL_NUM] == "RUN":
                            total_feeds = name_filling_algorithm(detailed_job_report, total_feeds, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 4)
                            total_run_hours = name_filling_algorithm(detailed_job_report, total_run_hours, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 1)

                if total_run_hours > 0:
                    average_run_speed_by_crew_array[crew_counter][0], average_run_speed_by_crew_array[crew_counter][1] = crew, total_feeds / total_run_hours
                else:
                    average_run_speed_by_crew_array[crew_counter][0], average_run_speed_by_crew_array[crew_counter][1] = crew, "N/A"

                crew_counter = crew_counter + 1

            print_table(average_run_speed_by_crew_array)

            if len(rows_with_no_name) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    print_additional_info(use_algo, rows_with_no_name, negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(average_run_speed_by_crew_array, len(average_run_speed_by_crew_array))
        else:
            print("There is nothing to show here")


# function for displaying the top three best or worst orders
# arguments: the detailed job report, an option and a sub option to determine what the user wishes to analyze,
# the start date as an integer and the end date as an integer
# returns nothing
def display_top_three_orders(detailed_job_report, option, sub_option, start_date_num, end_date_num):
    negative_num_rows = []
    excessive_num_rows = []
    unique_orders = []

    run_dict = {}
    setup_dict = {}
    ODT_dict = {}
    efficiency_dict = {}
    total_time_dict = {}

    # create list of unique orders
    for row in range(ROWS):
        if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                # only include valid orders that have both "SETUP" and "RUN" charge codes
                setup_check = False
                run_check = False
                for row_two in range(ROWS):
                    if detailed_job_report[row_two][ORDER_NUM_COL_NUM] == detailed_job_report[row][ORDER_NUM_COL_NUM]:
                        if start_date_num <= int(str(detailed_job_report[row_two][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_two][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_two][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                            if detailed_job_report[row_two][MACHINE_COL_NUM] == MACHINE:
                                if detailed_job_report[row_two][CHARGE_CODE_COL_NUM] == "SET UP":
                                    setup_check = True
                                if detailed_job_report[row_two][CHARGE_CODE_COL_NUM] == "RUN":
                                    run_check = True
                if setup_check and run_check:
                    append_element_in_array(unique_orders, detailed_job_report[row][ORDER_NUM_COL_NUM])

    # initialize all dictionaries
    for unique_order in unique_orders:
        run_dict[unique_order] = 0
        setup_dict[unique_order] = 0
        ODT_dict[unique_order] = 0
        efficiency_dict[unique_order] = 0
        total_time_dict[unique_order] = 0

    # fill in run, setup and ODT dictionaries
    for row in range(ROWS):
        if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                order_found = False
                for unique_order in unique_orders:
                    if detailed_job_report[row][ORDER_NUM_COL_NUM] == unique_order:
                        order_found = True
                        break

                if order_found:
                    if detailed_job_report[row][CHARGE_CODE_COL_NUM] == "RUN":
                        run_dict[detailed_job_report[row][ORDER_NUM_COL_NUM]] = variable_incrementer(detailed_job_report, row, run_dict[detailed_job_report[row][ORDER_NUM_COL_NUM]], negative_num_rows, excessive_num_rows, 1)
                    elif detailed_job_report[row][CHARGE_CODE_COL_NUM] == "SET UP":
                        setup_dict[detailed_job_report[row][ORDER_NUM_COL_NUM]] = variable_incrementer(detailed_job_report, row, setup_dict[detailed_job_report[row][ORDER_NUM_COL_NUM]], negative_num_rows, excessive_num_rows, 1)
                    else:
                        ODT_dict[detailed_job_report[row][ORDER_NUM_COL_NUM]] = variable_incrementer(detailed_job_report, row, ODT_dict[detailed_job_report[row][ORDER_NUM_COL_NUM]], negative_num_rows, excessive_num_rows, 1)

    # fill in efficiency and total time dictionaries
    for key in efficiency_dict.keys():
        if ODT_dict[key] + setup_dict[key] != 0 or run_dict[key] != 0:
            efficiency_dict[key] = run_dict[key] / (ODT_dict[key] + setup_dict[key] + run_dict[key]) * 100

        total_time_dict[key] = run_dict[key] + ODT_dict[key] + setup_dict[key]

    # produce resulting table
    top_three_orders_array = [[0 for x in range(2)] for y in range(4)]

    # create headers for resulting table
    if sub_option == "1":
        top_three_orders_array[0][0], top_three_orders_array[0][1] = "Order Number", "Efficiency (%)"
    elif sub_option == "2":
        top_three_orders_array[0][0], top_three_orders_array[0][1] = "Order Number", "ODT (hours)"
    elif sub_option == "3":
        top_three_orders_array[0][0], top_three_orders_array[0][1] = "Order Number", "Setup Time (hours)"
    else:
        top_three_orders_array[0][0], top_three_orders_array[0][1] = "Order Number", "Total Time (hours)"

    # modify initialized values for specific situations
    if (option == "1" and sub_option != "1") or (option != "1" and sub_option == "1"):
        for row in range(1, 4):
            top_three_orders_array[row][1] = 100

    # create rest of table
    if sub_option == "1":
        build_top_three_orders_array(top_three_orders_array, efficiency_dict, int(option), int(sub_option))
    elif sub_option == "2":
        build_top_three_orders_array(top_three_orders_array, ODT_dict, int(option), int(sub_option))
    elif sub_option == "3":
        build_top_three_orders_array(top_three_orders_array, setup_dict, int(option), int(sub_option))
    else:
        build_top_three_orders_array(top_three_orders_array, total_time_dict, int(option), int(sub_option))

    print_table(top_three_orders_array)

    if len(negative_num_rows) != 0 or len(excessive_num_rows) != 0:
        if yes_or_no(3):
            print_incorrect_hours(negative_num_rows, excessive_num_rows)

    if yes_or_no(2):
        write_to_excel(top_three_orders_array, len(top_three_orders_array))


##################
# global variables
##################
ROWS = 0
COLS = 0
MACHINE = ""

CHARGE_CODE_COL_NUM = 0
ELAPSED_HOURS_COL_NUM = 3
WORK_DATE_COL_NUM = 4
SHIFT_COL_NUM = 5
ORDER_NUM_COL_NUM = 8
ORDER_QTY_COL_NUM = 9
GROSS_FG_QTY_COL_NUM = 15
MACHINE_COL_NUM = 18
NUM_UPS_COL_NUM = 19
EMPLOYEE_NAME_COL_NUM = 26
DOWNTIME_COL_NUM = 28
NUM_COLORS_COL_NUM = 32

EXCESSIVE_THRESHOLD = 5

######
# main
######
preliminary_choices_made = False
while True:
    while not preliminary_choices_made:
        # obtaining preliminary information from the user
        djr_array = obtain_detailed_job_report()
        ROWS, COLUMNS = djr_array.shape
        MACHINE = obtain_machine_to_analyze(djr_array)
        preliminary_choices_made = True

    # obtaining instruction from the user
    user_input = obtain_instruction()
    first_date_string = ""
    second_date_string = ""

    #######################################
    # calculating open down time percentage
    #######################################
    if user_input == 1:
        print("\nEnter the start date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        user_choice = obtain_sub_instruction(1)

        if user_choice == "1":
            print("\nOverall open down time for all shifts/crews from " + first_date_string + " to " + second_date_string + ":\n")
        elif user_choice == "2":
            print("\nOpen down time from " + first_date_string + " to " + second_date_string + " based on shift:")
        elif user_choice == "3":
            print("\nOpen down time from " + first_date_string + " to " + second_date_string + " based on crew:")
        else:
            print("\nPareto chart of open down time from " + first_date_string + " to " + second_date_string + ":")

        display_ODT(djr_array, user_choice, first_date_string, second_date_string) # show calculated data

    #########################
    # calculating total feeds
    #########################
    elif user_input == 2:
        # obtain date frame from user
        print("Enter the start date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        user_choice = obtain_sub_instruction(2)

        if user_choice == "1":
            print("\nTotal feeds by shift from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("\nTotal feeds by crew from " + first_date_string + " to " + second_date_string + ":")

        display_total_feeds(djr_array, user_choice, start_date_num, end_date_num)

    ################################
    # calculating average setup time
    ################################
    elif user_input == 3:
        # obtain date frame from user
        print("Enter the start date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        user_choice = obtain_sub_instruction(3)

        if user_choice == "1":
            print("\nGeneral average setup time from " + first_date_string + " to " + second_date_string + ":")
        elif user_choice == "2":
            print("\nAverage setup time by shift from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("\nAverage setup time by crew from " + first_date_string + " to " + second_date_string + ":")

        display_average_setup_time(djr_array, user_choice, start_date_num, end_date_num)

    #############################
    # breaking down feeds per day
    #############################
    elif user_input == 4:
        # obtain date frame from user
        print("Enter the start date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        user_choice = obtain_sub_instruction(4)

        if user_choice == "1":
            print("\nAverage daily feeds for all shifts/crew from " + first_date_string + " to " + second_date_string + ":\n")
        elif user_choice == "2":
            print("\nFeeds per day by shift from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("\nFeeds per day by crew from " + first_date_string + " to " + second_date_string + ":")

        display_daily_feeds(djr_array, user_choice, start_date_num, end_date_num)

    #######################
    # analyze by order type
    #######################
    elif user_input == 5:
        # obtain date frame from user
        print("Enter the start date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        user_choice = obtain_sub_instruction(5)

        if user_choice == "1":
            print("\nAverage order size from " + first_date_string + " to " + second_date_string + ":")
        elif user_choice == "2":
            print("\nJobs by number of colours from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("\nJobs by number of ups from " + first_date_string + " to " + second_date_string + ":")

        display_order_type(djr_array, int(user_choice), start_date_num, end_date_num)

    #####################
    # calculate run speed
    #####################
    elif user_input == 6:
        # obtain date frame from user
        print("Enter the start date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        user_choice = obtain_sub_instruction(6)

        if user_choice == "1":
            print("\nAverage run speed by shift from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("\nAverage run speed by crew from " + first_date_string + " to " + second_date_string + ":")

        display_average_run_speed(djr_array, user_choice, start_date_num, end_date_num)

    #####################################
    # sort top three best or worst orders
    #####################################
    elif user_input == 7:
        # obtain date frame from user
        print("Enter the start date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        user_choice = obtain_sub_instruction(7)
        user_sub_choice = obtain_sub_instruction(8)

        if user_choice == "1":
            print("\nTop three best orders by ", end="")
        else:
            print("\nTop three worst orders by ", end="")

        if user_sub_choice == "1":
            print("efficiency from " + first_date_string + " to " + second_date_string + ":")
        elif user_sub_choice == "2":
            print("ODT from " + first_date_string + " to " + second_date_string + ":")
        elif user_sub_choice == "3":
            print("setup time from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("total time spent from " + first_date_string + " to " + second_date_string + ":")

        display_top_three_orders(djr_array, user_choice, user_sub_choice, start_date_num, end_date_num)

    ###############################################
    # analyze different machine/detailed job report
    ###############################################
    elif user_input == 8:
        preliminary_choices_made = False

    ######
    # exit
    ######
    elif user_input == 9:
        exit()
