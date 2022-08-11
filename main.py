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
    print("------------------------------------")
    print("| 1 - calculate open down time     |")
    print("| 2 - calculate total feeds        |")
    print("| 3 - calculate average setup time |")
    print("| 4 - break down feeds per day     |")
    print("| 5 - analyze order type           |")
    print("| 6 - calculate average run speed  |")
    print("| 8 - return to home page          |")
    print("| 9 - exit                         |")
    print("------------------------------------")

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
            choice = input("Enter (1) to calculate general average setup time or (2) to calculate average setup time by crew: ")
        elif option == 4:
            choice = input("Enter (1) to calculate overall daily average feeds for all shifts/crews, (2) to display feeds per day by shift or (3) to display feeds per day by crew: ")
        elif option == 5:
            choice = input("Enter (1) to calculate the average order size, (2) to analyze jobs by the number of colors or (3) to analyze jobs by the number of ups: ")
        else:
            choice = input("Enter (1) to calculate the average run speed by shift or (2) to calculate the average run speed by crew: ")

        if option == 1:
            if choice == "1" or choice == "2" or choice == "3" or choice == "4":
                error = False
            else:
                print("Please try again.")
        elif option == 4 or option == 5:
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


# function to obtain a second date as an input from the user and return that same date
# function performs error checking to ensure date is valid and in the future relative to the first date
# arguments: the detailed job report array and the first date as a string
# returns the second date (entered by the user) as a string
def obtain_second_date_string(detailed_job_report_array, first_date_string):
    second_date_string = ""
    error = True
    while error:
        print("Enter the second date (YYYY/MM/DD): ", end="")
        second_date_string = obtain_date_string(detailed_job_report_array)

        if int(second_date_string[5:7]) < int(first_date_string[5:7]) or int(second_date_string[8:10]) < int(first_date_string[8:10]):
            print("Error: second date precedes first date. Please try again.")
        else:
            error = False

    return second_date_string


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
        if option == 1: # increment on total machine hours
            if detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                assumed_name = assume_name(detailed_job_report, empty_name_rows, row)
                counter = update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, False)
        elif option == 2: # increment on ODT
            if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                assumed_name = assume_name(detailed_job_report, empty_name_rows, row)
                counter = update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, False)
        elif option == 3: # increment on total setup hours
            if detailed_job_report[row][DOWNTIME_COL_NUM] == "Setup" and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                assumed_name = assume_name(detailed_job_report, empty_name_rows, row)
                counter = update_counter(detailed_job_report, row, assumed_name, crew, counter, negative_num_rows, excessive_num_rows, False)
        else: # increment on total feeds
            if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
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
    if assumed_name == crew and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
        if is_total_feeds:
            counter = counter + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
        else:
            counter = counter + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
    elif assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
        append_element_in_array(excessive_num_rows, row)
    elif assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
        append_element_in_array(negative_num_rows, row)

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


# function to obtain spreadsheet name and write data from user's previous
# command to said spreadsheet
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
        spreadsheet_name = input("Enter the name for the new spreadsheet: ")
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


# function to print a long string or number according to a specific length
# arguments: the element to print as a string
# and the number of characters/digits requested as an int
# returns nothing
def print_element_long(element, num_digits):
    for digit in range(len(element)):
        if digit < num_digits: # note: the decimal point counts as a digit
            print(element[digit], end="")
        else:
            if element[digit] != "." and digit != len(element) - 1:
                if int(element[digit + 1]) > 5 and int(element[digit]) != 9:
                    print(int(element[digit]) + 1, end="")
                else:
                    print(element[digit], end="")
            break


# function to print either a number with few digits or a string with few characters
# according to a specific space to fill
# arguments: the element to print as a string
# and the number of total spaces to fill as an int
# returns nothing
def print_element_short(element, length_to_fill):
    for space in range(int((length_to_fill - len(element)) / 2)):
        print(" ", end="")
    print(element, end="")
    for space in range(int((length_to_fill - len(element)) / 2)):
        print(" ", end="")

    if len(element) % 2 != 0 and length_to_fill % 2 == 0:
        print(" ", end="")
    elif len(element) % 2 == 0 and length_to_fill % 2 != 0:
        print(" ", end="")


# function to print the table header when user wishes to display info by crew
# arguments: the list of crew members required for printing
# and an option (int) which reflects whether the header should print
# ODT, total feeds, average setup time or average feeds + opportunity
# returns length of longest name
def print_crew_header(list_of_crew_members, option):
    # find length of longest name first
    longest_name_length = 0
    for crew in list_of_crew_members:
        if len(crew) > longest_name_length:
            longest_name_length = len(crew)

    if longest_name_length > len(CREW_LABEL):
        if option == 1:
            for dash in range(longest_name_length + len(ODT_LABEL_PERCENTAGE) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(longest_name_length + len(TOTAL_FEEDS_LABEL) + 7):
                print("-", end="")
        elif option == 3:
            for dash in range(longest_name_length + len(AVERAGE_SETUP_TIME_LABEL) + 7):
                print("-", end="")
        elif option == 4:
            for dash in range(longest_name_length + len(AVERAGE_FEEDS_LABEL) + len(OPPORTUNITY_LABEL) + 10):
                print("-", end="")
        else:
            for dash in range(longest_name_length + len(AVERAGE_RUN_SPEED_LABEL) + 7):
                print("-", end="")
    else:
        if option == 1:
            for dash in range(len(ODT_LABEL_PERCENTAGE) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(len(TOTAL_FEEDS_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 3:
            for dash in range(len(AVERAGE_SETUP_TIME_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 4:
            for dash in range(len(AVERAGE_FEEDS_LABEL) + len(OPPORTUNITY_LABEL) + len(CREW_LABEL) + 10):
                print("-", end="")
        else:
            for dash in range(len(AVERAGE_RUN_SPEED_LABEL)  + 7):
                print("-", end="")

    print("\n| ", end="")

    if longest_name_length > len(CREW_LABEL):
        for space in range(int((longest_name_length - len(CREW_LABEL)) / 2)):
            print(" ", end="")
    print(CREW_LABEL, end="")
    if longest_name_length > 4:
        for space in range(int((longest_name_length - len(CREW_LABEL)) / 2)):
            print(" ", end="")
    if longest_name_length % 2 != 0:
        print(" ", end="")

    if option == 1:
        print(" | " + ODT_LABEL_PERCENTAGE + " |")
    elif option == 2:
        print(" | " + TOTAL_FEEDS_LABEL + " |")
    elif option == 3:
        print(" | " + AVERAGE_SETUP_TIME_LABEL + " |")
    elif option == 4:
        print(" | " + AVERAGE_FEEDS_LABEL + " | " + OPPORTUNITY_LABEL + " |")
    else:
        print(" | " + AVERAGE_RUN_SPEED_LABEL + " |")

    if longest_name_length > len(CREW_LABEL):
        if option == 1:
            for dash in range(longest_name_length + len(ODT_LABEL_PERCENTAGE) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(longest_name_length + len(TOTAL_FEEDS_LABEL) + 7):
                print("-", end="")
        elif option == 3:
            for dash in range(longest_name_length + len(AVERAGE_SETUP_TIME_LABEL) + 7):
                print("-", end="")
        elif option == 4:
            for dash in range(longest_name_length + len(AVERAGE_FEEDS_LABEL) + len(OPPORTUNITY_LABEL) + 10):
                print("-", end="")
        else:
            for dash in range(longest_name_length + len(AVERAGE_RUN_SPEED_LABEL) + 7):
                print("-", end="")
    else:
        if option == 1:
            for dash in range(len(ODT_LABEL_PERCENTAGE) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(len(TOTAL_FEEDS_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 3:
            for dash in range(len(AVERAGE_SETUP_TIME_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 4:
            for dash in range(len(AVERAGE_FEEDS_LABEL) + len(OPPORTUNITY_LABEL) + len(CREW_LABEL) + 10):
                print("-", end="")
        else:
            for dash in range(len(AVERAGE_RUN_SPEED_LABEL) + 7):
                print("-", end="")

    print()

    return longest_name_length


# function to continue printing rest of table
# arguments: the data array required for printing,
# the length of the longest crew name
# and an option (int) which reflects whether the function should print
# ODT (by crew or by charge code), total feeds or average setup time
# 1 = ODT by crew, 2 = ODT by charge code, 3 = total feeds, 4 = average setup time, 5 = average run speed
# returns nothing
def print_rest_of_table(array, longest_name_length, option):
    # first, find data with longest number of digits
    longest_number = 0
    for row in range(len(array) - 1):
        if option == 3:
            if len(str(int(array[row + 1][1]))) > longest_number:
                longest_number = len(str(int(array[row + 1][1])))
        else:
            if len(str(array[row + 1][1])) > longest_number:
                longest_number = len(str(array[row + 1][1]))

    for row in range(len(array) - 1):
        print("| ", end="")
        print_element_short(array[row + 1][0], longest_name_length)
        print(" | ", end="")
        if option == 1:
            if len(str(array[row + 1][1])) < len(ODT_LABEL_PERCENTAGE):
                print_element_short(str(array[row + 1][1]), len(ODT_LABEL_PERCENTAGE))
            else:
                print_element_long(str(array[row + 1][1]), len(ODT_LABEL_PERCENTAGE) - 1)
        elif option == 2:
            if len(str(array[row + 1][1])) < len(ODT_LABEL_HOURS):
                print_element_short(str(array[row + 1][1]), len(ODT_LABEL_HOURS))
            else:
                print_element_long(str(array[row + 1][1]), len(ODT_LABEL_HOURS) - 1)
        elif option == 3:
            if len(str(int(array[row + 1][1]))) < len(TOTAL_FEEDS_LABEL):
                print_element_short(str(array[row + 1][1]), len(TOTAL_FEEDS_LABEL))
            else:
                print_element_long(str(array[row + 1][1]), len(TOTAL_FEEDS_LABEL) - 1)
        elif option == 4:
            if len(str(array[row + 1][1])) < len(AVERAGE_SETUP_TIME_LABEL):
                print_element_short(str(array[row + 1][1]), len(AVERAGE_SETUP_TIME_LABEL))
            else:
                print_element_long(str(array[row + 1][1]), len(AVERAGE_SETUP_TIME_LABEL) - 1)
        else:
            if len(str(array[row + 1][1])) < len(AVERAGE_RUN_SPEED_LABEL):
                print_element_short(str(array[row + 1][1]), len(AVERAGE_RUN_SPEED_LABEL))
            else:
                print_element_long(str(array[row + 1][1]), len(AVERAGE_RUN_SPEED_LABEL) - 1)
        print(" |")

    if option == 1:
        for dash in range(longest_name_length + len(ODT_LABEL_PERCENTAGE) + 7):
            print("-", end="")
    elif option == 2:
        for dash in range(longest_name_length + len(ODT_LABEL_HOURS) + 7):
            print("-", end="")
    elif option == 3:
        for dash in range(longest_name_length + len(TOTAL_FEEDS_LABEL) + 7):
            print("-", end="")
    elif option == 4:
        for dash in range(longest_name_length + len(AVERAGE_SETUP_TIME_LABEL) + 7):
            print("-", end="")
    else:
        for dash in range(longest_name_length + len(AVERAGE_RUN_SPEED_LABEL) + 7):
            print("-", end="")
    print()


# function to print the table header when user wishes to display Pareto chart table
# arguments: the sorted pareto array
# returns length of longest charge code
def print_charge_code_header(pareto_array):
    # find length of longest charge code
    longest_charge_code_length = 0
    for row in range(len(pareto_array)):
        if len(pareto_array[row][0]) > longest_charge_code_length:
            longest_charge_code_length = len(pareto_array[row][0])

    if longest_charge_code_length > len(CHARGE_CODE_LABEL):
        for dash in range(longest_charge_code_length + len(ODT_LABEL_HOURS) + 7):
            print("-", end="")
    else:
        for dash in range(len(CHARGE_CODE_LABEL) + len(ODT_LABEL_HOURS) + 7):
            print("-", end="")

    print("\n| ", end="")
    if longest_charge_code_length > len(CHARGE_CODE_LABEL):
        for space in range(int((longest_charge_code_length - len(CHARGE_CODE_LABEL)) / 2)):
            print(" ", end="")
    print(CHARGE_CODE_LABEL, end="")
    if longest_charge_code_length > len(CHARGE_CODE_LABEL):
        for space in range(int((longest_charge_code_length - len(CHARGE_CODE_LABEL)) / 2)):
            print(" ", end="")
    if longest_charge_code_length % 2 == 0:
        print(" ", end="")

    print(" | " + ODT_LABEL_HOURS + " |")

    if longest_charge_code_length > len(CHARGE_CODE_LABEL):
        for dash in range(longest_charge_code_length + len(ODT_LABEL_HOURS) + 7):
            print("-", end="")
    else:
        for dash in range(len(CHARGE_CODE_LABEL) + len(ODT_LABEL_HOURS) + 7):
            print("-", end="")

    print()

    return longest_charge_code_length


# function for converting an integer corresponding to a date
# to its string equivalent
# arguments: the date as an integer
# returns the date as a string
def convert_date_int_to_string(date_num):
    day = ""
    month = ""
    year = ""

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


# function for printing feeds per day according to shift number
# arguments: the resulting table for feeds per day and the length of said table
# returns nothing
def print_feeds_per_day_by_shift(feeds_per_day_array, length_of_table):
    # first find length of longest number of total feeds
    longest_number_len = 0
    for row in range(len(feeds_per_day_array) - 1):
        for col in range(3):
            if len(str(feeds_per_day_array[row + 1][col + 1])) > longest_number_len:
                longest_number_len = len(str(feeds_per_day_array[row + 1][col + 1]))

    if longest_number_len > 7:
        for dash in range(3 * longest_number_len + len(WORK_DATE_LABEL) + 13):
            print("-", end="")
    else:
        for dash in range(len(WORK_DATE_LABEL) + (3 * 7) + 15):
            print("-", end="")
    print()

    print("|  Work Date  | ", end="")
    for shift in range(3):
        if longest_number_len > 7:
            print_element_short("Shift " + str(shift + 1), longest_number_len)
        else:
            print("Shift " + str(shift + 1), end="")
        print(" | ", end="")
    print()

    if longest_number_len > 7:
        for dash in range(3 * longest_number_len + len(WORK_DATE_LABEL) + 13):
            print("-", end="")
    else:
        for dash in range(len(WORK_DATE_LABEL) + (3 * 7) + 15):
            print("-", end="")
    print()

    for row in range(length_of_table - 1):
        print("| " + feeds_per_day_array[row + 1][0] + " | ", end="")

        for col in range(3):
            if longest_number_len > 7:
                print_element_long(str(feeds_per_day_array[row + 1][col + 1]), 7)
            else:
                print_element_short(str(feeds_per_day_array[row + 1][col + 1]), 7)
            print(" | ", end="")

        print()

    if longest_number_len > 7:
        for dash in range(3 * longest_number_len + len(WORK_DATE_LABEL) + 13):
            print("-", end="")
    else:
        for dash in range(len(WORK_DATE_LABEL) + (3 * 7) + 15):
            print("-", end="")
    print()


# function for printing feeds per day according to crew
# arguments: the resulting table for feeds per day and the length of said table
# returns nothing
def print_feeds_per_day_by_crew(feeds_per_day_array, length_of_table, list_of_crews):
    # first find length of longest number of total feeds
    longest_number_len = 0
    for row in range(len(feeds_per_day_array) - 1):
        for col in range(len(list_of_crews)):
            if len(str(feeds_per_day_array[row + 1][col + 1])) > longest_number_len:
                longest_number_len = len(str(feeds_per_day_array[row + 1][col + 1]))

    # find the longest crew name
    longest_name_len = 0
    for crew in list_of_crews:
        if len(str(crew)) > longest_name_len:
            longest_name_len = len(str(crew))

    print_dashes_by_crew(longest_number_len, longest_name_len, list_of_crews)

    print("|  Work Date  | ", end="")
    for crew in list_of_crews:
        if longest_number_len > longest_name_len:
            print_element_short(crew, longest_number_len)
        else:
            print_element_short(crew, longest_name_len)
        print(" | ", end="")
    print()

    print_dashes_by_crew(longest_number_len, longest_name_len, list_of_crews)

    for row in range(length_of_table - 1):
        print("| " + feeds_per_day_array[row + 1][0] + " | ", end="")

        for col in range(len(list_of_crews)):
            if longest_number_len > longest_name_len:
                print_element_long(str(feeds_per_day_array[row + 1][col + 1]), longest_number_len)
            else:
                print_element_short(str(feeds_per_day_array[row + 1][col + 1]), longest_name_len)
            print(" | ", end="")

        print()

    print_dashes_by_crew(longest_number_len, longest_name_len, list_of_crews)


# function for printing dashes when displaying feeds per day by crew
# arguments: length of longest number, length of longest name and
# the list of crew members
# returns nothing
def print_dashes_by_crew(length_of_longest_number, length_of_longest_name, list_of_crews):
    if length_of_longest_number > length_of_longest_name:
        for dash in range(len(list_of_crews) * length_of_longest_number + len(WORK_DATE_LABEL) + len(list_of_crews) * 3 + 6):
            print("-", end="")
    else:
        for dash in range(len(WORK_DATE_LABEL) + 3 * len(list_of_crews) + 6):
            print("-", end="")
        for dash in range(length_of_longest_name * len(list_of_crews)):
            print("-", end="")
    print()


# function for calculating the average feeds by shift
# arguments: the feeds per day by shift array and the length of said array
# returns nothing
def calculate_average_feeds_by_shift(feeds_per_day_array, len_feeds_per_day_array):
    resulting_average_table = [[0 for i in range(3)] for j in range(4)]

    resulting_average_table[0][0], resulting_average_table[0][1], resulting_average_table[0][2] = "Shift", AVERAGE_FEEDS_LABEL, OPPORTUNITY_LABEL
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

    print_average_feeds_by_shift(resulting_average_table)
    if yes_or_no(2):
        write_to_excel(resulting_average_table, len(resulting_average_table))


# function for calculating the average feeds by crew
# arguments: the array of feeds per day by crew, the length of the array of feeds per day by crew,
# and the list of all crew members considered
# returns nothing
def calculate_average_feeds_by_crew(feeds_per_day_array, len_feeds_per_day_array, list_of_crews):
    resulting_average_table = [[0 for i in range(3)] for j in range(len(list_of_crews) + 1)]

    # print header and crew names first
    resulting_average_table[0][0], resulting_average_table[0][1], resulting_average_table[0][2] = CREW_LABEL, AVERAGE_FEEDS_LABEL, OPPORTUNITY_LABEL
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

    print_average_feeds_by_crew(resulting_average_table, list_of_crews)
    if yes_or_no(2):
        write_to_excel(resulting_average_table, len(resulting_average_table))


# function for printing the average feeds by shift
# arguments: the array of average feeds by shift
# returns nothing
def print_average_feeds_by_shift(average_array):
    # find the longest average feeds per day
    longest_average_feeds_len = 0
    for row in range(len(average_array) - 1):
        if average_array[row + 1][1] != "N/A":
            if average_array[row + 1][1] > longest_average_feeds_len:
                longest_average_feeds_len = average_array[row + 1][1]

    # find the longest opportunity
    longest_opportunity_len = 0
    for row in range(len(average_array) - 1):
        if average_array[row + 1][2] != "N/A":
            if average_array[row + 1][2] > longest_opportunity_len:
                longest_opportunity_len = average_array[row + 1][2]

    # print header
    print("---------------------------------------------------")
    print("| Shift | " + AVERAGE_FEEDS_LABEL + " | " + OPPORTUNITY_LABEL + " |")
    print("---------------------------------------------------")

    # print rest of table
    for row in range(len(average_array) - 1):
        print("|   " + str(row + 1) + "   |", end="")

        if len(str(average_array[row + 1][1])) <= len(AVERAGE_FEEDS_LABEL):
            print_element_short(str(average_array[row + 1][1]), len(AVERAGE_FEEDS_LABEL) + 1)
        else:
            print_element_long(str(average_array[row + 1][1]), len(AVERAGE_FEEDS_LABEL) - 1)

        print(" | ", end="")
        if len(str(average_array[row + 1][2])) <= len(OPPORTUNITY_LABEL):
            print_element_short(str(average_array[row + 1][2]), len(OPPORTUNITY_LABEL))
        else:
            print_element_long(str(average_array[row + 1][2]), len(OPPORTUNITY_LABEL) - 1)
        print(" |")

    for dash in range(len("Shift") + len(AVERAGE_FEEDS_LABEL) + len(OPPORTUNITY_LABEL) + 10):
        print("-", end="")
    print()


# function for printing the average feeds by crew
# arguments: the array of average feeds by crew and the list of all crew members considered
# returns nothing
def print_average_feeds_by_crew(average_array, list_of_crews):
    # find the longest average feeds per day
    longest_average_feeds_len = 0
    for row in range(len(average_array) - 1):
        if average_array[row + 1][1] != "N/A":
            if average_array[row + 1][1] > longest_average_feeds_len:
                longest_average_feeds_len = average_array[row +1][1]

    # find the longest opportunity
    longest_opportunity_len = 0
    for row in range(len(average_array) - 1):
        if average_array[row + 1][2] != "N/A":
            if average_array[row + 1][2] > longest_opportunity_len:
                longest_opportunity_len = average_array[row + 1][2]

    # print header
    longest_crew_name = print_crew_header(list_of_crews, 4)

    # print rest of table
    for row in range(len(average_array) - 1):
        print("| ", end="")
        print_element_short(average_array[row + 1][0], longest_crew_name)
        print(" | ", end="")

        if len(str(average_array[row + 1][1])) <= len(AVERAGE_FEEDS_LABEL):
            print_element_short(str(average_array[row + 1][1]), len(AVERAGE_FEEDS_LABEL))
        else:
            print_element_long(str(average_array[row + 1][1]), len(AVERAGE_FEEDS_LABEL) - 1)

        print(" | ", end="")
        if len(str(average_array[row + 1][2])) <= len(OPPORTUNITY_LABEL):
            print_element_short(str(average_array[row + 1][2]), len(OPPORTUNITY_LABEL))
        else:
            print_element_long(str(average_array[row + 1][2]), len(OPPORTUNITY_LABEL) - 1)
        print(" |")

    for dash in range(longest_crew_name + len(AVERAGE_FEEDS_LABEL) + len(OPPORTUNITY_LABEL) + 10):
        print("-", end="")
    print()


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
        crew_ODT_by_charge_code_array[0][0], crew_ODT_by_charge_code_array[0][1] = CREW_LABEL, (charge_code + " (hours elapsed)")

        for index in range(len(crews_list)): # write all crew names to resulting table
            crew_ODT_by_charge_code_array[index + 1][0] = crews_list[index]

        for index in range(len(crews_list)): # compute ODT elapsed hours for each crew member based on charge code
            elapsed_hours_for_crew = 0
            for row in range(len(detailed_job_report)):
                if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE and detailed_job_report[row][CHARGE_CODE_COL_NUM] == charge_code:
                    if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if crews_list[index] == detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]:
                            if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                elapsed_hours_for_crew = elapsed_hours_for_crew + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]

                        elif use_algo and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                            if assume_name(detailed_job_report, rows_with_no_name, row) == crews_list[index]:
                                if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                    elapsed_hours_for_crew = elapsed_hours_for_crew + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                            if assume_name(detailed_job_report, rows_with_no_name, row) == "nan":
                                append_element_in_array(rows_with_no_name, row)

            crew_ODT_by_charge_code_array[index + 1][1] = elapsed_hours_for_crew

        print_ODT_by_crew_for_charge_code(crew_ODT_by_charge_code_array, crews_list)

        if use_algo and len(rows_with_no_name) > 0:
            if yes_or_no(3):
                print("\n")
                print_rows_with_no_name(rows_with_no_name)
                print("---------------------------------------------")

        if yes_or_no(2):
            write_to_excel(crew_ODT_by_charge_code_array, len(crew_ODT_by_charge_code_array))
    else:
        print("There is nothing to show here")


# function for printing the resulting array for the ODT by crew according to a specific charge code
# arguments: the resulting ODT by charge code array and the list of all relevant crew members
def print_ODT_by_crew_for_charge_code(crew_ODT_by_charge_code_array, list_of_crews):
    # find length of longest name first
    longest_name_length = 0
    for crew in list_of_crews:
        if len(crew) > longest_name_length:
            longest_name_length = len(crew)

    # print header
    if longest_name_length > len(CREW_LABEL):
        for dash in range(longest_name_length + len(str(crew_ODT_by_charge_code_array[0][1])) + 7):
            print("-", end="")
    else:
        for dash in range(len(CREW_LABEL) + len(str(crew_ODT_by_charge_code_array[0][1])) + 7):
            print("-", end="")

    print("\n| ", end="")

    if longest_name_length > len(CREW_LABEL):
        for space in range(int((longest_name_length - len(CREW_LABEL)) / 2)):
            print(" ", end="")
    print(CREW_LABEL, end="")
    if longest_name_length > len(CREW_LABEL):
        for space in range(int((longest_name_length - len(CREW_LABEL)) / 2)):
            print(" ", end="")
    if longest_name_length % 2 != 0:
        print(" ", end="")
    print(" | " + crew_ODT_by_charge_code_array[0][1] + " |")

    if longest_name_length > len(CREW_LABEL):
        for dash in range(longest_name_length + len(str(crew_ODT_by_charge_code_array[0][1])) + 7):
            print("-", end="")
    else:
        for dash in range(len(CREW_LABEL) + len(str(crew_ODT_by_charge_code_array[0][1])) + 7):
            print("-", end="")

    print()

    # print rest of table
    longest_number = 0
    for row in range(len(crew_ODT_by_charge_code_array) - 1):
        if len(str(crew_ODT_by_charge_code_array[row + 1][1])) > longest_number:
            longest_number = len(str(crew_ODT_by_charge_code_array[row + 1][1]))

    for row in range(len(crew_ODT_by_charge_code_array) - 1):
        print("| ", end="")
        print_element_short(crew_ODT_by_charge_code_array[row + 1][0], longest_name_length)
        print(" | ", end="")

        if len(str(crew_ODT_by_charge_code_array[row + 1][1])) < len(str(crew_ODT_by_charge_code_array[0][1])):
            print_element_short(str(crew_ODT_by_charge_code_array[row + 1][1]), len(str(crew_ODT_by_charge_code_array[0][1])))
            # print(len(str(crew_ODT_by_charge_code_array[0][1])))
        else:
            print_element_long(str(crew_ODT_by_charge_code_array[row + 1][1]), len(str(crew_ODT_by_charge_code_array[0][1])))

        print(" |")

    for dash in range(longest_name_length + len(str(crew_ODT_by_charge_code_array[0][1])) + 7):
        print("-", end="")
    print()


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


# function for printing the resulting array for order type
# works for both the number of colors and the number of ups
# arguments: the array to print and the maximum number of colors/ups found in the detailed job report
# returns nothing
def print_order_type_array(array_to_print, num_cols):
    for dash in range(len(array_to_print[0][0]) + 8 * (num_cols - 1) + 4):
        print("-", end="")
    print()

    print("| ", end="")
    print(array_to_print[0][0], end="")
    print(" |", end="")

    for num in range(num_cols - 1):
        print("   " + str(array_to_print[0][num + 1]), end="")
        if array_to_print[0][num + 1] < 10:
            print("   |", end="")
        else:
            print("  |", end="")
    print()

    for dash in range(len(array_to_print[0][0]) + 8 * (num_cols - 1) + 4):
        print("-", end="")
    print()

    print("| ", end="")
    for space in range(int((len(str(array_to_print[0][0])) - (len(TOTAL_ORDERS_LABEL))) / 2)):
        print(" ", end="")
    print(TOTAL_ORDERS_LABEL, end="")
    for space in range(int((len(str(array_to_print[0][0])) - (len(TOTAL_ORDERS_LABEL))) / 2)):
        print(" ", end="")
    if len(str(array_to_print[0][0])) % 2 != 0:
        print(" ", end="")
    print(" |", end="")

    for col in range(num_cols - 1):
        print(" ", end="")
        if len(str(array_to_print[1][col + 1])) > 5:
            print_element_long(str(array_to_print[1][col + 1]), 5)
        else:
            print_element_short(str(array_to_print[1][col + 1]), 5)
        print(" |", end="")

    print()

    for dash in range(len(array_to_print[0][0]) + 8 * (num_cols - 1) + 4):
        print("-", end="")
    print()


# function for displaying both the list of rows where no crew name could be attributed by the AI,
# the list of rows with negative elapsed hours and the list of excessive elapsed hours
# arguments: whether the algorithm was used, the list of rows with no names,
# the list of rows with negative elapsed hours and the list of rows with excessive elapsed hours
def display_additional_info(algorithm_used, rows_with_no_name, negative_num_rows, excessive_num_rows):
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
                    # calculate total machine hours for given date frame
                    if detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= 5:
                        total_machine_hours = total_machine_hours + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_element_in_array(negative_num_rows, row)
                    elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > 5:
                        append_element_in_array(excessive_num_rows, row)

                    # calculate total ODT hours for given date frame
                    if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= 5:
                        total_ODT = total_ODT + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_element_in_array(negative_num_rows, row)
                    elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > 5:
                        append_element_in_array(excessive_num_rows, row)

        total_machine_hours = total_machine_hours + total_ODT

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
        print("-------------------")
        print("| Shift | ODT (%) |")
        print("-------------------")

        # create & initialize array to hold data in case user wants to write data to an Excel spreadsheet
        ODT_by_shift_array = [[0 for x in range(2)] for y in range(4)]
        ODT_by_shift_array[0][0], ODT_by_shift_array[0][1] = "Shift", "ODT %"

        for shift in range(3):
            total_machine_hours = 0
            ODT = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                        # calculate total machine hours for specific shift
                        if detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                            total_machine_hours = total_machine_hours + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                        elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                            append_element_in_array(excessive_num_rows, row)
                        elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                            append_element_in_array(negative_num_rows, row)

                        # calculate total open downtime for specific shift
                        if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                            ODT = ODT + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                        elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                            append_element_in_array(excessive_num_rows, row)
                        elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                            append_element_in_array(negative_num_rows, row)

            total_machine_hours = total_machine_hours + ODT

            print("|   " + str(shift + 1) + "   | ", end="")
            if total_machine_hours != 0:
                result = (ODT/total_machine_hours) * 100
                if len(str(result)) > 6:
                    print_element_long(str(result), 6)
                else:
                    print_element_short(str(result), 7)
                print(" |")

                ODT_by_shift_array[shift + 1][0], ODT_by_shift_array[shift + 1][1] = shift + 1, result
            else:
                print("  N/A   |")
                ODT_by_shift_array[shift + 1][0], ODT_by_shift_array[shift + 1][1] = shift + 1, "N/A"

        print("-------------------")

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
            ODT_by_crew_array[0][0], ODT_by_crew_array[0][1] = "Crew", "ODT %"

            longest_name_len = print_crew_header(crews_list, 1)

            counter = 1
            for crew in crews_list:
                total_machine_hours = 0
                ODT = 0
                for row in range(ROWS):
                    if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                            if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew:
                                # calculate total machine hours for specific crew
                                if str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Run" and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                    total_machine_hours = total_machine_hours + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                                elif str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Run" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                    append_element_in_array(excessive_num_rows, row)
                                elif str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Run" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                                    append_element_in_array(negative_num_rows, row)

                                # calculate total ODT for specific crew
                                if str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Open Downtime" and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                    ODT = ODT + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                                elif str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Open Downtime" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                    append_element_in_array(excessive_num_rows, row)
                                elif str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Open Downtime" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                                    append_element_in_array(negative_num_rows, row)

                            elif use_algo:
                                total_machine_hours = name_filling_algorithm(detailed_job_report, total_machine_hours, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 1)
                                ODT = name_filling_algorithm(detailed_job_report, ODT, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 2)

                # add total run time with total ODT for total_machine_hours
                total_machine_hours = total_machine_hours + ODT

                if total_machine_hours > 0:
                    ODT_by_crew_array[counter][0], ODT_by_crew_array[counter][1] = crew, (ODT/total_machine_hours) * 100
                else:
                    ODT_by_crew_array[counter][0], ODT_by_crew_array[counter][1] = crew, "N/A"

                counter = counter + 1

            print_rest_of_table(ODT_by_crew_array, longest_name_len, 1)

            if len(rows_with_no_name) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    display_additional_info(use_algo, rows_with_no_name, negative_num_rows, excessive_num_rows)

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
                # calculate total open downtime for specific charge code
                ODT = 0
                for row_num in range(ROWS):
                    if start_date_num <= int(str(detailed_job_report[row_num][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_num][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_num][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if detailed_job_report[row_num][MACHINE_COL_NUM] == MACHINE:
                            if detailed_job_report[row_num][CHARGE_CODE_COL_NUM] == charge_code and 0 <= detailed_job_report[row_num][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                ODT = ODT + detailed_job_report[row_num][ELAPSED_HOURS_COL_NUM]
                            elif detailed_job_report[row_num][CHARGE_CODE_COL_NUM] == charge_code and detailed_job_report[row_num][ELAPSED_HOURS_COL_NUM] < 0:
                                append_element_in_array(negative_num_rows, row_num)
                            elif detailed_job_report[row_num][CHARGE_CODE_COL_NUM] == charge_code and detailed_job_report[row_num][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                append_element_in_array(excessive_num_rows, row_num)

                charge_code_dict[charge_code] = ODT

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

            longest_charge_code_len = print_charge_code_header(charge_code_array)
            print_rest_of_table(charge_code_array, longest_charge_code_len, 2)

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
        print("-----------------------")
        print("| Shift | Total Feeds |")
        print("-----------------------")

        total_feeds_by_shift_array = [[0 for x in range(2)] for y in range(4)]
        total_feeds_by_shift_array[0][0], total_feeds_by_shift_array[0][1] = "Shift", "Total Feeds"

        shift_counter = 1
        for shift in range(3):
            total_feeds = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                        if detailed_job_report[row][SHIFT_COL_NUM] == shift + 1 and 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                            total_feeds = total_feeds + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                        elif detailed_job_report[row][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                            append_element_in_array(negative_num_rows, row)
                        elif detailed_job_report[row][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                            append_element_in_array(excessive_num_rows, row)

            if total_feeds > 0:
                total_feeds_by_shift_array[shift_counter][0], total_feeds_by_shift_array[shift_counter][1] = shift + 1, total_feeds
            shift_counter = shift_counter + 1

        # find data with longest number of digits
        longest_number = 0
        for row in range(len(total_feeds_by_shift_array) - 1):
            if len(str(total_feeds_by_shift_array[row + 1][1])) > longest_number:
                longest_number = len(str(total_feeds_by_shift_array[row + 1][1]))

        table_empty = True
        for row in range(len(total_feeds_by_shift_array) - 1):
            if int(total_feeds_by_shift_array[row + 1][0]) != 0 and int(total_feeds_by_shift_array[row + 1][1]) != 0:
                table_empty = False
                print("|   " + str(total_feeds_by_shift_array[row + 1][0]) + "   | ", end="")
                if longest_number < len(TOTAL_FEEDS_LABEL):
                    print_element_short(str(int(total_feeds_by_shift_array[row + 1][1])), len(TOTAL_FEEDS_LABEL))
                else:
                    print_element_long(str(total_feeds_by_shift_array[row + 1][1]), len(TOTAL_FEEDS_LABEL))
                print(" |")

        if not table_empty:
            print("-----------------------")

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
        longest_name_len = print_crew_header(crews_list, 2)

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
                                if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                    total_feeds = total_feeds + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                                elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                                    append_element_in_array(negative_num_rows, row)
                                elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                    append_element_in_array(excessive_num_rows, row)

                            elif use_algo:
                                total_feeds = name_filling_algorithm(detailed_job_report, total_feeds, crew, row, empty_name_rows, negative_num_rows, excessive_num_rows, 4)

                total_feeds_by_crew_array[counter][0], total_feeds_by_crew_array[counter][1] = crew, total_feeds
                counter = counter + 1

            print_rest_of_table(total_feeds_by_crew_array, longest_name_len, 3)

            if len(empty_name_rows) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    display_additional_info(use_algo, empty_name_rows, negative_num_rows, excessive_num_rows)

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
            if detailed_job_report[elapsed_hours][DOWNTIME_COL_NUM] == "Setup":
                if start_date_num <= int(str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[elapsed_hours][MACHINE_COL_NUM] == MACHINE:
                        if detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] < 0:
                            append_element_in_array(negative_num_rows, elapsed_hours)
                        elif detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                            append_element_in_array(excessive_num_rows, elapsed_hours)
                        else:
                            total_elapsed_hours = total_elapsed_hours + detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM]
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
            longest_name_len = print_crew_header(crews_list, 3)

            average_setup_time_by_crew_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
            average_setup_time_by_crew_array[0][0], average_setup_time_by_crew_array[0][1] = "Crew", "Average Setup Time"

            counter = 1
            for crew in crews_list:
                total_elapsed_hours = 0  # counter for total elapsed hours
                unique_orders_list = []  # list to track all unique orders

                for elapsed_hours in range(ROWS):
                    if detailed_job_report[elapsed_hours][DOWNTIME_COL_NUM] == "Setup":
                        if start_date_num <= int(str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                            if detailed_job_report[elapsed_hours][MACHINE_COL_NUM] == MACHINE:
                                if detailed_job_report[elapsed_hours][EMPLOYEE_NAME_COL_NUM] == crew:
                                    if detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] < 0:
                                        append_element_in_array(negative_num_rows, elapsed_hours)
                                    elif detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                        append_element_in_array(excessive_num_rows, elapsed_hours)
                                    elif 0 <= detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                        total_elapsed_hours = total_elapsed_hours + detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] # increment elapsed hours
                                        append_element_in_array(unique_orders_list, detailed_job_report[elapsed_hours][ORDER_NUM_COL_NUM]) # append order if unique
                                elif use_algo:
                                    # increment elapsed hours
                                    total_elapsed_hours = name_filling_algorithm(detailed_job_report, total_elapsed_hours, crew, elapsed_hours, rows_with_no_name, negative_num_rows, excessive_num_rows, 3)

                                    # append order if unique
                                    if str(detailed_job_report[elapsed_hours][EMPLOYEE_NAME_COL_NUM]) == "nan":
                                        if assume_name(djr_array, rows_with_no_name, elapsed_hours) == crew:
                                            append_element_in_array(unique_orders_list, detailed_job_report[elapsed_hours][ORDER_NUM_COL_NUM])

                # final calculation
                total_unique_orders = len(unique_orders_list)
                if total_unique_orders != 0:
                    # print(str(total_elapsed_hours) + " " + str(total_unique_orders))
                    average_setup_time = (total_elapsed_hours / total_unique_orders) * 60
                    average_setup_time_by_crew_array[counter][0], average_setup_time_by_crew_array[counter][1] = crew, average_setup_time
                else:
                    average_setup_time_by_crew_array[counter][0], average_setup_time_by_crew_array[counter][1] = crew, "N/A"
                counter = counter + 1

            print_rest_of_table(average_setup_time_by_crew_array, longest_name_len, 4)

            if len(rows_with_no_name) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    display_additional_info(use_algo, rows_with_no_name, negative_num_rows, excessive_num_rows)

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
                    if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                        total_feeds = total_feeds + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                        append_element_in_array(unique_days, detailed_job_report[row][WORK_DATE_COL_NUM])
                    elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_element_in_array(negative_num_rows, row)
                    elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                        append_element_in_array(excessive_num_rows, row)

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
        resulting_table = [[0 for x in range(4)] for y in range(int(end_date_num - start_date_num + 2))]
        resulting_table[0][0], resulting_table[0][1], resulting_table[0][2], resulting_table[0][3] = "Work Date", "Shift 1", "Shift 2", "Shift 3"

        # generate rest of table
        for row_table in range(len(resulting_table) - 1):
            resulting_table[row_table + 1][0] = convert_date_int_to_string(start_date_num + row_table)
        for row_djr in range(ROWS):
            if detailed_job_report[row_djr][MACHINE_COL_NUM] == MACHINE:
                for row_table in range(len(resulting_table) - 1):
                    if str(start_date_num + row_table) == str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[8:10]:
                        for shift in range(3):
                            if detailed_job_report[row_djr][SHIFT_COL_NUM] == shift + 1 and 0 <= detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                resulting_table[row_table + 1][shift + 1] = resulting_table[row_table + 1][shift + 1] + detailed_job_report[row_djr][GROSS_FG_QTY_COL_NUM]
                            elif detailed_job_report[row_djr][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                append_element_in_array(excessive_num_rows, row_djr)
                            elif detailed_job_report[row_djr][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] < 0:
                                append_element_in_array(negative_num_rows, row_djr)

        table_length = len(resulting_table)

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

            table_length = location

        print_feeds_per_day_by_shift(resulting_table, table_length)

        if len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
            if yes_or_no(3):
                print_incorrect_hours(negative_num_rows, excessive_num_rows)

        if yes_or_no(2):
            write_to_excel(resulting_table, table_length)

        if yes_or_no(6):
            calculate_average_feeds_by_shift(resulting_table, table_length)

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
            resulting_table = [[0 for x in range(len(crews_list) + 1)] for y in range(int(end_date_num - start_date_num + 2))]
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
                                        if 0 <= detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                            resulting_table[row_table + 1][crew_counter + 1] = resulting_table[row_table + 1][crew_counter + 1] + detailed_job_report[row_djr][GROSS_FG_QTY_COL_NUM]
                                        elif detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                            append_element_in_array(excessive_num_rows, row_djr)
                                        elif detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] < 0:
                                            append_element_in_array(negative_num_rows, row_djr)
                            if str(detailed_job_report[row_djr][EMPLOYEE_NAME_COL_NUM]) == "nan" and use_algo:
                                for crew_index in range(len(crews_list)):
                                    feeds_calculated_by_AI = 0
                                    feeds_calculated_by_AI = name_filling_algorithm(detailed_job_report, feeds_calculated_by_AI, crews_list[crew_index], row_djr, empty_name_rows, negative_num_rows, excessive_num_rows, 4)
                                    resulting_table[row_table + 1][crew_index + 1] = resulting_table[row_table + 1][crew_index + 1] + feeds_calculated_by_AI

            table_length = len(resulting_table)

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

                table_length = location

            print_feeds_per_day_by_crew(resulting_table, table_length, crews_list)

            if len(empty_name_rows) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    display_additional_info(use_algo, empty_name_rows, negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(resulting_table, table_length)

            if yes_or_no(6):
                calculate_average_feeds_by_crew(resulting_table, table_length, crews_list)

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
            resulting_array[0][0], resulting_array[1][0] = NUM_COLORS_LABEL, TOTAL_ORDERS_LABEL
        else:
            resulting_array[0][0], resulting_array[1][0] = NUM_UPS_LABEL, TOTAL_ORDERS_LABEL

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

        print_order_type_array(resulting_array, counter)

        # remove unnecessary columns
        for col in range(largest_num_items - counter + 2):
            resulting_array = np.delete(resulting_array, counter, axis=1)
        # reconvert all numbers back to integers
        for row in range(2):
            for col in range(counter - 1):
                resulting_array[row][col + 1] = int(resulting_array[row][col + 1])

        if yes_or_no(2):
            write_to_excel(resulting_array, 2)


def display_average_run_speed(detailed_job_report, option, start_date_num, end_date_num):
    negative_num_rows = []
    excessive_num_rows = []

    if option == "1": # user wants average run speed by shift
        print("------------------------------------------")
        print("| Shift | Average Run Speed (feeds/hour) |")
        print("------------------------------------------")

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
                            if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                total_feeds = total_feeds + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                                total_run_hours = total_run_hours + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                            elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                append_element_in_array(excessive_num_rows, row)
                            elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                                append_element_in_array(negative_num_rows, row)

            print("|   " + str(shift + 1) + "   | ", end="")
            if total_run_hours != 0:
                average_run_speed = (total_feeds / total_run_hours)
                average_run_speed_by_shift_array[shift + 1][0], average_run_speed_by_shift_array[shift + 1][1] = shift + 1, average_run_speed

                if len(str(average_run_speed)) > 29:
                    print_element_long(str(average_run_speed), 29)
                else:
                    print_element_short(str(average_run_speed), 30)
            else:
                for spaces in range(int(len(AVERAGE_RUN_SPEED_LABEL) / 2) - 2):
                    print(" ", end="")
                print("N/A", end="")
                for spaces in range(int(len(AVERAGE_RUN_SPEED_LABEL) / 2) - 2):
                    print(" ", end="")
                if len(AVERAGE_RUN_SPEED_LABEL) % 2 == 0:
                    print(" ", end="")
            print(" |")

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
            average_run_speed_by_crew_array[0][0], average_run_speed_by_crew_array[0][1] = CREW_LABEL, AVERAGE_RUN_SPEED_LABEL

            longest_name_len = print_crew_header(crews_list, 5)

            crew_counter = 1
            for crew in crews_list:
                total_feeds = 0
                total_run_hours = 0
                for row in range(ROWS):
                    if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew and detailed_job_report[row][CHARGE_CODE_COL_NUM] == "RUN":
                            if detailed_job_report[row][MACHINE_COL_NUM] == MACHINE:
                                if 0 <= detailed_job_report[row][ELAPSED_HOURS_COL_NUM] <= EXCESSIVE_THRESHOLD:
                                    total_feeds = total_feeds + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                                    total_run_hours = total_run_hours + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                                elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] > EXCESSIVE_THRESHOLD:
                                    append_element_in_array(excessive_num_rows, row)
                                elif detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                                    append_element_in_array(negative_num_rows, row)

                        elif use_algo and detailed_job_report[row][CHARGE_CODE_COL_NUM] == "RUN":
                            total_feeds = name_filling_algorithm(detailed_job_report, total_feeds, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 4)
                            total_run_hours = name_filling_algorithm(detailed_job_report, total_run_hours, crew, row, rows_with_no_name, negative_num_rows, excessive_num_rows, 1)

                if total_run_hours > 0:
                    average_run_speed_by_crew_array[crew_counter][0], average_run_speed_by_crew_array[crew_counter][1] = crew, total_feeds / total_run_hours
                else:
                    average_run_speed_by_crew_array[crew_counter][0], average_run_speed_by_crew_array[crew_counter][1] = crew, "N/A"

                crew_counter = crew_counter + 1

            print_rest_of_table(average_run_speed_by_crew_array, longest_name_len, 5)

            if len(rows_with_no_name) > 0 or len(negative_num_rows) > 0 or len(excessive_num_rows) > 0:
                if yes_or_no(3):
                    display_additional_info(use_algo, rows_with_no_name, negative_num_rows, excessive_num_rows)

            if yes_or_no(2):
                write_to_excel(average_run_speed_by_crew_array, len(average_run_speed_by_crew_array))
        else:
            print("There is nothing to show here")


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


ODT_LABEL_HOURS = "ODT (hours)"
ODT_LABEL_PERCENTAGE = "ODT (%)"
TOTAL_FEEDS_LABEL = "Total Feeds"
AVERAGE_SETUP_TIME_LABEL = "Average Setup Time (minutes)"
CREW_LABEL = "Crew"
CHARGE_CODE_LABEL = "Charge Code"
WORK_DATE_LABEL = "Work Date"
AVERAGE_FEEDS_LABEL = "Average Feeds per Day"
OPPORTUNITY_LABEL = "Opportunity (%)"
NUM_COLORS_LABEL = "Number of Colors"
NUM_UPS_LABEL = "Number of Ups"
TOTAL_ORDERS_LABEL = "Total Orders"
AVERAGE_ORDER_QTY_LABEL = "Average Order Quantity"
AVERAGE_RUN_SPEED_LABEL = "Average Run Speed (feeds/hour)"

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
        print("\nEnter the first date (YYYY/MM/DD): ", end="")
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
        print("Enter the first date (YYYY/MM/DD): ", end="")
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
        print("Enter the first date (YYYY/MM/DD): ", end="")
        first_date_string = obtain_date_string(djr_array)
        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        user_choice = obtain_sub_instruction(3)

        if user_choice == "1":
            print("\nGeneral average setup time from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("\nAverage setup time from " + first_date_string + " to " + second_date_string + ":")

        display_average_setup_time(djr_array, user_choice, start_date_num, end_date_num)

    #############################
    # breaking down feeds per day
    #############################
    elif user_input == 4:
        # obtain date frame from user
        print("Enter the first date (YYYY/MM/DD): ", end="")
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
        print("Enter the first date (YYYY/MM/DD): ", end="")
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
        print("Enter the first date (YYYY/MM/DD): ", end="")
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
