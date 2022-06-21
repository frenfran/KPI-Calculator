import pandas as pd
import numpy as np

##################
# global variables
##################
ROWS = 0
COLUMNS = 0

CHARGE_CODE_COL_NUM = 0
ELAPSED_HOURS_COL_NUM = 3
WORK_DATE_COL_NUM = 4
SHIFT_COL_NUM = 5
ORDER_NUM_COL_NUM = 8
GROSS_FG_QTY_COL_NUM = 15
EMPLOYEE_NAME_COL_NUM = 26
DOWNTIME_COL_NUM = 28

ODT_LABEL = "ODT (%)"
TOTAL_FEEDS_LABEL = "Total Feeds"
AVERAGE_SETUP_TIME_LABEL = "Average Setup Time (minutes)"
CREW_LABEL = "Crew"
CHARGE_CODE_LABEL = "Charge Code"
WORK_DATE_LABEL = "Work Date"


##################
# helper functions
##################


# function to print all possible instructions and receive input from the user
# function also performs error checking to ensure user's input is valid
# arguments: none
# returns nothing
def obtain_instructions():
    print("\nWhat would you like to do?")
    print("-------------------------------------------")
    print("| 1 - calculate open down time percentage |")
    print("| 2 - calculate total feeds               |")
    print("| 3 - calculate average setup time        |")
    print("| 4 - break down feeds per day            |")
    print("| 5 - exit                                |")
    print("-------------------------------------------")

    user_error = True
    user_input = ""
    while user_error:
        user_input = input("Enter the number associated to your command: ")
        if user_input == "1" or user_input == "2" or user_input == "3" or user_input == "4" or user_input == "5":
            user_error = False
        else:
            print("Error: please try again.")

    return int(user_input)


# function to obtain a date as an input from the user and return that same date
# function performs error checking to ensure date is valid and within the detailed job report
# arguments: the detailed job report array and the total number of rows in the detailed job array
# returns the date entered by the user as a string
def obtain_date_string(detailed_job_report):
    day = ""
    month = ""
    year = ""
    result = ""

    error = True
    while error:
        try:
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
        print("Enter the second date: ", end="")
        second_date_string = obtain_date_string(detailed_job_report_array)

        if int(second_date_string[5:7]) < int(first_date_string[5:7]) or int(second_date_string[8:10]) < int(first_date_string[8:10]):
            print("Error: second date precedes first date. Please try again.")
        else:
            error = False

    return second_date_string


# algorithm to continue incrementing a counter variable
# (either total machine hours, ODT, total setup time or total feeds)
# when a missing name occurs in the employee name column
# arguments: the detailed job report array, the value of the counter variable so far, crew name from employee column,
# the start day as an integer, the end date as an integer,
# the array of rows with empty names still undetermined by the algorithm,
# the list of rows with negative elapsed hours
# and an int value (either 1, 2, 3 or 4) to determine which variable to increment
# returns new value of the counter variable
def name_filling_algorithm(detailed_job_report, counter, crew, start_date_num, end_date_num, empty_name_rows, negative_num_rows, option):
    if option == 1: # increment on total machine hours
        for row in range(ROWS):
            if detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    # first determine which employee name should fill this gap
                    assumed_name = assume_name(detailed_job_report, empty_name_rows, row)

                    # add to counter if name filled matches crew name
                    if assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        # print("-----------------\nTotal machine hours")
                        # print(detailed_job_report[row][ELAPSED_HOURS_COL_NUM])
                        # print(assumed_name)
                        counter = counter + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    elif assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

    elif option == 2: # increment on ODT
        for row in range(ROWS):
            if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    # first determine which name should fill this empty cell
                    assumed_name = assume_name(detailed_job_report, empty_name_rows, row)

                    # add to counter if name filled matches crew name
                    if assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        # print("-------\nODT")
                        # print(detailed_job_report[row][ELAPSED_HOURS_COL_NUM])
                        # print(assumed_name)
                        counter = counter + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    elif assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

    elif option == 3: # increment on total setup hours
        for row in range(ROWS):
            if detailed_job_report[row][DOWNTIME_COL_NUM] == "Setup" and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    # first determine which employee name should fill this gap
                    assumed_name = assume_name(detailed_job_report, empty_name_rows, row)

                    # add to counter if name filled matches crew name
                    if assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        # print("-------\nODT")
                        # print(detailed_job_report[row][ELAPSED_HOURS_COL_NUM])
                        # print(assumed_name)
                        counter = counter + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    elif assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

    else: # increment on total feeds
        for row in range(ROWS):
            if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    # first determine which employee name should fill this gap
                    assumed_name = assume_name(detailed_job_report, empty_name_rows, row)

                    if assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        counter = counter + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                    elif assumed_name == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

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


# function to print all negative numbers in a given list
# arguments: the number of negative numbers in an array and the array itself
# returns nothing
def print_neg_nums_list(num_negatives, negative_nums_list):
    print("Row(s) with negative elapsed hours:\n")

    counter = 0
    for row in range(num_negatives):
        print(negative_nums_list[row] + 2, end=" ")
        counter = counter + 1

        if counter > 15:
            print("")
            counter = 0

    if counter != 0:
        print()
    print("\nThese row(s) were omitted from calculations")
    print("---------------------------------------------")


# function to create array of unique crew members
# arguments: the detailed job report array, the start date as an integer,
# the end date as an integer and the empty list of crew members
# returns nothing
def generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list):
    for row in range(ROWS):
        if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
            already_included = False
            for crew in crews_list:
                if crew == str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]):
                    already_included = True
                    break

            if not already_included and str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) != "nan":
                crews_list.append(str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]))


# function for printing all the rows in which AI was unable to attribute a name
# arguments: the completed list of rows with no name
# returns nothing
def print_rows_with_no_name(rows_with_no_name):
    if len(rows_with_no_name) > 0:
        print("AI was unable to attribute names to the following row(s) due to insufficient information:\n")
        # reorganize rows_with_no_name from first row to last in numerical order
        sorting_algorithm(rows_with_no_name)
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


# function to determine whether user wants to ue AI
# arguments: none
# returns true or false depending on if user wants to use AI
def use_algorithm():
    while True:
        temp = input("Warning: gaps found in Employee Name column. Use AI to compute (y/n)? ")
        if temp == "y" or temp == "Y":
            return True
        elif temp == "n" or temp == "N":
            return False
        else:
            print("Error: please try again.")


# function to assume name to a particular empty Employee Name cell
# or to include the row as part of rows the AI could not assign a name to
# arguments: the detailed job report as an array,
# the list of empty names the AI could not attribute a name to
# and the row of the detailed job report being analyzed
# returns the assumed name
def assume_name(detailed_job_report, empty_name_rows, row):
    shift_num = detailed_job_report[row][SHIFT_COL_NUM]
    assumed_name = " "

    keep_going = True # check rows prior
    iterator = -1
    while keep_going:
        if detailed_job_report[row + iterator][SHIFT_COL_NUM] == shift_num and str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM]) != "nan" and int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) - int(str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[8:10]) < 2:
            assumed_name = str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM])
            keep_going = False
        elif detailed_job_report[row + iterator][SHIFT_COL_NUM] == shift_num and str(detailed_job_report[row + iterator][EMPLOYEE_NAME_COL_NUM]) == "nan" and int(str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator][WORK_DATE_COL_NUM])[8:10]) - int(str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row + iterator - 1][WORK_DATE_COL_NUM])[8:10]) < 2:
            iterator = iterator - 1
        else:
            keep_going = False

    if assumed_name == " ": # check rows after
        keep_going = True
        iterator = 1
        while keep_going:
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

    return assumed_name


# function to append rows with negative elapsed hours to an array
# ensures row isn't already included in array before appending it
# arguments: the array itself and the row to append
# returns nothing
def append_neg_num_row(array, neg_num_row):
    included = False
    for row in array:
        if row == neg_num_row:
            included = True

    if not included:
        array.append(neg_num_row)


# function to determine whether to write to an Excel spreadsheet
# asks the user if they'd like to write the data to an Excel spreadsheet
# arguments: none
# returns True or False
def to_excel():
    choice = ""
    error = True
    while error:
        choice = input("Would you like to write this data to Excel (y/n)? ")

        if choice == "y" or choice == "n" or choice == "Y" or choice == "N":
            error = False
        else:
            print("Please try again.")

    if choice == "y" or choice == "Y":
        return True
    else:
        return False


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
                print("\nError: no such file found in current directory")
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


# function to print a long number according to a specific number of digits
# arguments: the number to print as an int
# and the number of digits requested as an int
# returns nothing
def print_digit_long(result, num_digits):
    for digit in range(len(str(result))):
        if digit < num_digits: # note: the decimal point counts as a digit
            print(str(result)[digit], end="")
        else:
            if str(result)[digit] != "." and digit != len(str(result)) - 1:
                if int(str(result)[digit + 1]) > 5 and int(str(result)[digit + 1]) != 9:
                    print(int(str(result)[digit]) + 1, end="")
                else:
                    print(str(result)[digit], end="")
            break


# function to print a short number according to a specific space to fill
# arguments: the number to print as an int
# and the number of total spaces to fill as an int
# returns nothing
def print_digit_short(result, length_to_fill):
    for space in range(int((length_to_fill - len(str(result))) / 2)):
        print(" ", end="")
    print(result, end="")
    for space in range(int((length_to_fill - len(str(result))) / 2)):
        print(" ", end="")

    if len(str(result)) % 2 != 0 and length_to_fill % 2 == 0:
        print(" ", end="")
    elif len(str(result)) % 2 == 0 and length_to_fill % 2 != 0:
        print(" ", end="")


# function to print the table header when user wishes to display info by crew
# arguments: the list of crew members required for printing
# and an option (int) which reflects whether the header should print
# ODT, total feeds or average setup time
# returns length of longest name
def print_crew_header(list_of_crew_members, option):
    # find length of longest name first
    longest_name_length = 0
    for crew in list_of_crew_members:
        if len(crew) > longest_name_length:
            longest_name_length = len(crew)

    if longest_name_length > 4:
        if option == 1:
            for dash in range(longest_name_length + len(ODT_LABEL) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(longest_name_length + len(TOTAL_FEEDS_LABEL) + 7):
                print("-", end="")
        else:
            for dash in range(longest_name_length + len(AVERAGE_SETUP_TIME_LABEL) + 7):
                print("-", end="")
    else:
        if option == 1:
            for dash in range(len(ODT_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(len(TOTAL_FEEDS_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        else:
            for dash in range(len(AVERAGE_SETUP_TIME_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")

    print("\n| ", end="")

    if longest_name_length > 4:
        for space in range(int((longest_name_length - 4) / 2)):
            print(" ", end="")
    print(CREW_LABEL, end="")
    if longest_name_length > 4:
        for space in range(int((longest_name_length - 4) / 2)):
            print(" ", end="")
    if longest_name_length % 2 != 0:
        print(" ", end="")

    if option == 1:
        print(" | " + ODT_LABEL +" |")
    elif option == 2:
        print(" | " + TOTAL_FEEDS_LABEL + " |")
    else:
        print(" | " + AVERAGE_SETUP_TIME_LABEL + " |")

    if longest_name_length > 4:
        if option == 1:
            for dash in range(longest_name_length + len(ODT_LABEL) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(longest_name_length + len(TOTAL_FEEDS_LABEL) + 7):
                print("-", end="")
        else:
            for dash in range(longest_name_length + len(AVERAGE_SETUP_TIME_LABEL) + 7):
                print("-", end="")
    else:
        if option == 1:
            for dash in range(len(ODT_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        elif option == 2:
            for dash in range(len(TOTAL_FEEDS_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")
        else:
            for dash in range(len(AVERAGE_SETUP_TIME_LABEL) + len(CREW_LABEL) + 7):
                print("-", end="")

    print()

    return longest_name_length


# function to continue printing rest of table
# arguments: the data array required for printing,
# the length of the longest crew name
# and an option (int) which reflects whether the function should print
# ODT (by crew or by charge code), total feeds or average setup time
# 1 = ODT (by crew or by charge code), 2 = total feeds, 3 = average setup time
# returns length of longest name
def print_rest_of_table(array, longest_name_length, option):
    # first, find data with longest number of digits
    longest_number = 0
    for row in range(len(array) - 1):
        if option == 2:
            if len(str(int(array[row + 1][1]))) > longest_number:
                longest_number = len(str(int(array[row + 1][1])))
        else:
            if len(str(array[row + 1][1])) > longest_number:
                longest_number = len(str(array[row + 1][1]))

    for row in range(len(array) - 1):
        print("| ", end="")
        print_column_element(array[row + 1][0], longest_name_length)
        print(" | ", end="")
        if option == 1:
            if len(str(array[row + 1][1])) < len(ODT_LABEL):
                print_digit_short(array[row + 1][1], len(ODT_LABEL))
            else:
                print_digit_long(array[row + 1][1], len(ODT_LABEL) - 1)
        elif option == 2:
            if len(str(int(array[row + 1][1]))) < len(TOTAL_FEEDS_LABEL):
                print_digit_short(int(array[row + 1][1]), len(TOTAL_FEEDS_LABEL))
            else:
                print_digit_long(int(array[row + 1][1]), len(TOTAL_FEEDS_LABEL) - 1)
        else:
            if len(str(array[row + 1][1])) < len(AVERAGE_SETUP_TIME_LABEL):
                print_digit_short(array[row + 1][1], len(AVERAGE_SETUP_TIME_LABEL))
            else:
                print_digit_long(array[row + 1][1], len(AVERAGE_SETUP_TIME_LABEL) - 1)
        print(" |")

    if option == 1:
        for dash in range(longest_name_length + len(ODT_LABEL) + 7):
            print("-", end="")
    elif option == 2:
        for dash in range(longest_name_length + len(TOTAL_FEEDS_LABEL) + 7):
            print("-", end="")
    else:
        for dash in range(longest_name_length + len(AVERAGE_SETUP_TIME_LABEL) + 7):
            print("-", end="")
    print()


# function to print a column element according to a specific number of letters
# arguments: the word to print as a string
# and the length of the longest column element
# returns nothing
def print_column_element(word, max_length):
    # determine length of current column element first
    word_length = len(word)

    for space in range(int((max_length - word_length + 1) / 2)):
        print(" ", end="")
    print(word, end="")
    if (max_length - word_length) % 2 == 0:
        for space in range(int((max_length - word_length + 1) / 2)):
            print(" ", end="")
    else:
        for space in range(int((max_length - word_length + 1) / 2) - 1):
            print(" ", end="")


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
        for dash in range(longest_charge_code_length + len(ODT_LABEL) + 7):
            print("-", end="")
    else:
        for dash in range(len(CHARGE_CODE_LABEL) + len(ODT_LABEL) + 7):
            print("-", end="")

    print("\n| ", end="")
    if longest_charge_code_length > len(CHARGE_CODE_LABEL):
        for space in range(int((longest_charge_code_length - len(CHARGE_CODE_LABEL)) / 2)):
            print(" ", end="")
    print(CHARGE_CODE_LABEL, end="")
    if longest_charge_code_length > len(CHARGE_CODE_LABEL):
        for space in range(int((longest_charge_code_length - len(CHARGE_CODE_LABEL)) / 2)):
            print(" ", end="")

    print(" | " + ODT_LABEL + " |")

    if longest_charge_code_length > len(CHARGE_CODE_LABEL):
        for dash in range(longest_charge_code_length + len(ODT_LABEL) + 7):
            print("-", end="")
    else:
        for dash in range(len(CHARGE_CODE_LABEL) + len(ODT_LABEL) + 7):
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


# function to ask user to remove
# holidays and other days with zero feeds for all 3 shifts
# in the resulting table for feeds by day
# arguments: none
# returns either True or False
def remove_holidays():
    while True:
        option = input("Would you like to remove holidays and days off from the table (y/n)? ")
        if option == "y" or option == "Y":
            return True
        elif option == "n" or option == "N":
            return False
        else:
            print("Please try again.")


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
            print_column_element("Shift " + str(shift + 1), longest_number_len)
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
                print_digit_long(feeds_per_day_array[row + 1][col + 1], 7)
            else:
                print_digit_short(feeds_per_day_array[row + 1][col + 1], 7)
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
            print_column_element(crew, longest_number_len)
        else:
            print_column_element(crew, longest_name_len)
        print(" | ", end="")
    print()

    print_dashes_by_crew(longest_number_len, longest_name_len, list_of_crews)

    for row in range(length_of_table - 1):
        print("| " + feeds_per_day_array[row + 1][0] + " | ", end="")

        for col in range(len(list_of_crews)):
            if longest_number_len > longest_name_len:
                print_digit_long(feeds_per_day_array[row + 1][col + 1], longest_number_len)
            else:
                print_digit_short(feeds_per_day_array[row + 1][col + 1], longest_name_len)
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


# function for displaying both the list of rows where no crew name could be attributed by the AI
# and the list of rows with negative elapsed hours
# arguments: whether the algorithm was used, the list of rows with no names
# and the list of rows with negative elapsed hours
def display_additional_info(algorithm_used, rows_with_no_name, negative_num_rows):
    # print list of rows where no name was attributed
    if algorithm_used:
        print_rows_with_no_name(rows_with_no_name)
        if len(negative_num_rows) <= 0 and len(rows_with_no_name) > 0:
            print("-------------------------------------------------------------")

    # print list of rows with negative elapsed hours
    if len(negative_num_rows) > 0:
        sorting_algorithm(negative_num_rows)
        if len(rows_with_no_name) > 0:
            print("-------------------------------------------------------------")
        print_neg_nums_list(len(negative_num_rows), negative_num_rows)


# function to display ODT by either shift or crew
# arguments: the detailed job report array, which type of ODT the user wants,
# the start date and the end date
# returns nothing
def display_ODT(detailed_job_report, user_option, start_date, end_date):
    negative_num_rows = []
    start_date_num = int(start_date[0:4] + start_date[5:7] + start_date[8:10])
    end_date_num = int(end_date[0:4] + end_date[5:7] + end_date[8:10])

    if user_option == "1": # user wants ODT by shift
        print("-------------------")
        print("| Shift | ODT (%) |")
        print("-------------------")

        # create & initialize array to hold data in case user wants to write data to an Excel spreadsheet
        ODT_by_shift_array = [[0 for x in range(2)] for y in range(4)]
        ODT_by_shift_array[0][0], ODT_by_shift_array[0][1] = "Shift", "ODT %"
        counter = 1

        for shift in range(3):
            # calculate total machine hours for specific shift
            total_machine_hours = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        total_machine_hours = total_machine_hours + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Run" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

            # calculate total open downtime for specific shift
            ODT = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        ODT = ODT + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    elif detailed_job_report[row][DOWNTIME_COL_NUM] == "Open Downtime" and str(detailed_job_report[row][SHIFT_COL_NUM]) == str(shift + 1) and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

            if total_machine_hours != 0:
                result = (ODT/total_machine_hours) * 100
                print("|   " + str(shift + 1) + "   | ", end="")
                print_digit_long(result, 6)
                print(" |")

                ODT_by_shift_array[counter][0], ODT_by_shift_array[counter][1] = shift + 1, result
                counter = counter + 1

        print("-------------------")

        if len(negative_num_rows) > 0:
            sorting_algorithm(negative_num_rows)
            print_neg_nums_list(len(negative_num_rows), negative_num_rows)

        if to_excel():
            write_to_excel(ODT_by_shift_array, len(ODT_by_shift_array))

    elif user_option == "2": # user wants ODT by crew
        # check if there are gaps in data
        gap_name = False
        for row in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                    gap_name = True
                    break

        # ask user if they want to use name-filling algorithm
        use_algo = False
        if gap_name:
            use_algo = use_algorithm()
        rows_with_no_name = []

        # create list of crews
        crews_list = []
        generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list)

        # create & initialize array to hold data in case user wants to write data to an Excel spreadsheet
        ODT_by_crew_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
        ODT_by_crew_array[0][0], ODT_by_crew_array[0][1] = "Crew", "ODT %"
        counter = 1

        longest_name_len = print_crew_header(crews_list, 1)

        for crew in crews_list:
            # calculate total machine hours for specific crew
            total_machine_hours = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew:
                        if str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Run" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                            total_machine_hours = total_machine_hours + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                        elif str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Run" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                            append_neg_num_row(negative_num_rows, row)

            if use_algo:
                total_machine_hours = name_filling_algorithm(detailed_job_report, total_machine_hours, crew, start_date_num, end_date_num, rows_with_no_name, negative_num_rows, 1)

            # calculate total open downtime for specific crew
            ODT = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew:
                        if str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Open Downtime" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                            ODT = ODT + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                        elif str(detailed_job_report[row][DOWNTIME_COL_NUM]) == "Open Downtime" and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                            append_neg_num_row(negative_num_rows, row)

            if use_algo:
                ODT = name_filling_algorithm(detailed_job_report, ODT, crew, start_date_num, end_date_num, rows_with_no_name, negative_num_rows, 2)

            # add total run time with total ODT for total_machine_hours
            total_machine_hours = total_machine_hours + ODT

            # print(str(ODT) + " " + str(total_machine_hours))
            result = 0
            if total_machine_hours > 0:
                result = (ODT/total_machine_hours) * 100

            if result != 0:
                ODT_by_crew_array[counter][0], ODT_by_crew_array[counter][1] = crew, result
            counter = counter + 1

        print_rest_of_table(ODT_by_crew_array, longest_name_len, 1)

        # breakdown of additional info
        display_additional_info(use_algo, rows_with_no_name, negative_num_rows)

        if to_excel():
            write_to_excel(ODT_by_crew_array, len(ODT_by_crew_array))

    else: # user wants Pareto chart
        # create list of charge codes
        charge_code_list = []
        for row in range(ROWS):
            alreadyIncluded = False
            for charge_code in charge_code_list:
                if detailed_job_report[row][CHARGE_CODE_COL_NUM] == charge_code:
                    alreadyIncluded = True
                    break

            if not alreadyIncluded and detailed_job_report[row][CHARGE_CODE_COL_NUM] != "RUN" and detailed_job_report[row][CHARGE_CODE_COL_NUM] != "SET UP":
                charge_code_list.append(detailed_job_report[row][CHARGE_CODE_COL_NUM])

        # calculate total open downtime for all charge codes
        totalODT = 0
        for row in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                include = False
                for charge_code in charge_code_list:
                    if detailed_job_report[row][CHARGE_CODE_COL_NUM] == charge_code:
                        include = True

                if include and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                    totalODT = totalODT + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                elif include and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                    append_neg_num_row(negative_num_rows, row)

        # temporary dictionary to contain each charge code and their corresponding ODT
        charge_code_dict = {}
        for charge_code in charge_code_list:
            # calculate total open downtime for specific charge code
            ODT = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][CHARGE_CODE_COL_NUM] == charge_code and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        ODT = ODT + detailed_job_report[row][ELAPSED_HOURS_COL_NUM]
                    if detailed_job_report[row][CHARGE_CODE_COL_NUM] == charge_code and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

            charge_code_dict[charge_code] = ODT

        num_rows = 0
        for key in charge_code_dict.keys():
            if charge_code_dict[key] > 0:
                num_rows = num_rows + 1

        charge_code_array = [[0 for x in range(2)] for y in range(num_rows + 1)]
        charge_code_array[0][0], charge_code_array[0][1] = "Charge Code", "ODT (%)"

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
        print_rest_of_table(charge_code_array, longest_charge_code_len, 1)
        # print(charge_code_array)

        if len(negative_num_rows) > 0:
            sorting_algorithm(negative_num_rows)
            print_neg_nums_list(len(negative_num_rows), negative_num_rows)

        if to_excel():
            write_to_excel(charge_code_array, len(charge_code_array))


# function to display total feeds either by shift or by crew
# arguments: the detailed job report as an array, whether the user wants
# total feeds by shift or by crew, the start date as an integer and
# the end date as an integer
# returns nothing
def display_total_feeds(detailed_job_report, user_choice, start_date_num, end_date_num):
    if user_choice == "1": # user wants total feeds by shift
        print("-----------------------")
        print("| Shift | Total Feeds |")
        print("-----------------------")
        negative_num_rows = []
        total_feeds_by_shift_array = [[0 for x in range(2)] for y in range(4)]
        total_feeds_by_shift_array[0][0], total_feeds_by_shift_array[0][1] = "Shift", "Total Feeds"

        counter = 1
        for shift in range(3):
            total_feeds = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[row][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        total_feeds = total_feeds + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                    elif detailed_job_report[row][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

            if total_feeds > 0:
                total_feeds_by_shift_array[counter][0], total_feeds_by_shift_array[counter][1] = shift + 1, total_feeds
            counter = counter + 1

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
                    print_digit_short(int(total_feeds_by_shift_array[row + 1][1]), len(TOTAL_FEEDS_LABEL))
                else:
                    print_digit_long(int(total_feeds_by_shift_array[row + 1][1]), len(TOTAL_FEEDS_LABEL))
                print(" |")

        if not table_empty:
            print("-----------------------")

        if len(negative_num_rows) > 0:
            sorting_algorithm(negative_num_rows)
            print_neg_nums_list(len(negative_num_rows), negative_num_rows)

        if to_excel():
            write_to_excel(total_feeds_by_shift_array, len(total_feeds_by_shift_array))

    else: # user wants total feeds by crew
        # check for gaps in Employee Name column
        gap_name = False
        for row in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                    gap_name = True
                    break

        # ask if user would like to use AI to compute
        use_algo = False
        if gap_name:
            use_algo = use_algorithm()

        crews_list = []
        generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list)
        longest_name_len = print_crew_header(crews_list, 2)

        empty_name_rows = []
        negative_num_rows = []

        total_feeds_by_crew_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
        total_feeds_by_crew_array[0][0], total_feeds_by_crew_array[0][1] = "Crew", "Total Feeds"

        counter = 1
        for crew in crews_list:
            total_feeds = 0
            for row in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] >= 0:
                        total_feeds = total_feeds + detailed_job_report[row][GROSS_FG_QTY_COL_NUM]
                    elif str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == crew and detailed_job_report[row][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, row)

            if use_algo:
                total_feeds = name_filling_algorithm(detailed_job_report, total_feeds, crew, start_date_num, end_date_num, empty_name_rows, negative_num_rows, 4)

            total_feeds_by_crew_array[counter][0], total_feeds_by_crew_array[counter][1] = crew, total_feeds
            counter = counter + 1

        print_rest_of_table(total_feeds_by_crew_array, longest_name_len, 2)

        display_additional_info(use_algo, empty_name_rows, negative_num_rows)

        if to_excel():
            write_to_excel(total_feeds_by_crew_array, len(total_feeds_by_crew_array))


# function to display average setup time either by shift or by crew
# arguments: the detailed job report as an array, whether the user wants
# average setup time by shift or by crew, the start date as an integer and
# the end date as an integer
# returns nothing
def display_average_setup_time(detailed_job_report, user_choice, start_date_num, end_date_num):
    if user_choice == "1":  # user wants general average setup time
        # total elapsed hours
        negative_num_rows = [] # list to track rows with negative elapsed hours
        total_elapsed_hours = 0  # counter for total elapsed hours

        for elapsed_hours in range(ROWS):
            if detailed_job_report[elapsed_hours][DOWNTIME_COL_NUM] == "Setup":
                if start_date_num <= int(str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    # error checking for negative times
                    if detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] < 0:
                        append_neg_num_row(negative_num_rows, elapsed_hours)
                    else:
                        total_elapsed_hours = total_elapsed_hours + detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM]

        print("\nTotal Elapsed Hours: " + str(total_elapsed_hours))

        if len(negative_num_rows) > 0:
            sorting_algorithm(negative_num_rows)
            print_neg_nums_list(len(negative_num_rows), negative_num_rows)

        # total orders
        unique_orders_list = []  # list to track all unique orders

        for order in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[order][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[order][WORK_DATE_COL_NUM])[5:7] + str(djr_array[order][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                if detailed_job_report[order][DOWNTIME_COL_NUM] == "Setup":
                    unique = True
                    for item in unique_orders_list:
                        unique = True
                        if detailed_job_report[order][ORDER_NUM_COL_NUM] == item:
                            unique = False
                            break
                    if unique:
                        unique_orders_list.append(detailed_job_report[order][ORDER_NUM_COL_NUM])

        print("Total Unique Orders: " + str(len(unique_orders_list)))

        # final calculation
        if len(unique_orders_list) != 0:
            average_setup_time = total_elapsed_hours / len(unique_orders_list)
            print("Average Setup Time: " + str(average_setup_time * 60) + " minutes")  # display average setup time in minutes

        # display list of negative elapsed hours
        if len(negative_num_rows) > 0:
            sorting_algorithm(negative_num_rows)
            print_neg_nums_list(len(negative_num_rows), negative_num_rows)

    else:  # user wants average setup time by crew
        negative_num_rows = []  # list to track rows with negative elapsed hours

        # check if there are gaps in data
        gap_name = False
        for row in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(djr_array[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                    gap_name = True
                    break

        # ask user if they want to use name-filling algorithm
        use_algo = False
        if gap_name:
            use_algo = use_algorithm()
        rows_with_no_name = []

        crews_list = []
        # create list of crews
        generate_crews_list(djr_array, start_date_num, end_date_num, crews_list)

        longest_name_len = print_crew_header(crews_list, 3)

        # create and initialize array to write to Excel
        average_setup_time_by_crew_array = [[0 for x in range(2)] for y in range(len(crews_list) + 1)]
        average_setup_time_by_crew_array[0][0], average_setup_time_by_crew_array[0][1] = "Crew", "Average Setup Time"

        counter = 1
        for crew in crews_list:
            # total elapsed hours
            total_elapsed_hours = 0  # counter for total elapsed hours

            for elapsed_hours in range(ROWS):
                if detailed_job_report[elapsed_hours][DOWNTIME_COL_NUM] == "Setup":
                    if start_date_num <= int(str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[elapsed_hours][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                        # error checking for negative times
                        if detailed_job_report[elapsed_hours][EMPLOYEE_NAME_COL_NUM] == crew and detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] < 0:
                            append_neg_num_row(negative_num_rows, elapsed_hours)
                        elif detailed_job_report[elapsed_hours][EMPLOYEE_NAME_COL_NUM] == crew and detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM] >= 0:
                            total_elapsed_hours = total_elapsed_hours + detailed_job_report[elapsed_hours][ELAPSED_HOURS_COL_NUM]

            if use_algo:
                total_elapsed_hours = name_filling_algorithm(detailed_job_report, total_elapsed_hours, crew, start_date_num, end_date_num, rows_with_no_name, negative_num_rows, 3)

            # total orders
            total_unique_orders = 0
            unique_orders_list = []  # list to track all unique orders

            for order in range(ROWS):
                if start_date_num <= int(str(detailed_job_report[order][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[order][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[order][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                    if detailed_job_report[order][DOWNTIME_COL_NUM] == "Setup":
                        if detailed_job_report[order][EMPLOYEE_NAME_COL_NUM] == crew:
                            unique = True
                            for item in unique_orders_list:
                                unique = True
                                if detailed_job_report[order][ORDER_NUM_COL_NUM] == item:
                                    unique = False
                                    break
                            if unique:
                                unique_orders_list.append(detailed_job_report[order][ORDER_NUM_COL_NUM])
                                # print(detailed_job_report[order][ORDER_NUM_COL_NUM])

                        elif use_algo and str(detailed_job_report[order][EMPLOYEE_NAME_COL_NUM]) == "nan":
                            # first determine which employee name should fill this gap
                            assumed_name = assume_name(djr_array, rows_with_no_name, order)

                            if assumed_name == crew:
                                unique = True
                                for item in unique_orders_list:
                                    unique = True
                                    if detailed_job_report[order][ORDER_NUM_COL_NUM] == item and detailed_job_report[order][ELAPSED_HOURS_COL_NUM] >= 0:
                                        unique = False
                                        break
                                if unique:
                                    unique_orders_list.append(detailed_job_report[order][ORDER_NUM_COL_NUM])
                                    # print(detailed_job_report[order][ORDER_NUM_COL_NUM])

            # final calculation
            total_unique_orders = len(unique_orders_list)
            if total_unique_orders != 0:
                # print(str(total_elapsed_hours) + " " + str(total_unique_orders))
                average_setup_time = (total_elapsed_hours / total_unique_orders) * 60
                average_setup_time_by_crew_array[counter][0], average_setup_time_by_crew_array[counter][1] = crew, average_setup_time
            counter = counter + 1

        print_rest_of_table(average_setup_time_by_crew_array, longest_name_len, 3)

        display_additional_info(use_algo, rows_with_no_name, negative_num_rows)

        if to_excel():
            write_to_excel(average_setup_time_by_crew_array, len(average_setup_time_by_crew_array))


# function to display chart for feeds per day either by shift or by crew
# arguments: the detailed job report as an array, whether the user wants
# feeds per day by shift or by crew, the start date as an integer and
# the end date as an integer
# returns nothing
def display_feeds_per_day(detailed_job_report, user_choice, start_date_num, end_date_num):
    negative_num_rows = []

    if user_choice == "1": # user wants to display feeds per day by shift
        resulting_table = [[0 for x in range(4)] for y in range(int(end_date_num - start_date_num + 2))]
        resulting_table[0][0], resulting_table[0][1], resulting_table[0][2], resulting_table[0][3] = "Work Date", "Shift 1", "Shift 2", "Shift 3"

        for row_table in range(len(resulting_table) - 1):
            resulting_table[row_table + 1][0] = convert_date_int_to_string(start_date_num + row_table)
            for shift in range(3):
                total_feeds = 0
                for row_djr in range(ROWS):
                    if str(start_date_num + row_table) == str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[8:10]:
                        if detailed_job_report[row_djr][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] >= 0:
                            total_feeds = total_feeds + detailed_job_report[row_djr][GROSS_FG_QTY_COL_NUM]
                        elif detailed_job_report[row_djr][SHIFT_COL_NUM] == shift + 1 and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] < 0:
                            append_neg_num_row(negative_num_rows, row_djr)

                if total_feeds > 0:
                    resulting_table[row_table + 1][shift + 1] = int(total_feeds)

        table_length = len(resulting_table)

        if remove_holidays():
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

        if len(negative_num_rows) > 0:
            sorting_algorithm(negative_num_rows)
            print_neg_nums_list(len(negative_num_rows), negative_num_rows)

    else: # user wants to display feeds per day by crew
        # check if there are gaps in data
        gap_name = False
        for row in range(ROWS):
            if start_date_num <= int(str(detailed_job_report[row][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row][WORK_DATE_COL_NUM])[8:10]) <= end_date_num:
                if str(detailed_job_report[row][EMPLOYEE_NAME_COL_NUM]) == "nan":
                    gap_name = True
                    break

        # ask user if they want to use name-filling algorithm
        use_algo = False
        if gap_name:
            use_algo = use_algorithm()
        empty_name_rows = []

        crews_list = []
        generate_crews_list(detailed_job_report, start_date_num, end_date_num, crews_list)

        resulting_table = [[0 for x in range(len(crews_list) + 1)] for y in range(int(end_date_num - start_date_num + 2))]
        resulting_table[0][0] = "Work Date"
        for col in range(len(crews_list)):
            resulting_table[0][col + 1] = crews_list[col]

        for row_table in range(len(resulting_table) - 1):
            resulting_table[row_table + 1][0] = convert_date_int_to_string(start_date_num + row_table)
            crew_counter = 1
            for crew in crews_list:
                total_feeds = 0
                for row_djr in range(ROWS):
                    if str(start_date_num + row_table) == str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[0:4] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[5:7] + str(detailed_job_report[row_djr][WORK_DATE_COL_NUM])[8:10]:
                        if str(detailed_job_report[row_djr][EMPLOYEE_NAME_COL_NUM]) == crew and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] >= 0:
                            total_feeds = total_feeds + detailed_job_report[row_djr][GROSS_FG_QTY_COL_NUM]
                        elif str(detailed_job_report[row_djr][EMPLOYEE_NAME_COL_NUM]) == crew and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] < 0:
                            append_neg_num_row(negative_num_rows, row_djr)
                        elif use_algo and str(detailed_job_report[row_djr][EMPLOYEE_NAME_COL_NUM]) == "nan":
                            # first determine which employee name should fill this gap
                            assumed_name = assume_name(detailed_job_report, empty_name_rows, row_djr)

                            if assumed_name == crew and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] >= 0:
                                total_feeds = total_feeds + detailed_job_report[row_djr][GROSS_FG_QTY_COL_NUM]
                            elif assumed_name == crew and detailed_job_report[row_djr][ELAPSED_HOURS_COL_NUM] < 0:
                                append_neg_num_row(negative_num_rows, row_djr)

                if total_feeds > 0:
                    resulting_table[row_table + 1][crew_counter] = int(total_feeds)
                crew_counter = crew_counter + 1

        table_length = len(resulting_table)

        if remove_holidays():
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

        display_additional_info(use_algo, empty_name_rows, negative_num_rows)

    if to_excel():
        write_to_excel(resulting_table, table_length)


######################################################################
# obtaining the Detailed Job Report (.xlsx) spreadsheet from directory
######################################################################
djr_array = [] # initialize detailed job report array
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

        djr_dataframe = pd.read_excel(djr_name)
        djr_array = djr_dataframe.to_numpy()  # convert data frame to an array
        ROWS, COLUMNS = djr_array.shape

        file_found = True
        print("File found.")
    except:
        print("Error: no such file found in current directory")
        print("Please try again")


user_done = False
while not user_done:
    user_input = obtain_instructions()
    first_date_string = ""
    second_date_string = ""

    #######################################
    # calculating open down time percentage
    #######################################
    if user_input == 1:
        print("Enter the first date: ", end="")
        first_date_string = obtain_date_string(djr_array)

        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        error = True
        choice = " "
        while error:
            choice = input("Enter (1) to calculate ODT by shift, (2) to calculate ODT by crew or (3) to produce Pareto chart: ")

            if choice == "1" or choice == "2" or choice == "3":
                error = False
            else:
                print("Please try again.")

        if choice == "1":
            print("\nOpen down time from " + first_date_string + " to " + second_date_string + " based on shift:")
        elif choice == "2":
            print("\nOpen down time from " + first_date_string + " to " + second_date_string + " based on crew:")
        else:
            print("\nPareto chart of open down time from " + first_date_string + " to " + second_date_string + ":")

        display_ODT(djr_array, choice, first_date_string, second_date_string) # show calculated data

    #########################
    # calculating total feeds
    #########################
    elif user_input == 2:
        # obtain date frame from user
        print("Enter the first date: ", end="")
        first_date_string = obtain_date_string(djr_array)

        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        error = True
        user_choice = ""
        while error:
            user_choice = input("Enter (1) to calculate total feeds by shift or (2) to total feeds by crew: ")

            if user_choice == "1" or user_choice == "2":
                error = False
            else:
                print("Error: please try again.")

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
        print("Enter the first date: ", end="")
        first_date_string = obtain_date_string(djr_array)

        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        error = True
        user_choice = " "
        while error:
            user_choice = input("Enter (1) to calculate general average setup time or (2) to calculate average setup time by crew: ")

            if user_choice == "1" or user_choice == "2":
                error = False
            else:
                print("Error: please try again.")

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
        print("Enter the first date: ", end="")
        first_date_string = obtain_date_string(djr_array)

        second_date_string = obtain_second_date_string(djr_array, first_date_string)

        start_date_num = int(first_date_string[0:4] + first_date_string[5:7] + first_date_string[8:10])
        end_date_num = int(second_date_string[0:4] + second_date_string[5:7] + second_date_string[8:10])

        error = True
        user_choice = " "
        while error:
            user_choice = input("Enter (1) to display feeds per day by shift or (2) to display feeds per day by crew: ")

            if user_choice == "1" or user_choice == "2":
                error = False
            else:
                print("Error: please try again.")

        if user_choice == "1":
            print("\nFeeds per day by shift from " + first_date_string + " to " + second_date_string + ":")
        else:
            print("\nFeeds per day by crew from " + first_date_string + " to " + second_date_string + ":")

        display_feeds_per_day(djr_array, user_choice, start_date_num, end_date_num)

    ######
    # exit
    ######
    elif user_input == 5:
        exit()
