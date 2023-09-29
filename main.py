# Importing all the important libraries
from pandas import read_excel as rdxl
from pandas._libs.tslibs.timestamps import Timestamp as Ts


def attendance():
	"""This function prints the name of the employees who has attended office for 7 consecutive days. It first appends all the dates into an array and then iterate over this array to get the difference between the two elements(dates).
	for difference=1, it increments the attendance_counter by 1
	for difference=0, it skips any operation
	for difference>1 and attendance_counter>0, it makes the counter 0 again.
	When the counter reach 7 it breaks check for further dates and return the name and position of the employee back to the attendance function."""

	global data
	i = 0
	while i < data['Position ID'].count()-1:
		temp = data['Position ID'][i]  # position id for which all the dates should be checked
		attend_count = 0  # attendance_counter to check if employee has worked for 7 consecutive days
		attend_arr = []  # attendance_array to collect all the dates employee worked on

		while True:
			if type(data['Time'][i]) == Ts and data['Position ID'][i] == temp:  # check whether the value is a valid time stamp or not
				attend_arr.append(data['Time'][i].date())  # appending to and array if valid timestamp
				i += 1
			elif type(data['Time'][i]) != Ts and data['Position ID'][i] == temp:  # does nothing if it is not valid timestamp but is the same position id
				i += 1
			elif data['Position ID'][i] != temp:  # break the loop in order to display the result if iteration starts over another position id
				break
			else:
				i += 1
			if i == data['Position ID'].count()-1:  # break the loop if it reached the end of data
				break

		for d in range(len(attend_arr) - 1, 1, -1):  # iterating over the array to calculate difference between the sequential dates stored in it
			if (attend_arr[d] - attend_arr[d - 1]).days == 1:  # if diff is 1 day, increase attendance counter
				attend_count += 1
				if attend_count == 7:  # break the loop if attendance counter reaches 7
					break
			elif (attend_arr[d] - attend_arr[d - 1]).days == 0:  # if diff is 0 day, skip any operation
				continue
			else:  # if diff is >1 day, make attendance counter 0 again
				attend_count = 0
		if attend_count >= 7:  # if attendance counter is 7, display the employee name and position id
			print(f'Name: {data["Employee Name"][i-1]} | Position ID: {data["Position ID"][i-1]}')


def between_shifts():
	"""This function display all the employees who have taken more than 1hr and less than 10hr gap between their shifts. We would check the Time Out time for first shift and Time tie for second shift of the same day and then find the difference, which should be between 1-10."""
	global data
	i = 0
	while i < data['Position ID'].count()-3:  # iterating over rows
		temp = data['Position ID'][i]  # will check duration between shift for this employee
		while True:
			shift_end_time = shift_start_time = 0  # var for storing the timeout and time, time for the same day
			if data['Position ID'][i] == temp and data['Position ID'][i+1] == temp:  # check that we are checking time for the same employee
				if (type(data['Time Out'][i]) == Ts and type(data['Time'][i+1]) == Ts) and (data['Time Out'][i].date() == data['Time'][i+1].date()):  # check that the value is valid timestamp and the two times are of same date
					shift_end_time = data['Time Out'][i]
					shift_start_time = data['Time'][i+1]
					if 1 < (shift_start_time.time().hour + shift_start_time.time().minute/60) - (shift_end_time.time().hour + shift_end_time.time().minute/60) < 10:  # evaluate the difference and if in range 1-10, display the employee details
						print(f'Name: {data["Employee Name"][i]} | Position ID: {data["Position ID"][i]}')
						i += 2
						break  # we don't need to check further for this employee, so we break the loop
					else:  # if time difference is not in desired range, check difference for another day
						i += 2
				elif (type(data['Time Out'][i]) != Ts or type(data['Time'][i+1]) != Ts) and (data['Time Out'][i].date() == data['Time'][i+1].date()):  # if any of the time value is not valid timestamp but the dates are same, we start checking after skipping two entries for the same employee
					i += 2
				elif (type(data['Time Out'][i]) == Ts and type(data['Time'][i+1]) == Ts) and (data['Time Out'][i].date() != data['Time'][i+1].date()):  # if it is a valid timestamp for both but entries are of different dates, we start checking after skipping one entry for the same employee
					i += 1
				else:  # if nothing above is true, just read another entry
					i += 1
			elif data['Position ID'][i] != temp or data['Position ID'][i+1] != temp:  # if we read all the entries of one employee break the loop to read entries of other employees
				i += 1
				break
			else:  # if nothing above get satisfied, just read another entry
				i += 1
			
	
def over_work():
	"""This function display the employees who have worked for more than 14 hours in a single shift. We would just compare the value of Timecard column with 14, for greater or equal to it."""
	global data
	for i in range(data['Position ID'].count()):  # reading all the rows of Excel sheet
		if type(data['Time'][i]) == Ts and int(data['Timecard Hours (as Time)'][i].split(':')[0]) >= 14:  # checking whether value is valid timestamp and user and worked for more than 14hours
			print(f'Name: {data["Employee Name"][i]}| Position ID: {data["Position ID"][i]}')
					

def main():
	"""This is the driver function which would be display the menu and call different function according to the option selected."""
	global data
	menu_choice = input('\n-------------------------------------------------------\n1. Employees worked for 7 consecutive days\n2. Employees with gap in range 1 to 10 between shifts\n3. Employees worked more than 14 hours in single shift\nQ. Exit\n: ')  # displaying the option to choose the operation
	# calling different functions below according to the option chosen above
	if menu_choice == '1':
		attendance()
	elif menu_choice == '2':
		between_shifts()
	elif menu_choice == '3':
		over_work()
	elif menu_choice.upper() == 'Q':
		exit()
	else:
		main()


# driver code
inputFile = 'Assignment_Timecard.xlsx'
data = rdxl(inputFile)  # reading the Excel sheet
while True:
	main()
