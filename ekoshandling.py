'''
Copyright (c) 2022 Paul Marichal

permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
of the Software, and to permit persons to whom the Software is furnished to do
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''


#  pylint: disable=W0312, C0103

import os
import sys
import platform
import time
import pandas as pd
from datetime import date

InputDirectoryPath = './inputorder'
OutputDirectoryPath = './output'

''' this method processes Ekos report and creates spreasheet of items that needs to be ordered '''


def create_ingredients_order_excel(InputDirectoryPath, OutputDirectoryPath):
	try:  # get a list of all files with extension .xlsx from the input directory
		excel_names = []
		#  r=root, d=directories, f = files
		for r, d, f in os.walk(InputDirectoryPath):
			for file in f:
				if '.xlsx' in file:
					excel_names.append(os.path.join(r, file))
		if len(excel_names) == 0:
			print('\nNo files available to process')
			return 1
		# build input filename
		infilename = InputDirectoryPath + '/' + f[0]
		print(infilename)
		# build output filename
		outfilename = OutputDirectoryPath + "/order"
		# get today's date, so we can append it to output filename
		today = date.today()
		#  read the input from Ekos
		mydf = pd.read_excel(infilename, skiprows=12, header=None)
		# rename all column headers in dataframe
		mydf.rename(columns={0: 'Description', 1: "Required Quantity", 2: "UOM1", 3: "Inventory Quantity", 4: 'UOM2'}, inplace=True)
		# insert new column for calculation
		mydf.insert(5, 'Order Quantity', "NA")
		# loop thru each row in dataframe and subtract values
		loop = 0
		columnSeriesObjRequired = mydf['Required Quantity']
		columnSeriesObjInventory = mydf['Inventory Quantity']
		columnSeriesObjOrder = mydf['Order Quantity']
		#  loop thru row list so we can add day number.
		for i in columnSeriesObjRequired:
			required = columnSeriesObjRequired.values[loop]
			inventory = columnSeriesObjInventory.values[loop]
			# if value is not an integer the row don;t subtract
			if isinstance(required, int):
				if (inventory - required) < 0:
					columnSeriesObjOrder.values[loop] = inventory - required
			loop = loop + 1
		print(mydf)
		# get rid of all the entries that did not get updated
		mydf = mydf[~mydf['Order Quantity'].isin(['NA'])]
		#  create excel writer object
		d4 = today.strftime("%b-%d-%Y")
		# print("d4 =", d4)
		writer = pd.ExcelWriter(outfilename + d4 + '.xlsx', engine='xlsxwriter')
		mydf.to_excel(writer, index=False)
		#  Set the column width and format.
		#  Get the xlsxwriter workbook and worksheet objects.
		# workbook = writer.book
		worksheet = writer.sheets['Sheet1']
		worksheet.set_column('A:A', 40)
		worksheet.set_column('B:F', 20)
		# save Excel file
		writer.save()
		return 1
	except Exception as e:
		print("Entry in Ekos file is not valid")
		print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)
		# print("Exception", e)
		return (0)


''' this method processes Ekos report and creates spreadsheet of updated hop alpha '''


def update_hop_tracking_excel(InputDirectoryPath, OutputDirectoryPath):
	try:  # get a list of all files with extension .xlsx from the input directory
		# go get Ekos report filename from the inputhoptracing directory
		excel_names = []
		#  r=root, d=directories, f = files
		for r, d, f in os.walk(InputDirectoryPath):
			for file in f:
				if '.xlsx' in file:
					excel_names.append(os.path.join(r, file))
		if len(excel_names) == 0:
			print('\nNo files available to process')
			return 1
		# build input filename
		ekosfilename = InputDirectoryPath + '/' + f[0]
		# build output filename
		outSPfilename = OutputDirectoryPath + "/Hops Alpha Worksheet"
		# get today's date so we can append it to output filename
		# today = date.today()
		#  read the input from Ekos
		ekosdf = pd.read_excel(ekosfilename, header=None)
		newDate = ekosdf.iloc[5, 0]
		split_string = newDate.split(":", 1)
		newDate = split_string[1]
		# download SharePoint file into temporary working area
		excel_names = []
		#  r=root, d=directories, f = files
		for r, d, f in os.walk('./sharepointtemp/'):
			for file in f:
				if '.xlsx' in file:
					excel_names.append(os.path.join(r, file))
		if len(excel_names) == 0:
			print('\nNo files available to process')
			return 1
		# build input filename
		# ORIGINAL CODE to scan dierectory sharepointfilename = './sharepointtemp/' + '/' + f[0]
		# hard code the filename since it doesn't change
		sharepointfilename = './sharepointtemp/' + '/' + 'Hops Alpha Worksheet.xlsx'

		#  read the input from Ekos report
		sharepointdf = pd.read_excel(sharepointfilename, header=None)
		#  create pandas dataframes for both files that will be worked on
		# drop the top part of the ekos file
		modekosdf = ekosdf.drop(labels=range(0, 8), axis=0)
		modekosdf.rename(columns={0: 'Location', 1: 'Title', 2: "Lot"}, inplace=True)
		# drop the first part of the Sharepoint file
		modsharepointdf = sharepointdf.drop(labels=range(0, 4), axis=0)
		modsharepointdf.rename(columns={0: 'Title', 1: 'Lot', 2: "Date", 3: 'AA', 4: 'Foil'}, inplace=True)

		# create dataframe objects that we will use to parse and modify Share-point dataframe
		columnSeriesObjEkosTitle = modekosdf['Title']
		columnSeriesObjEkosLot = modekosdf['Lot']
		columnSeriesObjSPTitle = modsharepointdf['Title']
		columnSeriesObjSPLot = modsharepointdf['Lot']
		columnSeriesObjSPDate = modsharepointdf['Date']
		columnSeriesObjSPAA = modsharepointdf['AA']
		columnSeriesObjSPFoil = modsharepointdf['Foil']
		#  create a dataframe for the final Sharepoint file to be created
		column_names = ['Title', 'Lot', "Date", 'AA', 'Foil']
		finalsharepointdf = pd.DataFrame(columns=column_names)
		#  now look at Ekos file to see what entries are new and add them to SharePoint file
		loop = 0
		total_rows = modsharepointdf.shape[0]
		print('total rows in Sharepoint file', total_rows)
		# this loop pulls all entries from the Ekos report
		for i in columnSeriesObjEkosTitle:
			ekosTitle = columnSeriesObjEkosTitle.values[loop]
			ekosLot = columnSeriesObjEkosLot.values[loop]
			sploop = 0
			# this loop compares each Ekos entry the SharePoint file.
			for i in columnSeriesObjSPTitle:
				SPTitle = columnSeriesObjSPTitle.values[sploop]
				SPLot = columnSeriesObjSPLot.values[sploop]
				if (ekosTitle == SPTitle) and (ekosLot == SPLot):
					break
				# if there's no match for ekos entry we need to add it to new SharePoint dataframe
				elif sploop == total_rows - 1:

					adddict = {'Title': ekosTitle, 'Lot': ekosLot, 'Date': newDate}
					finalsharepointdf = finalsharepointdf.append(adddict, ignore_index=True)
					finalsharepointdf = finalsharepointdf.sort_values(by=['Title'])
					print('Found a new entry in Ekos report', ekosTitle, ekosLot)
				sploop = sploop + 1
			loop = loop + 1

		# now look thru the SharePoint file and delete entries that are no longer in Ekos file
		loop = 0
		total_rows = modekosdf.shape[0]
		print('total rows file in Ekos file', total_rows)
		for i in columnSeriesObjSPTitle:
			SPTitle = columnSeriesObjSPTitle.values[loop]
			SPLot = columnSeriesObjSPLot.values[loop]
			SPDate = columnSeriesObjSPDate.values[loop]
			SPAA = columnSeriesObjSPAA.values[loop]
			SPFoil = columnSeriesObjSPFoil.values[loop]
			sploop = 0
			# print('Looking in SP file for ', SPTitle, SPLot,loop,sploop)
			#  now see if each entry in the SharePoint file is in the new Ekos report.
			for i in columnSeriesObjEkosTitle:
				ekosTitle = columnSeriesObjEkosTitle.values[sploop]
				ekosLot = columnSeriesObjEkosLot.values[sploop]
				# if we find a match, keep the entry and add it to new dataframe
				if (ekosTitle == SPTitle) and (ekosLot == SPLot):
					# print('got one',ekosTitle,ekosLot, SPTitle,SPLot, loop, sploop,total_rows-1)
					adddict = {'Title': SPTitle, 'Lot': SPLot, 'Date': SPDate, 'AA': SPAA, 'Foil': SPFoil}
					finalsharepointdf = finalsharepointdf.append(adddict, ignore_index=True)
					finalsharepointdf = finalsharepointdf.sort_values(by=['Title'])
					break
				sploop = sploop + 1
			loop = loop + 1

		today = date.today()
		# add header info back into SharePoint file
		headerdf = pd.DataFrame({'Title': ['Hops Alpha Tracking By Lot', 'Generated on ', ''],
								'Lot': ['', '', ''],
								'Date': ['', today, ''],
								'AA': ['', '', ''],
								'Foil': ['', '', '']})
		total_rows = finalsharepointdf.shape[0]
		print('total rows in updated Sharepoint file', total_rows)
		# concatenate both dataframes to get final
		finalsharepointdf = headerdf.append(finalsharepointdf, ignore_index=True)
		# now write the newly create dataframe to an output file
		writer = pd.ExcelWriter(outSPfilename + '.xlsx', engine='xlsxwriter')
		finalsharepointdf.to_excel(writer, index=False)
		#  Get the xlsxwriter workbook and worksheet objects.
		# workbook = writer.book
		# worksheet = writer.sheets['Sheet1']
		# save Excel file
		writer.save()
		# now read the Excel file to see what it looks
		# newdf = pd.read_excel(outSPfilename+'.xlsx')

		return 1
	except Exception as e:
		print("Entry in Ekos file is not valid")
		print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)
		# print("Exception", e)
		return 0		


''' this method processes Ekos report and creates spreadsheet to order ingredients '''
def create_ingredients_order_csv(InputDirectoryPath, OutputDirectoryPath):
	try:  # get a list of all files with extension .xlsx from the input directory
		csv_names = []
		#  r=root, d=directories, f = files
		for r, d, f in os.walk(InputDirectoryPath):
			for file in f:
				if '.csv' in file:
					csv_names.append(os.path.join(r, file))
		if len(csv_names) == 0:
			print('\nNo files available to process')
			return 1
		# build input filename
		infilename = InputDirectoryPath + '/' + f[0]
		print(infilename)
		# build output filename
		outfilename = OutputDirectoryPath + "/order"
		# get today's date, so we can append it to output filename
		today = date.today()
		#  read the input from Ekos
		mydf = pd.read_csv(infilename, header=None, skiprows=1, error_bad_lines=False)
		# rename all column headers in dataframe
		mydf.rename(columns={0: 'Description', 1: "Required Quantity", 2: "UOM1", 3: "Inventory Quantity", 4: 'UOM2'}, inplace=True)
		# insert new column for calculation
		mydf.insert(5, 'Order Quantity', "NA")
		# delete rows that have a Required Quantity of NA since they are not valid entries like CSV headers
		mydf = mydf[mydf['Required Quantity'].notna()]
		# the CSV file has multiple duplicate headers for each type of ingredient so delete them
		mydf = mydf.drop_duplicates()
		# loop thru each row in dataframe and subtract values
		loop = 0
		# drop the first row of the dataframe since its a duplicate of the headers
		mydf = mydf.drop([0])
		# now change dtype from string to float for the columns we will do math on
		mydf[['Required Quantity', 'Inventory Quantity']] = mydf[['Required Quantity', 'Inventory Quantity']].apply(pd.to_numeric)
		columnSeriesObjRequired = mydf['Required Quantity']
		columnSeriesObjInventory = mydf['Inventory Quantity']
		columnSeriesObjOrder = mydf['Order Quantity']
		#  loop thru row list so we can add day number.
		for i in columnSeriesObjRequired:
			required = columnSeriesObjRequired.values[loop]
			inventory = columnSeriesObjInventory.values[loop]
			# if value is not a float the row don't try to subtract
			if isinstance(required, float):
				if (inventory - required) < 0:
					columnSeriesObjOrder.values[loop] = inventory - required
			loop = loop + 1
		print("All entries before deleting items not needing to be ordered")
		print(mydf)
		# get rid of all the entries that do not need to be ordered
		mydf = mydf[~mydf['Order Quantity'].isin(['NA'])]
		print("All entries to be ordered")
		print(mydf)
		#  create excel writer object
		d4 = today.strftime("%b-%d-%Y")
		# print("d4 =", d4)
		writer = pd.ExcelWriter(outfilename + d4 + '.xlsx', engine='xlsxwriter')
		mydf.to_excel(writer, index=False)
		#  Set the column width and format.
		#  Get the xlsxwriter workbook and worksheet objects.
		# workbook = writer.book
		worksheet = writer.sheets['Sheet1']
		worksheet.set_column('A:A', 40)
		worksheet.set_column('B:F', 20)
		# save Excel file
		writer.save()
		return 1
	except Exception as e:
		print("Entry in Ekos file is not valid")
		print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)
		# print("Exception", e)
		return(0)


''' this method processes Ekos report and creates spreadsheet of updated hop alpha '''

def update_hop_tracking_csv(InputDirectoryPath, OutputDirectoryPath):
	try:  # get a list of all files with extension .xlsx from the input directory
		# go get Ekos report filename from the inputhoptracing directory
		csv_names = []
		#  r=root, d=directories, f = files
		for r, d, f in os.walk(InputDirectoryPath):
			for file in f:
				if 'csv' in file:
					csv_names.append(os.path.join(r, file))
		if len(csv_names) == 0:
			print('\nNo files available to process')
			return 1
		# build input filename
		ekosfilename = InputDirectoryPath + '/' + f[0]
		# build output filename
		outSPfilename = OutputDirectoryPath + "/Hops Alpha Worksheet"
		#  read the input from Ekos
		ekosdf = pd.read_csv(ekosfilename, header=None)
		# get today's date so we can append it to output filename
		newDate = date.today()
		# build input filename
		# hard code the filename since it doesn't change
		sharepointfilename = './sharepointtemp/' + '/' + 'Hops Alpha Worksheet.xlsx'
		#  read the input from Ekos report
		sharepointdf = pd.read_excel(sharepointfilename, header=None)
		#  create pandas dataframes for both files that will be worked on
		ekosdf.rename(columns={0: 'Location', 1: 'Title', 2: "Lot"}, inplace=True)
		# drop the first part of the Sharepoint file
		modsharepointdf = sharepointdf.drop(labels=range(0, 4), axis=0)
		modsharepointdf.rename(columns={0: 'Title', 1: 'Lot', 2: "Date", 3: 'AA', 4: 'Foil'}, inplace=True)

		# create dataframe objects that we will use to parse and modify Share-point dataframe
		columnSeriesObjEkosTitle = ekosdf['Title']
		columnSeriesObjEkosLot = ekosdf['Lot']
		columnSeriesObjSPTitle = modsharepointdf['Title']
		columnSeriesObjSPLot = modsharepointdf['Lot']
		columnSeriesObjSPDate = modsharepointdf['Date']
		columnSeriesObjSPAA = modsharepointdf['AA']
		columnSeriesObjSPFoil = modsharepointdf['Foil']
		#  create a dataframe for the final Sharepoint file to be created
		column_names = ['Title', 'Lot', "Date", 'AA', 'Foil']
		finalsharepointdf = pd.DataFrame(columns=column_names)
		#  now look at Ekos file to see what entries are new and add them to SharePoint file
		loop = 0
		total_rows = modsharepointdf.shape[0]
		print('total rows in Sharepoint file', total_rows)
		# this loop pulls all entries from the Ekos report
		for i in columnSeriesObjEkosTitle:
			ekosTitle = columnSeriesObjEkosTitle.values[loop]
			ekosLot = columnSeriesObjEkosLot.values[loop]
			sploop = 0
			# this loop compares each Ekos entry the SharePoint file.
			for i in columnSeriesObjSPTitle:
				SPTitle = columnSeriesObjSPTitle.values[sploop]
				SPLot = columnSeriesObjSPLot.values[sploop]
				if (ekosTitle == SPTitle) and (ekosLot == SPLot):
					break
				# if there's no match for ekos entry we need to add it to new SharePoint dataframe
				elif sploop == total_rows - 1:

					adddict = {'Title': ekosTitle, 'Lot': ekosLot, 'Date': newDate}
					finalsharepointdf = finalsharepointdf.append(adddict, ignore_index=True)
					finalsharepointdf = finalsharepointdf.sort_values(by=['Title'])
					print('Found a new entry in Ekos report', ekosTitle, ekosLot)
				sploop = sploop + 1
			loop = loop + 1

		# now look thru the SharePoint file and delete entries that are no longer in Ekos file
		loop = 0
		total_rows = ekosdf.shape[0]
		print('total rows file in Ekos file', total_rows)
		for i in columnSeriesObjSPTitle:
			SPTitle = columnSeriesObjSPTitle.values[loop]
			SPLot = columnSeriesObjSPLot.values[loop]
			SPDate = columnSeriesObjSPDate.values[loop]
			SPAA = columnSeriesObjSPAA.values[loop]
			SPFoil = columnSeriesObjSPFoil.values[loop]
			sploop = 0
			# print('Looking in SP file for ', SPTitle, SPLot,loop,sploop)
			#  now see if each entry in the SharePoint file is in the new Ekos report.
			for i in columnSeriesObjEkosTitle:
				ekosTitle = columnSeriesObjEkosTitle.values[sploop]
				ekosLot = columnSeriesObjEkosLot.values[sploop]
				# if we find a match, keep the entry and add it to new dataframe
				if (ekosTitle == SPTitle) and (ekosLot == SPLot):
					# print('got one',ekosTitle,ekosLot, SPTitle,SPLot, loop, sploop,total_rows-1)
					adddict = {'Title': SPTitle, 'Lot': SPLot, 'Date': SPDate, 'AA': SPAA, 'Foil': SPFoil}
					finalsharepointdf = finalsharepointdf.append(adddict, ignore_index=True)
					finalsharepointdf = finalsharepointdf.sort_values(by=['Title'])
					break
				sploop = sploop + 1
			loop = loop + 1

		today = date.today()
		# add header info back into SharePoint file
		headerdf = pd.DataFrame({'Title': ['Hops Alpha Tracking By Lot', 'Generated on ', ''],
								'Lot': ['', '', ''],
								'Date': ['', today, ''],
								'AA': ['', '', ''],
								'Foil': ['', '', '']})
		total_rows = finalsharepointdf.shape[0]
		print('total rows in updated Sharepoint file', total_rows)
		# concatenate both dataframes to get final
		finalsharepointdf = headerdf.append(finalsharepointdf, ignore_index=True)
		# now write the newly create dataframe to an output file
		writer = pd.ExcelWriter(outSPfilename + '.xlsx', engine='xlsxwriter')
		finalsharepointdf.to_excel(writer, index=False)
		#  Get the xlsxwriter workbook and worksheet objects.
		# workbook = writer.book
		# worksheet = writer.sheets['Sheet1']
		# save Excel file
		writer.save()
		# now read the Excel file to see what it looks
		# newdf = pd.read_excel(outSPfilename+'.xlsx')

		return 1
	except Exception as e:
		print("Entry in Ekos file is not valid")
		print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)
		# print("Exception", e)
		return 0


def print_menu():
	""" print initial menu for user"""
	print(30 * "-" + "MENU" + 30 * "-")
	print("1   - Create Ingredients Order")
	print("2   - TBD")
	print("3   - TBD")
	print("4   - Quit")
	print(67 * "-")


if __name__ == "__main__":
	InputPath = InputDirectoryPath + '/'
	OutputDirectoryPath = OutputDirectoryPath + '/'

	print("Python Version from is " + platform.python_version())
	print("System Version is " + platform.platform())

	localtime = time.asctime(time.localtime(time.time()))
	print("Local current time :", localtime)

	while True:
		print_menu()
		try:
			choice = int(input("Enter your choice [1-4]: "))
		except ValueError:
			print("Not a valid number selection")
			#  better try again... Return to the start of the loop
			continue
		if choice < 1 or choice > 4:
			print("Selection not in range")
			continue
		if choice == 1:
			create_ingredients_order(InputDirectoryPath, OutputDirectoryPath)
		elif choice == 2:
			continue
		elif choice == 3:
			continue
		elif choice == 4:
			break
