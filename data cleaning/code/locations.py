

import xlrd
from collections import OrderedDict
import csv
 
source_filename = '../data/repossed.xlsx' 

path = "../output/locations.csv"

quotes = '''"'''

def get_locations_from_sheet1():
	# Open the workbook and select the first worksheet
	wb = xlrd.open_workbook(source_filename)
	sh = wb.sheet_by_index(0)
	loc_list = []
	# Iterate through each row in worksheet 
	for rownum in range(1, sh.nrows):
	    loc = OrderedDict()
	    row_values = sh.row_values(rownum)
	    loc['buildingnumber'] = str(row_values[0])
	    loc['buildingname'] = str(row_values[1]).strip().title()
	    loc['address'] = str(row_values[3]).strip().upper()
	    loc['address2'] = str(row_values[4]).strip().upper()
	    loc['city'] = str(row_values[5]).upper()
	    loc['state'] = str(row_values[7]).upper()
	    loc['zip'] = str(row_values[8])
	    loc['agency'] = str(row_values[9]).upper()
	    loc['lon'] = str(row_values[10])
	    loc['lat'] = str(row_values[11])
	    loc['region'] = str(row_values[12]).upper()
	    loc['phone'] = str(row_values[13])
	    loc['fax'] = str(row_values[14])

	    full_address = loc['address'].strip()
	    if len(loc['address2']) != 0:
	    	full_address = full_address + ", " + loc['address2'].strip()
	    full_address = full_address +", "+loc['city'].strip()+ ", "+loc['state'].strip()+ ", "+loc['zip'].strip()

	    loc['fulladdress'] = full_address

	    loc_list.append(loc)

	return loc_list

def sanitize_address(adstr):

	result = adstr

	result = result.replace(" W ", " WEST ")
	result = result.replace(" W. ", " WEST ")
	result = result.replace(" E. ", " EAST ")
	result = result.replace(" E ", " EAST ")
	result = result.replace(" S ", " SOUTH ")
	result = result.replace(" S. ", " SOUTH ")
	result = result.replace(" SO ", " SOUTH ")
	result = result.replace(" N ", " NORTH ")
	result = result.replace(" N. ", " NORTH ")
	result = result.replace(" NE ", " NORTH EAST")
	result = result.replace(" NW ", " NORTH WEST")
	result = result.replace(" N.W. ", " NORTH ")
	result = result.replace(" S.W. ", " SOUTH WEST ")
	result = result.replace(" SW ", " SOUTH WEST ")
	result = result.replace(" SE ", " SOUTH EAST ")
	result = result.replace(" S.E. ", " SOUTH EAST ")
	result = result.replace("U.S.A.", "USA")
	result = result.replace("U.S.", "US")

	return result

def get_locations_from_sheet2():
	# Open the workbook and select the first worksheet
	wb = xlrd.open_workbook(source_filename)
	sh = wb.sheet_by_index(1)
	loc_list = []
	# Iterate through each row in worksheet 
	for rownum in range(1, sh.nrows):
	    loc = OrderedDict()
	    row_values = sh.row_values(rownum)
	    loc['buildingnumber'] = str(row_values[0])
	    loc['buildingname'] = str(row_values[1]).strip().upper()
	    loc['address'] = sanitize_address(str(row_values[2]).upper())
	    loc['city'] = str(row_values[3]).upper()
	    loc['state'] = str(row_values[4]).upper()
	    loc['zip'] = str(row_values[5])
	    loc['lat'] = str(row_values[6])
	    loc['lon'] = str(row_values[7])
	    loc['region'] = str(row_values[8]).upper()
	    loc['agency'] = str(row_values[9]).upper()

	    full_address = loc['address'].strip()+", "+loc['city'].strip()+ ", "+loc['state'].strip()+ ", "+loc['zip'].strip()

	    loc['fulladdress'] = full_address

	    loc_list.append(loc)

	return loc_list

def merge_locations(loc1, loc2):
    loc3 = []

    loc = OrderedDict()

    for lo in loc1:
    	loc = OrderedDict()
    	loc['buildingnumber'] = lo['buildingnumber']
    	loc['buildingname'] = lo['buildingname']
    	loc['address'] = lo['address']
    	loc['city'] = lo['city']
    	loc['state'] = lo['state']
    	loc['zip'] = lo['zip']
    	loc['agency'] = lo['agency']
    	loc['lat'] = lo['lat'] 
    	loc['lon'] = lo['lon']
    	loc['region'] = lo['region']
    	loc['phone'] = lo['phone']
    	loc['fax'] = lo['fax']
    	loc['fulladdress'] = lo['fulladdress']
    	loc3.append(loc)

    for lo2 in loc2:
    	loc = OrderedDict()
    	loc['buildingnumber'] = lo2['buildingnumber']
    	loc['buildingname'] = lo2['buildingname']
    	loc['address'] = lo2['address']
    	loc['city'] = lo2['city']
    	loc['state'] = lo2['state']
    	loc['zip'] = lo2['zip']
    	loc['lat'] = lo2['lat']
    	loc['lon'] = lo2['lon']
    	loc['region'] = lo2['region']
    	loc['agency'] = lo2['agency']	
    	loc['phone'] = " - - "
    	loc['fax'] = " - - " 
    	loc['fulladdress'] = lo2['fulladdress']
    	loc3.append(loc)

    return loc3


def csv_writer(data, path):
	with open(path, "wb") as csv_file:
		writer = csv.writer(csv_file, delimiter=',')
		for line in data:
			writer.writerow(line)

	return

def quoted(inp):
	result = inp

	return result


def create_data(locat):
    
    result = []
    result.append("BuildingNumber,BuildingName,FullAddress,Agency,Phone,Fax,Address,City,State,ZipCode,Latitude,Longitude,Region".split(','))

    for loc in locat:
    	strg = (str(loc['buildingnumber'])+";"+
    		    str(loc['buildingname'])+";"+
    		    quoted(str(loc['fulladdress']))+";"+
    		    quoted(str(loc['agency']))+";"+
    		    str(loc['phone'])+";"+
    		    str(loc['fax'])+";"+
    		    str(loc['address'])+";"+
    		    str(loc['city'])+";"+
    		    str(loc['state'])+";"+
    		    str(loc['zip'])+";"+
    		    str(loc['lat'])+";"+
    		    str(loc['lon'])+";"+
    		    str(loc['region'])).split(';')
    	result.append(strg)

    return result



if __name__ == '__main__':
	locations2 = get_locations_from_sheet1()

	#for lo in locations2:
	#	print lo['fulladdress']

	print "first locations ", len(locations2)

	locations3 = get_locations_from_sheet2()

	#for lo in locations3:
	#	print lo['fulladdress']

	print "second locations ", len(locations3)

	locations = merge_locations(locations2, locations3)


	print " length merged list", len(locations)

	print "Total adding lens ", len(locations2) + len(locations3)

	d = create_data(locations)

	#print d
	csv_writer(d, path)



