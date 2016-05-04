import urllib.request, json
import time
import sys
import xlwt

# Here you can insert your api key and name the output .xls file
API_KEY = ""
OUTPUT_NAME = "LATLONG"
INPUT_NAME = "workfile"

# function to print the collection(in the database)
def printCollection(collection):
	printCursor = collection.find()
	for document in printCursor:
		print('---------')
		print(document)
		print('---------')

start_time = time.time()

# counter used for pausing every 10 requests (Google API limitation)
ctr = 0

# counter used to keep track of how many requests were made to Google
googleRequests = 0

# Open a file
f = open(INPUT_NAME, 'r')

# set up excel workbook and sheet for information storage
wb = xlwt.Workbook()

ws = wb.add_sheet('Lats & Longs')

row = 1
column = 0

# titles / headings
headings = ["Input Address", "Formatted Address", "Latitude", "Longitude"]
for x in range(0, len(headings)):
	ws.write(0, x, headings[x])

print("====================Lat Long Search Start====================")

# Go through each line in workfile
for line in f:

	# need to submit a request to Google for the information
	convert = line.replace(" ", "+")
	if API_KEY:
		url = "https://maps.googleapis.com/maps/api/geocode/json?address=" + convert + "&key=" + API_KEY
	else:
		url = "https://maps.googleapis.com/maps/api/geocode/json?address=" + convert

	# submit request & decode json
	response = urllib.request.urlopen(url)
	googleRequests += 1
	string = response.read().decode('utf-8')
	data = json.loads(string)

	# Parse the data, output the data
	if data.get('results'):
		faddress = data['results'][0]['formatted_address']
		latitude = data['results'][0]['geometry']['location']['lat']
		longitude = data['results'][0]['geometry']['location']['lng']
		# format: address lat, long
		print(faddress + " " + str(latitude) + ", " + str(longitude))
		# write to excel file
		ws.write(ctr + 1, 0, line)
		ws.write(ctr + 1, 1, faddress)
		ws.write(ctr + 1, 2, latitude)
		ws.write(ctr + 1, 3, longitude)

	elif data.get('error_message'):
		print(data['error_message'])
		ws.write(ctr + 1, 0, line)
	elif data.get('status') and data['status'] == 'ZERO_RESULTS':
		print("There are no results for " + line.strip())
		ws.write(ctr + 1, 0, line)
	else:
		print("====================DUMP====================")
		print(data)

	# increment counter
	ctr += 1

	# sleep for 1 second every 10 requests
	if ctr % 10 == 0:
		time.sleep(1)

# save the excel file
wb.save(OUTPUT_NAME + ".xls")

# print the db
#printCollection(coll)

#print the number of requests to Google were made
print("=====================Lat Long Search End=====================")
print("Number of request(s) made to Google: " + str(googleRequests))
print("Results have been saved in " + OUTPUT_NAME + ".xls")

# print the duration of time it took to execute
print("---Program runtime: %s seconds ---" % (time.time() - start_time))