#mojohelpdesk API scraper script

#Set some variables, need the url, "xxx.mojohelpdesk.com", also
#My access key, which is found on the profile page in Mojo.
#The queue ID can be found by going to the queue you want in Mojo, 
#then getting the queue id parameter from the URL.
Domain = "XXX"
AccessKey = "XXX"
QueueID = "XXX"


#Import necessary modules
import sys
import bs4
from urllib import urlopen as uReq
from bs4 import BeautifulSoup as soup
import os
import pygsheets
import pandas as pd
import datetime

#Create a list to contain the filenames, so we can 
#delete them when the script is done.
FileList = []

#Generate random number to rename the XML file we download
import random
for x in range(1):
	FileName = random.randint(1000,100000000000)*5
FileName = str(FileName)
FileList.append(FileName)


#Open the query URL
QueryURL = "http://" + Domain + ".mojohelpdesk.com/api/tickets/search.xml?query=status.id:\(\<50\)%20AND%20queue.id=" + QueueID + "\&sf=created_on\&per_page=100\&access_key=" + AccessKey
#I used curl here because it worked better than requests.
CurlCommand = "curl -s -o " + str(FileName) + " " + QueryURL
os.system(CurlCommand)
#Read the file
Contents = open(FileName).read()

#Use beautifulsoup to parse the data we receive
page_soup = soup(Contents, "html.parser")
#Get the ID of each ticket
IDContainers = page_soup.findAll("id")
#array definition 
CategoryArray = []
UniqueCategoryArray = []
for IDField in IDContainers:
	TicketID = IDField.text
	FileList.append(TicketID)
	#Read each individual ticket to get the category
	IndTicketString = "curl -s -o " + TicketID + " http://" + Domain + ".mojohelpdesk.com/api/tickets/" + TicketID + ".xml?access_key=" + AccessKey	
	os.system(IndTicketString)
	TicketContents = open(TicketID).read()
	TicketSoup = soup(TicketContents, "html.parser")
	CategoryContainers = TicketSoup.findAll("custom-field-type-of-device")	
	#Build a list of categories
	for CategoryContainer in CategoryContainers:
		Category = CategoryContainer.text
		CategoryArray.append(Category)
#Condense to a unique list of categories 
for Category in CategoryArray:
	if Category not in UniqueCategoryArray:
		UniqueCategoryArray.append(Category)



#Define parameters for the spreadsheet in Google Sheets. 
#You will need to create a service account from Google and 
#download the authorization file, then change the name to something
#shorter and put the filename here.
gc = pygsheets.authorize(service_file='/path/to/spreadsheet_auth.json')
sheet = gc.open('Mojo API Data')
wks = sheet[0]

#This function determines which column a number goes in.
def ColumnDistribution(CategoryName):
        Categories = ("Chromebook / Chromebase","Login / Password","Laptop","Desktop","iPad","Internet","User Setup","Phone","TeacherEase","Printer","Projector","Smart Board","Google Apps","Software","Website Block / Unblock","Website Change / Update","AV Equipment Setup","Security Camera Footage Request","New/Returning/Exiting Student","Please Select One")
        Columns = ("C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V")
        CategoryCount = 0
        for Categorytitle in Categories:
                if CategoryName == Categorytitle:
                        ThisColumn = Columns[CategoryCount]
                        #print(str(CategoryCount) + " " + ThisColumn)
                CategoryCount = CategoryCount + 1
	return ThisColumn

#This function determines the first empty row of column A of the spreadsheet
def find_empty_cell():
        letter = "A"
        for x in range(1, 1000):
            cell_coord = letter + str(x)
            if wks.cell(cell_coord).value == "":
                return(cell_coord)

EmptyCell = find_empty_cell()
EmptyRow = EmptyCell[1:]
#Define today's date,then put into the first empty row of Column A:
DateNow = datetime.datetime.now()
DateNow = DateNow.strftime("%D")
wks.cell(EmptyCell).value = DateNow

#Put Ticket count and counts of open ticket categories into the worksheet
TicketCount = str(len(IDContainers))
print(Domain + " has " + TicketCount + " open tickets:")
TotalCountCell = "B" + EmptyRow
print(TotalCountCell)
wks.cell(TotalCountCell).value = TicketCount

#Separate categories and fill columns with information
for Category in UniqueCategoryArray:
        print(Category)
	#sys.exit()
	Column = ColumnDistribution(Category)
	Cell = Column + EmptyRow
	wks.cell(Cell).value = str(CategoryArray.count(Category))
	#OutputString = Category + ": " + str(CategoryArray.count(Category))
        #print(OutputString)





























#Display ticket count and counts of open ticket categories
#TicketCount = str(len(IDContainers))
#print(Domain + " has " + TicketCount + " open tickets:")

#for Category in UniqueCategoryArray:
#	OutputString = Category + ": " + str(CategoryArray.count(Category))
#	print(OutputString)

#Cleanup
for File in FileList:
	os.remove(File)
	#print(File)
