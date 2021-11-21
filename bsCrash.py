"""
  My 1st Web Scrapping Program 
  
  Sources - youtube course: https://www.youtube.com/watch?v=XVv6mJpFOb0"
          - automate the boring stuff: https://automatetheboringstuff.com/2e/chapter12/  
          - automate the boring stuff; https://automatetheboringstuff.com/2e/chapter13/
  
  Scrapes data from : https://www.nijobfinder.co.uk/   (with selection: Northern Ireland / Python)
  Saves results to an excel file
  """


from bs4 import BeautifulSoup
import requests, os, openpyxl as xl

# excel: change location to save file to desktop
os.chdir("C:/Users/jslos/OneDrive/Desktop")

# excel: open new workbook and new worksheet
wb = xl.Workbook() 
sheet= wb.active
sheet.title="Ni Job Finder"

# excel: add header colums needed
sheet.cell(row=2, column=2).value = "ADVERTISER"
sheet.cell(row=2, column=3).value = "POSITION"
sheet.cell(row=2, column=4).value = "LOCATION"
sheet.cell(row=2, column=5).value = "SALARY"
sheet.cell(row=2, column=6).value = "POSTED"
sheet.cell(row=2, column=7).value = "CLOSING"

sheet.column_dimensions ['B', 'C'].width = 20 

# Scrap: Get html data to scrape
html_text = requests.get('https://www.nijobfinder.co.uk/search/?phrase=python&locations=Northern%20Ireland').text

# Scrap: Create soup object using html data to accomadate scraping 
soup = BeautifulSoup(html_text, 'lxml')

###  Scrapping  ###

jobs = soup.find_all('div', class_ = "c-result")

row = 3 # used to work out which row on excel sheet to save data to

for job in jobs:

  advertiser = job.find('p', class_ ='u-uppercase o-type-small u-marg-bottom-half u-light-text').text 

  position = job.find('h2', class_ ='c-result__job-title o-heading-two u-marg-bottom-small').text.strip()
  moreinfo = "https://www.nijobfinder.co.uk/" + (job.h2.a['href'])
  
  location = job.find('p', class_ = 'u-milli').text.strip().replace(' ','')

  info = job.find_all('li')
  salary = info[0].find_all('p')[1].text
  posted = info[1].find_all('p')[1].text
  closing = info[2].find_all('p')[1].text

  # excel: save data to excel sheet
  sheet.cell(row=row, column=2).value = advertiser
  sheet.cell(row=row, column=3).value = position
  sheet.cell(row=row, column=4).value = location
  sheet.cell(row=row, column=5).value = salary
  sheet.cell(row=row, column=6).value = posted
  sheet.cell(row=row, column=7).value = closing
  row = row +2 # moves to 2nd next row for next data save

  # not needed now : this prints the details to terminal

#  print(f"""
#  Advertiser:       {advertiser}
#  Position:         {position}
#  Location:         {location}
#  Salary:           {salary}
#  Posted:           {posted}
#  Closing:          {closing}
#  More Information: 

#  {moreinfo} 
#  """)

print("Completed: Data saved to Job Search.xlsx on Desktop")
# excel: final step - save the file
wb.save('Job Search.xlsx') 