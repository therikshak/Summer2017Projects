import requests
import csv
import urllib3
import queue
import os
from requests_ntlm import HttpNtlmAuth
from bs4 import BeautifulSoup

username = "Daryl"
password = "Beer1"
url = "http://worcs.lindnerlogistics.com/worcs/InventoryByCustomerResults.asp?show_stock_number=on&show_description=on&show_unit=on&show_tag_count=on&show_quantity=on&order_by=stock_number&stock_number=&description=&a1=&a2=&status=%25&ButtonPressed=Search"

r = requests.get(url, auth=HttpNtlmAuth(username, password))
soup = BeautifulSoup(r.content, "html.parser")
table = soup.find_all('table')


items = []
for row in table[1].find_all('tr')[1:]:
    this_item = []
    line = row.find_all('td')
    this_item.append("Lindner")
    this_item.append(line[0].text)  # ax number
    this_item.append(line[1].text + line[2].text)  # description with product size
    this_item.append(line[4].text)  # units
    items.append(this_item)

out_file_header = ['Brewery', 'AX #', 'Description', 'Quantity']
windows_username = os.getlogin()
start_path = r'C:\Users\\'
end_path = r'\Desktop\Lindner_Scrape\lindner.csv'
full_path = start_path + windows_username + end_path
# open the csv file
with open(full_path, 'w', newline='') as out:
    # create a csv writer
    csv_out = csv.writer(out)
    # write the header to the csv file
    csv_out.writerow(out_file_header)
    for row in items:
        csv_out.writerow(row)