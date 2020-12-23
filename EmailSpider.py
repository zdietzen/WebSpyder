#!/usr/bin/env python3

'''
EmailSpyder by Zak Dietzen                
Scrapes for emails on a website 
and places in spreadsheet.                                                                
'''

import re
from urllib.request import urlopen, Request
import os
from datetime import datetime

from openpyxl import Workbook
from bs4 import BeautifulSoup

save_excel = True #Change to "True" to save email into Excel

book = Workbook()
sheet = book.active


def start_scrape(page, name_the_file):

    print("\n\nWebpage is currently being scrapped... please wait...")
       
    scrape = BeautifulSoup(page, 'html.parser')
    scrape = scrape.get_text()
    
    
    emails = set(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,3}", scrape))

    nodupemail = len(list(emails))

    dupemail = len(list(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,3}", scrape)))

    number_of_dup_email = int(dupemail) - int(nodupemail)

    email_list = list(emails)

    if len(emails) == 0:
        print("No email address(es) found.")
        print("-----------------------------\n")
    else:
        count = 1
        for item in emails:
            print('Email address #' + str(count) + ': ' + item)
            count += 1

    if save_excel:
        for row in zip(email_list):
            sheet.append(row)
        excel_file = (name_the_file + ".xlsx")
        book.save(excel_file) 
       
    print("\nDuplicates have been removed from list.")
    print("Total email addresses: ", nodupemail)
    print("There were " + str(number_of_dup_email) + " duplicate email addresses.")

    if save_excel:
        print("\n\nData has been stored inside of an Excel spreadsheet named: "
              + excel_file + " in this directory: " + os.getcwd())
        mod_time = os.stat(excel_file).st_mtime
        print("\nCompleted at: " + str(datetime.fromtimestamp(mod_time)))
        print("\nSize of file: " + str(os.stat(excel_file).st_size) + " KB")

def main():

    webpage = input("Paste the webpage you would like to scrape (include http/https): ")

    if save_excel:
        name_the_file = input("Name the file you would like to save the data in (don't include .xlsx): ")

    try:
        page = urlopen(webpage) 
        start_scrape(page)
    except:
        hdr = {'User-Agent': 'Mozilla/5.0'}
        req = Request(webpage, headers=hdr)
        page = urlopen(req)
        start_scrape(page, name_the_file)

if __name__ == "__main__":
    main()