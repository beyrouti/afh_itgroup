#!/usr/bin/python3
#
# Author: Michael Beyrouti
# Atlanta Fine Homes Sotheby's International Realty, IT Group
#
# This script is intended to parse the Network Setup form used for onboarding new members to AFH
# Next iteration will send parsed data to Active Directory, creating the user with the data

import sys
import mysql.connector
import csv
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1

# member_name = ""
# member_license = ""
# member_address = ""
# member_email = ""
# member_cell = ""
# member_status = "" #Agent, Staff, Assistant, Broker, etc...
# member_copyCode = ""
member_stats = {}
# FIELD KEYS FOR member_stats FROM NSF
#    "uEmail": "",
#    "uName": "",
#    "uCell": "",
#    "uEmergencyContactNum": "",
#    "uBirthday": "",
#    "uStartDate": "",
#    "uAfhEmail": "",
#    "uNetworkLogin": "",
#    "uCopyCode": "",
#    "uCopyCode": "",
#    "uDID": "",
#    "uFaxNum": "",
#    "uExtension": "",
#    "uMobileDevice": "",
#    "uTablet": "",
#    "uComputer": "",
#    "uTeam": "",
#    "uBoard": "",
#    "uFormerBrokerCode": "",
#    "uRenewalDate": "",
#    "uHomeAddress1": "",
#    "uHomeAddress2": "",
#    "uNightVMExt": "",
#    "ulicenseNum": "",

#Connect to AWS & creates SQL cursor
host = "aft-test-db.cx72tu6umfui.us-east-1.rds.amazonaws.com"
user = "afhit"
passwd = "X7kG!553o5sV"
database = "AFH_users"
mydb = mysql.connector.connect(
  host=host,
  user=user,
  passwd=passwd,
  database=database
)
cursor = mydb.cursor()

# Parses the NSF data into a dictonary, member_stats
# https://stackoverflow.com/questions/3984003/how-to-extract-pdf-fields-from-a-filled-out-form-in-python
file_name = "Network Setup Form - Mace Windu.pdf"
fp = open(file_name,'rb')
parser = PDFParser(fp)
network_form = PDFDocument(parser)
fields = resolve1(network_form.catalog['AcroForm'])['Fields']
for i in fields:
    field = resolve1(i)
    name, value = field.get('T').decode('utf-8'), field.get('V')
    if value == None:
        value = ""
    else:
        value = value.decode('utf-8')
    member_stats[str(name)] = str(value)
    print('{0}: {1}'.format(name, value))
member_stats['uAfhEmail'] = member_stats['uAfhEmail'] + '@atlantafinehomes.com'
fp.close()

sql = "INSERT IGNORE INTO UserInfo (first_name, last_name, network_username, copier_code, afh_email, outside_email, mobile_number, did_number, fax_number, office_ext) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
val = (member_stats["uName"].split(" ")[0], 
        member_stats['uName'].split(" ")[len(member_stats['uName'].split(" ")) -1], 
        member_stats['uNetworkLogin'], 
        member_stats['uCopyCode'], 
        member_stats['uAfhEmail'], 
        member_stats['uEmail'],
        member_stats['uCell'],
        member_stats['uDID'],
        member_stats['uFaxNum'],
        member_stats['uExtension'])
cursor.execute(sql,val)
mydb.commit()
mydb.close()

print("Inserted Row " + member_stats['uAfhEmail'])