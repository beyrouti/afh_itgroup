#!/usr/bin/python3
# ============================================================
# Author: Michael Beyrouti
# Atlanta Fine Homes Sotheby's International Realty, IT Group
# ============================================================
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

import sys
import mysql.connector
import csv
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from pyad import *

member_stats = {}

# Parses the NSF data into a dictonary, member_stats, using pdfminer
# https://stackoverflow.com/questions/3984003/how-to-extract-pdf-fields-from-a-filled-out-form-in-python
file_name = input("Enter the file name: ")
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

# Sets default credential information for AD
pyad.set_defaults(ldap_server="10.0.254.3", username="asterisk", password="Shit$andwich747")