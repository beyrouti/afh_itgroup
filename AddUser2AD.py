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
import time
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from pyad import *

member_stats = {}
staff_or_agent = ""
office_loc = ""
ou = ""

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

# ===========================================
# Sets default credential information for AD
# ===========================================
pyad.set_defaults(ldap_server="10.0.254.3", username="asterisk", password="Shit$andwich747")

# ===================
# Set ou of new user
# ===================
staff_or_agent = input("New user [staff] or [agent]? : ").lower()
office_loc = input("which office [bh],[na],[in],[co]? :").lower()
if office_loc == "bh":
    if staff_or_agent == "staff":
        ou = "OU=BH Staff, OU=AFH Staff"
    elif staff_or_agent == "agent":
        ou = pyad.adcontainer.ADContainer.from_dn("OU=BH Agents,OU=AFH Agents,DC=AFH,DC=pri")
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
elif office_loc == "na":
    if staff_or_agent == "staff":
        ou = "OU=NA Staff, OU=AFH Staff"
    elif staff_or_agent == "agent":
        ou = "NA Agents, AFH Agents"
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
elif office_loc == "in":
    if staff_or_agent == "staff":
        ou = "OU=IN Staff, OU=AFH Staff"
    elif staff_or_agent == "agent":
        ou = "IN Agents, AFH Agents"
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
elif office_loc == "co":
    if staff_or_agent == "staff":
        ou = "OU=Cobb Staff, OU=AFH Staff"
    elif staff_or_agent == "agent":
        ou = "Cobb Agents, AFH Agents"
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
else:
    print("Sorry, looks like you spelled the office abreviation wrong, please run the program again")
    quit()

# ===================
# Create user in AD!
# ===================
new_user = pyad.aduser.ADUser.create(member_stats["uNetworkLogin"],ou,password="Changeme1",upn_suffix=None,enable=True,optional_attributes={
    "mail" : member_stats["uAfhEmail"],
    "givenName" : member_stats["uName"].split(" ")[0],
    "displayName" : member_stats["uName"],
    "mobile" : member_stats["uCell"],
    "sn" : member_stats['uName'].split(" ")[len(member_stats['uName'].split(" ")) -1],
    "userPrincipalName" : member_stats["uNetworkLogin"] + "@AFH.pri",
    "company" : "NADA",
    "proxyAddresses" : "SMTP:" + member_stats["uAfhEmail"]
})
# ADD - description, office, logon script, department, Job title
#time.sleep(5)
#wrap in try catch...
new_user.rename(member_stats["uName"],set_sAMAccountName=False)
# ADD user to groups

print(member_stats["uNetworkLogin"]+ " added Successfully")
