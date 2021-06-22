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
#     uFirstName,
#     uLastName,

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
member_ou = ""
officeDesc = {"bh":"Buckhead","na":"North Atlanta","in":"Intown","co":"Cobb"}
member_groups = ["#AllAtlantaFineHomes","PowerUser","FreePBX Users"]
logon_script = ""

# =========================================================================================================
# Parses the NSF data into a dictonary member_stats using pdfminer
# https://stackoverflow.com/questions/3984003/how-to-extract-pdf-fields-from-a-filled-out-form-in-python
# =========================================================================================================
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
    #print('{0}: {1}'.format(name, value))
member_stats['uAfhEmail'] = member_stats['uAfhEmail'] + '@atlantafinehomes.com'
fp.close()
member_stats["uFirstName"] = member_stats["uName"].split(" ")[0]
member_stats["uLastName"] = member_stats['uName'].split(" ")[len(member_stats['uName'].split(" ")) -1]

# ===========================================
# Sets default credential information for AD
# ===========================================
pyad.set_defaults(ldap_server="10.0.254.3", username="asterisk", password="Shit$andwich747")

# =========================================================================================
# Set OU and group memberships of new user based on inputs [SHOULD BE A LOOP UNTIL CORRECT]
# =========================================================================================
staff_or_agent = input("New user [staff] or [agent]? : ").lower()
office_loc = input("which office [bh],[na],[in],[co]? :").lower()
if office_loc == "bh":
    if staff_or_agent == "staff":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=BH Staff,OU=AFH Staff,DC=AFH,DC=pri")
        logon_script = "BH_STAFF.vbs"
    elif staff_or_agent == "agent":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=BH Agents,OU=AFH Agents,DC=AFH,DC=pri")
        logon_script = "BH_AGENT.vbs"
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
elif office_loc == "na":
    if staff_or_agent == "staff":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=NA Staff,OU=AFH Staff,DC=AFH,DC=pri")
        logon_script = "NA_STAFF.VBS"
    elif staff_or_agent == "agent":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=NA Agents,OU=AFH Agents,DC=AFH,DC=pri")
        logon_script = "NA_AGENT.vbs"
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
elif office_loc == "in":
    if staff_or_agent == "staff":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=IN Staff,OU=AFH Staff,DC=AFH,DC=pri")
        logon_script = "IN_STAFF.vbs"
    elif staff_or_agent == "agent":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=IN Agents,OU=AFH Agents,DC=AFH,DC=pri")
        logon_script = "IN_AGENT.vbs"
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
elif office_loc == "co":
    if staff_or_agent == "staff":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=Cobb Staff,OU=AFH Staff,DC=AFH,DC=pri")
        logon_script = "CB_Staff.vbs"
    elif staff_or_agent == "agent":
        member_ou = pyad.adcontainer.ADContainer.from_dn("OU=Cobb Agents,OU=AFH Agents,DC=AFH,DC=pri")
        logon_script = "CB_AGENT.vbs"
    else:
        print("Sorry looks like you spelled [staff] or [agent] wrong, please run the program again")
        quit()
else:
    print("Sorry, looks like you spelled the office abreviation wrong, please run the program again")
    quit()
    #this needs to loop back to the input prompt

# ===================
# Create user in AD!
# ===================
new_user = pyad.aduser.ADUser.create(member_stats["uNetworkLogin"],member_ou,password="Changeme1",upn_suffix=None,enable=True,optional_attributes={
    "mail" : member_stats["uAfhEmail"],
    "givenName" : member_stats["uFirstName"],
    "displayName" : member_stats["uName"],
    "sn" : member_stats["uLastName"],
    "userPrincipalName" : member_stats["uNetworkLogin"] + "@AFH.pri",
    "mobile" : member_stats["uCell"],
    "company" : "NADA",
    "proxyAddresses" : "SMTP:" + member_stats["uAfhEmail"],
    "description" : officeDesc[office_loc],
    "scriptPath" : logon_script,
    "title" : member_stats["uFirstName"] + "." + member_stats["uLastName"] + "@sothebysrealty.com"
})
# ADD - office, department

try:
    new_user.rename(member_stats["uName"],set_sAMAccountName=False)
except:
    print(member_stats["uNetworkLogin"]+ " added successfully")
# ADD user to groups


