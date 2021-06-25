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
#     uState,
#     uZip,
#     uCity

import sys
import csv
import time
import subprocess
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from pyad import *
from pyad import pyad

member_stats = {}
staff_or_agent = ""
office_loc = ""
member_ou = ""
officeDesc = {"bh":"Buckhead","na":"North Atlanta","in":"Intown","co":"Cobb"}
logon_script = ""
new_user= ""
notValidInput = True
cityStateZip = []

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
cityStateZip = member_stats["uHomeAddress2"].split()
member_stats["uZip"] = cityStateZip[len(cityStateZip) - 1]
member_stats["uState"] = "GA"
member_stats["uCity"] = cityStateZip[0] #check for commas

# ===========================================
# Sets default credential information for AD
# ===========================================
pyad.set_defaults(ldap_server="10.0.254.3", username="asterisk", password="Shit$andwich747")

# =====================
# Create user function
# =====================
def createUser():
    the_user = pyad.aduser.ADUser.create(member_stats["uNetworkLogin"],member_ou,password="Changeme1",upn_suffix=None,enable=True,optional_attributes={
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
        "title" : member_stats["uFirstName"] + "." + member_stats["uLastName"] + "@sothebysrealty.com",
        "streetAddress" : member_stats["uHomeAddress1"],
        "st" : member_stats["uState"],
        "l" : member_stats["uCity"],
        "postalCode" : member_stats["uZip"]
    })
    # ADD -> office, department, extension, office phone (if designated)
    # Powershell:
    # C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
    subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#AllAtlantaFineHomes' -Members " + member_stats["uNetworkLogin"], shell=True)
    subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'PowerUser' -Members " + member_stats["uNetworkLogin"], shell=True)
    subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'Docusign' -Members " + member_stats["uNetworkLogin"], shell=True)
    subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'FreePBX Users' -Members " + member_stats["uNetworkLogin"], shell=True)
    return the_user

# ==========================================================================
# Set OU, create user, and add group memberships of new user based on inputs
# ==========================================================================
while notValidInput:
    staff_or_agent = input("New user [staff] or [agent]? : ").lower()
    office_loc = input("which office [bh],[na],[in],[co]? :").lower()
    if staff_or_agent == "staff":
        if office_loc == "bh":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=BH Staff,OU=AFH Staff,DC=AFH,DC=pri")
            logon_script = "BH_STAFF.vbs"
            new_user = createUser()
            # BH STAFF GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#BuckheadOffice' -Members " + member_stats["uNetworkLogin"], shell=True)
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#BuckheadStaff' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        elif office_loc == "na":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=NA Staff,OU=AFH Staff,DC=AFH,DC=pri")
            logon_script = "NA_STAFF.VBS"
            new_user = createUser()
            # NA STAFF GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#NorthAtlantaOffice' -Members " + member_stats["uNetworkLogin"], shell=True)
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#NorthAtlantaStaff' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        elif office_loc == "in":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=IN Staff,OU=AFH Staff,DC=AFH,DC=pri")
            logon_script = "IN_STAFF.vbs"
            new_user = createUser()
            # IN STAFF GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#IntownOffice' -Members " + member_stats["uNetworkLogin"], shell=True)
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#IntownStaff' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        elif office_loc == "co":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=Cobb Staff,OU=AFH Staff,DC=AFH,DC=pri")
            logon_script = "CB_Staff.vbs"
            new_user = createUser()
            # CO STAFF GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#CobbStaff' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        else:
            print("Sorry, looks like you spelled the office abreviation wrong, please try again")
        #STAFF GROUPS
        subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'AFH Staff' -Members " + member_stats["uNetworkLogin"], shell=True)
        subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'Accounting' -Members " + member_stats["uNetworkLogin"], shell=True)
        subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'Brokerwolf' -Members " + member_stats["uNetworkLogin"], shell=True)
        subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'Listings' -Members " + member_stats["uNetworkLogin"], shell=True)
        subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'Administrators' -Members " + member_stats["uNetworkLogin"], shell=True)
        # Should get user input for above staff groups y/n
    elif staff_or_agent == "agent":
        if office_loc == "bh":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=BH Agents,OU=AFH Agents,DC=AFH,DC=pri")
            logon_script = "BH_AGENT.vbs"
            new_user = createUser()
            # BH AGENT GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#BuckheadOffice' -Members " + member_stats["uNetworkLogin"], shell=True)
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#BuckheadAgents' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        elif office_loc == "na":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=NA Agents,OU=AFH Agents,DC=AFH,DC=pri")
            logon_script = "NA_AGENT.vbs"
            new_user = createUser()
            # NA AGENT GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#NorthAtlantaOffice' -Members " + member_stats["uNetworkLogin"], shell=True)
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#NorthAtlantaAgents' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        elif office_loc == "in":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=IN Agents,OU=AFH Agents,DC=AFH,DC=pri")
            logon_script = "IN_AGENT.vbs"
            new_user = createUser()
            # IN AGENT GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#IntownOffice' -Members " + member_stats["uNetworkLogin"], shell=True)
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#IntownAgents' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        elif office_loc == "co":
            member_ou = pyad.adcontainer.ADContainer.from_dn("OU=Cobb Agents,OU=AFH Agents,DC=AFH,DC=pri")
            logon_script = "CB_AGENT.vbs"
            new_user = createUser()
            # CO AGENT GROUPS
            subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#CobbAgents' -Members " + member_stats["uNetworkLogin"], shell=True)
            notValidInput = False
        else:
            print("Sorry, looks like you spelled the office abreviation wrong, please try again")
        #AGENT GROUPS
        subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity '#AtlantaFineHomesAgents' -Members " + member_stats["uNetworkLogin"], shell=True)
        subprocess.call("C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe Add-ADGroupMember -Identity 'AFH Agents' -Members " + member_stats["uNetworkLogin"], shell=True)
    else:
        print("Sorry looks like you spelled 'staff' or 'agent' wrong, please try again")

# ========================================
# renames the user to first and last name
# ========================================
try:
    new_user.rename(member_stats["uName"],set_sAMAccountName=False)
except:
    print(member_stats["uNetworkLogin"]+ " added successfully as a " + officeDesc[office_loc] + " " + staff_or_agent)

# =========================================
# runs the .bat file to add user to egnyte
# =========================================
subprocess.call(r'C:/egnyte_win32_ad_kit_4.15.1_r18/egnyte_win32_ds_kit/run.bat')
