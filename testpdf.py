# #!/usr/bin/python3
#
# Author: Michael Beyrouti
# Atlanta Fine Homes Sotheby's International Realty, IT Group
#
# This script is intended to parse the Network Setup form used for onboarding new members to AFH
# Next iteration will send parsed data to Active Directory, creating the user with the data

import sys
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1

member_name = ""
member_license = ""
member_address = ""
member_email = ""
member_cell = ""
member_status = "" #Agent, Staff, Assistant, Broker, etc...
member_copyCode = ""

# https://stackoverflow.com/questions/3984003/how-to-extract-pdf-fields-from-a-filled-out-form-in-python
file_name = input("Enter Network Form file name: ")
fp = open(file_name,'rb')
parser = PDFParser(fp)
network_form = PDFDocument(parser)
fields = resolve1(network_form.catalog['AcroForm'])['Fields']
for i in fields:
    field = resolve1(i)
    name, value = field.get('T'), field.get('V')
    print('{0}: {1}'.format(name, value))

