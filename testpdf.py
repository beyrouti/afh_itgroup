# #!/usr/bin/python3
#
# Author: Michael Beyrouti
# Atlanta Fine Homes Sotheby's International Realty, IT Group
#
# This script is intended to parse the Network Setup form used for enboarding new members to AFH
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
