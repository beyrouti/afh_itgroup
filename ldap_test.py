#!/usr/bin/python3

from pyad import *

pyad.set_defaults(ldap_server="afh.pri", username="asterisk", password="Shit$andwich747")

user = aduser.ADUser.from_cn("jbravo", options=dict(ldap_server="10.0.254.3"))

print(user)

