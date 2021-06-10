#!/usr/bin/python3

from pyad import *

pyad.set_defaults(ldap_server="10.0.254.3", username="asterisk", password="Shit$andwich747")

#user1 = aduser.ADUser.from_cn("Caroline Davenport")
user1 = adobject.ADObject.from_dn("CN=Caroline Davenport,OU=BH Agents,OU=AFH Agents,DC=AFH,DC=pri")
#OU=AFH Agents,OU=BH Agents,DC=AFH,DC=pri"
print(user1)
# user1.update_attribute("Company","anda")
print(user1.get_attribute("Company"))

user2 = adobject.ADObject.from_dn("CN=Mace Windu,OU=BH Agents,OU=AFH Agents,DC=AFH,DC=pri")