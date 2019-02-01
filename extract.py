import os
import xml.etree.ElementTree as ET

folders = ["Inbox", "Sent Items"]
blacklist = ["asana.com", "no-reply", "noreply", "registration"]

email_list = []
for folder in folders:
    print("> Looking for: "+folder)
    current_folder = "Local/com.microsoft.__Messages/"+folder
    for filename in os.listdir(current_folder):
        if not filename.endswith(".xml"):
            print(current_folder+"/"+filename+" is not xml")
            continue

        try:
            tree = ET.parse(current_folder+"/"+filename)
        except:
            print(current_folder+"/"+filename+" couldn't parsed")
        finally:
            root = tree.getroot()
            for item in root.iter('emailAddress'):
                if 'OPFContactEmailAddressAddress' in item.attrib:
                    if any(to_check in item.attrib['OPFContactEmailAddressAddress'] for to_check in blacklist):
                        print(">>"+item.attrib['OPFContactEmailAddressAddress']+" is blacklisted")
                        continue
                    if 'OPFContactEmailAddressName' in item.attrib:
                        to_append = item.attrib['OPFContactEmailAddressName']+" <"+item.attrib['OPFContactEmailAddressAddress']+">"
                    else:
                        to_append = item.attrib['OPFContactEmailAddressAddress']
                    email_list.append(to_append)

# make every email unique
email_list = list(set(email_list))

f = open("export.txt", "w")
f.write("\n".join(email_list))
print("operation completed")
