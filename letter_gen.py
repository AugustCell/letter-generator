#!/usr/bin/python3

import docx
import pandas as pd
import os

path_to_save = os.path.abspath("F:/Documents/Real Estate/Letter Campaign/3_2_22")

def secondNameValid(name):
    # Check if this cell value is null
    if pd.isnull(name):
        return False
	
    # Check if this cell value name is multiple names.
    # Attempting to filter out trusts.
    strName = str(name)
    if len(strName.split()) > 1:
        return False

    return True

if __name__ == "__main__":
    data = pd.read_excel("property_list.xlsx")
    info = pd.DataFrame(data, columns=['First Name', 'Property Address'])
    second_names = pd.DataFrame(data, columns=["Owner 2 First Name"]).to_dict('dict')

    doc = docx.Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Lucida Handwriting'
    font.size = docx.shared.Pt(14)

    fileData = None
    with open('letter_template.txt') as f:
        fileData = f.read()

    fileData = fileData.replace("%%%", "Augusto")
    fileData = fileData.replace("###", "Amanda")
    fileData = fileData.replace("$$$", "3177 140th street")
    print(fileData)

    paragraph = doc.add_paragraph(fileData)
    doc.save(os.path.join(path_to_save, "temp.docx"))

    for idx, (name, propAddress) in enumerate(zip(info.iloc[:, 0], info.iloc[:, 1])):
        second_name = second_names.get('Owner 2 First Name').get(idx)
        valid_name = secondNameValid(second_name)


        # if valid_name:

        #     print("Cell Number: ", idx, "| First name: ", name, "| Address: ", propAddress, "| Second name: ", second_name)
        # else:
        #     print("Cell Number: ", idx, "| First name: ", name, "| Address: ", propAddress)
        # print()

    #   if len(name.split()) > 1:
    # 		print("MARKED AS TRUST!")
    # 	else:
    # 		pass
		# 	if not pd.isnull(second_names.get('Owner 2 First Name').get(idx)) or len(second_names.get('Owner 2 First Name').get(idx).split()) < 2:
		# 		print("Cell Number: ", idx, "| First name: ", name, "| Address: ", propAddress, "| Second name: ", second_names.get('Owner 2 First Name').get(idx))
		# 	else:
		# 		print("Cell Number: ", idx, "| First name: ", name, "| Address: ", propAddress)
		# print()
