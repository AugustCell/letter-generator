#!/usr/bin/python3

import docx
import pandas as pd
import os
from docx2pdf import convert

#path_to_save = os.path.abspath("F:/Documents/Real Estate/Letter Campaign/3_2_22")
path_to_save = os.path.abspath("C:/Users/alast/Documents/Real Estate/saved_letters")

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
    info = pd.DataFrame(data, columns=['First Name', 'Property Address', 'Last Name'])
    second_names = pd.DataFrame(data, columns=["Owner 2 First Name"]).to_dict('dict')

    for idx, (name, propAddress, lastName) in enumerate(zip(info.iloc[:, 0], info.iloc[:, 1], info.iloc[:, 2])):

        # Get doc object setup
        doc = docx.Document()
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Lucida Handwriting'
        font.size = docx.shared.Pt(14)

        # Get intial letter template data inside of fileData object
        fileData = None
        with open('letter_template.txt') as f:
            fileData = f.read()

        # Check if we have a valid second name
        second_name = second_names.get('Owner 2 First Name').get(idx)
        valid_second_name = secondNameValid(second_name)

        if not valid_second_name:
            print_name = " ".join(name.split()) + ","
            print_address = propAddress
            fileData = fileData.replace("%%%", print_name)
            fileData = fileData.replace("||", "")
            fileData = fileData.replace("###", "")
            fileData = fileData.replace("$$$", print_address)
            print("MARK!")
            print("Cell Number: ", idx, "| First name: ", name, "| Address: ", propAddress)

        else:
            print_name = " ".join(name.split())
            print_address = propAddress
            second_name = second_name + ","
            fileData = fileData.replace("%%%", print_name)
            fileData = fileData.replace("||", "and")
            fileData = fileData.replace("###", second_name)
            fileData = fileData.replace("$$$", print_address)
            print("Cell Number: ", idx, "| First name: ", name, "| Address: ", propAddress, "| Second name: ", second_name)
        print()

        paragraph = doc.add_paragraph(fileData)
        save_name = name + "_" + lastName + "_" + "letter.docx"
        doc.save(os.path.join(path_to_save, save_name))

    convert(path_to_save, os.path.join(path_to_save, "pdf"))
