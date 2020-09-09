import tabula
import pdfplumber
import pandas as pd
import sys
import os
from datetime import datetime
import re

def write_into_file(content):
    # my_file.write("{\n")
    my_file.write(content)

# Change multiline of a cell to json
def multiline_to_json(s):
    rr = ""
    for ss in s.splitlines():
        if ss.find(":") > -1:
            sss = '"' + remove_special_characters(ss.split(":", 1)[0]) + '": "' + remove_special_characters(ss.split(":", 1)[1]) + '"'
        else:
            sss = '"' + remove_special_characters(ss) + '": ""'
        if rr != "" and sss != "": rr += ", "
        rr += sss
    return rr

def remove_special_characters(content):
    # return ''.join(e for e in content if e.isalnum() or e == ' ')
    return content.replace('"', '\\"')

# datetime object containing current date and time
now = datetime.now()
dt_string = now.strftime("%Y-%m-%d_%H-%M-%S")

# Get pdf file name from command line
if len(sys.argv) == 1:
    print("INPUT TYPE: python extract_pdf.py pdf_file_name [output_path]")
    exit()
elif len(sys.argv) >= 2:
    pdf_file = sys.argv[1]
    output_path = "."
    if len(sys.argv) > 2: output_path = sys.argv[2]

filename = os.path.basename(pdf_file)
filename = filename[:filename.rfind(".")]

# Get row data of the tables in a pdf file using tabula and Save them as array
df = tabula.read_pdf(pdf_file, pages="all", guess = False, multiple_tables = True) 
t_word = []
# first_row = []
for a in df:
    # first_row_tmp = []
    t_word_page = []
    for i in range(len(a.index)):
        count = 0
        t_word_tmp = []
        for c in a:
            ss = str(a.loc[i, c])
            if ss != "nan" and ss.strip() != "":
                t_word_tmp.append(" ".join(ss.splitlines()))
        t_word_page.append(t_word_tmp)
    #     if len(first_row_tmp) <2 and len(t_word_tmp) == 1 :
    #         first_row_tmp.append(t_word_tmp[0])
        
    #     print(str(i) + "  " + str(c) + "  " + "::".join(t_word_tmp))
    # first_row.append(first_row_tmp[0])
    # print("first_row :: " + first_row_tmp[0])
    t_word.append(t_word_page)


# Parse a pdf file with pdfplumber
pdf = pdfplumber.open(pdf_file) 

# Open output text file
# filename += "_" + f"{dt_string}.txt"
filename = "output.txt"
my_file = open(f"{output_path}\{filename}", "w", encoding="utf8")
write_into_file("{\n")


# Main Working Flow
page_num = 0 # current page number (1 ~ )
write_started = False # Whether writing started


# Iterate through pages
pre_ww_2 = ""
start_time_started = False
vendor_started = False
product_started = False
support_started = False

for p0 in pdf.pages:
    page_num += 1 # Increase current page
    cell = [] # An array of index of cell data
    word = [] # An array of cell data
    # Parse the tables of PDF
    table = p0.extract_table()    
    # Change table to DataFrame of Pandas
    df = pd.DataFrame(table[0:], columns=table[1])

    # Set an array of index of cell data and an array of cell data
    # Iterate through rows
    bit_data_flag = False
    add_row_count = 0
    
    # first row
    # cell_temp = []
    # word.append(first_row[page_num - 1])
    # for j in  range(len(df.columns)):
    #     cell_temp.append(len(word) - 1)
    # cell.append(cell_temp)
    skip_count = 0
    for i in range(len(df.index)):
        if skip_count > 0:
            skip_count -= 1
            continue

        # my_file.write("line " + str(i) + " : \n")
        count = 1
        cell_temp = []        
        print(str(len(df.index)) + ":::" + str(len(df.columns)))
        # Iterate through columns
        afe_flag = False
        for j in  range(len(df.columns)):
            
            ss = str(df.iloc[i, j]) 
            # my_file.write(str(j) + ":" + ss + "  ")
            if ss != "None" :  
                if ss.find("From - To") == 0 :
                    skip_count = 4     
                    add_row_count -= 4  
                    word_cell = []
                    cell_temp = []
                    for jj in  range(len(df.columns)):             
                        ss = str(df.iloc[i, jj])
                        word_cell.append(ss)
                    for jj in range(4):
                        for jjj in  range(len(df.columns)):             
                            ss = str(df.iloc[i + 1 + jj, jjj])
                            if ss != "" and ss != "None":
                                if ss in ["Code", "Status"]:
                                    word_cell[jjj] += " " + ss
                                else:
                                    word_cell[jjj] += ss
                    for jj in  range(len(df.columns)):
                        if word_cell[jj] != "None":
                            word.append(word_cell[jj])
                        cell_temp.append(len(word) - 1) 
                    cell.append(cell_temp)
                    
                    skip_count += 1
                    add_row_count -= 1
                    cell_temp = []
                    word_cell = []
                    for jj in  range(len(df.columns)):             
                        word_cell.append("")
                    is_empty_word_cell = True
                    for jj in range(len(str(df.iloc[i + 5, 0]).splitlines())):
                        if str(df.iloc[i + 5, 0]).splitlines()[jj].strip() != "":                            
                            for jjj in  range(len(df.columns)):
                                if not is_empty_word_cell:
                                    if word_cell[jjj] != "None" and word_cell[jjj] != "":
                                        word.append(word_cell[jjj])
                                    cell_temp.append(len(word) - 1)

                                if len(str(df.iloc[i + 5, jjj]).splitlines()) >= jj + 1:
                                    word_cell[jjj] = str(df.iloc[i + 5, jjj]).splitlines()[jj].strip()
                                else:
                                    word_cell[jjj] = ""
                            if not is_empty_word_cell:
                                add_row_count += 1
                                cell.append(cell_temp)
                                cell_temp = []
                            is_empty_word_cell = False
                        else:
                            for jjj in  range(len(df.columns)):
                                if len(str(df.iloc[i + 5, jjj]).splitlines()) >= jj + 1:
                                    ss = str(df.iloc[i + 5, jjj]).splitlines()[jj].strip()
                                    if ss != "" and ss != "None":
                                        word_cell[jjj] += " " + ss

                    for jjj in  range(len(df.columns)):
                        if word_cell[jjj] != "None" and word_cell[jjj] != "":
                            word.append(word_cell[jjj])
                        cell_temp.append(len(word) - 1)
                    add_row_count += 1


                    break
                word.append(ss)
                count += 1
                
            else:
                if j == 0 or (str(df.iloc[i, 0]) == "Output" and (j == 1 or j == 15)) or (str(df.iloc[i, 0]) == "Output" and (j == 1 or j == 15)):
                    word.append("")
                    count += 1

            cell_temp.append(len(word) - 1)
        cell.append(cell_temp)
        if afe_flag:
            cell_temp = []
            word.append('AFE & Field Estimated Cost')
            for j in  range(len(df.columns)):
               cell_temp.append( len(word) - 1) 
            cell.append(cell_temp)
        # my_file.write("\n")

    # Main Working Flow
    # Iterate through rows
    for i in range(len(df.index) + add_row_count): 
        # Iterate through columns
        for j in  range(len(df.columns)):
            ss = ""
            if cell[i][j] == -1: continue
            # Get cell number of data
            k = j + 1
            for k in range(j + 1, len(df.columns)): 
                if cell[i][k] != cell[i][j]: break
            else:
                if k == len(df.columns) - 1: k = len(df.columns)
            k -= 1
            # data (or key of group)
            ww_2 = " ".join(word[cell[i][j]].splitlines())           
            # print("ww_2::  " + ww_2 + "   i = " + str(i) + "  j = " + str(j) + "  k = " + str(k) + "  df.index = " + str(len(df.index) + add_row_count))
            # Whether data is group
            if i < len(df.index) + add_row_count-1 and k > j: 
                # print("ww_2:::::  " + ww_2 + "   i = " + str(i) + "  j = " + str(j) + "  k = " + str(k))
                if not (ww_2 in ["00:00 - 06:00 Update"]) and (j == 0 or (j > 0 and cell[i+1][j-1] != cell[i+1][j])) and (k == len(df.columns) - 1 or (k < len(df.columns) - 1 and cell[i+1][k] != cell[i+1][k+1])) and (cell[i+1][j] != cell[i+1][k] or ww_2 in ["Last Survey", "Team Leader, Supervisor, & Engineers", "AFE & Field Estimated Cost", "Safety", "Observation Card Summary", "Safety Comments", "Operations Summary", "Mud Data", "Fann Data", "Weather", "Remarks"]) :
                    # Get the number of rows in a group
                    for ii in range(i+2, len(df.index) + add_row_count): 
                        if ww_2 == "Last Survey" and word[cell[ii][j]] == "Team Leader, Supervisor, & Engineers": break
                        if ww_2 == "Team Leader, Supervisor, & Engineers" and word[cell[ii][j]] == "AFE & Field Estimated Cost": break
                        # if ww_2 == "Safety": print("aaa = " + str(word[cell[ii][j]]) + " ii = " + str(ii) + " j = " + str(j) + " cell = " + str(cell[ii][j]))  
                        if ww_2 == "AFE & Field Estimated Cost" and word[cell[ii][j]] == "Safety": break
                        if ww_2 == "Safety" and word[cell[ii][j]] == "Observation Card Summary": break
                        if ww_2 == "Observation Card Summary" and word[cell[ii][j]] == "Safety Comments": break
                        if ww_2 == "Mud Data" and word[cell[ii][j]] == "Fann Data": break
                        if ww_2 == "Fann Data" and word[cell[ii][j]].find('"Pump / Hydraulics"') == 0: break  
                        if ww_2 == "Weather" and word[cell[ii][j]].find('"Anchor Tension"') == 0: break

                        if not ((j == 0 or (j > 0 and cell[ii][j-1] != cell[ii][j])) and (k == len(df.columns) - 1 or (k < len(df.columns) - 1 and cell[ii][k] != cell[ii][k+1])) and (cell[ii][j] != cell[ii][k] or ww_2 in ["Last Survey", "Team Leader, Supervisor, & Engineers", "AFE & Field Estimated Cost", "Safety", "Observation Card Summary", "Safety Comments", "Operations Summary", "Mud Data", "Fann Data", "Weather", "Remarks"])) : break
                        if not(ww_2 in ["Last Survey", "Team Leader, Supervisor, & Engineers", "AFE & Field Estimated Cost", "Safety", "Observation Card Summary", "Safety Comments", "Operations Summary", "Mud Data", "Fann Data", "Weather", "Remarks"]):
                            for jj in range(j+1, k+1): 
                                if cell[ii][jj] - cell[ii-1][jj] != cell[ii][j] - cell[ii-1][j]:break
                            else:
                                continue
                            break
                    else:
                        if ii < i + 2: ii = i + 2                        
                        if ii == len(df.index) + add_row_count - 1 : ii += 1
                    if ww_2 == "Remarks": print(" ii = " + str(ii))
              
                    if ww_2 == "Time Log":
                        pre_cell = -2
                        header = []
                        # Get the column headers in a group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        for iii in range(i + 2, ii): 
                            group_column_num = 0
                            tmp_list = []
                            pre_cell = -2
                            # Get cell data
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    tmp_list.append( word[cell[iii][jj]])
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1

                            # Split cell data to multiple rows
                            while len(tmp_list) > 0 and len(tmp_list[0].splitlines()) > 0:

                                # Get data of the first line of the first column
                                comp_str = tmp_list[0].splitlines()[-1].strip()  
                                
                                # Compare data of word array with comp_str and Set data
                                for t_a in t_word[page_num - 1]:
                                    for t_b in t_a:
                                        # Whether Is there comp_str in data of word array
                                        if t_b.find(comp_str.strip()) > -1 :
                                            f = False
                                            for c in tmp_list:  
                                                if len(c.splitlines()) == 0: continue                                           
                                                for t_c in t_a:
                                                    if t_c.find(c.splitlines()[-1].strip()) > -1:  
                                                        f = True
                                                        break
                                                if f == False: break
                                            # If there is comp_str in data of word array, check other data
                                            if f == True:                                                
                                                for t_c in t_a:
                                                    if len(tmp_list) == 0: break
                                                    if len(tmp_list[0].splitlines()) ==0: break
                                                    if t_c.find(tmp_list[0].splitlines()[-1].strip()) == -1: continue
                                                    ssss = ""
                                                    tt = tmp_list[:]
                                                    for c in range(len(tmp_list)): 
                                                        sss = ""                                               
                                                        if len(tmp_list[c].splitlines()) == 0 :
                                                            sss = header[c] + ":   "
                                                            continue    

                                                        tt[c] =  tmp_list[c]                                                 
                                            
                                                        for c_len in range(len(tt[c].splitlines())):
                                                            
                                                            for t_d in t_a:
                                                                if t_d.find(tt[c].splitlines()[len(tt[c].splitlines())- c_len -1].strip()) > -1:
                                                                    cc = "  ".join(tt[c].splitlines()[len(tt[c].splitlines())- c_len -1:])
                                                                    sss = '"' + header[c] + '": "' + remove_special_characters(cc) + '" '
                                                                    tt[c] = "\n".join(tt[c].splitlines()[: len(tt[c].splitlines())- c_len -1])
                                                                    break
                                                            else:
                                                                continue
                                                            break
                                                        else:
                                                            break
                                                        if ssss != "": ssss += ", "
                                                        ssss += sss
                                                    else:
                                                        if ss == "":
                                                            ss = '{' + ssss + '}'
                                                        else:    
                                                            ss += ', {' + ssss + '}'
                                                        tmp_list = tt[:]

                                                if ss != "": ss += "\n"
                                                if len(tmp_list) == 0: 
                                                    write_into_file("\n")
                                                    write_into_file(ss)
                                                    break
                    elif ww_2 in ["OPERATION SUMMARY"]:
                        # col header
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        if header[0] == "Title Job Contact":
                            header[0] ="Title"
                            header.append("Job Contact")
                        # Get the columns data in group
                        ss += "[\n"
                        for iii in range(i + 2, ii): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        if sss != "": sss += ", "
                                        

                                        sss += '"' + header[header_cc] + '": "' + remove_special_characters(" ".join(word[cell[iii][jj]].splitlines())) + '" '
                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            sss = '{' + sss + '}'
                            if ss != "[\n": ss += ", \n"
                            ss += sss
                        ss += sss + '\n]'

                    elif ww_2 in ["Last Survey", "AFE & Field Estimated Cost", "Safety Comments", "Mud Data", "Fann Data", "Weather", "Remarks"]:
                        print("www = " + ww_2) 
                        pre_cell = -2
                        ss += "{"
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "  
                                    if word[cell[iii][jj]].find("Mud/Fluid Check Depth") > -1 or word[cell[iii][jj]].find("BHA # 4") > -1:
                                        word[cell[iii][jj]] = "\n".join(word[cell[iii][jj]].split(": "))
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    if len(word[cell[iii][jj]].splitlines()) > 0 :
                                        sss += '"' + remove_special_characters(word[cell[iii][jj]].splitlines()[0]) + '": "'
                                        if len(word[cell[iii][jj]].splitlines()) > 1 : 
                                            sss += remove_special_characters(" ".join(word[cell[iii][jj]].splitlines()[1:]))
                                        sss += '"'
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                if ss != "{": ss += ", "
                                ss += sss
                        ss += "}"

                    elif ww_2 in ["Safety"]:
                        # print("www = " + ww_2) 
                        pre_cell = -2
                        ss += "{\n"
                        if ww_2 == "Safety":
                            add_row = 1
                        elif ww_2 == "Operations Summary":
                            add_row = 2
                        for iii in range(i + 1, i + 1 + add_row): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "  
                                    if word[cell[iii][jj]].find("Mud/Fluid Check Depth") > -1 or word[cell[iii][jj]].find("BHA # 4") > -1:
                                        word[cell[iii][jj]] = "\n".join(word[cell[iii][jj]].split(": "))
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    if len(word[cell[iii][jj]].splitlines()) > 0 :
                                        sss += '"' + remove_special_characters(word[cell[iii][jj]].splitlines()[0]) + '": "'
                                        if len(word[cell[iii][jj]].splitlines()) > 1 : 
                                            sss += remove_special_characters(" ".join(word[cell[iii][jj]].splitlines()[1:]))
                                        sss += '"'
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                if ss != "{\n": ss += ", "
                                ss += sss
                        ss += ", \n[\n"
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1 + add_row][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1 + add_row][jj]].splitlines()))
                                pre_cell = cell[i+1 + add_row][jj]
                            cell[i+1 + add_row][jj] = -1
                        # Get the columns data in group
                        for iii in range(i + 2 + add_row, ii): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        if sss != "": sss += ", "
                                        sss += '"' + header[header_cc] + '": "' + remove_special_characters(" ".join(word[cell[iii][jj]].splitlines())) + '" '
                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            sss = '{' + sss + '}'
                            if ss[-2] != "[": ss += ", \n"

                            ss += sss

                        ss += "\n]}"

                    elif ww_2 in ["Operations Summary"]:
                        # print("www = " + ww_2)
                        start_time_started = False 
                        pre_cell = -2
                        ss += "{\n"
                        
                        add_row = 2
                        for iii in range(i + 1, i + 1 + add_row): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "  
                                    if word[cell[iii][jj]].find("Mud/Fluid Check Depth") > -1 or word[cell[iii][jj]].find("BHA # 4") > -1:
                                        word[cell[iii][jj]] = "\n".join(word[cell[iii][jj]].split(": "))
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    if len(word[cell[iii][jj]].splitlines()) > 0 :
                                        sss += '"' + remove_special_characters(word[cell[iii][jj]].splitlines()[0]) + '": "'
                                        if len(word[cell[iii][jj]].splitlines()) > 1 : 
                                            sss += remove_special_characters(" ".join(word[cell[iii][jj]].splitlines()[1:]))
                                        sss += '"'
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                if ss != "{\n": ss += ", \n"
                                ss += sss
                        ss += ", \n[\n"
                        ss += word[cell[i + 1 + add_row][0]]
                        cell[i + 1 + add_row][0] = -1

                    elif ww_2.find('{"start_time":') == 0:
                        ss += ", \n"
                        ss += word[cell[i + 1 + add_row][0]]
                        cell[i + 1 + add_row][0] = -1

                        
                    else: # etc group
                        # print("ww_s = " + ww_2)
                        pre_cell = -2
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    sss += word[cell[iii][jj]]
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                sss = '{' + sss + '}'
                                
                                if ss != "": ss += ", "
                                ss += sss
            # Output to output file
            if ww_2 != "": 
                ww_2_= ww_2
                
                if ww_2.find('{"start_time"') == 0 or ww_2.find('"Offline Activity Dates":') == 0 or ww_2.find('"Offline Activity Time Logs":') == 0 or ww_2.find('"Pump / Hydraulics":') == 0 or ww_2.find('"Bits":') == 0 or ww_2.find('"BHA #<stringno>, <des>":') == 0 or ww_2.find('"Assembly Components":') == 0 or ww_2.find('"Surveys":') == 0 or ww_2.find('"Anchor Tension":') == 0:
                    if ww_2.find('{"start_time"') == 0 and not start_time_started:
                        ww_2 = word[cell[i][j]]
                        start_time_started = True                        
                    else:
                        ww_2 = ", \n" + word[cell[i][j]]
                elif ww_2.find('{"Vendor":') == 0:
                    if ww_2.find('{"Vendor":') == 0 and not vendor_started:
                        ww_2 = word[cell[i][j]]
                        vendor_started = True                        
                    else:
                        ww_2 = ", \n" + word[cell[i][j]]  # Product Name Unit Label Received Consumed Cum On Loc
                elif ww_2.find('{"Product Name":') == 0:
                    if ww_2.find('{"Product Name":') == 0 and not product_started:
                        ww_2 = word[cell[i][j]]
                        product_started = True                        
                    else:
                        ww_2 = ", \n" + word[cell[i][j]] 
                elif ww_2.find('{"Vessel Name":') == 0:
                    if ww_2.find('{"Vessel Name":') == 0 and not support_started:
                        ww_2 = word[cell[i][j]]
                        support_started = True                        
                    else:
                        ww_2 = ", \n" + word[cell[i][j]]     
                elif ww_2 == "00:00 - 06:00 Update":                    
                    if pre_ww_2 != ww_2_:
                        ww_2 = ', \n"' + ww_2 + '": [\n'
                        start_time_started = False
                    else:
                        start_time_started = True
                        ww_2 = ""
                elif ww_2 == "Personnel":                    
                    if pre_ww_2 != ww_2_:
                        ww_2 = ', \n"' + ww_2 + '": [\n'
                        vendor_started = False
                    else:
                        vendor_started = True
                        ww_2 = ""
                elif ww_2 == "Material - Bulk":                    
                    print("pre = " + pre_ww_2 + "  \nww_2 = " + ww_2 + "\nww_2_ = " + ww_2_)
                    if pre_ww_2 != ww_2_:
                        ww_2 = ', \n"' + ww_2 + '": [\n'
                        product_started = False
                    else:
                        product_started = True
                        ww_2 = ""
                elif ww_2 == "Support Vessels":                    
                    print("pre = " + pre_ww_2 + "  \nww_2 = " + ww_2 + "\nww_2_ = " + ww_2_)
                    if pre_ww_2 != ww_2_:
                        ww_2 = ', \n"' + ww_2 + '": [\n'
                        support_started = False
                    else:
                        support_started = True
                        ww_2 = ""
                else: 
                    ww_2 = remove_special_characters(ww_2)
                    ww_2_0 = remove_special_characters(word[cell[i][j]].splitlines()[0])

                    if ww_2 in ["Weather", "Remarks"]:
                        write_into_file("\n]")
                    if write_started:
                        write_into_file(", \n")
                    if ww_2.find("Time Log Total Hours (hr)") == 0:
                        ww_2 = '],\n"' + word[cell[i][j]].splitlines()[0] + '": "' + str(word[cell[i][j]].splitlines()[1]) + '"\n}'
                    elif ww_2.find("Head Count") == 0:
                        ww_2 = '],\n"' + word[cell[i][j]].splitlines()[0] + '": "' + str(word[cell[i][j]].splitlines()[1]) + '"\n}'
                    elif ww_2_0 in ["Spud Date", "Depth Progress (m)", "Current Depth (mKB)", "Current Depth (TVD) (…", "Authorized MD (mKB)", "Water Depth (m)", "Orig KB Elev (m)", "KB-MudLn (m)", "PTTEP Field Name", "Block", "Country", "State/Province", "District", "Latitude (°)", "Longitude (°)", "Contractor", "Rig Name/No", "Rig Phone/Fax Number", "BHA Hrs of Service (hr)", "Leak Off Equivalent Fluid Density (lb/gal)", "Last Casing String", "Next Casing String"]:
                        ww_2 = '"' + ww_2_0 + '": "'
                        if len(word[cell[i][j]].splitlines()) > 1:
                            ww_2 += remove_special_characters(" ".join(word[cell[i][j]].splitlines()[1:]))
                        ww_2 += '"'
                    elif ss != "" and (ss[0] == "{" or ss[0] == "["): 
                        # ww_2 = '"' + ww_2 + '": [\n'
                        ww_2 = '"' + ww_2 + '": '
                    elif ww_2.find(":") > -1:
                        ww_2 = '"' + remove_special_characters(ww_2.split(":", 1)[0]) + '": "' + ww_2.split(":", 1)[1] + '"'
                    elif len(ww_2.splitlines()) > 1:
                        ww_2 = '"' + remove_special_characters(ww_2.splitlines()[0]) + '": "' + "  ".join(ww_2.splitlines()[1:]) + '"'
                    else:
                        # Let's remove special characters if it's header
                        ww_2 = remove_special_characters(ww_2)
                        ww_2 = '"' + ww_2 + '": ""'

                

                write_into_file(ww_2)
                write_started = True


            if ss != "": write_into_file(ss)

            # if ww_2 != "" and ss != "" and ss[0] == "{": my_file.write('\n}')
            # if ww_2 != "" and ss != "" and ss[0] == "{": write_into_file('\n]')
            for jj in range(j, k+1):
                cell[i][jj] = -1
            if ww_2_.find('{"start_time"') != 0 and ww_2_.find('{"Vendor"') != 0 and ww_2_.find('{"Vessel Name"') != 0:
                pre_ww_2 = ww_2_
    print("page number = " + str(page_num) + "  pre_ww_2 = " + ww_2)  
          
    # my_file.write("\n")
# my_file.write("}")
write_into_file("\n")
write_into_file("}")
my_file.close()
