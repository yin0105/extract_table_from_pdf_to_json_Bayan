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

def multiline_to_json_2(s):
    rr = ""
    for ss in s.splitlines():
        if ss.find(":") > -1:
            sss = '"' + remove_special_characters(ss.split(":", 1)[0]) + '": "' + remove_special_characters(ss.split(":", 1)[1]) + '"'
            if rr != "" and sss != "": rr += ", "
            rr += sss
        else:
            rr = rr[:-1] + " " + remove_special_characters(ss) + '"'
        
    return rr

def multiline_to_json_num(s):
    rr = ""
    for ss in s.splitlines():
        ss = ss.strip()
        if ss == "" : continue
        if ss[-1] >= '0' and ss[-1] <= '9':
            sss = '"' + remove_special_characters(" ".join(ss.split(" ")[:-1])) + '": "' + remove_special_characters(ss.split(" ")[-1]) + '"'
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
operation_started = False

start_time_started = False
vendor_started = False
product_started = False
support_started = False
operation_rows_count = [
    [0, 4, 30, 35],
    [4, 5, 9, 14, 16, 19, 21, 22, 24, 28, 32, 36, 38, 42, 45, 57],
    [1, 12, 15, 18, 51],
    [9]
]
for p0 in pdf.pages:
    page_num += 1 # Increase current page
    cell = [] # An array of index of cell data
    word = [] # An array of cell data
    # Parse the tables of PDF
    table = p0.extract_table()    
    # Change table to DataFrame of Pandas
    if page_num == 1:
        df = pd.DataFrame(table[1:], columns=table[1])
    else:
        df = pd.DataFrame(table[2:], columns=table[1])

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
                        if str(df.iloc[i + 5, jj]) != "":        
                            word_cell.append(str(df.iloc[i + 5, jj]).splitlines()[0])
                        else:
                            word_cell.append("")
                        for jjj in str(df.iloc[i + 5, jj]).splitlines()[1:]:    
                            if jjj.strip() != "" and jjj.strip() != "None":
                                word_cell[jj] += "\n" + jjj
                    if len(word_cell[0].splitlines()) > 0:
                        for jj in range(len(word_cell[0].splitlines())):
                            add_row_count += 1
                            cell_temp = []
                            last_jjj = 0
                            for jjj in  range(len(df.columns)):
                                if len(word_cell[jjj].splitlines()) > jj:                                
                                    www = word_cell[jjj].splitlines()[jj]
                                    if www != "None" and www != "":
                                        last_jjj = jjj
                                        word.append(www)
                                cell_temp.append(len(word) - 1)
                            
                            if jj == 0 and operation_rows_count[page_num - 1][0] > 0:
                                word[len(word) - 1] = " ".join(word_cell[last_jjj].splitlines()[0 : operation_rows_count[page_num - 1][jj]])
                                word[len(word) - 1] += "#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*"
                                word[len(word) - 1] += " ".join(word_cell[last_jjj].splitlines()[operation_rows_count[page_num - 1][jj] : operation_rows_count[page_num - 1][jj + 1]])
                            else:
                                word[len(word) - 1] = " ".join(word_cell[last_jjj].splitlines()[operation_rows_count[page_num - 1][jj] : operation_rows_count[page_num - 1][jj + 1]])
                            if jj < len(word_cell[0].splitlines()) - 1: cell.append(cell_temp)
                    else:
                        add_row_count += 1
                        cell_temp = []
                        last_jjj = 0
                        word.append("")
                        cell_temp.append(len(word) - 1)
                        for jjj in  range(1, len(df.columns)):
                            if len(word_cell[jjj].splitlines()) > 0:                                
                                word.append(" ".join(word_cell[jjj].splitlines()))
                            cell_temp.append(len(word) - 1)

                    break

                elif ss.find("SURVEYS") == 0 :
                    skip_count = 2    
                    add_row_count -= 1
                    cell_temp = []
                    word.append("SURVEYS")
                    for jj in  range(len(df.columns)):
                        cell_temp.append(len(word) - 1)
                    cell.append(cell_temp)

                    word_cell = []
                    cell_temp = []
                    for jj in  range(len(df.columns)):             
                        ss = str(df.iloc[i + 1, jj])
                        word_cell.append(ss)
                    for jjj in  range(len(df.columns)):             
                        ss = str(df.iloc[i + 2, jjj])
                        if ss != "" and ss != "None":
                            word_cell[jjj] += " " + ss

                    for jj in  range(len(df.columns)):
                        if word_cell[jj] != "None":
                            word.append(word_cell[jj])
                        cell_temp.append(len(word) - 1) 

                    break

                elif ss.find("Daily NPT") == 0 :
                    word.append(ss)
                    cell_temp = []
                    for jj in range(19):
                        cell_temp.append(len(word) - 1)
                    word.append("")
                    for jj in range(19, len(df.columns)):
                        cell_temp.append(len(word) - 1)
                    break

                word.append(ss)
                count += 1
                if str(df.iloc[i, 0]) != "BIT DATA" and bit_data_flag and j > 0 and j < 9 : 
                    if ss != "":
                        bit_data_flag = False
                
            else:
                if str(df.iloc[i, 0]) == "GAS READINGS": bit_data_flag = False
                if str(df.iloc[i, 0]) == "BIT DATA": 
                    bit_data_flag = True
                if str(df.iloc[i, 0]) != "BIT DATA" and bit_data_flag and (j == 8 or j == 9):
                    word.append("")
                    count += 1
                elif j == 0 or (str(df.iloc[i, 0]) == "Output" and (j == 1 or j == 15)) or (str(df.iloc[i, 0]) == "Output" and (j == 1 or j == 15)):
                    word.append("")
                    count += 1

            cell_temp.append(len(word) - 1)
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
            # print("ww_2::  " + ww_2 + "   i = " + str(i) + "  j = " + str(j) + "  k = " + str(k) + "  df.index = " + str(len(df.index) + add_row_count) + " df.columns = " +str(len(df.columns)))
            # Whether data is group
            if i < len(df.index) + add_row_count-1 and k > j: 
                # print("ww_2:::::  " + ww_2 + "   i = " + str(i) + "  j = " + str(j) + "  k = " + str(k))
                if (j == 0 or (j > 0 and cell[i+1][j-1] != cell[i+1][j])) and (k == len(df.columns) - 1 or (k < len(df.columns) - 1 and cell[i+1][k] != cell[i+1][k+1])) and (cell[i+1][j] != cell[i+1][k] or ww_2 in ["WELL INFO", "DEPTH DAYS", "COSTS()", "STATUS", "HYDROCYCLONE", "WEATHER",   "STATUS", "MUD CHECK", "BIT DATA", "GAS READINGS", "MUD VOLUME", "PERSONNEL", "PUMP/HYDRAULICS"]) :
                    # Get the number of rows in a group
                    for ii in range(i+2, len(df.index) + add_row_count): 
                        if ww_2 == "WELL INFO" and word[cell[ii][j]] == "DEPTH DAYS": break
                        if ww_2 == "DEPTH DAYS" and word[cell[ii][j]] == "STATUS": break
                        if ww_2 == "COSTS()" and word[cell[ii][j]] == "STATUS": break
                        if ww_2 == "STATUS" and word[cell[ii][j]] == "OPERATION SUMMARY": break
                        if ww_2 == "SHAKER" and word[cell[ii][j]] == "LOT/FIT": break
                        if ww_2 == "SAFETY CARDS" and word[cell[ii][j]] == "ANCHOR TENSION": break
                        if ww_2 == "ANCHOR TENSION" and word[cell[ii][j]] == "SAFETY": break


                        # if ww_2 == "AFE & Field Estimated Cost" and word[cell[ii][j]] == "Safety": break
                        # if ww_2 == "Safety" and word[cell[ii][j]] == "Observation Card Summary": break
                        # if ww_2 == "Observation Card Summary" and word[cell[ii][j]] == "Safety Comments": break
                        # if ww_2 == "Mud Data" and word[cell[ii][j]] == "Fann Data": break
                        # if ww_2 == "Fann Data" and word[cell[ii][j]].find('"Pump / Hydraulics"') == 0: break  
                        # if ww_2 == "Weather" and word[cell[ii][j]].find('"ANCHOR TENSION"') == 0: break
                        # if ww_2 == "ANCHOR TENSION" and word[cell[ii][j]].find('"SAFETY"') == 0: break
                        if not(ww_2 in ["SAFETY CARDS", "ANCHOR TENSION"]):
                            if not ((j == 0 or (j > 0 and cell[ii][j-1] != cell[ii][j])) and (k == len(df.columns) - 1 or (k < len(df.columns) - 1 and (cell[ii][k] != cell[ii][k+1] or ww_2 in ["DEPTH DAYS", "COSTS()", "CENTRIFUGE", "HYDROCYCLONE", "SHAKER"]))) and (cell[ii][j] != cell[ii][k] or ww_2 in ["WELL INFO", "DEPTH DAYS", "COSTS()", "STATUS", "CENTRIFUGE", "HYDROCYCLONE", "SHAKER", "WEATHER",       "STATUS", "MUD CHECK", "BIT DATA", "GAS READINGS", "MUD VOLUME", "PERSONNEL", "PUMP/HYDRAULICS"])) : break
                        if not(ww_2 in ["WELL INFO", "DEPTH DAYS", "COSTS()", "STATUS", "OPERATION SUMMARY", "CENTRIFUGE", "HYDROCYCLONE", "SHAKER", "WEATHER", "SAFETY CARDS", "ANCHOR TENSION",              "COSTS", "STATUS", "MUD CHECK", "BIT DATA", "GAS READINGS", "MUD VOLUME", "PERSONNEL", "PUMP/HYDRAULICS", "SHAKER"]):
                            for jj in range(j+1, k+1): 
                                if cell[ii][jj] - cell[ii-1][jj] != cell[ii][j] - cell[ii-1][j]:break
                            else:
                                continue
                            break
                    else:
                        if ii < i + 2: ii = i + 2                        
                        if ii == len(df.index) + add_row_count - 1 : ii += 1
                                  
                    if ww_2 in ["OPERATION SUMMARY"]:
                        # col header
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        if not operation_started:
                            ss += "[\n"
                        for iii in range(i + 2, ii): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                if word[cell[iii][0]] == "":
                                    ss += " " + remove_special_characters(" ".join(word[cell[iii][0] + 1].splitlines()))
                                    for jjj in range(j, k+1): cell[iii][jjj] = -1
                                    break
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        if header[header_cc] == "Operation" and word[cell[iii][jj]].find("#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*") > -1:
                                            ss += " " + remove_special_characters(" ".join(word[cell[iii][jj]][:word[cell[iii][jj]].find("#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*")].splitlines())) + '"}'
                                        if sss != "": sss += ", "                                        

                                        if header[header_cc] == "Operation" and word[cell[iii][jj]].find("#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*") > -1:
                                            sss += '"' + header[header_cc] + '": "' + remove_special_characters(" ".join(word[cell[iii][jj]][word[cell[iii][jj]].find("#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*#*") + 40:].splitlines())) + '" '
                                        else:
                                            sss += '"' + header[header_cc] + '": "' + remove_special_characters(" ".join(word[cell[iii][jj]].splitlines())) + '" '
                                        if header_cc < len(header) - 1 and header[header_cc + 1] == "NPT":
                                            sss += ', "' + header[header_cc + 1] + '": "" '
                                            header_cc += 1
                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            if sss != "": sss = '{' + sss + '}'
                            if ss != "[\n" and sss != "": ss += ", \n"
                            ss += sss
                        # ss += '\n]'
                    
                    elif ww_2 in ["SAFETY CARDS", "ANCHOR TENSION"]:
                        # col header
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        ss += "[\n"
                        for iii in range(i + 2, ii): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                ssss = remove_special_characters(word[cell[iii][jj]])
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] == "":
                                        if ssss == "": 
                                            for jjj in range(jj, k + 1):
                                                cell[iii][jjj] = -1
                                            break
                                        if ssss.find("Total Cards") == 0:
                                            sss += '"": "' + " ".join(re.split(" +", ssss)[:-2]) + '", "' + header[1] + '": "' + str(re.split(" +", ssss)[-2]) + '", "' + header[2] + '": "' + str(re.split(" +", ssss)[-1]) + '" '
                                            for jjj in range(jj, k + 1):
                                                cell[iii][jjj] = -1
                                            break
                                        # sss += '"' + ssss + '": {'
                                    if len(sss) > 0 and sss[-1] != "{": sss += ", "      
                                    sss += '"' + header[header_cc] + '": "' + ssss + '" '

                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            if sss != "": sss = '{' + sss + '}'
                            if ss != "[\n" and sss != "": ss += ", \n"
                            ss += sss
                        ss += '\n]'
                    
                    elif ww_2 in ["SUPPORT CRAFT"]:
                        # col header
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        ss += "[\n"
                        for iii in range(len(word[cell[i + 2][j]].splitlines())): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[i + 2][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        if sss != "": sss += ", "    
                                        if len(word[cell[i + 2][jj]].splitlines()) > iii:                                    
                                            sss += '"' + header[header_cc] + '": "' + remove_special_characters(word[cell[i + 2][jj]].splitlines()[iii]) + '" '
                                        else:
                                            sss += '"' + header[header_cc] + '": "" '
                                    pre_cell = cell[i + 2][jj]
                                    header_cc += 1
                                # cell[iii][jj] = -1
                            if sss != "": sss = '{' + sss + '}'
                            if ss != "[\n" and sss != "": ss += ", \n"
                            ss += sss
                        for jj in range(j, k + 1):
                            cell[i + 2][jj] = -1
                        ss += '\n]'
                    
                    elif ww_2 in ["PERSONNEL"]: 
                        # Double Column Header
                        pre_cell = -2
                        sss = ""
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                if sss != "": sss += ", "
                                sss += '"' + '": "'.join(" ".join(word[cell[i+1][jj]].splitlines()).split(":")) + '"'
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # sss = '"' + '": "'.join(sss.split(":")) + '"'
                        sss = '{' + sss + '}'
                        if ss != "": ss += ", "
                        ss += sss + ', \n"": [\n'
                        # Get the columns data in group
                        pre_cell = -2
                        header = []
                        for jj in range(j, k + 1):
                            if cell[i+2][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+2][jj]].splitlines()))
                                pre_cell = cell[i+2][jj]
                            cell[i+2][jj] = -1
                        # Get the columns data in group
                        for iii in range(i + 3, ii): 
                            header_cc = 0
                            sss = ""
                            hh_count = 0
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        hh_count += 1
                                        if sss != "": 
                                            if hh_count == 3:
                                                sss += "}, {"
                                            else:
                                                sss += ", "
                                        sss += '"' + header[header_cc] + '": "' + remove_special_characters(word[cell[iii][jj]]) + '" '
                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            sss = '{' + sss + '}'
                            if ss[-2] != "[": ss += ", "
                            ss += sss
                        ss += "\n]"
                        
                    elif ww_2 in ["BULKS", "PUMP/HYDRAULICS", "BIT DATA", "CENTRIFUGE"]: 
                        print("www = " + ww_2)
                        # Row Header
                        ss += "[\n"
                        for jj in range(j + 1, k + 1):
                            if cell[i+1][jj] == cell[i+1][jj-1]: continue
                            sss = ""
                            for iii in range(i + 1, ii):                                 
                                if ' '.join(word[cell[iii][j]].splitlines()) == '' : continue
                                if sss != "" : sss += ", "
                                sss += '"' + ' '.join(word[cell[iii][j]].splitlines()) + '": "' + remove_special_characters(' '.join(word[cell[iii][jj]].splitlines())) + '"'
                            sss = '{' + sss + '}'
                            
                            if ss != "[\n": ss += ", "
                            ss += sss
                        ss += "\n]"
                        for iii in range(i + 1, ii):
                            for jj in range(j, k+1):
                                cell[iii][jj] = -1

                    elif ww_2 in ["PIPE DATA", "ANN. VELOCITIES"]: 
                        # Multiple Row Header
                        jj_count = 0
                        base_j = j
                        sss = ""
                        for jj in range(j + 1, k + 1):
                            if cell[i+1][jj] == cell[i+1][jj-1]: continue
                            jj_count += 1
                            if jj_count % 2 == 0:
                                base_j = jj
                                continue

                            
                            for iii in range(i + 1, ii): 
                                if sss != "" : sss += ", "
                                sss += '"' + remove_special_characters(' '.join(word[cell[iii][base_j]].splitlines())) + '": "' + remove_special_characters(' '.join(word[cell[iii][jj]].splitlines())) + '"'
                                # print(str(iii) + "  " + str(jj) + "  " + " ".join(word[cell[iii][j]].splitlines()) + ": " +" ".join(word[cell[iii][jj]].splitlines()))
                        sss = '{' + sss + '}'
                            
                        if ss != "": ss += ", "
                        ss += sss
                        for iii in range(i +1, ii):
                            for jj in range(j, k+1):
                                cell[iii][jj] = -1
                    
                    elif ww_2 in ["SHAKER", "HYDROCYCLONE", "LOT/FIT", "FORMATION DATA"]: # if name of group is ...
                        print("www = " + ww_2)
                        pre_cell = -2
                        header = []
                        # Get the columns data in group
                        for jj in range(j, k + 1):
                            if cell[i+1][jj] != pre_cell:
                                header.append(" ".join(word[cell[i+1][jj]].splitlines()))
                                if ww_2 == "HYDROCYCLONE": print(" ".join(word[cell[i+1][jj]].splitlines()))
                                pre_cell = cell[i+1][jj]
                            cell[i+1][jj] = -1
                        # Get the columns data in group
                        if ww_2 == "HYDROCYCLONE": ii = i + 3
                        ss += "[\n"
                        for iii in range(i + 2, ii): 
                            header_cc = 0
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if header[header_cc] != "":
                                        if sss != "": sss += ", "
                                        sss += '"' + header[header_cc] + '": "' + remove_special_characters(word[cell[iii][jj]]) + '" '
                                    pre_cell = cell[iii][jj]
                                    header_cc += 1
                                cell[iii][jj] = -1
                            while header_cc < len(header) :
                                sss += ', "' + header[header_cc] + '": ""'
                                header_cc += 1
                                
                            sss = '{' + sss + '}'
                            if ss != "[\n": ss += ", "
                            ss += sss
                        ss += "\n]"

                    elif ww_2 in ["DEPTH DAYS", "COSTS()", "MUD VOLUME"]:
                        pre_cell = -2
                        ss += "{"
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:                                      
                                    for ssss in word[cell[iii][jj]].splitlines():
                                        if sss != "": sss += ", "
                                        sss += '"' + remove_special_characters(ssss.split(":")[0].strip()) + '": "'
                                        if len(ssss.split(":")) > 1 :
                                            if ssss.find("Daily NPT") == 0:
                                                sss += remove_special_characters(ssss[:ssss.find("Cumm NPT")].split(":")[1].strip())
                                                sss += '", "' + remove_special_characters(ssss[ssss.find("Cumm NPT"):].split(":")[0].strip()) + '": "'
                                                sss += remove_special_characters(ssss[ssss.find("Cumm NPT"):].split(":")[1].strip())
                                            else:
                                                sss += remove_special_characters(ssss.split(":")[1].strip())
                                        sss += '"'
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                if ss != "{": ss += ", "
                                ss += sss
                        ss += "}"

                    elif ww_2 in ["WELL INFO"]:
                        pre_cell = -2
                        ss += "{"
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "  
                                    
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    if len(word[cell[iii][jj]].splitlines()) > 0 :
                                        sss += '"' + remove_special_characters(ww.split(":")[0].strip()) + '": "'
                                        if len(ww.split(":")) > 1 : 
                                            sss += remove_special_characters(" ".join(ww.split(":")[1:]).strip())
                                        sss += '"'
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                if ss != "{": ss += ", "
                                ss += sss
                        ss += "}"                    

                                            
                    elif ww_2 in ["STATUS"]:
                        # print("ww_s = " + ww_2)
                        pre_cell = -2
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "
                                    # ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    sss += multiline_to_json_2(word[cell[iii][jj]]) #sss += word[cell[iii][jj]]
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                sss = '{' + sss + '}'
                                
                                if ss != "": ss += ", "
                                ss += sss

                    else: # etc group
                        print("ww_etc = " + ww_2)
                        pre_cell = -2
                        for iii in range(i + 1, ii): 
                            sss = ""
                            for jj in range(j, k + 1):
                                if cell[iii][jj] != pre_cell:
                                    if sss != "" and word[cell[iii][jj]] != "": sss += ", "
                                    ww = " ".join(word[cell[iii][jj]].splitlines()) 
                                    if ww_2 == "MUD CHECK" and iii == i + 2:                                        
                                        sss += multiline_to_json_num(word[cell[iii][jj]]) #sss += word[cell[iii][jj]]
                                    else:
                                        sss += multiline_to_json(word[cell[iii][jj]]) #sss += word[cell[iii][jj]]
                                    pre_cell = cell[iii][jj]
                                cell[iii][jj] = -1
                            if sss != "":
                                if ww_2 == "MUD CHECK":
                                    sss = '{' + sss
                                else:
                                    sss = '{' + sss + '}'
                                
                                if ss != "": ss += ", "
                                if ww_2 == "MUD CHECK" and iii == i + 2: ss += '"": ['
                                ss += sss
                        if ww_2 == "MUD CHECK": ss += "}]}"

                    
            # Output to output file
            if ww_2 != "": 
                ww_2_= ww_2
                # if write_started and ww_2.find("Time Log Total Hours (hr)") != 0 and ww_2.find('{"Vendor":') != 0 and ww_2.find("Head Count") != 0 and not(ww_2 in ["Mud Data"]):
                #         write_into_file(", \n")
                if ww_2 == "OPERATION SUMMARY" and operation_started:
                    ww_2 = ""
                elif ww_2 == "MUD CHECK" and pre_ww_2 == "MUD CHECK":
                    ww_2 = ", \n"
                    print(ww_2)
                elif ww_2 == "BIT DATA":
                    ww_2 = '"}\n], \n"' + ww_2 + '": '
                
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

                if ww_2_ == "OPERATION SUMMARY": 
                    operation_started = True
                

                write_into_file(ww_2)
                write_started = True


            if ss != "": 
                if ii == len(df.index) + add_row_count and ww_2_ == "OPERATION SUMMARY":
                    ss = ss[:-3]
                elif ww_2_ == "MUD CHECK":
                    if pre_ww_2 != "MUD CHECK":
                        ss = "[\n" + ss
                    else:
                        ss += "]"
                write_into_file(ss)

            for jj in range(j, k+1):
                cell[i][jj] = -1
            if ww_2_.find('{"start_time"') != 0 and ww_2_.find('{"Vendor"') != 0 and ww_2_.find('{"Vessel Name"') != 0:
                pre_ww_2 = ww_2_
          
write_into_file("\n")
write_into_file("}")
my_file.close()
