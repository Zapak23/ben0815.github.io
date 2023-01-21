import xlrd
import xlwt

def create_actives_xl_file(actives_file):     
    all_members = xlrd.open_workbook(actives_file)
    headshots_wb = xlwt.Workbook()
    headshots_wb_file = "spring2023dummy.xlsx"
    headshots_ws = headshots_wb.add_sheet("spring2023dummy.xlsx")

    # Go throught the current memebers and if they are active, add them to a data structure to sort and write to a new excel

    active_members_by_rows = []

    actives_sheet = all_members.sheet_by_name("Actives") # grabs the current actives sheet
    for member in reversed(range(0, actives_sheet.nrows)):
        #print(actives_sheet.cell_value(member, 1)) # Prints out the active members in the actives sheet
        active_members_by_rows.append(actives_sheet.row_values(member))
        
    print("\n")
    
    
    headshots_ws.write(0, 0, "Full Name")
    headshots_ws.write(0, 1, "Pic Name")
    headshots_ws.write(0, 2, "On EC?") # Write out the headers

    for row in range(len(active_members_by_rows) - 1):
        #print(active_members_by_rows[row]) # Prints out the row associates with an active in the active_members_by_rows arr
        is_eboard = False
        active_title = ""
        active_name = active_members_by_rows[row][1] # should be a string variable
        if active_members_by_rows[row][7] != "":
            is_eboard = True
            active_title = active_members_by_rows[row][7]
        
        #print(active_name)
        headshots_ws.write(row + 1, 0, active_name) # writes the name of the active in first cell
        headshots_ws.write(row + 1, 1, active_name.replace(" ", "") + '.JPG') # writes the picture of the active here
        if is_eboard:
            headshots_ws.write(row + 1, 2, active_title) # writes whether the active is on EBoard
        else:
            headshots_ws.write(row + 1, 2, "False")
        #print(f"{active_name}, {active_title}, {row}")
    headshots_wb.save(headshots_wb_file)
    return headshots_wb_file