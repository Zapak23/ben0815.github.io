from getactives import create_actives_xl_file
import xlrd
import xlwt
#This is a python script I made because I was bored and sometimes manually rearranging all the code for actives and their headshots
#  can be time consuming. It takes the active member data from HeadshotList.xlsx and prints out the html code in a seperate file 
# for their headshots and names in a html file. 

#Run using 'python3 update.py' on mac/linux, I didnt mess with windows because managing the imports is a pain. 
# !!Must install xlrd using 'pip install xlrd' in unix terminal!! If it throws a NOT SUPPORTED kind of error than
#make sure that you have the right version with 'pip install xlrd==1.2.0' and try again

# ADAM'S EDITS: So hopfully ive made a script that will make the html easier. Since it takes time to edit the excel sheet with the 
# actives and the all the photos, I figured it would be easiest to make a script to make the excel. The OT-Scribe has a sheet that has a 
# list of all the actives. Just download the list and run the 'update.py' file. Below are some variables to help with the updates process. 
# Just update the variable to update the entire the list. The following should include:
#
# imagepath: the path to all the images your looking to update
# headshot_file: a string variable that holds the file name from the 'create_actives_xl_file'. 
# file_of_active_members: just to make the process a little more intuitive, there is an input variable to type the filename for active members.
# 
# No one ever said I was creative with variable/function names, but hopefully this makes it easier.
# Make sure the image .JPGs are FirstnameLastname


file_of_active_memebers = "spring2023actives.xlsx"

imagepath = "../images/Headshots/oldHeadshots/" # Change this variable to change the path where the images are located
# ^ with respect to where the 'active.html' file is. 

headshots_file = create_actives_xl_file(file_of_active_memebers)

print(headshots_file)

wb = xlrd.open_workbook(headshots_file)
sheet = wb.sheet_by_index(0)

def populateEC(sheet, EC):
    for active in range(1,sheet.nrows):
        if sheet.cell_value(active, 2) in EC:
            EC[sheet.cell_value(active,2)] = active
            print(f"{sheet.cell_value(active, 0)} : {sheet.cell_value(active, 1)} : {sheet.cell_value(active, 2)} < from row || location >  {EC[sheet.cell_value(active,2)]}")
            #print(EC[sheet.cell_value(active,2)])


file = open("activehtmlcode_new.html", "w+")

file.write("<div id=\"gallery\">\n  <figure>\n\n    <h5 class=\"center\" style=\"font-size:200%;\">Executive Committee</h5>\n    <hr><br>\n    <ul class=\"nospace clear\">\n\n") #write EC header

EC = {          #Dictionary of EC member entry index, keys line up with the key on excel sheet
    "R":"",
    "VR":"",
    "T":"",
    "M":"",
    "S":"",
    "CS":"",
    "PM":""
} # EBoard Pos : Row
ECListOrder = ["R","VR","T","M","S","CS","PM"]  #dictionary doesnt stay in order so theres the order
populateEC(sheet, EC)

for i in range(0,4):    #write the images of EC members
    section = "one_quarter"
    if i == 0:
        section+=" first" 
    #print(f"imagepath + sheet.cell_value(EC[ECListOrder[i]],1):{imagepath}{1} what is i:{i}->{type(EC[ECListOrder[i]])}")
    file.write( "<li class=\""   +section+   "\"><img src=\"" + imagepath + sheet.cell_value(int(EC[ECListOrder[i]]),1)  +  "\"></li>\n")

file.write("\n\n")
for i in range(0,4):    #write the Names of EC members
    section = "one_quarter"
    if i == 0:
        section+=" first"
    file.write("<li style=\"text-align:center;\" class=\""+section+"\">\n      <h2>"+sheet.cell_value(EC[ECListOrder[i]],0)+"</h2>\n</li>\n")

#I got lazy so this just prints the titles statically
file.write("\n\n<li style=\"text-align:center;\" class=\"one_quarter first\">\n  <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Regent</h4>\n</li>\n<li style=\"text-align:center;\" class=\"one_quarter\">\n  <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Vice Regent</h4>\n</li>\n<li style=\"text-align:center;\" class=\"one_quarter\">\n  <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Treasurer</h4>\n</li>\n<li style=\"text-align:center;\" class=\"one_quarter\">\n  <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Marshal</h4>\n</li><br>\n\n")

for i in range(4,7):    #write the images of EC members
    section = "one_third"
    if i == 4:
        section+=" first"
    file.write( "<li class=\""   +section+   "\"><img src=\"" + imagepath + sheet.cell_value(EC[ECListOrder[i]],1)  +  "\"></li>\n")

file.write("\n\n")
for i in range(4,7):    #write the Names of EC members
    section = "one_third"
    if i == 4:
        section+=" first"
    file.write("<li style=\"text-align:center;\" class=\""   +section+   "\">\n\t<h2>"   +sheet.cell_value(EC[ECListOrder[i]],0)+   "</h2>\n</li>\n")

file.write("\n<li style=\"text-align:center;\" class=\"one_third first\">\n  <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Scribe</h4>\n</li>\n<li style=\"text-align:center;\" class=\"one_third\">\n  <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Corresponding Secretary</h4>\n</li>\n<li style=\"text-align:center;\" class=\"one_third\">\n  <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Pledge Marshal</h4>\n</li>\n\n")

file.write("<!-- ################################################### -->\n")
file.write("<!-- ################################################### -->\n")
file.write("<!-- ################################################### -->\n\n\n")



file.write("</ul>\n<br><br>\n<h5 class=\"center\" style=\"font-size:200%;\">Active Members</h5>\n<hr><br>\n<ul class=\"nospace clear\">\n") #write active member header
# FullName, imageURL, EC?
img_index = 1   #used to keep track of where youre at in the total collection
name_index = 1

img_count = 1   #used to keep track of where youre at in groups of 4, reset at the end of outer while loop
name_count = 1
print(f"img_index before loop: {img_index}")
while(img_index < sheet.nrows):
    #print(f"img_index right before the 2nd loop: {img_index}")
    while img_count<=4:
        #print(f"img_index right before the if: {img_index}")
        if img_index == sheet.nrows:
            #print(f"{img_index} || {sheet.nrows}")
            break
        #print(f"{img_count} || {sheet.cell_value(name_index,2)} == \"False\" || {sheet.name}")
        if (sheet.cell_value(img_index,2) == "False") | (sheet.cell_value(img_index,2) == 0):
            section = "one_quarter"
            if img_count == 1:
                #print("here")
                section +=" first"
            #print(f"{sheet.cell_value(img_index,2) == False} or {sheet.cell_value(img_index,2) == 0} results in")
            active_img_code = "<li class=\""  +section+  "\"> <img src=\""  + imagepath+sheet.cell_value(img_index, 1) + "\"></li>\n" 

            img_count+=1
            file.write(active_img_code)
        #print(f"{img_index}")
        img_index+=1
    file.write("\n")
    
    while name_count<=4:
        if name_index == sheet.nrows:
            #print(f"{name_index} || {sheet.nrows}")
            break
        #print(f"{sheet.cell_value(name_index,2)} == \"False\" but like this is the second time around")
        if (sheet.cell_value(name_index,2) == "False") | (sheet.cell_value(name_index,2) == 0):
            #print(f"{sheet.cell_value(name_index,2) == False} or {sheet.cell_value(name_index,2) == 0}")

            section = "one_quarter"
            if name_count == 1:
                section+=" first"
            active_name_code = "<li style=\"text-align:center;\" class=\""+section+"\">\n\t<h2 style=\"margin-bottom:20px;margin-top:-5%;\">"+sheet.cell_value(name_index,0)+ "</h2>\n</li>\n\n"
            name_count+=1

            file.write(active_name_code)
        name_index+=1
    img_count = 1
    name_count=1

file.write("\n<!-- Writing complete -->\n<!-- ################################################### -->\n\n")
