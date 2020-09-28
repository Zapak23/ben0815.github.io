import xlrd

#This is a python script I made because I was bored and sometimes manually rearranging all the code for actives and their headshots can be time consuming. It takes the active member data from HeadshotList.xlsx and prints out the html code in a seperate file for their headshots and names in a txt file. It doesn't generate code for EC at the moment.

#Run using 'python3 update.py' on mac/linux, I didnt mess with windows because managing the imports is a pain. Must install xlrd using 'pip install xlrd'

wb = xlrd.open_workbook('HeadshotList.xlsx')
sheet= wb.sheet_by_index(0)

def populateEC(sheet, EC):
    for active in range(1,sheet.nrows):
        if sheet.cell_value(active, 2) in EC:
            EC[sheet.cell_value(active,2)] = active


file = open("activehtmlcode.html", "w+")

file.write("<div id=\"gallery\">\n  <figure>\n\n    <h5 class=\"center\" style=\"font-size:200%;\">Executive Committee</h5>\n    <hr><br>\n    <ul class=\"nospace clear\">\n\n") #write EC header

EC = {          #Dictionary of EC member entry index, keys line up with the key on excel sheet
    "R":"",
    "VR":"",
    "T":"",
    "M":"",
    "S":"",
    "CS":"",
    "PM":""
}
ECListOrder = ["R","VR","T","M","S","CS","PM"]  #dictionary doesnt stay in order so theres the order
populateEC(sheet, EC)

for i in range(0,4):    #write the images of EC members
    section = "one_quarter"
    if i == 0:
        section+=" first"
    file.write("<li class=\""+section+"\"><img src=\"../images/Headshots/"+ sheet.cell_value(EC[ECListOrder[i]],1)+"\"></li>\n")

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
    file.write("<li class=\""+section+"\"><img src=\"../images/Headshots/"+ sheet.cell_value(EC[ECListOrder[i]],1)+"\"></li>\n")

file.write("\n\n")
for i in range(4,7):    #write the Names of EC members
    section = "one_third"
    if i == 4:
        section+=" first"
    file.write("<li style=\"text-align:center;\" class=\""+section+"\">\n      <h2>"+sheet.cell_value(EC[ECListOrder[i]],0)+"</h2>\n</li>\n")

file.write("<li style=\"text-align:center;\" class=\"one_third first\">\n   <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Scribe</h4>\n </li>\n <li style=\"text-align:center;\" class=\"one_third\"><h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Corresponding Secretary</h4></li>\n <li style=\"text-align:center;\" class=\"one_third\">\n   <h4 style=\"margin-top:-15px;font-size:120%;color:#601313;\">Pledge Marshal</h4>\n </li>\n\n")

file.write("<!-- ################################################### -->\n")
file.write("<!-- ################################################### -->\n")
file.write("<!-- ################################################### -->\n\n\n")



file.write("</ul>\n<br><br>\n<h5 class=\"center\" style=\"font-size:200%;\">Active Members</h5>\n<hr><br>\n<ul class=\"nospace clear\">\n") #write active member header
#FullName, imageURL, EC?
img_index = 1   #used to keep track of where youre at in the total collection
name_index = 1

img_count = 1   #used to keep track of where youre at in groups of 4, reset at the end of outer while loop
name_count = 1
while img_index < sheet.nrows:
    while img_count<=4:
        if img_index == sheet.nrows:
            break;
        if (sheet.cell_value(img_index,2) == False) | (sheet.cell_value(img_index,2) == 0):
            section = "one_quarter"
            if img_count == 1:
                section+=" first"

            active_img_code = "<li class=\""+section+"\"> <img src=\"../images/Headshots/" + sheet.cell_value(img_index, 1) + "\" alt=\"\"></li>\n"

            img_count+=1

            file.write(active_img_code)
        img_index+=1

    file.write("\n")
    while name_count<=4:
        if name_index == sheet.nrows:
            break
        if (sheet.cell_value(name_index,2) == False) | (sheet.cell_value(name_index,2) == 0):
            section = "one_quarter"
            if name_count == 1:
                section+=" first"
            active_name_code = "<li style=\"text-align:center;\" class=\""+section+"\">\n\t<h2 style=\"margin-bottom:20px;margin-top:-5%;\">"+sheet.cell_value(name_index,0)+ "</h2>\n</li>\n"

            name_count+=1
            file.write(active_name_code)
        name_index+=1

    img_count = 1
    name_count=1
    file.write("<!-- ################################################### -->\n\n")
