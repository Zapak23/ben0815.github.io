
Congrats, its your turn to update the photos with new members and without alumni

this process can be kind of a pain with formatting everything so that it all fits right and looks good so I wrote a script that goes throught an excel list of active members and generates the right html code to be copy/pasted into the actives.html page. The generated code is simply written in the html file in this folder and includes all members, including EC. You also have to update the spreadsheet with lists of new people but all in all that takes a lot less time than formatting everything manually. You can use it or not, maybe make it better, up to you, but it can make things go a lot faster.

The only thing that you need to have installed to get this to run is the xlrd library, installed using "pip install xlrd" on mac/linux. To run it, navigate to this directory and type "python3 update.py" to run it and generate the file. Always check it locally before you push it to git.


Have fun,
Austin P
