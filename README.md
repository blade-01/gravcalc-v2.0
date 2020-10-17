## Gravcalc-v2.0 

An automated spreadsheet built to calculate gravity readings for geophysical surveys. 

## FEATURES
-   This Gravity spreadsheet is built with Microsoft Excel and automated with Python.
-   It consists of a single sheet which solves for gravity corrections such as Bouguer anomaly, Free Air Correction, Latitude Correction etc

## REQUIREMENTS
To run this program, you need to have the following installed on your local computer:
-   A recent version of [Python](www.python.org)
-   Microsoft Excel

## WORKFLOW

To run this spreadsheet, adhere to the following instructions:
- Switch your directory to the location of this project on your computer 
  
-   run the following comand in your terminal (Powershell, Gitbash, etc)
```
pip install requirements.txt
```
 
-   then run,
```
python sheet1.py
```
-   This activates the excel sheet named 'GravCalc.xlsx'. Type all your measured values gotten from the field under the appropriate headings (**DO NOT EDIT THE HEADINGS**), then save your work by pressing CTRL + S.

-   Open the file 'GravCalc.xlsx' from your file directory (the sheet will be saved to the same location as this script on your local computer)
  
-   To view the final result, run
```
python test2.py
```
-   Open the file 'GravCalcResult.xlsx' from your files directory to view your measured and corrected values.

## Note
-   If after running the 'python sheet1.py' or 'python sheet2.py' command, you notice you made mistakes in your readings, kindly delete the file and run the commands from scratch. 
  
## Remarks
- Feel free to submit pull requests if you notice any bugs or want to improve on this project.