## Gravcalc-v2.0 

An automated spreadsheet built to calculate gravity readings for geophysical surveys. 

## FEATURES
-   This Gravity spreadsheet is built with Microoft Excel and automated with Python.
-   It consists of a single sheet which solves for gravity corrections such as Bouguer anomaly, Free Air Corrections etc

## WORKFLOW

To run this spreadsheet, adhere to the following instructions
-   run 
```
pip install requirements.txt
```
-   The python files are in the 'app' folder 

-   In your terminal 'cd to the app folder',
![image](https://user-images.githubusercontent.com/49791498/94964773-3142f300-04f2-11eb-8e2a-f7a2dbc3d909.png) 
*git bash*
 
-   then run,
```
python test1.py
```
-   This activates the excel sheet named 'GravCalc.xlsx'. Type all your measured values gotten from the field (**DO NOT EDIT THE HEADINGS**).

-   Open the file 'GravCalc.xlsx' from your files directory
  
-   To view the final result, run
```
python test2.py
```
-   Open the file 'GravCalcResult.xlsx' from your files directory to view your corrected values.