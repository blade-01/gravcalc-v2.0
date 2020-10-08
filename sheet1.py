#this file appends out titles to the spreadsheet

from openpyxl import Workbook, load_workbook
from pandas import DataFrame 
import pandas as pd

#activate a spreadsheet
gravcalc = Workbook()
sheet_one = gravcalc.active
gravcalc.save(filename= "GravCalc.xlsx")

#append values
edited_file = load_workbook(filename="GravCalc.xlsx")
data = edited_file.active
data.auto_filter.ref = "A1:O100"

#append the cell titles to the loaded spread sheet
data["A1"] = "ROCK DENSITY"
data["A2"] = "(kg/m^3)"
data["B1"] = "GRAVITY STATION"
data["C1"] = "TIME (hr)"
data["C2"] = "(hr)"
data["D1"] = "TIME(min)"
data["D2"] = "(min)" 
data["E1"] = "ELEVATION(cm)"
data["E2"] = "(cm)"
data["F1"] = "ELEVATION(m)"
data["F2"] = "(m)"
data["G1"] = "LATITUDE (degrees)"
data["G2"] = "Degrees"
data["H1"] = "LATITUDE(min)"
data["H2"] = "Minutes"
data["I1"] = "LATITUDE(sec)"
data["I2"] = "Seconds"
data["J1"] = "LONGITUDE(degrees)"
data["J2"] = "Degrees"
data["K1"] = "LONGITUDE(min)"
data["K2"] = "Minutes"
data["L1"] = "LONGITUDE(sec)"
data["L2"] = "Seconds"
data["M1"] = "TOTAL TIME(min)"
data["M2"] = "(min)"
data["N1"] = "LATITUDE(DECIMAL DEGREES)"
data["N2"] = "(DECIMAL DEGREES)"
data["O1"] = "LONGITUDE (DECIMAL DEGREES)"
data["O2"] = "(DECIMAL DEGREES)"
data["P1"] = "LATITUDE CORRECTION"
data["Q1"] = "GRAVITY READING (dial)"
data["Q2"] = "(dial)"
data["R1"] = "GRAVITY READING (mGal)"
data["R2"] = "(mGal)"
data["S1"] = "OBSERVED GRAVITY (mGal)"
data["S2"] = "(mGal)"
data["T1"] = "FREE AIR CORRECTION (mGal)"
data["T2"] = "(mGal)"
data["U1"] = "FREE AIR ANOMALY (mGal)"
data["U2"] = "(mGal)"
data["V1"] = "BOUGUER CORRECTION (mGal)"
data["V2"] = "(mGal)"
data["W1"] = "BOUGUER ANOMALY(mGal)"
data["W2"] = "(mGal)"


#save the edited spreadsheet
edited_file.save(filename="GravCalc.xlsx")