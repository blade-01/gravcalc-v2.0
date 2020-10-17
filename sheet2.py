#the page to display results

from openpyxl import Workbook, load_workbook
from pandas import DataFrame 
from openpyxl.styles import NamedStyle, Font, Color, Border, colors, Alignment, Side
import pandas as pd
import math

#format the spreadsheet
bold_font = Font(bold=True, color=colors.BLUE, size=10)
font_alignment = Alignment(horizontal='center', vertical='center')
font_border = Border(bottom=Side(border_style='thin'))

#load the first worksheet
edited_file = load_workbook(filename="GravCalc.xlsx")
data = edited_file.active
data.title = 'Corrected Readings'
filename = r'GravCalc.xlsx'
df = pd.read_excel(filename)

#get the user's input values from the columns in the spreadsheet as python lists
gravity_station = list(df["GRAVITY STATION"][1:])
time_hr = list(df["TIME(hr)"][1:])
time_min = list(df["TIME(min)"][1:])
elevation = list(df["ELEVATION(m)"][1:])
elevation_cm = list(df["ELEVATION(cm)"][1:])
latitude_degrees = list(df["LATITUDE(degrees)"][1:])
latitude_minutes = list(df["LATITUDE(min)"][1:])
latitude_seconds = list(df["LATITUDE(sec)"][1:])
longitude_degrees = list(df["LONGITUDE(degrees)"][1:])
longitude_minutes = list(df["LONGITUDE(min)"][1:])
longitude_seconds = list(df["LONGITUDE(sec)"][1:])
total_time = list(df["TOTAL TIME"][1:])
latitude_dd = list(df["LATITUDE(DD)"][1:])
longitude_dd = list(df["LONGITUDE(DD)"][1:])
latitude_correction = list(df["LATITUDE CORRECTION"][1:])
gravity_reading_dial = list(df["GRAVITY READING(dial)"][1:])
gravity_reading_mgal = list(df["GRAVITY READING(mgal)"][1:])
observed_gravity = list(df["OBSERVED GRAVITY"][1:])
fac = list(df["FREE AIR CORRECTION"][1:])
faa = list(df["FREE AIR ANOMALY"][1:])
bc = list(df["BOUGUER CORRECTION"][1:])
ba = list(df["BOUGUER ANOMALY"][1:])

#applying formulae to the values gotten from the spreadsheet
#convert the time in hour and minutes to minutes
total_time = [(time_hr[i] * 60) + time_min[i] for i in range(len(time_hr)) and range(len(time_min))]

#check if the value of elevatin was entered in cm or m
if elevation_cm:
    elevation = [(elevation_cm[i] / 100) for i in range(len(elevation_cm))]
else:
    elevation_cm = list(df["ELEVATION(cm)"][1:])

#convert all latitudes readings to latitude(decimal degrees)
latitude_dd = [((latitude_degrees[i]*1)) + ((latitude_minutes[i]/60)) + ((latitude_seconds[i]/3600)) for i in range(len(latitude_degrees)) and range(len(latitude_minutes)) and range(len(latitude_seconds))]

#convert all longitudes readings to longitudes (decimal degrees)
longitude_dd = [((longitude_degrees[i]*1)) + ((longitude_minutes[i]/60)) + ((longitude_seconds[i]/3600)) for i in range(len(longitude_degrees)) and range(len(longitude_minutes)) and range(len(longitude_seconds))]

#solve for latitude correction 
latitude_correction = [978031.85 * (1.0 + (0.005278895 * (math.sin(latitude_dd[i])**2)) + (0.000023462 * (math.sin(latitude_dd[i])**2) * (math.sin(latitude_dd[i])**2))) for i in range(len(latitude_dd))]

#convert readings in dial to mgal
if gravity_reading_dial:
    gravity_reading_mgal = [(gravity_reading_dial[i] * 1.01388) for i in range(len(gravity_reading_dial))]
else:
    gravity_reading_dial = list(df["GRAVITY READING(dial)"][1:])

#solve for observed gravity
observed_gravity = [0]+[(gravity_reading_mgal[i + 1] - gravity_reading_mgal[i]) for i in range(len(gravity_reading_mgal)-1)]

#solve for free air correction
fac = [(elevation[i] * 0.3086) for i in range(len(elevation))]

#solve for free air anomaly
faa = [(observed_gravity[i] - latitude_correction[i] + fac[i]) for i in range(len(observed_gravity)) and range(len(latitude_correction)) and range(len(fac))]

#solve for bouguer correction
bc = [(-0.04193 * 6.67428 * (math.pow(10,-11)) * elevation[i]) for i in range(len(elevation))]

#solve for bouguer anomaly
ba = [observed_gravity[i] - latitude_correction[i] + fac[i] for i in range(len(observed_gravity)) and range(len(latitude_correction)) and range(len(fac))]

#display the results of the above lists as excel columns
df = pd.DataFrame()

#to display final results gotten after performing calculations
df["GRAVITY STATION"] = gravity_station[::]
df["TIME(hr)"] = time_hr[::]
df["TIME(min)"] = time_min[::]
df["ELEVATION(cm)"] = elevation_cm[::]
df["ELEVATION(m)"] = elevation[::]
df["LATITUDE(degrees)"] = latitude_degrees[::]
df["LATITUDE(min)"] = latitude_minutes[::]
df["LATITUDE(sec)"] = latitude_seconds[::]
df["LONGITUDE(degrees)"] = longitude_degrees[::]
df["LONGITUDE(min)"] = longitude_minutes[::]
df["LONGITUDE(sec)"] = longitude_seconds[::]
df["TOTAL TIME(min)"] = total_time[::]
df["LATITUDE(DD)"] = latitude_dd[::]
df["LONGITUDE(DD)"] = longitude_dd[::]
df["LATITUDE CORRECTION"] = latitude_correction[::]
df["GRAVITY READING(dial)"] = gravity_reading_dial[::]
df["GRAVITY READING(mGal)"] = gravity_reading_mgal[::]
df["OBSERVED GRAVITY"] = observed_gravity[::]
df["FREE AIR CORRECTION"] = fac[::]
df["FREE AIR ANOMALY"] = faa[::]
df["BOUGUER CORRECTION"] = bc[::]
df["BOUGUER ANOMALY"] = ba[::]

#then rewrite the values to a new sheet
df.to_excel('GravCalcResult.xlsx', index=False)
