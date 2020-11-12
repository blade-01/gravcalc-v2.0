#the page to display results

from openpyxl import Workbook, load_workbook
from pandas import DataFrame 
import pandas as pd
import math

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
elevation_cm = list(df["ELEVATION(cm)"][1:])
elevation = list(df["ELEVATION(m)"][1:])
latitude_degrees = list(df["LATITUDE(degrees)"][1:])
latitude_minutes = list(df["LATITUDE(min)"][1:])
latitude_seconds = list(df["LATITUDE(sec)"][1:])
longitude_degrees = list(df["LONGITUDE(degrees)"][1:])
longitude_minutes = list(df["LONGITUDE(min)"][1:])
longitude_seconds = list(df["LONGITUDE(sec)"][1:])
observed_gravity_dial = list(df["OBSERVED GRAVITY"][1:])
observed_gravity_mgal = list(df["OBSERVED GRAVITY(mgal)"][1:])
change_in_observed_gravity = list(df["CHANGE IN OBSERVED GRAVITY(mgal)"][1:])
total_time = list(df["TOTAL TIME"][1:])
change_in_time = list(df["CHANGE IN TIME(min)"][1:])
change_in_elevation = list(df["CHANGE IN ELEVATION(m)"][1:])
latitude_dd = list(df["LATITUDE(DD)"][1:])
longitude_dd = list(df["LONGITUDE(DD)"][1:])
latitude_correction = list(df["LATITUDE CORRECTION"][1:])
drift_correction = list(df["DRIFT CORRECTION"][1:])
drift_rate = list(df["DRIFT RATE"][1:])
corrected_gravity = list(df["CORRECTED GRAVITY"][1:])
fac = list(df["FREE AIR CORRECTION"][1:])
faa = list(df["FREE AIR ANOMALY"][1:])
bc = list(df["BOUGUER CORRECTION"][1:])
ba = list(df["BOUGUER ANOMALY"][1:])

#applying formulae to the values gotten from the spreadsheet

#convert the time in hour and minutes to minutes
total_time = [(time_hr[i] * 60) + time_min[i] for i in range(len(time_hr)) and range(len(time_min))]

#change in time
time_change = [0]+ [total_time[i+1] - total_time[i] for i in range(len(total_time)-1)]

#check if the value of elevation was entered in cm or m
if elevation_cm:
    elevation = [(elevation_cm[i] / 100) for i in range(len(elevation_cm))]
else:
    elevation_cm = list(df["ELEVATION(cm)"][1:])

#change in elevation
change_in_elevation = [0] + [(elevation[i+1] - elevation[i]) for i in range(len(elevation)-1)]

#drift rate
drift_rate = [(elevation[3] - elevation[4]) / (time_change[3] - time_change[4]) for i in range(len(elevation)) and range(len(time_change))]

#drift correction
drift_correction =  [(time_change[i] * drift_rate) for i in range(len(time_change)) and drift_rate]

#convert all latitudes readings to latitude(decimal degrees)
latitude_dd = [((latitude_degrees[i]*1)) + ((latitude_minutes[i]/60)) + ((latitude_seconds[i]/3600)) for i in range(len(latitude_degrees)) and range(len(latitude_minutes)) and range(len(latitude_seconds))]

#convert all longitudes readings to longitudes (decimal degrees)
longitude_dd = [((longitude_degrees[i]*1)) + ((longitude_minutes[i]/60)) + ((longitude_seconds[i]/3600)) for i in range(len(longitude_degrees)) and range(len(longitude_minutes)) and range(len(longitude_seconds))]

#solve for latitude correction 
latitude_correction = [978031.85 * (1.0 + (0.005278895 * (math.sin(latitude_dd[i])**2)) + (0.000023462 * (math.sin(latitude_dd[i])**2) * (math.sin(latitude_dd[i])**2))) for i in range(len(latitude_dd))]

#convert readings in dial to mgal
if observed_gravity_dial:
    observed_gravity_mgal = [(observed_gravity_dial[i] * 1.01388) for i in range(len(observed_gravity_dial))]
else:
    observed_gravity_dial = list(df["OBSERVED GRAVITY"][1:])

#change in observed gravity
change_in_observed_gravity = [0] + [(observed_gravity_mgal[i+1] - observed_gravity_mgal[i]) for i in range(len(observed_gravity_mgal)-1)]

#solve for corrected gravity
corrected_gravity = [0]+[(observed_gravity_mgal[i] - drift_correction) for i in range(len(observed_gravity_mgal)-1) and drift_correction]

#solve for free air correction
fac = [(elevation[i] * 0.3086) for i in range(len(elevation))]

#solve for free air anomaly
faa = [(corrected_gravity[i] - latitude_correction[i] + fac[i]) for i in range(len(corrected_gravity)) and range(len(latitude_correction)) and range(len(fac))]

#solve for bouguer correction
bc = [(-0.04193 * 6.67428 * (math.pow(10,-11)) * elevation[i]) for i in range(len(elevation))]

#solve for bouguer anomaly
ba = [corrected_gravity[i] - latitude_correction[i] + fac[i] for i in range(len(corrected_gravity)) and range(len(latitude_correction)) and range(len(fac))]

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
df["OBSERVED GRAVITY"] = observed_gravity_dial[::]
df["OBSERVED GRAVITY(mGal)"] = observed_gravity_mgal[::]
df["CHANGE IN OBSERVED GRAVITY(mgal)"] = change_in_observed_gravity[::]
df["TOTAL TIME(min)"] = total_time[::]
df["CHANGE IN TIME(min)"] = time_change[::]
df["CHANGE IN ELEVATION(m)"] =  change_in_elevation[::]
df["LATITUDE(DD)"] = latitude_dd[::]
df["LONGITUDE(DD)"] = longitude_dd[::]
df["LATITUDE CORRECTION"] = latitude_correction[::]
df["DRIFT CORRECTION"] = drift_correction[::]
df["DRIFT RATE"] = drift_rate[::]
df["CORRECTED GRAVITY"] = corrected_gravity[::]
df["FREE AIR CORRECTION"] = fac[::]
df["FREE AIR ANOMALY"] = faa[::]
df["BOUGUER CORRECTION"] = bc[::]
df["BOUGUER ANOMALY"] = ba[::]

#then rewrite the values to a new sheet
df.to_excel('GravCalcResult.xlsx', index=False)
