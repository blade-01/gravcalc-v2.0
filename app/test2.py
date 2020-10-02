#the page to sidaplay results

from openpyxl import Workbook, load_workbook

from pandas import DataFrame 
import pandas as pd

#load the first worksheet
edited_file = load_workbook(filename="GravCalc.xlsx")
data = edited_file.active

filename = r'GravCalc.xlsx'
df = pd.read_excel(filename)
gravity_station = []

#get the user's input values from the columns in the spreadsheet as python lists
gravity_station = list(df['GRAVITY STATION'][1:])
date = list(df['DATE (mm/dd/yyyy)'][1:])
time = list(df['TIME (HOURS)'][1:])
min_time = list(df['TIME(MINUTES)'][1:])
duration = list(df['DURATION(h)'][1:])
latitude_degrees = list(df["LATITUDE (degrees)"][1:])
latitude_minutes = list(df["LATITUDE(minutes)"][1:])
latitude_seconds = list(df["LATITUDE(seconds)"][1:])
longitude_degrees = list(df["LONGITUDE(degrees)"][1:])
longitude_minutes = list(df["LONGITUDE(minutes)"][1:])
longitude_seconds = list(df["LONGITUDE(seconds)"][1:])
corrected_latitude = list(df["LATITUDE (DD)"][1:])
corrected_longitude = list(df["LONGITUDE (DD)"][1:])
elipsoid = list(df["ELLIPSOID HEIGHT (m)"][1:])
lacoste_romberg = list(df["LACOSTE-ROMBERG METER"][1:])
counter_and_calibrated = list(df["COUNTER AND CALIBRATED (mGal)"][1:])
measured_lacoste_romberg = list(df[ "LACOSTE-ROMBERG, MEASURED (mGal)"][1:])
worden = list(df["WORDEN METER, COUNTER (DIAL READING)"][1:])
calibrated = list(df["CALIBRATED (mGal)"][1:])
scintrex_meter = list(df["SCINTREX METER, DRIFT AND TIDE CORRECTED (mGal)"][1:])
raw_observed_gravity = list(df["RAW OBSERVED GRAVITY (mGal)"][1:])
tide = list(df["TIDE (mGal)"][1:])
corrected_tide = list(df["TIDE CORRECTED (mGal)"][1:])
meter_drift = list(df["METER DRIFT (mGal)"][1:])
drift_tide_corrected = list(df["DRIFT AND TIDE CORRECTED (mGal)"][1:])
dc_shift = list(df["DC SHITF (mGal)"][1:])
theoretical_gravity = list(df["THEORETICAL GRAVITY (mGal)"][1:])
height_corrected = list(df["HEIGHT CORRECTED (mGal)"][1:])
atm_effect = list(df["ATM EFFECT (mGal)"][1:])
bouger_spherical_cap = list(df["BOUGUER SPHERICAL CAP (mGal)"][1:])
terrain_correction = list(df["TERRAIN CORRECTION (mGal)"][1:])
observed_abs_gravity = list(df["OBSERVED ABSOLUTE GRAVITY (mGal)"][1:])
complete_bouger_anomaly = list(df["COMPLETE BOUGUER ANOMALY (mGal)"][1:])

#display the results of the above lists as excel columns
df = pd.DataFrame()
#to display final results gotten after performing calculations
df['GRAVITY STATION'] = gravity_station[::]
df['DATE (mm/dd/yyyy)'] = date[::]
df['TIME (HOURS)'] = time[::]
df['TIME(MINUTES)'] = min_time[::]
df['DURATION(h)'] = duration[::]
df["LATITUDE (degrees)"] = latitude_degrees[::]
df["LATITUDE(minutes)"] = latitude_minutes[::]
df["LATITUDE(seconds)"] = latitude_seconds[::]
df["LONGITUDE(degrees)"] = longitude_degrees[::]
df["LONGITUDE(minutes)"] = longitude_minutes[::]
df["LONGITUDE(seconds)"] = longitude_seconds[::]
df["LATITUDE (DD)"] = corrected_latitude[::]
df["LONGITUDE (DD)"] = corrected_longitude[::]
df["ELLIPSOID HEIGHT (m)"] = elipsoid[::]
df["LACOSTE-ROMBERG METER"] = lacoste_romberg[::]
df["COUNTER AND CALIBRATED (mGal)"] = counter_and_calibrated[::]
df[ "LACOSTE-ROMBERG, MEASURED (mGal)"] = measured_lacoste_romberg[::]
df["WORDEN METER, COUNTER (DIAL READING)"] = worden[::]
df["CALIBRATED (mGal)"] = calibrated[::]
df["SCINTREX METER, DRIFT AND TIDE CORRECTED (mGal)"] = scintrex_meter[::]
df["RAW OBSERVED GRAVITY (mGal)"] = raw_observed_gravity[::]
df["TIDE (mGal)"] = tide[::]
df["TIDE CORRECTED (mGal)"] = corrected_tide[::]
df["METER DRIFT (mGal)"] = meter_drift[::]
df["DRIFT AND TIDE CORRECTED (mGal)"] = drift_tide_corrected[::]
df["DC SHITF (mGal)"] = dc_shift[::]
df["THEORETICAL GRAVITY (mGal)"] = theoretical_gravity[::]
df["HEIGHT CORRECTED (mGal)"] = height_corrected[::]
df["ATM EFFECT (mGal)"] = atm_effect[::]
df["BOUGUER SPHERICAL CAP (mGal)"] = bouger_spherical_cap[::]
df["TERRAIN CORRECTION (mGal)"] = terrain_correction[::]
df["OBSERVED ABSOLUTE GRAVITY (mGal)"] = observed_abs_gravity[::]
df["COMPLETE BOUGUER ANOMALY (mGal)"] = complete_bouger_anomaly[::]

#then rewrite the values to a new sheet
df.to_excel('GravCalcResult.xlsx', index=False)


#edited_file.save(filename="GravCalcResult.xlsx")