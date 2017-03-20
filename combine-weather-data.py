'''
    File name: combine-weather-data.py
    Author: Patrick Puecher
    Python Version: 2.7
'''

import pandas as pd
import glob

outputfile = "Klima_LT_N_daily_combine.csv"
open(outputfile, 'w').close()

df = pd.DataFrame(data=[], columns=["date", "rainfall", "max", "min", "station", "number", "xutm", "yutm", "alt"])
df.to_csv(outputfile, mode="a", header=True, index=False)

for file in glob.glob("*.xls"):
    df = pd.read_excel(file, skiprows=4, parse_cols="C", header=None)
    station = df.loc[0].values
    number = df.loc[1].values
    xutm = df.loc[2].values
    yutm = df.loc[3].values
    alt = df.loc[4].values

    df = pd.read_excel(file, skiprows=13, parse_cols="B:E", header=None)

    s = pd.Series(station)
    df[4] = s[0].encode('utf-8')

    s = pd.Series(number)
    df[5] = s[0]

    s = pd.Series(xutm)
    df[6] = s[0]

    s = pd.Series(yutm)
    df[7] = s[0]

    s = pd.Series(alt)
    df[8] = s[0]

    df[1] = [str(x).strip().replace(',', '.').replace('---', 'null') for x in df[1]]
    df[2] = [str(x).strip().replace(',', '.').replace('---', 'null') for x in df[2]]
    df[3] = [str(x).strip().replace(',', '.').replace('---', 'null') for x in df[3]]

    df.to_csv(outputfile, mode="a", header=False, index=False)
