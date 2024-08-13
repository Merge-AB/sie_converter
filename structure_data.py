import pandas as pd
import numpy as np

path = "Output-filer/TRANSACTIONS.csv"

df = pd.read_csv(path, delimiter=",", encoding="IBM437")

"""
length = len(df)

half_lenght = length // 4

df_1 = df[:half_lenght]

df_2 = df[half_lenght:]

#print(len(df[:half_lenght]) + len(df[half_lenght:]))
max_date = str(max(df_1["ver_datum"]))
min_date = str(min(df_1["ver_datum"]))
name = f"TRANSACTIONS_{max_date[0:4]}_{min_date[4:6]}_TO_{max_date[4:6]}.xlsx"
"""

df_split = np.array_split(df,4)

def save_output(df):
    max_date = str(max(df["ver_datum"]))
    min_date = str(min(df["ver_datum"]))
    name = f"TRANSACTIONS_{min_date}{max_date}.csv"
    df.to_csv(name)

save_output(df_split[0])
save_output(df_split[1])
save_output(df_split[2])
save_output(df_split[3])