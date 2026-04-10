import pandas as pd
from  logic  import is_data_clean


df = pd.read_excel("source.xlsx")
print(df.columns)

if is_data_clean(df):
    print("hello world")