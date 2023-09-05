
import pandas as pd

import dataframe_image as di


df=pd.read_excel("excel/finale.xlsx",sheet_name="AGADIR")

di.export(df,filename="test.jpg")

print(df)