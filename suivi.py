import json

import pandas as pd
from excel_fonctions import Excel
import dataframe_image as dfi


class Suivi:

    def __init__(self, vendeur_name, total_to_work, days_work):
        self.vendeur_name = vendeur_name
        self.total_to_work = total_to_work
        self.day_work = days_work

        try:
            with open("days.json", "r") as jsonFile:
                global json_data
                json_data = json.load(jsonFile)
        except Exception as e:
            print("An error occurred:", e)

        df = pd.read_excel("excel/finale.xlsx", sheet_name="AGADIR")

        list_of_string_H: list = df[df['H'].apply(lambda x: isinstance(x, str))]

        for i in list_of_string_H.index:
            df.loc[i, "H"] = 0

        df = df.astype(
            {
                "REAL": "int",
                "OBJ": "int",
                "EnCours": "int",
                "Real 2023": "int",
                "Historique 2022": "int",

            }
        )

        df.loc[:, "Percent"] = df["Percent"].map("{:.1%}".format)

        df.loc[:, "H"] = df["H"].map('{:.1%}'.format)

        df_quali = pd.read_excel("excel/finale.xlsx", sheet_name="QUALI NV")
        # df_quali = df_quali.replace("%", 0)
        df_quali = df_quali.fillna(1)
        df_quali = df_quali.astype(
            {
                "ACM": "float",
                "TSM": "float"

            }
        )
        df_quali.loc[:, "ACM"] = df_quali.ACM.map("{:.2%}".format)
        df_quali.loc[:, "LINE"] = df_quali.LINE.map("{:.2%}".format)
        df_quali["RAF"] = round((df_quali['CLT PROGRAMME'] - (df_quali['CLT PROGRAMME'] * df_quali['TSM'])) / 3)
        df_quali.loc[:, "TSM"] = df_quali.TSM.map("{:.2%}".format)

        df["OBJ ttc"] = df.OBJ.apply(lambda x: (x * (json_data["from_file"]["t"] /json_data["from_file"]["d"])) * 1.2)



        raf: int = df["OBJ ttc"] - ((df["REAL"] * 1.2) + (df["EnCours"] * 1.2))
        df["RAF"] = round(raf / (int(total_to_work) - int(days_work)))

        global df_by_vendeur_quantitatif
        global df_by_vendeur_qualitatif
        global df_obj_vendeur
        df_by_vendeur_quantitatif = df.query("Vendeur==@vendeur_name")
        df_by_vendeur_qualitatif = df_quali.query("Vendeur==@vendeur_name")


    def send_quantitatif(self):
        dfi.export(df_by_vendeur_quantitatif, f"excel/{self.vendeur_name}.jpeg")
        print(df_by_vendeur_quantitatif)

    def send_qualitatif(self):
        dfi.export(df_by_vendeur_qualitatif, f"excel/{self.vendeur_name}.jpeg")
        print(df_by_vendeur_qualitatif)

