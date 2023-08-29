from openpyxl import load_workbook
from openpyxl.styles import Font, Fill, PatternFill, GradientFill


class Excel:

    def __init__(self, path: str):
        self.__day_work = 24
        self.path = path

    def get_day_work(self) -> tuple:
        """
        return tuple as total days of month and day works
        """
        wb = load_workbook(self.path)
        sheet_ranges = wb["AGADIR"]
        # wb.active
        day_work = sheet_ranges["C6"].value
        day_work = day_work.split(" ", 2)
        a: str = day_work[0].replace("/", "")
        b: str = day_work[1]
        print(f"day work is : {int(a), int(b)}")

        return int(a), int(b.strip())

    def __fix_quantitatif(self):
        wb = load_workbook(self.path)
        sheet_ranges = wb["AGADIR"]

        try:

            sheet_ranges.unmerge_cells("A8:A9")
            sheet_ranges.unmerge_cells("B8:B9")
            sheet_ranges.unmerge_cells("D8:D9")
            sheet_ranges.unmerge_cells("F8:J8")
            sheet_ranges.unmerge_cells("K8:O8")
            sheet_ranges.delete_cols(1, 2)
            sheet_ranges.delete_cols(3, 1)
            sheet_ranges.delete_cols(6, 2)
            sheet_ranges.delete_cols(7, 2)
            sheet_ranges.delete_cols(9, 1)
            sheet_ranges.delete_cols(10, 1)

            sheet_ranges.delete_rows(1, 8)
            sheet_ranges.delete_rows(2, 32)
            sheet_ranges['A1'] = "Vendeur"
            sheet_ranges['B1'] = "Famille"
            sheet_ranges['E1'] = "Percent"
            sheet_ranges['F1'] = "Real 2023"
            sheet_ranges['G1'] = "Historique 2022"
            sheet_ranges['H1'] = "H"
            print("max", sheet_ranges.max_row)
            for i in range(1, sheet_ranges.max_row + 1):
                if sheet_ranges[f'D{i}'].value <= 1.0 or sheet_ranges[f'D{i}'].value is None:
                    sheet_ranges[f'E{i}'].value = 0.0
                    sheet_ranges[f'D{i}'].value = 0.0

                if sheet_ranges[f'G{i}'].value <= 1.0 or sheet_ranges[f'D{i}'].value is None:
                    sheet_ranges[f'H{i}'].value = 0.0

            wb.save("excel/finale.xlsx")
            print("quantitatif saved")

        except Exception as e:
            print(e)

    def __fix_qualitatif(self):
        wb = load_workbook("excel/finale.xlsx")
        sheet_ranges = wb["QUALI NV"]
        try:
            sheet_ranges.unmerge_cells("E1:K2")
            sheet_ranges.delete_rows(1, 7)
            sheet_ranges.delete_rows(2, 4)
            sheet_ranges.delete_cols(1, 3)
            sheet_ranges.delete_cols(2, 3)
            sheet_ranges.delete_cols(3, 3)
            sheet_ranges.delete_cols(4, 11)
            sheet_ranges.delete_cols(7, 2)
            sheet_ranges['A1'] = "Vendeur"
            sheet_ranges['C1'] = "ACM"
            sheet_ranges['F1'] = "LINE"
            sheet_ranges['G1'] = "TSM"
            sheet_ranges['G1'].fill = PatternFill("solid", fgColor="4cbb17")
            sheet_ranges["F1"].fill = PatternFill("solid", fgColor="4cbb17")

            wb.save("excel/finale.xlsx")
            print("qualitatif saved")
        except Exception as e:
            print(e)

    def fix_sheets(self):
        self.__fix_quantitatif()
        self.__fix_qualitatif()
