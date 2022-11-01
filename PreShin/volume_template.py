import os
import re

import openpyxl
import pandas as pd
from PySide2.QtWidgets import QWidget
from openpyxl.styles import Border, PatternFill, borders

from PreShin.loggers import logger


class Vol_Template_UI(QWidget):
    def __init__(self):
        super().__init__()

    def initUI(self):
        pass


class Vol_Template:
    def __init__(self):
        super().__init__()
        self.lbl_diff = list
        self.pre_diff = list
        self.lbl_pre_diff = list
        self.label_id_list = list
        self.predict_id_list = list

        self.thin_border = Border(left=borders.Side(style='thin'),
                                  right=borders.Side(style='thin'),
                                  top=borders.Side(style='thin'),
                                  bottom=borders.Side(style='thin'))
        self.blue_color = PatternFill(start_color='b3d9ff', end_color='b3d9ff', fill_type='solid')
        self.green_color = PatternFill(start_color='c1f0c1', end_color='c1f0c1', fill_type='solid')
        self.red_color = PatternFill(start_color='ffcccc', end_color='ffcccc', fill_type='solid')
        self.yellow_color = PatternFill(start_color='ffffb3', end_color='ffffb3', fill_type='solid')
        self.gray_color = PatternFill(start_color='bfbfbf', end_color='bfbfbf', fill_type='solid')
        self.orange_color = PatternFill(start_color='ff9900', end_color='ff9900', fill_type='solid')
        self.outlier_color = PatternFill(start_color='C0504D', end_color='C0504D', fill_type='solid')
        self.white_color = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

    def main(self):
        pass

    def label(self, loc: str):  # 사람 기준
        # logger.info('label data transform start')
        label_dict = {}
        self.label_id_list = os.listdir(loc)

        for i in self.label_id_list:
            data_list = os.listdir(f'{loc}/{i}')
            data = [0, 0, 0]
            for j in range(len(data_list)):
                if j == 1:
                    data[j] = int(re.sub(r'[^0-9]', '', data_list[j])) + 128  # 정규 표현식 문자열 제거
                elif j == 2:
                    data[j] = int(re.sub(r'[^0-9]', '', data_list[j])) - 128  # 정규 표현식 문자열 제거
                else:
                    data[j] = int(re.sub(r'[^0-9]', '', data_list[j]))

            label_dict[i] = data  # dict 에 추가

        df_label = pd.DataFrame(label_dict, index=['Air', 'Hard Tissue', 'Soft Tissue'])

        return df_label
        # logger.info('label data transform end')

    def predict(self, loc: str):  # ai 데이터 기준
        # logger.info('predict data transform start')

        predict_dict = {}
        self.predict_id_list = os.listdir(loc)

        for i in self.predict_id_list:
            data_list = os.listdir(f'{loc}/{i}')
            data = [0, 0, 0]

            for j in range(len(data_list)):
                txt = open(f'{loc}/{i}/{data_list[j]}', 'r')
                lines = txt.readlines()

                for k in range(len(lines)):
                    line = lines[k].split(',')
                    line_float = line[1].split('\n')
                    data[k] = float(line_float[0])

            predict_dict[i] = data  # dict 에 추가

        df_predict = pd.DataFrame(predict_dict, index=['Air', 'Hard Tissue', 'Soft Tissue'])

        return df_predict

        # logger.info('predict data transform end')

    def percent(self, lbl: pd.DataFrame, pre: pd.DataFrame, float_outlier: float):
        lbl_percent = lbl.div(384)
        result_percent = abs(lbl_percent - pre)
        result_percent_outlier = result_percent[result_percent < float_outlier]

        result_percent_average = result_percent.mean(axis=1)
        result_percent_std = result_percent.std(axis=1, ddof=0)

        result_percent_outlier_average = result_percent_outlier.mean(axis=1)
        result_percent_outlier_std = result_percent_outlier.std(axis=1, ddof=0)

        result_percent.insert(0, 'Outlier_Std', result_percent_outlier_std)
        result_percent.insert(0, 'Outlier_Aver', result_percent_outlier_average)
        result_percent.insert(0, 'Std', result_percent_std)
        result_percent.insert(0, 'Aver', result_percent_average)

        return result_percent

    def sheet(self, lbl: pd.DataFrame, pre: pd.DataFrame, int_outlier: int):
        pre_sheet = pre.mul(384).round(0)
        result_sheet = abs(lbl - pre_sheet)
        result_sheet_outlier = result_sheet[result_sheet < int_outlier]

        result_sheet_average = result_sheet.mean(axis=1)
        result_sheet_std = result_sheet.std(axis=1, ddof=0)

        result_sheet_outlier_average = result_sheet_outlier.mean(axis=1)
        result_sheet_outlier_std = result_sheet_outlier.std(axis=1, ddof=0)

        result_sheet.insert(0, 'Outlier_Std', result_sheet_outlier_std)
        result_sheet.insert(0, 'Outlier_Aver', result_sheet_outlier_average)
        result_sheet.insert(0, 'Std', result_sheet_std)
        result_sheet.insert(0, 'Aver', result_sheet_average)

        return result_sheet

    def pre_lbl_compare(self):
        self.lbl_diff = list(set(self.label_id_list) - set(self.predict_id_list))  # 차집합
        self.pre_diff = list(set(self.predict_id_list) - set(self.label_id_list))

    def to_xlsx(self, loc: str, lbl: pd.DataFrame, pre: pd.DataFrame, float_outlier: float, int_outlier: int):

        writer = pd.ExcelWriter(loc, engine='openpyxl')
        percent_result = self.percent(lbl, pre, float_outlier)
        sheet_result = self.sheet(lbl, pre, int_outlier)
        percent_result.to_excel(writer, startcol=0, startrow=3)
        sheet_result.to_excel(writer, startcol=0, startrow=20)
        writer.close()

    def xlsx_style(self, loc: str, float_outlier: float, int_outlier: int, float_error: float, int_error: int):
        wb = openpyxl.load_workbook(filename=loc)
        ws = wb['Sheet1']

        ws.cell(row=3, column=6).value = 'Patient_ID'
        ws.cell(row=20, column=6).value = 'Patient_ID'
        ws.merge_cells(start_row=3, start_column=6, end_row=3, end_column=ws.max_column)
        ws.merge_cells(start_row=20, start_column=6, end_row=20, end_column=ws.max_column)
        ws.cell(row=3, column=6).fill = self.blue_color
        ws.cell(row=20, column=6).fill = self.blue_color

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['D'].width = 15

        wb.save(filename=loc)





vol = Vol_Template()
df_lbl = vol.label(r'D:\temp_data\label')
df_pre = vol.predict(r'D:\temp_data\predict')

vol.to_xlsx(r'D:\temp_data\temp.xlsx', df_lbl, df_pre, 0.05, 15)
vol.xlsx_style(r'D:\temp_data\temp.xlsx')
