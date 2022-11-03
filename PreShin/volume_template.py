import os
import re

import numpy as np
import openpyxl
import pandas as pd
from PySide2.QtCore import Qt
from PySide2.QtGui import QIntValidator
from PySide2.QtWidgets import QWidget, QTableWidget, QTableWidgetItem, QPushButton, QLabel, QLineEdit, QPlainTextEdit, QFileDialog, QDialog, QMessageBox
from matplotlib import pyplot as plt
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, PatternFill, borders, Font
import seaborn as sns

from PreShin.loggers import logger

xlsx_name = 'result.xlsx'


def messagebox(text: str, i: str):
    signBox = QMessageBox()
    signBox.setWindowTitle(text)
    signBox.setText(i)

    signBox.setIcon(QMessageBox.Information)
    signBox.setStandardButtons(QMessageBox.Ok)
    signBox.exec_()


class Vol_Template_UI(QWidget):
    def __init__(self):
        super().__init__()
        self.dialog = QDialog()
        self.initUI()

    def initUI(self):
        batch = '4'
        rate = '2e-4'
        optimizer = 'adam'
        aug = '0'
        comment = 'write comment'
        safe_zone_sheet = '3'
        outlier_sheet = '5'

        self.table = QTableWidget(4, 2, self.dialog)
        self.table.setSortingEnabled(False)  # 정렬 기능
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(0, 80)  # checkbox 열 폭 강제 조절.
        self.table.setColumnWidth(1, 80)

        self.table.setItem(0, 0, QTableWidgetItem('Batch size'))
        self.table.setItem(0, 1, QTableWidgetItem(batch))
        self.table.setItem(1, 0, QTableWidgetItem('Learning rate'))
        self.table.setItem(1, 1, QTableWidgetItem(rate))
        self.table.setItem(2, 0, QTableWidgetItem('Optimizer'))
        self.table.setItem(2, 1, QTableWidgetItem(optimizer))
        self.table.setItem(3, 0, QTableWidgetItem('Aug'))
        self.table.setItem(3, 1, QTableWidgetItem(aug))

        self.table.setHorizontalHeaderLabels(["Name", "Value"])
        self.table.setGeometry(20, 195, 180, 145)

        btn_pre_path = QPushButton(self.dialog)
        btn_lbl_path = QPushButton(self.dialog)
        btn_export = QPushButton(self.dialog)
        btn_manual = QPushButton(self.dialog)

        btn_pre_path.setText('Predict Path')
        btn_lbl_path.setText('Label Path')
        btn_export.setText('Export Excel')
        btn_manual.setText('Open Manual')

        self.edt_pre = QLineEdit(self.dialog)
        self.edt_lbl = QLineEdit(self.dialog)
        lbl_error = QLabel(self.dialog)
        lbl_outlier = QLabel(self.dialog)

        lbl_error_percent = QLabel(self.dialog)
        lbl_error_sheet = QLabel(self.dialog)
        lbl_outlier_percent = QLabel(self.dialog)
        lbl_outlier_sheet = QLabel(self.dialog)

        lbl_comment = QLabel(self.dialog)
        lbl_xlsx_name = QLabel(self.dialog)
        lbl_xlsx = QLabel(self.dialog)
        self.edt_error_sheet = QLineEdit(self.dialog)
        self.edt_error_sheet.setAlignment(Qt.AlignRight)
        self.lbl_error_percent = QLabel(self.dialog)
        self.lbl_error_percent.setAlignment(Qt.AlignRight)
        self.edt_outlier_sheet = QLineEdit(self.dialog)
        self.edt_outlier_sheet.setAlignment(Qt.AlignRight)
        self.lbl_outlier_percent = QLabel(self.dialog)
        self.lbl_outlier_percent.setAlignment(Qt.AlignRight)
        self.edt_xlsx_name = QLineEdit(self.dialog)
        self.edt_xlsx_name.setAlignment(Qt.AlignRight)  # 엑셀명

        self.edt_lbl.setGeometry(130, 35, 230, 20)
        self.edt_pre.setGeometry(130, 60, 230, 20)
        lbl_error.setGeometry(220, 200, 100, 20)
        lbl_outlier.setGeometry(220, 269, 100, 20)

        lbl_error_percent.setGeometry(275, 246, 100, 20)
        lbl_error_sheet.setGeometry(275, 223, 100, 20)
        lbl_outlier_percent.setGeometry(275, 315, 100, 20)
        lbl_outlier_sheet.setGeometry(275, 292, 100, 20)

        lbl_comment.move(20, 90)
        lbl_xlsx_name.move(20, 355)
        lbl_xlsx.move(173, 355)

        self.edt_error_sheet.setGeometry(220, 223, 50, 20)
        self.lbl_error_percent.setGeometry(220, 250, 50, 20)
        self.edt_outlier_sheet.setGeometry(220, 292, 50, 20)
        self.lbl_outlier_percent.setGeometry(220, 319, 50, 20)
        self.edt_xlsx_name.setGeometry(70, 350, 103, 20)

        lbl_error_percent.setText('%')
        lbl_error_sheet.setText('장')
        lbl_outlier_percent.setText('%')
        lbl_outlier_sheet.setText('장')

        lbl_error.setText('Error Safe Zone')
        lbl_outlier.setText('Remove Outlier')
        lbl_comment.setText('Comment')
        lbl_xlsx_name.setText('파일명 : ')
        lbl_xlsx.setText('.xlsx')

        self.onlyInt = QIntValidator()
        self.edt_error_sheet.setValidator(self.onlyInt)
        self.edt_outlier_sheet.setValidator(self.onlyInt)
        self.edt_error_sheet.setText(safe_zone_sheet)
        self.edt_error_sheet.textChanged[str].connect(self.lbl_error_changed)
        self.lbl_error_percent.setText(str(round((int(self.edt_error_sheet.text()) / 384 * 100), 5)))
        self.edt_outlier_sheet.setText(outlier_sheet)
        self.edt_outlier_sheet.textChanged[str].connect(self.lbl_error_changed)
        self.lbl_outlier_percent.setText(str(round((int(self.edt_outlier_sheet.text()) / 384 * 100), 5)))
        btn_manual.setGeometry(20, 10, 100, 20)
        btn_lbl_path.setGeometry(20, 35, 100, 20)
        btn_pre_path.setGeometry(20, 60, 100, 20)
        btn_export.setGeometry(220, 345, 120, 30)

        btn_lbl_path.clicked.connect(self.btn_lbl_clicked)
        btn_pre_path.clicked.connect(self.btn_pre_clicked)
        btn_export.clicked.connect(self.btn_export_clicked)
        # btn_manual.clicked.connect(btn_manual_clicked)

        self.edt = QPlainTextEdit(self.dialog)
        self.edt.setPlainText(comment)
        self.edt.setGeometry(20, 105, 340, 80)

        self.dialog.setWindowTitle('AI')
        self.dialog.setGeometry(500, 300, 370, 420)
        self.dialog.exec()

    def lbl_error_changed(self, text):
        self.lbl_error_percent.setText(str(round((int(text) / 384 * 100), 5)))

    def lbl_outlier_changed(self, text):
        self.lbl_outlier_percent.setText(str(round((int(text) / 384 * 100), 5)))

    def btn_lbl_clicked(self):
        logger.info('lbl_btn in')
        loc = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(),
                                               QFileDialog.ShowDirsOnly)  # 창 title, 주소 나중에 변경

        # 폴더 경로 입력
        if loc != '':
            self.edt_lbl.setText(str(loc))
        else:
            self.edt_lbl.setText('')
        logger.info('lbl_btn out')

    def btn_pre_clicked(self):
        logger.info('pre_btn in')
        loc = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(),
                                               QFileDialog.ShowDirsOnly)  # 창 title, 주소 나중에 변경

        # 폴더 경로 입력
        if loc != '':
            self.edt_pre.setText(str(loc))
        else:
            self.edt_pre.setText('')
        logger.info('pre_btn out')

    def btn_export_clicked(self):
        if self.edt_lbl.text() != '' and self.edt_pre.text() != '':

            if self.edt_xlsx_name.text() != '':  # 파일명 입력 했을때

                loc_xlsx = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(),
                                                            QFileDialog.ShowDirsOnly)
                if loc_xlsx != '':  # 폴더 선택 했을때
                    file = os.listdir(loc_xlsx)  # 엑셀 저장 위치에 있는 파일 읽기

                    if fr'{self.edt_xlsx_name.text()}.xlsx' not in file:  # 동일한 파일명이 없을때
                        vol = Vol_Template()
                        df_lbl = vol.label(self.edt_lbl.text())
                        df_pre = vol.predict(self.edt_pre.text())

                        img_folder = vol.create_img_folder(loc_xlsx)

                        vol.to_xlsx(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', df_lbl, df_pre, percent_outlier=float(self.lbl_outlier_percent.text()),
                                    sheet_outlier=int(self.edt_outlier_sheet.text()))
                        vol.sheet1_xlsx_style(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', percent_error=float(self.lbl_error_percent.text()),
                                              percent_outlier=float(self.lbl_outlier_percent.text()), sheet_outlier=int(self.edt_outlier_sheet.text()),
                                              sheet_error=int(self.edt_error_sheet.text()),
                                              image_folder_loc=img_folder)
                        vol.sheet2_xlsx_style(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', percent_error=float(self.lbl_error_percent.text()),
                                              percent_outlier=float(self.lbl_outlier_percent.text()),
                                              image_folder_loc=img_folder)
                        self.insert_comment(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', 'Sheet1')
                        self.insert_comment(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', 'Sheet')
                        messagebox('notice', 'Excel 생성이 완료 되었습니다.')

                    else:
                        messagebox('Warning', "동일한 파일명이 존재합니다. 다시 입력하세요")

                else:
                    pass

            else:
                messagebox('Warning', "파일명을 입력하세요")

        else:
            messagebox('Warning', "label 또는 predict 경로를 확인 하세요.")

    def insert_comment(self, loc: str, xlsx: str, sheet: str):
        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb[sheet]
        # table 에 default 값 출력
        ws.cell(row=1, column=6).value = f'Hyperparameter Batch size = {self.table.item(0, 1).text()}' \
                                         f', Learning rate = {self.table.item(1, 1).text()}' \
                                         f', optimizer = {self.table.item(2, 1).text()}' \
                                         f', aug = {self.table.item(3, 1).text()} '
        ws.cell(row=2, column=6).value = f'comment : {self.edt.toPlainText()}'

        ws.cell(row=1, column=2).font = Font(bold=True)
        ws.cell(row=1, column=6).font = Font(bold=True)
        ws.cell(row=2, column=6).font = Font(bold=True)
        if sheet == 'Sheet1':
            ws.cell(row=1, column=2).value = f'Error Safe Zone : {self.lbl_error_percent.text()} %  {self.edt_error_sheet.text()} 장'
            ws.cell(row=2, column=2).value = f'Remove Outlier : {self.lbl_outlier_percent.text()} %  {self.edt_outlier_sheet.text()} 장'
        elif sheet == 'Sheet':
            ws.cell(row=1, column=2).value = f'Error Safe Zone : {100 - float(self.lbl_error_percent.text())} %  '
            ws.cell(row=2, column=2).value = f'Remove Outlier : {100 - float(self.lbl_outlier_percent.text())} %  '

        ws.cell(row=2, column=2).font = Font(bold=True)

        wb.save(filename=f'{loc}/{xlsx}')


class Vol_Template:
    def __init__(self):
        super().__init__()
        self.right_percent_aver_std = pd.DataFrame()
        self.sheet_aver_std = pd.DataFrame()
        self.percent_aver_std = pd.DataFrame()
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
        self.yellow_color2 = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        self.gray_color2 = PatternFill(start_color='e0e0eb', end_color='e0e0eb', fill_type='solid')
        self.gray_color = PatternFill(start_color='bfbfbf', end_color='bfbfbf', fill_type='solid')
        self.blue_color2 = PatternFill(start_color='ccf5ff', end_color='ccf5ff', fill_type='solid')
        self.orange_color = PatternFill(start_color='ff9900', end_color='ff9900', fill_type='solid')
        self.outlier_color = PatternFill(start_color='C0504D', end_color='C0504D', fill_type='solid')
        self.white_color = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        self.air_color = PatternFill(start_color='8DD3C7', end_color='8DD3C7', fill_type='solid')
        self.hts_color = PatternFill(start_color='FFFFB3', end_color='FFFFB3', fill_type='solid')
        self.sts_color = PatternFill(start_color='BEBADA', end_color='BEBADA', fill_type='solid')

    def main(self):
        pass

    def label(self, loc: str):  # 정답 기준
        # logger.info('label data transform start')
        label_dict = {}
        self.label_id_list = os.listdir(loc)  # 환자 list

        for i in self.label_id_list:
            data_list = os.listdir(f'{loc}/{i}')  # air, hts, sts 순서
            data = [-999999, -999999, -999999]
            for j in range(len(data_list)):
                if 'hts' in data_list[j]:  # hts
                    data[1] = int(re.sub(r'[^0-9]', '', data_list[j])) + 128  # 정규 표현식 문자열 제거, 정답을 ai 기준으로 맞추기 위해서 128씩 더하고 뺌
                elif 'sts' in data_list[j]:  # sts
                    data[2] = int(re.sub(r'[^0-9]', '', data_list[j])) - 128
                elif 'air' in data_list[j]:
                    data[0] = int(re.sub(r'[^0-9]', '', data_list[j]))

            label_dict[i] = data  # dict 에 추가

        df_label = pd.DataFrame(label_dict, index=['Air', 'Hard Tissue', 'Soft Tissue'])
        return df_label
        # logger.info('label data transform end')

    def predict(self, loc: str):  # ai 데이터 기준
        # logger.info('predict data transform start')

        predict_dict = {}
        self.predict_id_list = os.listdir(loc)  # 환자 list

        for i in self.predict_id_list:
            data_list = os.listdir(f'{loc}/{i}')  # txt 하나, 2개 이상 일 때 잘못된 것으로 나중에 error 코드 추가
            data = [0, 0, 0]

            for j in range(len(data_list)):
                txt = open(f'{loc}/{i}/{data_list[j]}', 'r')
                lines = txt.readlines()  # txt 한줄씩 읽기

                for k in range(len(lines)):
                    line = lines[k].split(',')  # air,0.35156316 형식, 숫자만 남김
                    line_float = line[1].split('\n')
                    data[k] = float(line_float[0])

            predict_dict[i] = data  # dict 에 추가

        df_predict = pd.DataFrame(predict_dict, index=['Air', 'Hard Tissue', 'Soft Tissue'])
        return df_predict

        # logger.info('predict data transform end')

    def percent(self, lbl: pd.DataFrame, pre: pd.DataFrame, *args):
        if 'right' in args:
            lbl_percent = lbl.div(384)
            lbl_percent = lbl_percent.mul(100)
            pre = pre.mul(100)
            lbl_percent = lbl_percent.sub(100)
            pre = pre.sub(100)
            result_percent = abs(lbl_percent - pre).sub(100)  # 퍼센티지에 대한 결과
        else:
            lbl_percent = lbl.div(384)
            lbl_percent = lbl_percent.mul(100)
            pre = pre.mul(100)
            result_percent = abs(lbl_percent - pre)  # 퍼센티지에 대한 결과

        result_percent = result_percent[result_percent < 10000]
        return result_percent

    def sheet(self, lbl: pd.DataFrame, pre: pd.DataFrame):
        pre_sheet = pre.mul(384).round(0)
        result_sheet = abs(lbl - pre_sheet)  # 장수에 대한 결과
        result_sheet = result_sheet[result_sheet < 10000]

        return result_sheet

    def percent_sheet_result(self, outlier, method: pd.DataFrame):

        result = method

        result_outlier = result[result < outlier]  # outlier 값을 버림 ( nan 으로 만듬 )

        result_average = result.mean(axis=1)
        result_std = result.std(axis=1, ddof=0)

        result_outlier_average = result_outlier.mean(axis=1)
        result_outlier_std = result_outlier.std(axis=1, ddof=0)

        result_aver_std = pd.DataFrame()  # 환자 측정과 평균,표준편차 df 나눔
        result = result.transpose()
        result_aver_std.insert(0, 'Out_Std', result_outlier_std)
        result_aver_std.insert(0, 'Out_Aver', result_outlier_average)
        result_aver_std.insert(0, 'Std', result_std)
        result_aver_std.insert(0, 'Aver', result_average)

        result = round(result, 6)
        # result = result.fillna('None')
        result_aver_std = round(result_aver_std, 6)

        return result, result_aver_std

    def pre_lbl_compare(self):
        self.lbl_diff = list(set(self.label_id_list) - set(self.predict_id_list))  # 차집합
        self.pre_diff = list(set(self.predict_id_list) - set(self.label_id_list))

    def to_xlsx(self, loc: str, xlsx: str, lbl: pd.DataFrame, pre: pd.DataFrame, percent_outlier: float, sheet_outlier: int):
        writer = pd.ExcelWriter(f'{loc}/{xlsx}', engine='openpyxl')  # pandas 엑셀 작성

        # 결과 값 불러옴(pd.dataframe 구성)
        right_percent_result, self.right_percent_aver_std = self.percent_sheet_result(percent_outlier, self.percent(lbl, pre, 'right'))
        percent_result, self.percent_aver_std = self.percent_sheet_result(percent_outlier, self.percent(lbl, pre))
        sheet_result, self.sheet_aver_std = self.percent_sheet_result(sheet_outlier, self.sheet(lbl, pre))

        right_percent_result = abs(right_percent_result)
        self.right_percent_aver_std = abs(self.right_percent_aver_std)
        # df 엑셀에 입력
        right_percent_result.to_excel(writer, startcol=0, startrow=3, sheet_name='Sheet')
        self.right_percent_aver_std.to_excel(writer, startcol=5, startrow=3, sheet_name='Sheet')
        percent_result.to_excel(writer, startcol=0, startrow=3, sheet_name='Sheet1')
        self.percent_aver_std.to_excel(writer, startcol=5, startrow=3, sheet_name='Sheet1')
        sheet_result.to_excel(writer, startcol=12, startrow=3, sheet_name='Sheet1')
        self.sheet_aver_std.to_excel(writer, startcol=17, startrow=3, sheet_name='Sheet1')
        writer.close()

    def create_img_folder(self, loc: str):
        img_folder = f'{loc}/graph_image'
        os.mkdir(img_folder)  # 폴더 생성

        return img_folder

    def accept_outlier_error_sh1(self, ws, row, column, error, outlier):
        if ws.cell(row=row, column=column).value is None:
            ws.cell(row=row, column=column).value = 'None'
            ws.cell(row=row, column=column).fill = self.gray_color2

        elif ws.cell(row=row, column=column).value > error:
            ws.cell(row=row, column=column).fill = self.yellow_color2

            if ws.cell(row=row, column=column).value > outlier:
                ws.cell(row=row, column=column).fill = self.outlier_color

    def accept_outlier_error_sh2(self, ws, row, column, error, outlier):
        if ws.cell(row=row, column=column).value is None:
            ws.cell(row=row, column=column).value = 'None'
            ws.cell(row=row, column=column).fill = self.gray_color2

        elif ws.cell(row=row, column=column).value < error:
            ws.cell(row=row, column=column).fill = self.yellow_color2

            if ws.cell(row=row, column=column).value < outlier:
                ws.cell(row=row, column=column).fill = self.outlier_color

    def sheet1_xlsx_style(self, loc: str, xlsx: str, percent_outlier: float, sheet_outlier: int, percent_error: float, sheet_error: int, image_folder_loc: str):
        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb['Sheet1']

        ws.cell(row=4, column=1).value = 'Patient_ID'
        ws.cell(row=4, column=13).value = 'Patient_ID'

        # patient 결과 outlier, error 적용 색상
        for row in range(5, ws.max_row + 1):

            for column in range(2, 5):  # percent
                self.accept_outlier_error_sh1(ws, row, column, percent_error, percent_outlier)

            for column in range(14, 17):  # sheet
                self.accept_outlier_error_sh1(ws, row, column, sheet_error, sheet_outlier)

            for column in [7, 9]:  # percent aver_std
                if ws.cell(row=row, column=column).value is None:
                    pass
                else:
                    self.accept_outlier_error_sh1(ws, row, column, percent_error, percent_outlier)

            for column in [19, 21]:  # percent aver_std
                if ws.cell(row=row, column=column).value is None:
                    pass
                else:
                    self.accept_outlier_error_sh1(ws, row, column, sheet_error, sheet_outlier)

        # air, hard tissue, soft tissue 색상
        for column in range(1, ws.max_column + 1):
            if ws.cell(row=4, column=column).value is not None and ws.cell(row=4, column=column).value != 'Patient_ID':
                ws.cell(row=4, column=column).fill = self.blue_color

        # # patient_id title 색상
        # ws.cell(row=4, column=1).fill = self.blue_color2
        # ws.cell(row=4, column=1).border = self.thin_border
        # ws.cell(row=4, column=13).fill = self.blue_color2
        # ws.cell(row=4, column=13).border = self.thin_border
        for row in range(1, ws.max_row + 1):
            for column in range(1, ws.max_column + 1):
                if ws.cell(row=row, column=column).value is not None:
                    ws.cell(row=row, column=column).border = self.thin_border
        # patient_id 색상
        for row in range(5, ws.max_row + 1):
            ws.cell(row=row, column=1).fill = self.yellow_color
            ws.cell(row=row, column=13).fill = self.yellow_color

        col_number = [6, 18]
        for i in col_number:
            ws.cell(row=5, column=i).fill = self.air_color
            ws.cell(row=6, column=i).fill = self.hts_color
            ws.cell(row=7, column=i).fill = self.sts_color

        # 이미지 삽입

        self.graph(self.percent_aver_std, 'percent', '', loc, percent_error)
        self.graph(self.percent_aver_std, 'remove_outlier_percent', 'Out_', loc, percent_error)
        self.graph(self.sheet_aver_std, 'sheet', '', loc, sheet_error)
        self.graph(self.sheet_aver_std, 'remove_outlier_sheet', 'Out_', loc, sheet_error)

        img_list = os.listdir(f'{loc}/graph_image')
        for i in img_list:
            if '.png' in i:
                img = Image(image_folder_loc + f'/{i}')
                if i == 'remove_outlier_percent.png':
                    ws.add_image(img, 'F23')
                elif i == 'remove_outlier_sheet.png':
                    ws.add_image(img, 'R23')
                elif i == 'sheet.png':
                    ws.add_image(img, 'R9')
                elif i == 'percent.png':
                    ws.add_image(img, 'F9')

        ws.auto_filter.ref = f'A4:D{ws.max_row}'  # 엑셀 필터 적용

        # column 사이즈
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['F'].width = 11
        ws.column_dimensions['N'].width = 14
        ws.column_dimensions['O'].width = 14
        ws.column_dimensions['P'].width = 14
        ws.column_dimensions['R'].width = 11

        wb.save(filename=f'{loc}/{xlsx}')

    def sheet2_xlsx_style(self, loc: str, xlsx: str, percent_outlier: float, percent_error: float, image_folder_loc: str):
        percent_outlier = abs(percent_outlier - 100)
        percent_error = abs(percent_error - 100)

        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb['Sheet']
        ws.cell(row=4, column=1).value = 'Patient_ID'
        # patient 결과 outlier, error 적용 색상
        for row in range(5, ws.max_row + 1):
            for column in range(2, 5):  # percent
                self.accept_outlier_error_sh2(ws, row, column, percent_error, percent_outlier)
            for column in [7, 9]:  # percent aver_std
                if ws.cell(row=row, column=column).value is None:
                    pass
                else:
                    self.accept_outlier_error_sh2(ws, row, column, percent_error, percent_outlier)
        # air, hard tissue, soft tissue 색상
        for column in range(1, ws.max_column + 1):
            if ws.cell(row=4, column=column).value is not None and ws.cell(row=4, column=column).value != 'Patient_ID':
                ws.cell(row=4, column=column).fill = self.blue_color

        for row in range(1, ws.max_row + 1):
            for column in range(1, ws.max_column + 1):
                if ws.cell(row=row, column=column).value is not None:
                    ws.cell(row=row, column=column).border = self.thin_border
        # patient_id 색상
        for row in range(5, ws.max_row + 1):
            ws.cell(row=row, column=1).fill = self.yellow_color
        col_number = [6]
        for i in col_number:
            ws.cell(row=5, column=i).fill = self.air_color
            ws.cell(row=6, column=i).fill = self.hts_color
            ws.cell(row=7, column=i).fill = self.sts_color
        # 이미지 삽입

        self.graph(self.right_percent_aver_std, 'right_percent', '', loc, percent_error)
        self.graph(self.right_percent_aver_std, 'right_remove_outlier_percent', 'Out_', loc, percent_outlier)
        img_list = os.listdir(f'{loc}/graph_image')

        for i in img_list:
            if '.png' in i:
                img = Image(image_folder_loc + f'/{i}')
                if i == 'right_remove_outlier_percent.png':
                    ws.add_image(img, 'F23')
                elif i == 'right_percent.png':
                    ws.add_image(img, 'F9')

        ws.auto_filter.ref = f'A4:D{ws.max_row}'  # 엑셀 필터 적용
        # column 사이즈
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['F'].width = 14
        wb.save(filename=f'{loc}/{xlsx}')

    def graph(self, df: pd.DataFrame, title: str, outlier: str, loc: str, error_line):  # 그래프 제작
        graph = df
        graph_dict = graph.to_dict('list')  # dataframe list 제작

        fig = plt.figure(figsize=(5, 3))  # Figure 생성 사이즈
        ax = fig.add_subplot()  # Axes 추가
        colors = sns.color_palette('Set3', len(list(graph.index)))  # 색상
        xtick_label_position = list(range(len(list(graph.index))))  # x 축에 글시 넣을 위치

        # if 'percent' in title:    # sheet 와 percent 에 따른 축 범위 변경
        #     plt.ylim([0,1])
        # elif 'sheet' in title:
        #     plt.ylim([0,1])

        plt.xticks(xtick_label_position, list(graph.index))  # x 축에 삽입
        plt.axhline(y=float(error_line), color='red', linestyle='--')  # error 라인 그리기
        bars = plt.bar(xtick_label_position, graph_dict[f'{outlier}Aver'], color=colors, edgecolor='black')  # 그래프 생성
        plt.title(title, fontsize=10)  # 타이틀 입력
        plt.errorbar(x=list(graph.index), y=graph_dict[f'{outlier}Aver'], yerr=np.array(graph_dict[f'{outlier}Std']) / 2, color='black', ecolor='black', fmt='.',
                     alpha=0.5, elinewidth=2)  # 에러바 삽입

        for i, b in enumerate(bars):  # 바에 결과 값 추가
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() / 2, graph_dict[f'{outlier}Aver'][i], ha='center', fontsize=10)

        plt.xticks(rotation=0)
        plt.savefig(f'{loc}/graph_image/{title}.png')


if __name__ == "__main__":
    vol = Vol_Template()
    df_lbl = vol.label(r'D:\temp_data\label')
    df_pre = vol.predict(r'D:\temp_data\predict')

    image_folder = vol.create_img_folder(r'D:\temp_data')
    vol.to_xlsx(fr'D:\temp_data', xlsx_name, df_lbl, df_pre, 5, 15)
    vol.sheet1_xlsx_style(fr'D:\temp_data', xlsx_name, 5, 13, 3, 7, image_folder)
    vol.sheet2_xlsx_style(fr'D:\temp_data', xlsx_name, 5, 3, image_folder)  # ----------------------------------------------------------- 평균치 이하 표시 수정해야함.
fr'D:\temp_data\graph_image'
