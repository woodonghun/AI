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
        self.edt_error_sheet.setAlignment(Qt.AlignRight)    # 우측 정렬
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

        self.onlyInt = QIntValidator(0, 384)    # 숫자만 입력, 범위 -> 왜 384까지가 아니라 999가 되는지 모르겠음
        self.edt_error_sheet.setValidator(self.onlyInt)
        self.edt_outlier_sheet.setValidator(self.onlyInt)

        self.edt_error_sheet.setText(safe_zone_sheet)
        self.edt_error_sheet.textChanged[str].connect(self.lbl_error_changed)
        self.lbl_error_percent.setText(str(round((int(self.edt_error_sheet.text()) / 384 * 100), 5)))
        self.edt_outlier_sheet.setText(outlier_sheet)
        self.edt_outlier_sheet.textChanged[str].connect(self.lbl_outlier_changed)
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

    def zero_none_edt(self):
        if self.edt_error_sheet.text() == '':
            self.edt_error_sheet.setText(0)

    def lbl_error_changed(self, text):    # edt 에 입력한 장수를 label 에 % 로 변환하고 실시간 전환, 00, 000, ... , None 에 맞춰서 변환 하는건 못함
        self.lbl_error_percent.setText(str(round((int(text) / 384 * 100), 3)))

    def lbl_outlier_changed(self, text):
        self.lbl_outlier_percent.setText(str(round((int(text) / 384 * 100), 3)))

    def btn_lbl_clicked(self):
        logger.info('lbl_btn in')
        loc = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(), QFileDialog.ShowDirsOnly)  # 창 title, 주소 나중에 변경

        # 폴더 경로 입력
        if loc != '':
            self.edt_lbl.setText(str(loc))
        else:
            self.edt_lbl.setText('')
        logger.info('lbl_btn out')

    def btn_pre_clicked(self):
        logger.info('pre_btn in')
        loc = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(), QFileDialog.ShowDirsOnly)  # 창 title, 주소 나중에 변경

        # 폴더 경로 입력
        if loc != '':
            self.edt_pre.setText(str(loc))
        else:
            self.edt_pre.setText('')
        logger.info('pre_btn out')

    def btn_export_clicked(self):

        logger.info('btn_export_clicked')

        if self.edt_lbl.text() != '' and self.edt_pre.text() != '':

            if self.edt_xlsx_name.text() != '':  # 파일명 입력 했을때

                loc_xlsx = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(), QFileDialog.ShowDirsOnly)
                if loc_xlsx != '':  # 폴더 선택 했을때
                    file = os.listdir(loc_xlsx)  # 엑셀 저장 위치에 있는 파일 읽기

                    if fr'{self.edt_xlsx_name.text()}.xlsx' not in file:  # 동일한 파일명이 없을때

                        try:  # lbl pre 폴더 경로 확인
                            os.listdir(self.edt_pre.text())
                            os.listdir(self.edt_lbl.text())
                        except FileNotFoundError:
                            messagebox('Warning', 'predict 또는 lbl 의 폴더 경로가 올바르지 않습니다.')
                            logger.error(" 폴더 경로 에러")

                        vol = Vol_Template()
                        vol.pre_lbl_compare(self.edt_lbl.text(), self.edt_pre.text())  # lbl, pre 폴더에 존재 하는 폴더 목록 비교
                        df_lbl = vol.label(self.edt_lbl.text())  # label loc 읽기
                        df_pre = vol.predict(self.edt_pre.text())  # predict loc 읽기

                        img_folder = vol.create_img_folder(loc_xlsx, self.edt_xlsx_name.text())  # image 폴더 생성

                        vol.to_xlsx(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', df_lbl, df_pre, percent_outlier=float(self.lbl_outlier_percent.text()),
                                    sheet_outlier=int(self.edt_outlier_sheet.text()))  # 엑셀 생성
                        vol.sheet1_xlsx_style(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', percent_error=float(self.lbl_error_percent.text()),
                                              percent_outlier=float(self.lbl_outlier_percent.text()), sheet_outlier=int(self.edt_outlier_sheet.text()),
                                              sheet_error=int(self.edt_error_sheet.text()),
                                              image_folder_loc=img_folder, file_name=self.edt_xlsx_name.text())  # sheet1 설정
                        vol.sheet2_xlsx_style(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', percent_error=float(self.lbl_error_percent.text()),
                                              percent_outlier=float(self.lbl_outlier_percent.text()),
                                              image_folder_loc=img_folder, file_name=self.edt_xlsx_name.text())  # sheet 설정

                        self.insert_comment(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', 'Sheet1')  # ui 에 있는 comment 삽입
                        self.insert_comment(loc_xlsx, fr'{self.edt_xlsx_name.text()}.xlsx', 'Sheet2')
                        messagebox('notice', 'Excel 생성이 완료 되었습니다.')

                    else:
                        messagebox('Warning', "동일한 파일명이 존재합니다. 다시 입력하세요")
                        logger.error("same file name exist")

                else:
                    pass  # 폴더 선택 하지 않았을 때 pass

            else:
                messagebox('Warning', "파일명을 입력하세요")
                logger.error("no file name")

        else:
            messagebox('Warning', "label 또는 predict 경로를 확인 하세요.")
            logger.error("label, predict location error")
        logger.info("btn_export out")

    def insert_comment(self, loc: str, xlsx: str, sheet: str):  # comment 삽입
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
        elif sheet == 'Sheet2':
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

    def label(self, loc: str):  # 정답 기준
        logger.info('label data transform start')
        label_dict = {}
        label_id_list = os.listdir(loc)  # 환자 list

        for i in label_id_list:
            if '.' in i:  # 폴더가 아닌 확장자인 경우 제외
                logger.error(f'폴더가 아닌 파일이 존재합니다. File Name : {i}')
                continue
            data_list = os.listdir(f'{loc}/{i}')  # air, hts, sts 순서

            data = [-999999, -999999, -999999]  # -999999 로 표현해서 나중에 없는 사진이 있을 경우 이상치는 삭제해서 없는 이미지 판단함
            if len(data_list) > 3 and '.png' not in data_list:  # png 파일이 아니거나 3개 초과일 때 제외, 3개 미만인 경우에는 진행하고 결측치로 표현한다.
                logger.error(f'파일 구성이 올바르지 않습니다. File Name : {i} - {data_list}')
                continue

            for j in range(len(data_list)):

                if 'hts' in data_list[j]:  # hts
                    data[1] = int(re.sub(r'[^0-9]', '', data_list[j])) + 128  # 정규 표현식 문자열 제거, 정답을 ai 기준으로 맞추기 위해서 128씩 더하고 뺌
                elif 'sts' in data_list[j]:  # sts
                    data[2] = int(re.sub(r'[^0-9]', '', data_list[j])) - 128
                elif 'air' in data_list[j]:
                    data[0] = int(re.sub(r'[^0-9]', '', data_list[j]))

            label_dict[i] = data  # dict 에 추가

        df_label = pd.DataFrame(label_dict, index=['Air', 'Hard Tissue', 'Soft Tissue'])
        logger.info('label data transform end')

        return df_label

    def predict(self, loc: str):  # ai 데이터 기준
        logger.info('predict data transform start')

        predict_dict = {}
        predict_id_list = os.listdir(loc)  # 환자 list

        for i in predict_id_list:
            if '.' in i:  # 폴더가 아닌 확장자가 들어간 파일인 경우 제외
                logger.error(f'폴더가 아닌 파일이 존재합니다. File Name : {i}')
                continue

            data_list = os.listdir(f'{loc}/{i}')  # txt 하나, 2개 이상 일 때 잘못된 것으로 나중에 error 코드 추가
            data = [-999999, -999999, -999999]
            if len(data_list) > 1 and '.dat' not in data_list:  # .dat 파일이 아닌 경우, 폴더안에 파일이 2개 이상인 경우 제외
                logger.error(f'파일 구성이 올바르지 않습니다. File Name : {i} - {data_list}')
                continue

            for j in range(len(data_list)):
                txt = open(f'{loc}/{i}/{data_list[j]}', 'r')
                lines = txt.readlines()  # txt 한줄씩 읽기

                for k in range(len(lines)):
                    line = lines[k].split(',')  # air,0.35156316 형식, 숫자만 남김
                    line_float = line[1].split('\n')
                    data[k] = float(line_float[0])

            predict_dict[i] = data  # dict 에 추가

        df_predict = pd.DataFrame(predict_dict, index=['Air', 'Hard Tissue', 'Soft Tissue'])
        logger.info('predict data transform end')

        return df_predict

    def percent(self, lbl: pd.DataFrame, pre: pd.DataFrame, *args) -> pd.DataFrame:  # 퍼센티지의 오차, 셩공률 에 대한 결과 값 함수
        logger.info('percent data transform start')

        # lbl = 장수 -> 소수점 , pre = 소수점 -> 소수점
        if 'right' in args:
            lbl_percent = lbl.div(384)  # 나누기
            lbl_percent = lbl_percent.mul(100)  # 곱
            pre = pre.mul(100)
            lbl_percent = lbl_percent.sub(100)
            pre = pre.sub(100)
            result_percent = abs(abs(lbl_percent - pre).sub(100))  # 성공 퍼센티지에 대한 결과 lbl-pre 음수 되는 경우 절대값, 100 뺄셈 절대값
        else:
            lbl_percent = lbl.div(384)
            lbl_percent = lbl_percent.mul(100)
            pre = pre.mul(100)
            result_percent = abs(lbl_percent - pre)  # 오차 퍼센티지에 대한 결과

        result_percent = result_percent[result_percent < 10000]

        logger.info('percent data transform end')

        return result_percent

    def sheet(self, lbl: pd.DataFrame, pre: pd.DataFrame) -> pd.DataFrame:  # 장수 차이에 대한 결과 값
        # lbl = 장수 -> 장수, pre = 소수점 -> 소수점
        pre_sheet = pre.mul(384).round(0)
        result_sheet = abs(lbl - pre_sheet)  # 장수에 대한 결과
        result_sheet = result_sheet[result_sheet < 10000]

        return result_sheet

    def percent_sheet_result(self, outlier, method: pd.DataFrame, *args: str):  # 결과 값을 토대로 std, aver 를 outlier 적용 한 값도 같이 2개 생성
        logger.info('make dataframe average, std with outlier start')

        result = method
        print(result)
        if 'right' in args:
            result_outlier = result[result > outlier]  # outlier 값을 버림 ( nan 으로 만듬 )
        else:
            result_outlier = result[result < outlier]  # outlier 값을 버림 ( nan 으로 만듬 )
        print(result_outlier)
        result_average = result.mean(axis=1)  # 평균
        result_std = result.std(axis=1, ddof=0)  # 표준편차

        result_outlier_average = result_outlier.mean(axis=1)
        result_outlier_std = result_outlier.std(axis=1, ddof=0)

        result_aver_std = pd.DataFrame()  # 환자 측정과 평균,표준편차 df 나눔
        result = result.transpose()  # column, row 전환
        result_aver_std.insert(0, 'Out_Std', result_outlier_std)
        result_aver_std.insert(0, 'Out_Aver', result_outlier_average)
        result_aver_std.insert(0, 'Std', result_std)
        result_aver_std.insert(0, 'Aver', result_average)

        result = round(result, 6)  # 소수점 자리수 6
        result_aver_std = round(result_aver_std, 6)  # 소수점 자리수 6
        result_aver_std = result_aver_std.fillna(0)
        logger.info('make dataframe average, std with outlier end')

        return result, result_aver_std

    def pre_lbl_compare(self, label_loc: str, predict_loc: str):  # 서로 없는 번호 나중에 출력 용

        label_id_list = os.listdir(label_loc)  # 환자 list
        predict_id_list = os.listdir(predict_loc)  # 환자 list
        print(label_id_list, predict_id_list)
        lbl_diff = list(set(label_id_list) - set(predict_id_list))  # 차집합
        pre_diff = list(set(predict_id_list) - set(label_id_list))
        if len(lbl_diff) != 0 or len(pre_diff) != 0:
            logger.error(f'label 폴더에 {pre_diff}가 없습니다. predict 폴더에 {lbl_diff}가 없습니다. ')

    def to_xlsx(self, loc: str, xlsx: str, lbl: pd.DataFrame, pre: pd.DataFrame, percent_outlier: float, sheet_outlier: int):  # 엑셀 생성, 결과값 삽입
        logger.info('make xlsx start')

        writer = pd.ExcelWriter(f'{loc}/{xlsx}', engine='openpyxl')  # pandas 엑셀 작성
        # 결과 값 불러옴(pd.dataframe 구성)
        right_percent_result, self.right_percent_aver_std = self.percent_sheet_result(100-percent_outlier, self.percent(lbl, pre, 'right'), 'right')  # 성공률
        percent_result, self.percent_aver_std = self.percent_sheet_result(percent_outlier, self.percent(lbl, pre))  # 오차율
        sheet_result, self.sheet_aver_std = self.percent_sheet_result(sheet_outlier, self.sheet(lbl, pre))  # 장수 차이

        # df 엑셀에 입력
        right_percent_result.to_excel(writer, startcol=0, startrow=3, sheet_name='Sheet2')
        self.right_percent_aver_std.to_excel(writer, startcol=5, startrow=3, sheet_name='Sheet2')
        percent_result.to_excel(writer, startcol=0, startrow=3, sheet_name='Sheet1')
        self.percent_aver_std.to_excel(writer, startcol=5, startrow=3, sheet_name='Sheet1')
        sheet_result.to_excel(writer, startcol=12, startrow=3, sheet_name='Sheet1')
        self.sheet_aver_std.to_excel(writer, startcol=17, startrow=3, sheet_name='Sheet1')
        writer.close()

        logger.info('make xlsx end')

    def create_img_folder(self, loc: str, file_name: str) -> str:  # 폴더 생성, 위치 값 출력
        img_folder = f'{loc}/{file_name}_graph_image'
        os.mkdir(img_folder)  # 폴더 생성

        return img_folder

    def accept_outlier_error_sh1(self, ws, row, column, error, outlier):  # 오차율, 장수 차이에 색상, none 값 적용
        if ws.cell(row=row, column=column).value is None:
            ws.cell(row=row, column=column).value = 'None'
            ws.cell(row=row, column=column).fill = self.gray_color2

        elif ws.cell(row=row, column=column).value > error:  # error 가 outlier 보다 범위가 크기 떄문에 상위에 있음
            ws.cell(row=row, column=column).fill = self.yellow_color2

            if ws.cell(row=row, column=column).value > outlier:
                ws.cell(row=row, column=column).fill = self.outlier_color

    def accept_outlier_error_sh2(self, ws, row, column, error, outlier):  # 성공률에 색상, none 값 적용
        if ws.cell(row=row, column=column).value is None:
            ws.cell(row=row, column=column).value = 'None'
            ws.cell(row=row, column=column).fill = self.gray_color2

        elif ws.cell(row=row, column=column).value < error:
            ws.cell(row=row, column=column).fill = self.yellow_color2

            if ws.cell(row=row, column=column).value < outlier:
                ws.cell(row=row, column=column).fill = self.outlier_color

    # sheet1 에 스타일 적용
    def sheet1_xlsx_style(self, loc: str, xlsx: str, percent_outlier: float, sheet_outlier: int, percent_error: float, sheet_error: int, image_folder_loc: str,
                          file_name: str):
        logger.info('sheet1 style start')

        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb['Sheet1']

        ws.cell(row=4, column=1).value = 'Patient_ID'
        ws.cell(row=4, column=13).value = 'Patient_ID'

        # patient 결과 outlier, error 적용 색상
        for row in range(5, ws.max_row + 1):

            for column in range(2, 5):  # percent 범위
                self.accept_outlier_error_sh1(ws, row, column, percent_error, percent_outlier)

            for column in range(14, 17):  # sheet 범위
                self.accept_outlier_error_sh1(ws, row, column, sheet_error, sheet_outlier)

            for column in [7, 9]:  # percent aver_std 범위
                if ws.cell(row=row, column=column).value is None:
                    pass
                else:
                    self.accept_outlier_error_sh1(ws, row, column, percent_error, percent_outlier)

            for column in [19, 21]:  # percent aver_std 범위
                if ws.cell(row=row, column=column).value is None:
                    pass
                else:
                    self.accept_outlier_error_sh1(ws, row, column, sheet_error, sheet_outlier)

        # air, hard tissue, soft tissue 색상
        for column in range(1, ws.max_column + 1):
            if ws.cell(row=4, column=column).value is not None and ws.cell(row=4, column=column).value != 'Patient_ID':
                ws.cell(row=4, column=column).fill = self.blue_color

        # # patient_id title 색상  조잡해서 뺌
        # ws.cell(row=4, column=1).fill = self.blue_color2
        # ws.cell(row=4, column=1).border = self.thin_border
        # ws.cell(row=4, column=13).fill = self.blue_color2
        # ws.cell(row=4, column=13).border = self.thin_border

        # 테두리 적용 -> 값이 있는 곳에만 적용
        for row in range(1, ws.max_row + 1):
            for column in range(1, ws.max_column + 1):
                if ws.cell(row=row, column=column).value is not None:
                    ws.cell(row=row, column=column).border = self.thin_border

        # patient_id 색상
        for row in range(5, ws.max_row + 1):
            ws.cell(row=row, column=1).fill = self.yellow_color
            ws.cell(row=row, column=13).fill = self.yellow_color

        # air hts sts 색상 적용
        col_number = [6, 18]
        for i in col_number:
            ws.cell(row=5, column=i).fill = self.air_color
            ws.cell(row=6, column=i).fill = self.hts_color
            ws.cell(row=7, column=i).fill = self.sts_color

        # 이미지 삽입
        self.graph(self.percent_aver_std, 'percent', '', loc, percent_error, file_name)
        self.graph(self.percent_aver_std, 'remove_outlier_percent', 'Out_', loc, percent_error, file_name)
        self.graph(self.sheet_aver_std, 'sheet', '', loc, sheet_error, file_name)
        self.graph(self.sheet_aver_std, 'remove_outlier_sheet', 'Out_', loc, sheet_error, file_name)

        img_list = os.listdir(f'{loc}/{file_name}_graph_image')
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

        logger.info('sheet1 style end')

    def sheet2_xlsx_style(self, loc: str, xlsx: str, percent_outlier: float, percent_error: float, image_folder_loc: str, file_name: str):
        logger.info('sheet2 style start')

        percent_outlier = abs(percent_outlier - 100)
        percent_error = abs(percent_error - 100)

        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb['Sheet2']

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

        self.graph(self.right_percent_aver_std, 'right_percent', '', loc, percent_error, file_name)
        self.graph(self.right_percent_aver_std, 'right_remove_outlier_percent', 'Out_', loc, percent_error, file_name)

        img_list = os.listdir(f'{loc}/{file_name}_graph_image')
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

        logger.info('sheet2 style end')

    def graph(self, df: pd.DataFrame, title: str, outlier: str, loc: str, error_line, file_name: str):  # 그래프 제작
        logger.info(f'{title} - make graph start')
        graph = df
        graph_dict = graph.to_dict('list')  # dataframe list 제작

        fig = plt.figure(figsize=(5, 3))  # Figure 생성 사이즈
        ax = fig.add_subplot()  # Axes 추가
        colors = sns.color_palette('Set3', len(list(graph.index)))  # 색상
        xtick_label_position = list(range(len(list(graph.index))))  # x 축에 글시 넣을 위치

        if 'percent' in title:  # sheet 와 percent 에 따른 축 범위 변경
            plt.ylim([0, 10])
            if 'right' in title:
                plt.ylim([0, 100])

        elif 'sheet' in title:
            plt.ylim([0, 20])

        plt.xticks(xtick_label_position, list(graph.index))  # x 축에 삽입
        plt.axhline(y=float(error_line), color='red', linestyle='--')  # error 라인 그리기
        bars = plt.bar(xtick_label_position, graph_dict[f'{outlier}Aver'], color=colors, edgecolor='black')  # 그래프 생성
        plt.title(title, fontsize=10)  # 타이틀 입력
        plt.errorbar(x=list(graph.index), y=graph_dict[f'{outlier}Aver'], yerr=np.array(graph_dict[f'{outlier}Std']) / 2, color='black', ecolor='black', fmt='.',
                     alpha=0.5, elinewidth=2)  # 에러바 삽입

        for i, b in enumerate(bars):  # 바에 결과 값 추가
            ax.text(b.get_x() + b.get_width() / 2, b.get_height() / 2, graph_dict[f'{outlier}Aver'][i], ha='center', fontsize=10)

        plt.xticks(rotation=0)
        plt.savefig(f'{loc}/{file_name}_graph_image/{title}.png')
        logger.info(f'{title} - make graph end')


if __name__ == "__main__":
    vol = Vol_Template()
    xlsx_name = 'result.xlsx'
    df_lbl = vol.label(r'D:\temp_data\label')
    df_pre = vol.predict(r'D:\temp_data\predict')

    image_folder = vol.create_img_folder(r'D:\temp_data', 'ff')
    vol.to_xlsx(fr'D:\temp_data', xlsx_name, df_lbl, df_pre, 5, 15)
    vol.sheet1_xlsx_style(fr'D:\temp_data', xlsx_name, 5, 13, 3, 7, image_folder, 'ff')
    vol.sheet2_xlsx_style(fr'D:\temp_data', xlsx_name, 5, 3, image_folder, 'ff')
# fr'D:\temp_data\graph_image'
