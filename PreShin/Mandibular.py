import os
import nrrd
import numpy as np
import openpyxl
import pandas as pd
from PySide2.QtCore import Qt

from PySide2.QtWidgets import QWidget, QDialog, QMessageBox, QLabel, QLineEdit, QPushButton, QTableWidgetItem, QTableWidget, QPlainTextEdit, QFileDialog
from matplotlib import pyplot as plt
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Font, Border, borders

from PreShin.loggers import logger

# id 에 해당되는 nrrd 파일 필요 ( class )
# 현재 class 가 몇개로 나누어지는지 nrrd 파일 명 또한 정해진게 없음.

batch = '4'
rate = '2e-4'
optimizer = 'adam'
aug = '0'
comment = 'write comment'

safe_zone_error = '10'
outlier_error = '13'

voxel = '180000'
outlier_voxel = '250000'


def messagebox(text: str, i: str):
    signBox = QMessageBox()
    signBox.setWindowTitle(text)
    signBox.setText(i)

    signBox.setIcon(QMessageBox.Information)
    signBox.setStandardButtons(QMessageBox.Ok)
    signBox.exec_()


class Mandibular_UI(QWidget):
    def __init__(self):
        super().__init__()
        self.dialog = QDialog()
        self.initUI()

    def initUI(self):
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

        lbl_error_rate = QLabel(self.dialog)
        lbl_outlier_rate = QLabel(self.dialog)
        lbl_error_voxel = QLabel(self.dialog)
        lbl_outlier_voxel = QLabel(self.dialog)

        lbl_comment = QLabel(self.dialog)
        lbl_xlsx_name = QLabel(self.dialog)
        lbl_xlsx = QLabel(self.dialog)

        self.edt_error_rate = QLineEdit(self.dialog)
        self.edt_error_rate.setAlignment(Qt.AlignRight)
        self.edt_outlier_rate = QLineEdit(self.dialog)
        self.edt_outlier_rate.setAlignment(Qt.AlignRight)
        self.edt_error_voxel = QLineEdit(self.dialog)
        self.edt_error_voxel.setAlignment(Qt.AlignRight)
        self.edt_outlier_voxel = QLineEdit(self.dialog)
        self.edt_outlier_voxel.setAlignment(Qt.AlignRight)

        self.edt_xlsx_name = QLineEdit(self.dialog)
        self.edt_xlsx_name.setAlignment(Qt.AlignRight)  # 엑셀명

        self.edt_lbl.setGeometry(130, 35, 230, 20)
        self.edt_pre.setGeometry(130, 60, 230, 20)
        lbl_error.setGeometry(220, 200, 100, 20)
        lbl_outlier.setGeometry(220, 269, 100, 20)

        lbl_error_rate.setGeometry(275, 223, 100, 20)
        lbl_outlier_rate.setGeometry(275, 292, 100, 20)
        lbl_error_voxel.setGeometry(275, 245, 100, 20)
        lbl_outlier_voxel.setGeometry(275, 315, 100, 20)

        lbl_comment.move(20, 90)
        lbl_xlsx_name.move(20, 355)
        lbl_xlsx.move(173, 355)

        self.edt_error_rate.setGeometry(220, 223, 50, 20)  # 퍼센티지 에러 입력
        self.edt_outlier_rate.setGeometry(220, 292, 50, 20)  # 퍼센티지 아웃라이어 입력
        self.edt_error_voxel.setGeometry(220, 245, 50, 20)  # voxel 에러 입력
        self.edt_outlier_voxel.setGeometry(220, 315, 50, 20)  # voxel 아웃라이어 입력
        self.edt_xlsx_name.setGeometry(70, 350, 103, 20)

        lbl_error_rate.setText('%')
        lbl_outlier_rate.setText('%')
        lbl_error_voxel.setText('복셀')
        lbl_outlier_voxel.setText('복셀')

        lbl_error.setText('Error Safe Zone')
        lbl_outlier.setText('Remove Outlier')
        lbl_comment.setText('Comment')
        lbl_xlsx_name.setText('파일명 : ')
        lbl_xlsx.setText('.xlsx')

        # 퍼센티지 에서 장수로 변환 했기 때문에 입력도 퍼센티지만 입력하게 함.
        # 장수를 퍼센테지로 변환 할 경우에는 반올림으로 인해서 값의 변동 폭이 있어 그래프가 일치하지 않을 경우가 생김
        self.edt_error_rate.setText(safe_zone_error)
        self.edt_outlier_rate.setText(outlier_error)
        self.edt_error_voxel.setText(voxel)
        self.edt_outlier_voxel.setText(outlier_voxel)

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

    def btn_lbl_clicked(self):
        logger.info('Label Button IN')
        loc = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(), QFileDialog.ShowDirsOnly)  # 창 title, 주소 나중에 변경

        # 폴더 경로 입력
        if loc != '':
            self.edt_lbl.setText(str(loc))
        else:
            self.edt_lbl.setText('')
        logger.info('Label Button OUT')

    def btn_pre_clicked(self):
        logger.info('Predict Button IN')
        loc = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(), QFileDialog.ShowDirsOnly)  # 창 title, 주소 나중에 변경

        # 폴더 경로 입력
        if loc != '':
            self.edt_pre.setText(str(loc))
        else:
            self.edt_pre.setText('')
        logger.info('Predict Button OUT')

    def btn_export_clicked(self):
        # label, predict 경로 없을때, 파일명 입력되지 않았을 때. 동일한 파일명 존재 할 때

        logger.info('Export Button IN')

        if self.edt_lbl.text() != '' and self.edt_pre.text() != '':  # label, predict 경로 입력

            try:  # lbl pre 폴더 경로 확인
                os.listdir(self.edt_pre.text())
                os.listdir(self.edt_lbl.text())
            except FileNotFoundError:
                messagebox('Warning', 'predict 또는 lbl 의 폴더 경로가 올바르지 않습니다.')
                logger.error(" 폴더 경로 에러")
            else:
                if self.edt_xlsx_name.text() != '':  # 파일명 입력 했을때
                    loc_xlsx = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(), QFileDialog.ShowDirsOnly)  # 선택한 경로 str 저장

                    if loc_xlsx != '':  # 폴더 선택 했을때
                        file = os.listdir(loc_xlsx)  # 엑셀 저장 위치에 있는 파일 읽기

                        if fr'{self.edt_xlsx_name.text()}' not in file:  # 동일한 파일명이 없을때
                            os.mkdir(f'{loc_xlsx}/{self.edt_xlsx_name.text()}')  # 엑셀 생성
                            mandibular = Mandibular()

                            mandibular.progress(
                                loc_xlsx=loc_xlsx,
                                xlsx_name=self.edt_xlsx_name.text(),
                                lbl_loc=self.edt_lbl.text(),
                                pre_loc=self.edt_pre.text(),
                                error_rate=int(self.edt_error_rate.text()),
                                outlier_rate=int(self.edt_outlier_rate.text()),
                                error_voxel=int(self.edt_error_voxel.text()),
                                outlier_voxel=int(self.edt_outlier_voxel.text())
                            )

                            # ui 에 있는 comment 삽입
                            self.insert_comment(f'{loc_xlsx}/{self.edt_xlsx_name.text()}', fr'{self.edt_xlsx_name.text()}.xlsx', 'Sheet1')
                            self.insert_comment(f'{loc_xlsx}/{self.edt_xlsx_name.text()}', fr'{self.edt_xlsx_name.text()}.xlsx', 'Sheet2')

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

        logger.info("Export Button OUT")

    def insert_comment(self, loc: str, xlsx: str, sheet: str):  # comment 삽입
        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb[sheet]
        # table 에 default 값 출력
        ws.cell(row=1, column=9).value = f'Hyperparameter Batch size = {self.table.item(0, 1).text()}' \
                                         f', Learning rate = {self.table.item(1, 1).text()}' \
                                         f', optimizer = {self.table.item(2, 1).text()}' \
                                         f', aug = {self.table.item(3, 1).text()} '
        ws.cell(row=2, column=9).value = f'comment : {self.edt.toPlainText()}'

        # sheet1 에는 error rate, sheet, accuracy
        if sheet == 'Sheet1':
            ws.cell(row=1, column=2).value = f'Safe Error rate : {100 - int(self.edt_error_rate.text())} % '
            ws.cell(row=2, column=2).value = f'Remove Outlier Rate : {100 - int(self.edt_outlier_rate.text())} %  '

        elif sheet == 'Sheet2':
            ws.cell(row=1, column=2).value = f'Error Safe Zone : {self.edt_error_voxel.text()} Error safe Zone : {int(self.edt_error_voxel.text()) * 0.3}'
            ws.cell(row=2, column=2).value = f'Remove Outlier voxel : {self.edt_outlier_voxel.text()} Remove Outlier mm : {int(self.edt_outlier_voxel.text()) * 0.3}'

        ws.cell(row=1, column=1).fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')  # 색상 노랑
        ws.cell(row=2, column=1).fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')  # 빨강

        ws.cell(row=1, column=2).font = Font(bold=True)  # 글씨 굵게
        ws.cell(row=2, column=2).font = Font(bold=True)
        ws.cell(row=1, column=9).font = Font(bold=True)
        ws.cell(row=2, column=9).font = Font(bold=True)

        wb.save(filename=f'{loc}/{xlsx}')


class Mandibular:
    def __init__(self):
        super().__init__()
        self.predict_id_list: list = []
        self.label_id_list: list = []
        self.voxel_dict: dict = {}
        self.thin_border = Border(left=borders.Side(style='thin'),
                                  right=borders.Side(style='thin'),
                                  top=borders.Side(style='thin'),
                                  bottom=borders.Side(style='thin'))
        self.blue_color = PatternFill(start_color='b3d9ff', end_color='b3d9ff', fill_type='solid')
        self.yellow_color = PatternFill(start_color='ffffb3', end_color='ffffb3', fill_type='solid')
        self.yellow_color2 = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        self.gray_color2 = PatternFill(start_color='e0e0eb', end_color='e0e0eb', fill_type='solid')
        self.outlier_color = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')

    def label(self, loc: str):
        self.label_id_list = os.listdir(loc)  # 환자 list

    def predict(self, loc: str):
        self.predict_id_list = os.listdir(loc)  # 환자 list


    def compare_lbl_pre(self, lbl_list: list, pre_list: list):
        intersection = list(set(lbl_list) & set(pre_list))  # 두개 다있는 폴더
        complement_lbl = list(set(lbl_list) - set(pre_list))
        complement_pre = list(set(pre_list) - set(lbl_list))

        return intersection, complement_lbl, complement_pre

    def numpy_change(self, lbl_file: str, pre_file: str):
        label_read, label_header = nrrd.read(lbl_file)  # nrrd 파일 numpy 로 읽음
        predict_read, predict_header = nrrd.read(pre_file)

        # 특정값을 곱해서 나중에 pyvista 를 사용하게 되면
        # label, predict 의 합이 총 3가지의 형태로 나타남 -> 3d 상에서 3가지 색상이 나타남
        label_read = label_read * 50  # 붉은색
        predict_read = predict_read * 200  # 파랑색

        total_data = label_read + predict_read  # numpy 의 합

        len_total = len(total_data[total_data > 0])  # 0을 제외한 부분 합집합 = 총 voxel 개수
        intersection_total = len(total_data[total_data == 250])  # 교집합 (성공), 파랑색

        accuracy = intersection_total / len_total  # 성공률
        self.voxel_dict[lbl_file.split('/')[-1]] = len_total
        print(list(self.voxel_dict)[-1])
        return accuracy

    def mk_dataframe(self, lbl_loc: str, lbl_list: list, pre_loc: str, pre_list: list):
        logger.info("Make Dataframe start")

        df_result = pd.DataFrame()
        df_error = pd.DataFrame()

        inter, complement_lbl, complement_pre = self.compare_lbl_pre(lbl_list, pre_list)  # 교집합 id 목록만 가져옴 inter
        for i in inter:
            lbl_nrrd_list = os.listdir(lbl_loc + '/' + i)
            pre_nrrd_list = os.listdir(pre_loc + '/' + i)
            accuracy_list = []
            error_list = []

            if lbl_nrrd_list != pre_nrrd_list:  # pre
                continue

            else:
                for j in range(len(lbl_nrrd_list)):
                    accuracy = self.numpy_change(lbl_loc + '/' + i + '/' + lbl_nrrd_list[j], pre_loc + '/' + i + '/' + pre_nrrd_list[j])
                    accuracy_list.append(accuracy)
                    error = (1 - accuracy) * self.voxel_dict[lbl_nrrd_list[j]]
                    error_list.append(error)

            df_result[i] = accuracy_list
            df_error[i] = error_list

        df_result = df_result * 100

        df_result = df_result.round(4)
        df_error = df_error.round(4)

        logger.info("Make Dataframe end")

        return df_result, df_error
        # inter 에서 목록을 가져와 for 문 돌려서 id 폴더 안에 있는 class 4개 ( 아직 개수 미정 )
        # nrrd 파일 읽어오고 연산 까지 한 뒤에 dataframe 에 저장 한다.

    def dataframe_avr_std(self, result_df: pd.DataFrame, outlier: float, *args):
        """dataframe 평균,표준 편차 제작"""
        if args[0] == 'accuracy':
            result_df_outlier = result_df[result_df > outlier]
        else:
            result_df_outlier = result_df[result_df < outlier]

        result_df_average = result_df.mean(axis=1)
        result_df_std = result_df.std(axis=1, ddof=0)

        result_df_average_outlier = result_df_outlier.mean(axis=1)
        result_df_std_outlier = result_df_outlier.std(axis=1, ddof=0)

        result_aver_std = pd.DataFrame()  # 환자 측정과 평균,표준편차 df 나눔

        result_aver_std['Aver'] = result_df_average
        result_aver_std['Std'] = result_df_std
        result_aver_std['Out_Aver'] = result_df_average_outlier
        result_aver_std['Out_Std'] = result_df_std_outlier

        result_aver_std = result_aver_std.round(4)
        return result_aver_std

    # dataframe outlier, std, aver 모두 적용.
    def remake_df(self, accuracy_result: pd.DataFrame, voxel: pd.DataFrame, outlier: int, voxel_outlier: int):
        df_mm = voxel * 0.3  # 복셀계산

        df_mm_only_aver = self.dataframe_avr_std(df_mm, voxel_outlier * 0.3, '')  # id 세로축 std, aver 만
        df_mm = df_mm.transpose()  # id 세로축
        df_mm_aver_std = self.dataframe_avr_std(df_mm, voxel_outlier * 0.3, '')  # id 세로축 std, aver 만
        df_sum_result_mm = pd.concat([df_mm, df_mm_aver_std], axis=1)  # 최종 합침 id 세로 축 std, aver

        df_voxel_only_aver = self.dataframe_avr_std(voxel, voxel_outlier, '')  # id 세로축 std, aver 만
        df_voxel = voxel.transpose()  # id 세로축
        df_voxel_aver_std = self.dataframe_avr_std(df_voxel, voxel_outlier, '')  # id 세로축 std, aver 만
        df_sum_result_voxel = pd.concat([df_voxel, df_voxel_aver_std], axis=1)  # 최종 합침 id 세로 축 std, aver

        df_only_aver = self.dataframe_avr_std(accuracy_result, 100 - outlier, 'accuracy')  # id 세로축 std, aver 만
        df_accuracy_result = accuracy_result.transpose()  # id 세로축
        df_aver_std = self.dataframe_avr_std(df_accuracy_result, 100 - outlier, 'accuracy')  # id 세로축 std, aver 만
        df_sum_result = pd.concat([df_accuracy_result, df_aver_std], axis=1)  # 최종 합침 id 세로 축 std, aver

        return df_sum_result, df_only_aver, df_sum_result_voxel, df_voxel_only_aver, df_sum_result_mm, df_mm_only_aver

    # accuracy,aver, error, aver, voxel, aver 순서대로 6개
    def to_xlsx(self, loc: str, xlsx: str, *args: pd.DataFrame):
        """엑셀 생성, 결과값 삽입"""
        logger.info('make xlsx start')

        writer = pd.ExcelWriter(f'{loc}/{xlsx}', engine='openpyxl')  # pandas 엑셀 작성

        args[0].to_excel(writer, startcol=0, startrow=3, sheet_name='Sheet1')
        args[1].to_excel(writer, startcol=10, startrow=3, sheet_name='Sheet1')
        args[2].to_excel(writer, startcol=0, startrow=3, sheet_name='Sheet2')
        args[3].to_excel(writer, startcol=10, startrow=3, sheet_name='Sheet2')
        args[4].to_excel(writer, startcol=17, startrow=3, sheet_name='Sheet2')
        args[5].to_excel(writer, startcol=27, startrow=3, sheet_name='Sheet2')

        writer.close()

        logger.info('make xlsx end')

    def create_img_folder(self, loc: str, file_name: str) -> str:  # 폴더 생성, 위치 값 출력
        img_folder = f'{loc}/{file_name}_graph_image'
        os.mkdir(img_folder)  # 폴더 생성

        return img_folder

    def graph(self, df: pd.DataFrame, title: str, outlier: str, loc: str, error_line, file_name: str, ):
        """그래프 제작"""
        logger.info(f'"{title}" - Make Graph Start')
        graph = df
        graph_dict = graph.to_dict('list')  # dataframe list 제작

        fig = plt.figure(figsize=(5, 3))  # Figure 생성 사이즈
        ax = fig.add_subplot()  # Axes 추가
        colors = ['#c1f0c1']  # 초록색
        xtick_label_position = list(range(len(list(graph.index))))  # x 축에 글시 넣을 위치

        if 'accuracy' in title:  # 성공률 일때 적용
            plt.ylim([0, 100])

            for j in range(len(list(graph.index))):
                if float(graph_dict[f'{outlier}Aver'][j]) <= float(error_line):  # error_line 값보다 낮으면 색변환 // 성공률 낮으면~
                    colors[j] = '#FFCCCC'  # error 빨강

        else:
            for j in range(len(list(graph.index))):
                if float(graph_dict[f'{outlier}Aver'][j]) >= float(error_line):  # error_line 값보다 낮으면 색변환 // 오차율 높으면~
                    colors[j] = '#FFCCCC'  # error 빨강

        # if 'accuracy' in title:  # sheet 와 error 에 따른 축 범위 변경
        #     # plt.gca().yaxis.set_major_formatter(mticker.FormatStrFormaater('%i %'))
        # elif 'accuracy' in title:
        #     plt.ylim([0, accuracy_range])
        #
        # elif 'sheet' in title:
        #     plt.ylim([0, sheet_range])

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
        logger.info(f'"{title}" - Make Graph End')

    # sheet1 오차율, 장수 차이에 색상, none 값 적용
    def accept_outlier_error_sh2(self, ws, row, column, error, outlier):
        if ws.cell(row=row, column=column).value is None:
            ws.cell(row=row, column=column).value = 'None'
            ws.cell(row=row, column=column).fill = self.gray_color2
        elif ws.cell(row=row, column=column).value >= error:

            if ws.cell(row=4, column=column).value == 'Std' or ws.cell(row=4, column=column).value == 'Out_Std':
                pass
            else:
                ws.cell(row=row, column=column).fill = self.yellow_color2
                if ws.cell(row=row, column=column).value >= outlier:
                    ws.cell(row=row, column=column).fill = self.outlier_color

    # sheet1 성공률에 색상, none 값 적용
    def accept_outlier_error_sh1(self, ws, row, column, error, outlier):
        if ws.cell(row=row, column=column).value is None:
            ws.cell(row=row, column=column).value = 'None'
            ws.cell(row=row, column=column).fill = self.gray_color2
        elif ws.cell(row=row, column=column).value <= error:

            if ws.cell(row=4, column=column).value == 'Std' or ws.cell(row=4, column=column).value == 'Out_Std':
                pass
            else:
                ws.cell(row=row, column=column).fill = self.yellow_color2
                if ws.cell(row=row, column=column).value <= outlier:
                    ws.cell(row=row, column=column).fill = self.outlier_color

    def sheet1_xlsx_style(self, loc: str, xlsx: str, error_outlier: float, error_rate: float, image_folder_loc: str,
                          file_name: str):
        logger.info('Sheet1 Apply Style Start')

        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb['Sheet1']

        ws.cell(row=4, column=1).value = 'Patient_ID'

        # patient 결과 outlier, error 적용 색상
        for row in range(5, ws.max_row + 1):

            for column in range(2, 7):  # error
                self.accept_outlier_error_sh1(ws, row, column, error_rate, error_outlier)

        # max row 로 하면 전체 row 가 걸려서 안됨.
        for row in [5]:
            for column in range(12, 15):  # error aver_std
                self.accept_outlier_error_sh1(ws, row, column, error_rate, error_outlier)

        # class, aver, std 색상
        for column in range(2, ws.max_column + 1):
            if ws.cell(row=4, column=column).value is not None and ws.cell(row=4, column=column).value != 'Patient_ID':  # 4행 None, patient id 제외
                ws.cell(row=4, column=column).fill = self.blue_color

        # 존재하지 않는 값 제외 하고 테두리 적용
        for row in range(3, ws.max_row + 1):
            for column in range(1, ws.max_column + 1):
                if ws.cell(row=row, column=column).value is not None:
                    ws.cell(row=row, column=column).border = self.thin_border

        # patient_id 색상
        for row in range(5, ws.max_row + 1):
            ws.cell(row=row, column=1).fill = self.yellow_color

        # 이미지 삽입
        img_list = os.listdir(f'{loc}/{file_name}_graph_image')
        for i in img_list:
            if '.png' in i:  # png 파일 일때만
                img = Image(image_folder_loc + f'/{i}')
                if i == 'remove_outlier_accuracy.png':
                    ws.add_image(img, 'k25')
                elif i == 'accuracy.png':
                    ws.add_image(img, 'k11')

        ws.auto_filter.ref = f'B4:B{ws.max_row}'  # 엑셀 필터 적용

        # column 사이즈
        wb.save(filename=f'{loc}/{xlsx}')

        logger.info('Sheet1 Apply Style End')

    def sheet2_xlsx_style(self, loc: str, xlsx: str, error_outlier: float, error_rate: float, image_folder_loc: str,
                          file_name: str):
        logger.info('Sheet2 Apply Style Start')

        wb = openpyxl.load_workbook(filename=f'{loc}/{xlsx}')
        ws = wb['Sheet2']

        ws.cell(row=4, column=1).value = 'Patient_ID'

        # patient 결과 outlier, error 적용 색상
        for row in range(5, ws.max_row + 1):

            for column in range(2, 7):  # error
                self.accept_outlier_error_sh2(ws, row, column, error_rate, error_outlier)

            for column in range(19, 24):  # error
                self.accept_outlier_error_sh2(ws, row, column, error_rate*0.3, error_outlier*0.3)

        # max row 로 하면 전체 row 가 걸려서 안됨.
        for row in [5]:
            for column in range(12, 15):  # error aver_std
                self.accept_outlier_error_sh2(ws, row, column, error_rate, error_outlier)

            for column in range(29, 33):  # error
                self.accept_outlier_error_sh2(ws, row, column, error_rate*0.3, error_outlier*0.3)

        # class, aver, std 색상
        for column in range(2, ws.max_column + 1):
            if ws.cell(row=4, column=column).value is not None and ws.cell(row=4, column=column).value != 'Patient_ID':  # 4행 None, patient id 제외
                ws.cell(row=4, column=column).fill = self.blue_color

        # 존재하지 않는 값 제외 하고 테두리 적용
        for row in range(3, ws.max_row + 1):
            for column in range(1, ws.max_column + 1):
                if ws.cell(row=row, column=column).value is not None:
                    ws.cell(row=row, column=column).border = self.thin_border

        # patient_id 색상
        for row in range(5, ws.max_row + 1):
            ws.cell(row=row, column=1).fill = self.yellow_color

        # 이미지 삽입
        img_list = os.listdir(f'{loc}/{file_name}_graph_image')
        for i in img_list:
            if '.png' in i:  # png 파일 일때만
                img = Image(image_folder_loc + f'/{i}')
                if i == 'remove_outlier_error_voxel.png':
                    ws.add_image(img, 'k25')
                elif i == 'error_mm.png':
                    ws.add_image(img, 'AB11')
                elif i == 'remove_outlier_error_mm.png':
                    ws.add_image(img, 'AB25')
                elif i == 'error_voxel.png':
                    ws.add_image(img, 'k11')

        ws.auto_filter.ref = f'B4:B{ws.max_row}'  # 엑셀 필터 적용

        # column 사이즈
        wb.save(filename=f'{loc}/{xlsx}')

        logger.info('Sheet2 Apply Style End')

    def progress(self, loc_xlsx, xlsx_name, lbl_loc, pre_loc, error_rate, outlier_rate, error_voxel, outlier_voxel):
        logger.info('Progress Start')

        img_folder = self.create_img_folder(f'{loc_xlsx}/{xlsx_name}', xlsx_name)  # image 폴더 생성

        self.label(lbl_loc)
        mn_predict = self.predict(pre_loc)

        df_result_accuracy, df_voxel = self.mk_dataframe(lbl_loc, mn_label, pre_loc, mn_predict)  # 성공률, id가 가로축

        # dataframe 제작
        df_sum_result, df_only_aver, df_sum_result_error, df_error_only_aver, df_sum_result_voxel, df_voxel_only_aver = self.remake_df(
            df_result_accuracy, df_voxel, int(outlier_rate), int(outlier_voxel))

        # 엑셀 제작
        logger.info('Make Xlsx starts')
        self.to_xlsx(f'{loc_xlsx}/{xlsx_name}', fr'{xlsx_name}.xlsx', df_sum_result, df_only_aver,
                     df_sum_result_error, df_error_only_aver, df_sum_result_voxel, df_voxel_only_aver)

        logger.info('Make Xlsx end')


        # 그래프 이미지 제작

        self.graph(df_only_aver, 'accuracy', '',
                   f'{loc_xlsx}/{xlsx_name}', 100 - outlier_rate, xlsx_name)
        self.graph(df_only_aver, 'remove_outlier_accuracy', 'Out_',
                   f'{loc_xlsx}/{xlsx_name}', 100 - outlier_rate, xlsx_name)
        self.graph(df_error_only_aver, 'error_voxel', '',
                   f'{loc_xlsx}/{xlsx_name}', outlier_voxel, xlsx_name)
        self.graph(df_error_only_aver, 'remove_outlier_error_voxel', 'Out_',
                   f'{loc_xlsx}/{xlsx_name}', outlier_voxel, xlsx_name)
        self.graph(df_voxel_only_aver, 'error_mm', '',
                   f'{loc_xlsx}/{xlsx_name}', outlier_voxel * 0.3, xlsx_name)
        self.graph(df_voxel_only_aver, 'remove_outlier_error_mm', 'Out_',
                   f'{loc_xlsx}/{xlsx_name}', outlier_voxel * 0.3, xlsx_name)

        self.sheet1_xlsx_style(f'{loc_xlsx}/{xlsx_name}', f'{xlsx_name}.xlsx',100 - outlier_rate, 100 - error_rate, img_folder, xlsx_name)
        self.sheet2_xlsx_style(f'{loc_xlsx}/{xlsx_name}', f'{xlsx_name}.xlsx', outlier_voxel, error_voxel, img_folder, xlsx_name)

        logger.info('Progress end')

if __name__ == '__main__':
    lbl_loc = r'C:\Users\3DONS\Desktop\sample\label'
    pre_loc = r'C:\Users\3DONS\Desktop\sample\predict'
    xlsx_loc = r'C:\Users\3DONS\Desktop\temp'
    mandibular = Mandibular()
    mn_label = mandibular.label(lbl_loc)
    mn_predict = mandibular.predict(pre_loc)

    df_result_accuracy, df_voxel = mandibular.mk_dataframe(lbl_loc, mn_label, pre_loc, mn_predict)  # 성공률, id가 가로축

    df_sum_result, df_only_aver, df_sum_result_error, df_error_only_aver, df_sum_result_voxel, df_voxel_only_aver = mandibular.remake_df(
        df_result_accuracy, df_voxel, 3, 20000)
    mandibular.to_xlsx(f'xlsx_loc/ㅇㅇㅇ', fr'ddd.xlsx', df_sum_result, df_only_aver,
                       df_sum_result_error, df_error_only_aver, df_sum_result_voxel, df_voxel_only_aver)
