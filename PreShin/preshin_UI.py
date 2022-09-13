import ast
import os

from PySide2.QtCore import Qt
from matplotlib import pyplot as plt
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, borders
import openpyxl
from PySide2.QtWidgets import QWidget, QPushButton, QFileDialog, QDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QMessageBox, QPlainTextEdit, QLineEdit
import pandas as pd
from openpyxl.styles import PatternFill, Font
import seaborn as sns

def btn_manual_clicked():
    os.startfile('C:/woo_project/AI/AI_manual.pdf')  # 메뉴얼 오픈


class Preshin_UI(QWidget):
    def __init__(self):
        super().__init__()

        self.dialog = QDialog()
        self.initUI()
        self.dialog_close()

    def initUI(self):
        batch = '4'
        rate = '2e-4'
        optimizer = 'adam'
        aug = '0'
        comment = 'write comment'
        safe_zone = '3'

        self.table = QTableWidget(4, 2, self.dialog)
        self.table.setSortingEnabled(False)  # 정렬기능
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

        self.lbl_pre = QLabel(self.dialog)
        self.lbl_lbl = QLabel(self.dialog)
        lbl_error = QLabel(self.dialog)
        lbl_mm = QLabel(self.dialog)
        lbl_comment = QLabel(self.dialog)
        lbl_xlsx_name = QLabel(self.dialog)
        lbl_xlsx = QLabel(self.dialog)
        self.edt_error = QLineEdit(self.dialog)
        self.edt_error.setAlignment(Qt.AlignRight)
        self.edt_xlsx_name = QLineEdit(self.dialog)
        self.edt_xlsx_name.setAlignment(Qt.AlignRight)

        self.lbl_lbl.setGeometry(125, 31, 250, 30)
        self.lbl_pre.setGeometry(125, 56, 250, 30)
        lbl_error.setGeometry(220, 195, 100, 20)
        lbl_mm.setGeometry(270, 220, 50, 20)
        lbl_comment.move(20, 90)
        lbl_xlsx_name.move(20, 355)
        lbl_xlsx.move(173, 355)
        self.edt_error.setGeometry(220, 220, 50, 20)
        self.edt_xlsx_name.setGeometry(70, 350, 103, 20)

        lbl_error.setText('Error Safe Zone')
        lbl_mm.setText('mm')
        lbl_comment.setText('Comment')
        lbl_xlsx_name.setText('파일명 : ')
        lbl_xlsx.setText('.xlsx')

        self.edt_error.setText(safe_zone)
        btn_manual.setGeometry(20, 10, 100, 20)
        btn_lbl_path.setGeometry(20, 35, 100, 20)
        btn_pre_path.setGeometry(20, 60, 100, 20)
        btn_export.setGeometry(220, 345, 120, 30)

        btn_lbl_path.clicked.connect(self.btn_lbl_clicked)
        btn_pre_path.clicked.connect(self.btn_pre_clicked)
        btn_export.clicked.connect(self.btn_export_clicked)
        btn_manual.clicked.connect(btn_manual_clicked)

        self.edt = QPlainTextEdit(self.dialog)
        self.edt.setPlainText(comment)
        self.edt.setGeometry(20, 105, 300, 80)

        self.dialog.setWindowTitle('AI')
        self.dialog.setGeometry(500, 300, 370, 420)
        self.dialog.exec()

    def btn_lbl_clicked(self):
        self.landmark()

        self.lbl_id = QFileDialog.getExistingDirectory(self, self.tr("Open file"), 'C:/woo_project/AI/root',
                                                       QFileDialog.ShowDirsOnly)  # 창 title, 주소 나중에 변경
        if self.lbl_id != '':
            self.lbl_lbl.setText(str(self.lbl_id))
            self.lbl_list = os.listdir(str(self.lbl_id))  # 경로에 있는 파일 읽기
            self.compare_landmark()
        else:
            pass

    def btn_pre_clicked(self):
        self.pre_id = QFileDialog.getExistingDirectory(self, self.tr("Open file"), 'C:/woo_project/AI/root',
                                                       QFileDialog.ShowDirsOnly)  # 주소 나중에 변경

        if self.pre_id != '':
            self.lbl_pre.setText(self.pre_id)
            self.pre_list = os.listdir(str(self.pre_id))  # 경로에 있는 파일 읽기
        else:
            pass

    # id 정렬
    def set_pre_lbl(self):
        set_lbl = set(self.lbl_list)
        set_pre = set(self.pre_list)

        id = list(set_lbl & set_pre)  # id 두개다 있는 것만 추려냄
        id = [b.split('.')[0] for b in id]
        id.sort()  # id 정렬 set 함수는 정렬 안되서 나옴
        id.reverse()
        self.id_list = [(j + '.txt') for j in id]

    def compare_landmark(self):
        label = open(str(self.lbl_id + '/' + self.lbl_list[0]), "r", encoding="UTF-8")
        lines = label.read()
        lines = lines.replace("\n", ",")
        lines = lines.split(",")
        if lines[-1] == '':
            del lines[-1]  # 마지막 빈 칸 제거

        self.lines_chunk = [lines[i * 4:(i + 1) * 4] for i in
                            range((len(lines) + 4 - 1) // 4)]

        lines_chunk_num = []
        for i in range(len(self.lines_chunk)):
            lines_chunk_num.append(self.lines_chunk[i][0])  # 빈 리스트에 label에서 가져온 landmark번호만 따로 저장

        set_lines_chunk_num = set(lines_chunk_num)
        set_landmark_value = set(self.landmark_value)
        empty = set_lines_chunk_num - set_landmark_value
        empty_list = list(empty)  # 집합을 만들어 차집합 으로 landmark.dat 에 없는 num 를 찾음

        self.landmark_name = []  # 빈 리스트 생성

        # landmark 저장
        for i in range(len(self.lines_chunk)):
            for j in range(len(self.landmark_value)):
                if self.lines_chunk[i][0] == self.landmark_value[j]:  # 비교후 같은 값을 landmark_name 에 리스트로 추가
                    self.landmark_name.append(self.landmark_key[j])  # landmark key : id, value : number
                    continue
                if j > len(self.landmark_value) - 2:
                    for k in range(len(empty_list)):
                        if empty_list[k] == self.lines_chunk[i][0]:  # 없는 num 와 비교후 같으면 empty 저장
                            self.landmark_name.append('None')

    def error_id(self, loc, name):
        set_lbl = set(self.lbl_list)
        set_pre = set(self.pre_list)
        only_lbl = list(set_lbl - set_pre)
        only_pre = list(set_pre - set_lbl)

        if only_lbl == [] and only_lbl == []:
            pass
        else:
            f = open(f"{loc}/{name}_error.txt", 'w')
            f.write(f'label 폴더에 {only_lbl} : id가 존재하지 않습니다.\n')
            f.write(f'Predict 폴더에 {only_pre} : id가 존재하지 않습니다.')
            f.close()
            self.messagebox("label 또는 predict에 존재하지 않는 id가 있습니다. error.txt를 확인하세요")

    def btn_export_clicked(self):
        # lbl, pre 둘다 선택
        if self.lbl_lbl.text() != '' and self.lbl_pre.text() != '':
            # self.file_name = QFileDialog.getSaveFileName(self, self.tr("Save Data file"), "C:/woo_project/AI/root",
            #                                             self.tr("Data Files(*.xlsx)"))  # 창 title, 위치, 확장자
            # 저장 파일 선택 했을 때

            if self.edt_xlsx_name.text() != '':  # 파일명 입력 했을때
                self.loc_xlsx = QFileDialog.getExistingDirectory(self, self.tr("Open file"), 'C:/woo_project/AI/root',
                                                                 QFileDialog.ShowDirsOnly)

                if self.loc_xlsx != '':  # 폴더 선택 했을때
                    file = os.listdir(self.loc_xlsx)
                    if self.edt_xlsx_name.text() + '.xlsx' not in file:  # 동일한 파일명이 없을때
                        df_sheet = pd.DataFrame()
                        self.set_pre_lbl()  # id 정렬
                        wb = openpyxl.Workbook()
                        self.new_xlsx = self.loc_xlsx + f'/{self.edt_xlsx_name.text()}.xlsx'
                        wb.save(self.new_xlsx)

                        self.sheet2_value()

                        for i in range(len(self.id_list)):
                            name = self.id_list[i].split('/')
                            label = open(str(self.lbl_id + '/' + self.id_list[i]), "r", encoding="UTF-8")  # 파일 하나씩하기 위해선?
                            lines = label.read()
                            lines = lines.replace("\n", ",")
                            lines = lines.split(",")
                            if lines[-1] == '':
                                del lines[-1]  # 마지막 빈 칸 제거

                            lines_chunk = [lines[i * 4:(i + 1) * 4] for i in
                                           range((len(lines) + 4 - 1) // 4)]  # 4개 단위로 리스트 나눔 (id,x,y,z)

                            predict = open(str(self.pre_id + '/' + self.id_list[i]), "r", encoding="UTF-8")
                            lines2 = predict.read()
                            lines2 = lines2.replace("\n", ",")
                            lines2 = lines2.split(",")
                            if lines2[-1] == '':
                                del lines2[-1]  # 마지막 빈 칸 제거
                            lines_chunk2 = [lines2[i * 4:(i + 1) * 4] for i in
                                            range((len(lines2) + 4 - 1) // 4)]

                            self.sheet_style()
                            writer = pd.ExcelWriter(self.new_xlsx, engine='openpyxl')
                            df = pd.DataFrame(lines_chunk, columns=['Landmark_num', 'x', 'y', 'z'])  # label 데이터 프레임

                            df['x'] = df['x'].astype(float)  # 타입 변경 안하면 연산 안됨
                            df['y'] = df['y'].astype(float)
                            df['z'] = df['z'].astype(float)
                            df['Landmark_num'] = df['Landmark_num'].astype(int)
                            df = df.replace(to_replace=-99999.00, value=pd.NA)  # float 라서 -99999.00 -> 결측치로 변경
                            df = df.sort_values(by='Landmark_num')  # 데이터 정렬

                            df2 = pd.DataFrame(lines_chunk2,
                                               columns=['Landmark_num', 'x', 'y', 'z'])  # predict 데이터 프래임

                            df2['x'] = df2['x'].astype(float)
                            df2['y'] = df2['y'].astype(float)
                            df2['z'] = df2['z'].astype(float)
                            df2 = df2.replace(to_replace=-99999.00, value=pd.NA)
                            df2['Landmark_num'] = df2['Landmark_num'].astype(int)
                            df2 = df2.sort_values(by='Landmark_num')

                            result = df.sub(df2)  # 결과값 데이터 프레임 df-df2
                            result['Landmark_num'] = df['Landmark_num']  # result[landmark_num] = 0이되서 정렬된 df[landmark_num] 넣음

                            df_landmark = pd.DataFrame(lines_chunk2, columns=['Landmark_num', 'x', 'y',
                                                                              'z'])  # 랜드마크 번호, 이름에 대한 dataframe 생성
                            df_landmark['Landmark_num'] = df_landmark['Landmark_num'].astype(int)
                            df_landmark.insert(1, 'Landmark_name', self.landmark_name)
                            df_landmark.drop('x', axis=1, inplace=True)
                            df_landmark.drop('y', axis=1, inplace=True)
                            df_landmark.drop('z', axis=1, inplace=True)
                            df_landmark = df_landmark.sort_values(by='Landmark_num')  # 데이터 정렬

                            new_df = pd.DataFrame({'Landmark_num': [0], 'Landmark_name': ['Aver']})

                            df_landmark = pd.concat([df_landmark, new_df], ignore_index=True)

                            result[name[0]] = (result['x'].pow(2) + result['y'].pow(2) + result['z'].pow(2)).pow(
                                1 / 2)  # name[-2] 파일명 뒤에 있는 환자 번호, 두 점 사이의 거리 공식 적용
                            result.drop('x', axis=1, inplace=True)
                            result.drop('y', axis=1, inplace=True)
                            result.drop('z', axis=1, inplace=True)  # x,y,z제거
                            result[name[0]].loc[-1] = result[name[0]].mean(axis=0)  # 평균 axis = 0 : 행방향, axis =1 : 열방향
                            list_land = result[name[0]].tolist()  # 다음 df에 넣기 위해 list로 만듬

                            # 엑셀에 id 처음 입력
                            patient_id = name[0].split('.')[0]

                            df_sheet.insert(0, patient_id, list_land)  # 새로운 데이터 프레임 첫번째에 추가됨. (0, 이름, 결과)

                        data_sum = df_sheet.sum()  # 각각 id의 sum
                        id_sum = data_sum.sum()  # data sum 의 sum
                        data_count = df_sheet.count()  # df_sheet 의 각각 id의 value 개수
                        data_count = data_count.sum()  # id의 value 개수 합
                        aver = id_sum / data_count  # 전체 평균

                        df_sheet['Aver'] = df_sheet.mean(axis=1)  # 마지막 열에 평균 추가

                        df_result = pd.concat([df_landmark, df_sheet], axis=1)  # 랜드마크, value 데이터 프레임 합치기

                        group = sum(self.group_num, [])
                        self.number = [int(i.split('[')[0]) for i in group]
                        self.number.append(0)
                        df_result = df_result.query(f'Landmark_num == {self.number}')

                        # 측정값과 num,name 나누어서 다시 평균 만들기 현재 평균은 group을 제거한 값도 같이 평균한 값임
                        df_result1 = df_result.iloc[:, 0:2]
                        df_result2 = df_result.iloc[:, 2: len(df_result.columns)]
                        df_result2.drop(df_result2.tail(1).index)
                        df_result2.loc[df.shape[0]] = df_result2.mean(axis=0)

                        self.df_result = pd.concat([df_result1, df_result2], axis=1)
                        self.df_result = self.df_result.fillna(-99999)  # 결측치에 -99999 입력 -> 엑셀에서 색상 변경시 숫자일 때만 가능 하기 때문
                        self.df_result.iat[-1, -1] = aver  # 마지막 행,열에 전체 aver 추가
                        self.df_result.reset_index(inplace=True, drop='index')
                        self.df_result.to_excel(writer, startcol=0, startrow=3,
                                                index=False, sheet_name='Sheet1')  # 0,3부터 엑셀로 저장, 인덱스 제거, Sheet1에 저장

                        # 시트 2

                        df_sheet2_name_aver = pd.DataFrame()
                        df_sheet2_name_aver['Name'] = self.df_result['Landmark_num'].astype(str) + '[' + self.df_result[
                            'Landmark_name'] + ']'  # 2[Sella] 형식으로 dataframe 만듬
                        df_sheet2_name_aver['Aver'] = self.df_result['Aver']
                        df_sheet2_name_aver = df_sheet2_name_aver.drop(
                            df_sheet2_name_aver.index[len(df_sheet2_name_aver) - 1])  # 마지막 줄 제거

                        df_sheet2 = self.df_sheet2_name.merge(df_sheet2_name_aver, on='Name',
                                                              how='left')  # 2[Sella] aver 형태로 합침 # 빈 칸 Nan 으로 합쳐짐

                        df_sheet2.to_excel(writer, startcol=0, startrow=3,
                                           index=False, sheet_name='Sheet2')

                        new_df = pd.DataFrame({'Name': ['Total_aver'], 'Aver': [aver]})

                        df_sheet2 = pd.concat([new_df, df_sheet2])
                        df_sheet2 = df_sheet2.fillna('None')
                        df_sheet2.to_excel(writer, startcol=0, startrow=2,
                                           index=False, sheet_name='Sheet2')

                        writer.save()  # Sheet1, Sheet2 저장

                        self.sheet1_setting()
                        self.sheet2_setting()
                        ############
                        self.error_id(self.loc_xlsx, self.edt_xlsx_name.text())
                    else:
                        self.messagebox("동일한 파일명이 존재합니다. 다시 입력하세요")
                else:
                    pass
            else:
                self.messagebox("파일명을 입력하세요")
        # label, predict 선택 되지 않았을 때
        elif self.lbl_lbl.text() == '' or self.lbl_pre.text() == '':
            self.messagebox("label 또는 predict 경로를 확인 하세요.")

    # landmark.dat 구조 변경 후 number - key, name - value 로 지정
    def landmark(self):

        txt = open('C:/woo_project/AI/root/landmark.dat', 'r')
        landmark = txt.read()
        txt.close()
        landmark = landmark.replace(',', ' ')
        landmark = landmark.replace('\t', ',')
        landmark = landmark.replace('\n', ',')
        landmark = landmark.replace('   ', ',')
        landmark = landmark.split(',')

        landmark_chunk = [landmark[i:i + 12] for i in
                          range(0, len(landmark), 12)]

        landmark_dict = {}
        for i in range(len(landmark_chunk) - 1):
            landmark_dict[landmark_chunk[i][2]] = landmark_chunk[i][1]  # id : key , number : value

        self.landmark_name_value = []
        self.landmark_key = list(landmark_dict.keys())
        self.landmark_value = list(landmark_dict.values())
        for i in range(len(self.landmark_key)):
            self.landmark_name_value.append(str(self.landmark_value[i]) + '[' + str(self.landmark_key[i]) + ']')

    # 시트 색상, 테두리 스타일
    def sheet_style(self):
        self.yellow_color = PatternFill(start_color='ffffb3', end_color='ffffb3', fill_type='solid')
        self.red_color = PatternFill(start_color='ffcccc', end_color='ffcccc', fill_type='solid')
        self.green_color = PatternFill(start_color='c1f0c1', end_color='c1f0c1', fill_type='solid')
        self.blue_color = PatternFill(start_color='b3d9ff', end_color='b3d9ff', fill_type='solid')
        self.blue_color2 = PatternFill(start_color='ccf5ff', end_color='ccf5ff', fill_type='solid')
        self.gray_color = PatternFill(start_color='bfbfbf', end_color='bfbfbf', fill_type='solid')
        self.gray_color2 = PatternFill(start_color='e0e0eb', end_color='e0e0eb', fill_type='solid')
        self.orange_color = PatternFill(start_color='ff9900', end_color='ff9900', fill_type='solid')

        self.thin_border = Border(left=borders.Side(style='thin'),
                                  right=borders.Side(style='thin'),
                                  top=borders.Side(style='thin'),
                                  bottom=borders.Side(style='thin'))

    # 시트 색상,테두리 설정
    def sheet1_setting(self):
        wb = openpyxl.load_workbook(filename=self.new_xlsx)
        ws = wb['Sheet1']

        # table에 default값 출력
        ws.cell(row=1, column=3).value = f'Error Safe Zone : {self.edt_error.text()}mm'
        ws.cell(row=1, column=6).value = f'Hyperparameter : Batch size = {self.table.item(0, 1).text()}' \
                                         f', Learning rate = {self.table.item(1, 1).text()}' \
                                         f', optimizer = {self.table.item(2, 1).text()}' \
                                         f', aug = {self.table.item(3, 1).text()} '
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 20

        # comment에 default값 출력
        ws.cell(row=2, column=3).value = f'comment : {self.edt.toPlainText()}'
        ws.cell(row=3, column=3).value = 'Patient_ID'
        ws.cell(3, 3).fill = self.blue_color
        ws.cell(4, 1).fill = self.blue_color
        ws.cell(4, 2).fill = self.blue_color

        for j in range(ws.max_row - 5):
            ws.cell(5 + j, 1).border = self.thin_border
            ws.cell(5 + j, 1).fill = self.green_color
            ws.cell(5 + j, 2).border = self.thin_border
            ws.cell(5 + j, 2).fill = self.yellow_color

        for row in range(5, ws.max_row):
            ws.cell(row=row, column=ws.max_column).fill = self.blue_color2
            ws.cell(row=row, column=ws.max_column).border = self.thin_border

        for col in range(3, ws.max_column):
            ws.cell(row=ws.max_row, column=col).fill = self.blue_color2
            ws.cell(row=ws.max_row, column=col).border = self.thin_border

        for col in range(3, ws.max_column + 1):
            for row in range(5, ws.max_row + 1):
                data = float(ws.cell(row=row, column=col).value)
                if data > float(self.edt_error.text()):  # 특정 수치 이상 이면 색상 변함
                    ws.cell(row=row, column=col).fill = self.red_color
                    ws.cell(row=row, column=col).border = self.thin_border
                elif data == -99999:
                    ws.cell(row=row, column=col).value = ' '
                    ws.cell(row=row, column=col).fill = self.gray_color2
                    ws.cell(row=row, column=col).border = self.thin_border
                ws.cell(row=4, column=col).fill = self.gray_color

        for row in range(5, ws.max_row):
            if ws.cell(row=row, column=2).value == 'None':
                ws.cell(row=row, column=2).fill = self.red_color

        ws.cell(4, ws.max_column).fill = self.blue_color
        ws.cell(4, ws.max_column).border = self.thin_border
        ws.cell(ws.max_row, 2).fill = self.blue_color
        ws.cell(ws.max_row, 2).border = self.thin_border

        ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=ws.max_column - 1)
        wb.save(filename=self.new_xlsx)

    def messagebox(self, i: str):
        signBox = QMessageBox()
        signBox.setWindowTitle("Warning")
        signBox.setText(i)

        signBox.setIcon(QMessageBox.Information)
        signBox.setStandardButtons(QMessageBox.Ok)
        signBox.exec_()

    def sheet2_value(self):

        with open('C:/woo_project/AI/root/group_points_preShin.json', 'r') as inf:
            group = ast.literal_eval(inf.read())  # 그룹 포인트 프리신 dict 로 변환

        group_key = list(group.keys())
        group_value = list(group.values())
        self.group_num = list(group.values())

        sheet2_group = []
        for i in range(len(group_key)):
            for j in range(len(group_value[i])):
                for k in range(len(self.landmark_name_value)):  # landmark_name_value = 2[Sella] 형태
                    name = self.landmark_name_value[k].split('[')
                    if str(group_value[i][j]) == name[0]:  # name[0] = 랜드마크 번호
                        group_value[i][j] = self.landmark_name_value[k]  # value 즉 num 가 2[Sella] 형태가 됨

        for i in range(len(group_value)):
            for j in range(len(group_value[i])):
                if ']' in str(group_value[i][j]):
                    pass
                else:
                    group_value[i][j] = str(group_value[i][j]) + '[None]'

        for i in range(len(group_key)):
            sheet2_group.append(group_key[i])
            for j in range(len(group_value[i])):
                sheet2_group.append(group_value[i][j])

        self.df_sheet2_name = pd.DataFrame()
        self.df_sheet2_name.insert(0, 'Name', sheet2_group)

    def sheet2_setting(self):

        with open('C:/woo_project/AI/root/group_points_preShin.json', 'r') as inf:
            group = ast.literal_eval(inf.read())  # 그룹 포인트 프리신 dict 로 변환

        group_key = list(group.keys())
        group_value = list(group.values())

        wb = openpyxl.load_workbook(filename=self.new_xlsx)
        ws = wb['Sheet2']

        # table에 default값 출력
        ws.cell(row=1, column=3).value = f'Error Safe Zone : {self.edt_error.text()}mm'
        ws.cell(row=1, column=6).value = f'Hyperparameter Batch size = {self.table.item(0, 1).text()}' \
                                         f', Learning rate = {self.table.item(1, 1).text()}' \
                                         f', optimizer = {self.table.item(2, 1).text()}' \
                                         f', aug = {self.table.item(3, 1).text()} '

        ws.cell(row=2, column=3).value = f'comment : {self.edt.toPlainText()}'
        ws['F4'].fill = self.blue_color
        ws['F4'] = 'Total_aver'
        ws['H4'].fill = self.green_color
        ws['H4'] = 'Group_Name'
        ws['J4'].fill = self.yellow_color
        ws['J4'] = 'Landmark'
        ws['L4'].fill = self.red_color
        ws['L4'] = 'Error'

        for row in range(5, ws.max_row+1):
            data = ws.cell(row=row, column=2).value
            if data == -99999:
                ws.cell(row=row, column=2).value = ' '
                ws.cell(row=row, column=2).fill = self.gray_color2
                ws.cell(row=row, column=2).border = self.thin_border
            elif data == 'None':
                ws.cell(row=row, column=2).fill = self.orange_color
                ws.cell(row=row, column=2).border = self.thin_border
            ws.cell(row=row, column=1).fill = self.yellow_color
            ws.cell(row=row, column=1).border = self.thin_border


        ws.cell(row=4, column=1).fill = self.blue_color
        ws.cell(row=4, column=1).border = self.thin_border
        ws.cell(row=4, column=2).fill = self.blue_color
        ws.cell(row=4, column=2).border = self.thin_border

        a = 5  # 시작줄
        b = 6

        # group 색상 초록색
        for i in range(len(group_key)):  # Group 색상
            ws.cell(row=a, column=1).fill = self.green_color
            ws.cell(row=a, column=1).border = self.thin_border
            ws.cell(row=a, column=2).fill = self.green_color
            ws.cell(row=a, column=2).border = self.thin_border
            ws.cell(row=a, column=2).font = Font(bold=True)
            ws.cell(row=a, column=1).font = Font(bold=True)
            for j in range(len(group_value)):  # Group 평균
                num = 0
                result = 0
                for row in range(b, b + len(group_value[i])):    # None, ' ' 을 제외한 합
                    data = ws.cell(row=row, column=2).value
                    if data == 'None' or data == ' ':
                        pass
                    else:
                        result += ws.cell(row=row, column=2).value
                        num += 1

                ws[f'B{a}'] = result / num  # Group 평균 삽입

            a += len(group_value[i]) + 1
            b += len(group_value[i]) + 1

        for row in range(4, ws.max_row):  # 결측치 제외 특정 수치 이상 빨강색
            data = ws.cell(row=row, column=2).value
            if data == 'None' or data == ' ':
                pass
            elif float(data) > float(self.edt_error.text()):
                ws.cell(row=row, column=2).fill = self.red_color
                ws.cell(row=row, column=2).border = self.thin_border

        c = 5
        n = 0
        for h in range(len(group_value)):  # Group 바뀔때 마다 줄 띄우기
            ws.insert_rows(c + n)
            ws.insert_rows(c + n + 1)
            n += 1
            c += len(group_value[h]) + 2

        ws.column_dimensions['A'].width = 23
        ws.column_dimensions['B'].width = 15

        wb.save(filename=self.new_xlsx)
        df = pd.read_excel(self.new_xlsx, sheet_name='Sheet2', header=2, usecols=[0, 1])    # 그래프가 없는 시트2 불러옴 (시트에서 평균값을 만들어서 다시 불러와
        df = df.replace(to_replace=' ', value=-0.0001)  # 결측치 -> -0.0001    오차값은 음수가 없어서 음수값으로 결측치와 존재하지 않은것을 확인한다
        df = df.replace(to_replace='None', value=-0.0002)  # 아에 없는거 -> -0.0002
        df = df.dropna(axis=0)    # 줄 띄워서 생긴 결측치 제거
        df = round(df, 4)  # 소수 자리수
        self.matploblib(df)

    def matploblib(self, df):  # 그래프 생성

        with open('C:/woo_project/AI/root/group_points_preShin.json', 'r') as inf:
            group = ast.literal_eval(inf.read())  # 그룹 포인트 프리신 dict 로 변환

        key = list(group.keys())
        value = list(group.values())
        group_list = []


        for i in range(len(key)):  # 그룹별 value 개수
            group_list.append(1 + len(value[i]))
        graph = df
        graph_dict = graph.to_dict('list')
        graph_value = list(graph_dict.values())

        start_row = 0
        image_insert = 7    # 이미지 삽입 시작 셀

        location = self.loc_xlsx + f'/{self.edt_xlsx_name.text()}_image'
        os.mkdir(location)

        wb = openpyxl.load_workbook(filename=self.new_xlsx)
        ws = wb['Sheet2']

        group_total_name = []
        group_total_value = []
        group_total_name.append(graph_value[0][0])
        group_total_value.append(graph_value[1][0])
        # total_aver 이름, 측정값 추가

        for j in range(len(group_list)):
            group_total_name.append(graph_value[0][1 + start_row])
            group_total_value.append(graph_value[1][1 + start_row])

            group_name = graph_value[0][1 + start_row: 1 + group_list[j] + start_row]
            group_value = graph_value[1][1 + start_row: 1 + group_list[j] + start_row]

            start_row += group_list[j]
            self.bar_style(group_name, group_value, location)

            img = Image(location + f'/{group_name[0]}.png')    # 파일저장
            img.width = 800  # 픽셀 단위 사이즈 변경
            img.height = 225
            ws.add_image(img, f'D{image_insert}')
            image_insert += group_list[j] + 2  # 2칸씩 띄워서 삽입

        # total 이랑 group graph
        plt.figure(figsize=(12, 20))
        plt.xlim([-3, 15])  # 범위
        plt.axvline(x=0, color='black', linestyle='--')
        plt.yticks(fontsize=12)
        plt.xticks(fontsize=12)

        colors = ['#B3D9FF']    # total aver 파랑
        for j in range(len(group_total_name)-1):
            if group_total_value[j+1] >float(self.edt_error.text()):
                colors.append('#FFCCCC')    # error 빨강
            else:
                colors.append('#C1F0C1')    # group 초록
        sns.set_palette(sns.color_palette(colors))
        bar = sns.barplot(x=group_total_value, y=group_total_name, edgecolor = 'black')
        bar.set(title=group_total_name[0])
        for p in bar.patches:  # 바에 내용 추가
            width = p.get_width()
            if width == -0.0001:
                bar.text(-2, p.get_y() + p.get_height() / 2, 'N/A', ha='center', size=10, color='r')
            elif width == -0.0002:
                bar.text(-2, p.get_y() + p.get_height() / 2, 'None', ha='center', size=10, color='orange')
            else:
                bar.text(-2, p.get_y() + p.get_height() / 2, width, ha='center', size=12)

        plt.savefig(location + f'/{group_total_name[0]}.png')    # save랑 show의 위치가 바뀌면 save는 실행되지 않는다, 파일저장
        plt.close()
        img = Image(location + f'/{group_total_name[0]}.png')    # 파일 불러옴
        img.width = 650  # 픽셀 단위 사이즈 변경
        img.height = 1000
        ws.add_image(img, 'O2')

        wb.save(filename=self.new_xlsx)
        # 최대치 ----- 20 ~ 30
        # 소수점 3자리

    def bar_style(self, x, y, location):

        plt.figure(figsize=(13, 3))
        plt.ylim([-3, 15])  # 범위
        plt.axhline(y=0, color='black', linestyle='--')
        plt.xticks(fontsize=8, rotation=-5)
        colors = ['#C1F0C1']    # group 초록색
        for j in range(len(x)-1):
            if float(y[j+1]) >= float(self.edt_error.text()):
                colors.append('#FFCCCC')    # error 빨강
            else:
                colors.append('#FFFFB3')    # 기본 노랑
        sns.set_palette(sns.color_palette(colors))
        bar = sns.barplot(x=x, y=y,edgecolor='black')
        bar.set(title=x[0])
        for p in bar.patches:  # 바에 내용 추가
            height = p.get_height()
            if height == -0.0001:    # 결측치 일때
                bar.text(p.get_x() + p.get_width() / 2., -2, 'N/A', ha='center', size=7, color='r')
            elif height == -0.0002:    # group 에 값이 없을 때
                bar.text(p.get_x() + p.get_width() / 2., -2, 'None', ha='center', size=7, color='orange')
            else:
                bar.text(p.get_x() + p.get_width() / 2., -2, height, ha='center', size=7)

        plt.savefig(location + f'/{x[0]}.png')  # save랑 show의 위치가 바뀌면 save는 실행되지 않는다
        plt.close()

    def dialog_close(self):
        self.dialog.close()
