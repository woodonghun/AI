import ast
import os

from PySide2.QtCore import Qt
from openpyxl.styles import Border, borders
import openpyxl
from PySide2.QtWidgets import QWidget, QPushButton, QFileDialog, QDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QMessageBox, QPlainTextEdit, QLineEdit
import pandas as pd
from openpyxl.styles import PatternFill

# 엑셀이 비었을 때, 셋팅이 되어 있을 때
# label, predict 번호 일치 하지 않을때
# 이미 존재 하는 id 넣었을 때
# label, predict 하나 라도 없을 때
# 허용 오차 범위


class Preshin_UI(QWidget):
    def __init__(self):
        super().__init__()

        self.dialog = QDialog()
        self.initUI()
        self.dialog_close()

    def initUI(self):

        txt = open("C:/woo_project/AI/comment.txt", 'r')
        value = txt.read()
        self.value = value.split(',')
        txt.close()

        self.table = QTableWidget(4, 2, self.dialog)
        self.table.setSortingEnabled(False)         # 정렬기능
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(0, 80)            # checkbox 열 폭 강제 조절.
        self.table.setColumnWidth(1, 80)

        self.table.setItem(0, 0, QTableWidgetItem('Batch size'))
        self.table.setItem(0, 1, QTableWidgetItem(self.value[0]))
        self.table.setItem(1, 0, QTableWidgetItem('Learning rate'))
        self.table.setItem(1, 1, QTableWidgetItem(self.value[1]))
        self.table.setItem(2, 0, QTableWidgetItem('Optimizer'))
        self.table.setItem(2, 1, QTableWidgetItem(self.value[2]))
        self.table.setItem(3, 0, QTableWidgetItem('Aug'))
        self.table.setItem(3, 1, QTableWidgetItem(self.value[3]))

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
        self.edt_error = QLineEdit(self.dialog)
        self.edt_error.setAlignment(Qt.AlignRight)

        self.lbl_lbl.setGeometry(125, 31, 250, 30)
        self.lbl_pre.setGeometry(125, 56, 250, 30)
        lbl_error.setGeometry(220, 195, 100, 20)
        lbl_mm.setGeometry(270, 220, 50, 20)
        lbl_comment.move(20, 90)
        self.edt_error.setGeometry(220, 220, 50, 20)

        lbl_error.setText('Error Safe Zone')
        lbl_mm.setText('mm')
        lbl_comment.setText('Comment')

        self.edt_error.setText(self.value[7])
        btn_manual.setGeometry(20, 10, 100, 20)
        btn_lbl_path.setGeometry(20, 35, 100, 20)
        btn_pre_path.setGeometry(20, 60, 100, 20)
        btn_export.setGeometry(20, 350, 180, 30)

        btn_lbl_path.clicked.connect(self.btn_lbl_clicked)
        btn_pre_path.clicked.connect(self.btn_pre_clicked)
        btn_export.clicked.connect(self.btn_export_clicked)
        btn_manual.clicked.connect(self.btn_manual_clicked)

        self.edt = QPlainTextEdit(self.dialog)
        self.edt.setPlainText(self.value[4])
        self.edt.setGeometry(20, 105, 300, 80)

        self.dialog.setWindowTitle('AI')
        self.dialog.setGeometry(500, 300, 370, 420)
        self.dialog.exec()

    def btn_lbl_clicked(self):
        self.landmark()

        self.lbl_id = QFileDialog.getOpenFileName(self, self.tr("Openfile"), self.value[6])    # 창 title, 처음 위치
        self.lbl_lbl.setText(self.lbl_id[0])
        name = self.lbl_id[0].replace('.', '/')
        self.lbl_name = name.split('/')    # 파일명 뒤에 환자 id를 가져오기 위해 사용

        label = open(self.lbl_id[0], "r", encoding="UTF-8")
        lines = label.read()
        lines = lines.replace("\n", ",")
        lines = lines.split(",")
        if lines[-1] == '':
            del lines[-1]    # 마지막 빈 칸 제거

        self.lines_chunk = [lines[i * 4:(i + 1) * 4] for i in
                            range((len(lines) + 4 - 1) // 4)]    # 4개 단위로 리스트 나눔 (id,x,y,z)

        lines_chunk_num = []
        for i in range(len(self.lines_chunk)):
            lines_chunk_num.append(self.lines_chunk[i][0])    # 빈 리스트에 label에서 가져온 landmark번호만 따로 저장

        set_lines_chunk_num = set(lines_chunk_num)
        set_landmark_value = set(self.landmark_value)
        empty = set_lines_chunk_num - set_landmark_value
        empty_list = list(empty)    # 집합을 만들어 차집합 으로 landmark.dat 에 없는 num 를 찾음

        self.landmark_name = []    # 빈 리스트 생성

        # landmark 저장
        for i in range(len(self.lines_chunk)):
            for j in range(len(self.landmark_value)):
                if self.lines_chunk[i][0] == self.landmark_value[j]:     # 비교후 같은 값을 landmark_name 에 리스트로 추가
                    self.landmark_name.append(self.landmark_key[j])      # landmark key : id, value : number
                    continue
                if j > len(self.landmark_value)-2:
                    for k in range(len(empty_list)):
                        if empty_list[k] == self.lines_chunk[i][0]:      # 없는 num 와 비교후 같으면 empty 저장
                            self.landmark_name.append('Empty')

    def btn_pre_clicked(self):
        self.pre_id = QFileDialog.getOpenFileName(self, self.tr("Open file"), self.value[5])
        self.lbl_pre.setText(self.pre_id[0])
        name = self.pre_id[0].replace('.', '/')
        self.pre_name = name.split('/')

        # predict open 하고 dataframe 에 저장
        predict = open(self.pre_id[0], "r", encoding="UTF-8")
        lines2 = predict.read()
        lines2 = lines2.replace("\n", ",")
        lines2 = lines2.split(",")
        if lines2[-1] == '':
            del lines2[-1]  # 마지막 빈 칸 제거
        self.lines_chunk2 = [lines2[i * 4:(i + 1) * 4] for i in
                             range((len(lines2) + 4 - 1) // 4)]

    def btn_manual_clicked(self):
        os.startfile('C:/woo_project/AI/AI_manual.pdf') # 메뉴얼 오픈

    def btn_export_clicked(self):
        # lbl, pre 둘다 선택
        if self.lbl_lbl.text() != '' and self.lbl_pre.text() != '':
            # 환자 id 일치 x
            if self.lbl_name[-2] != self.pre_name[-2]:
                self.messagebox("label 과 predict의 환자 id가 일치하지 않습니다.")
            # 환자 id 일치 o
            else:
                self.file_name = QFileDialog.getSaveFileName(self, self.tr("Save Data file"), "C:/woo_project/AI/root",
                                                             self.tr("Data Files(*.xlsx)"))  # 창 title, 위치, 확장자
                # 저장 파일 선택 했을 때
                if self.file_name[0] != '':

                    df_sheet = pd.read_excel(self.file_name[0],sheet_name='Sheet1', header=3,
                                             engine='openpyxl')  # 4행부터 저장해서 header=3 부터 읽음

                    # 엑셀에 id 이미 존재함
                    if self.lbl_name[-2] in df_sheet.columns:
                        self.messagebox("이미 존재 하는 id 입니다")
                    # 엑셀에 id 없음
                    else:
                        self.sheet_style()
                        writer = pd.ExcelWriter(self.file_name[0], engine='openpyxl')
                        df = pd.DataFrame(self.lines_chunk, columns=['Landmark_num', 'x', 'y', 'z'])  # label 데이터 프레임

                        df['x'] = df['x'].astype(float)  # 타입 변경 안하면 연산 안됨
                        df['y'] = df['y'].astype(float)
                        df['z'] = df['z'].astype(float)
                        df['Landmark_num'] = df['Landmark_num'].astype(int)
                        df = df.replace(to_replace=-99999.00, value=pd.NA)  # float 라서 -99999.00 -> 결측치로 변경
                        df = df.sort_values(by='Landmark_num')  # 데이터 정렬

                        df2 = pd.DataFrame(self.lines_chunk2,
                                           columns=['Landmark_num', 'x', 'y', 'z'])  # predict 데이터 프래임

                        df2['x'] = df2['x'].astype(float)
                        df2['y'] = df2['y'].astype(float)
                        df2['z'] = df2['z'].astype(float)
                        df2 = df2.replace(to_replace=-99999.00, value=pd.NA)
                        df2['Landmark_num'] = df2['Landmark_num'].astype(int)
                        df2 = df2.sort_values(by='Landmark_num')

                        result = df.sub(df2)  # 결과값 데이터 프레임 df-df2
                        result['Landmark_num'] = df['Landmark_num']  # result[landmark_num] = 0이되서 정렬된 df[landmark_num] 넣음

                        df_landmark = pd.DataFrame(self.lines_chunk2, columns=['Landmark_num', 'x', 'y',
                                                                               'z'])  # 랜드마크 번호, 이름에 대한 dataframe 생성
                        df_landmark['Landmark_num'] = df_landmark['Landmark_num'].astype(int)
                        df_landmark.insert(1, 'Landmark_name', self.landmark_name)
                        df_landmark.drop('x', axis=1, inplace=True)
                        df_landmark.drop('y', axis=1, inplace=True)
                        df_landmark.drop('z', axis=1, inplace=True)
                        df_landmark = df_landmark.sort_values(by='Landmark_num')  # 데이터 정렬

                        df_landmark = df_landmark.append({'Landmark_num': 0, 'Landmark_name': 'Aver'},
                                                         ignore_index=True)

                        result[self.lbl_name[-2]] = (result['x'].pow(2) + result['y'].pow(2) + result['z'].pow(2)).pow(
                            1 / 2)  # name[-2] 파일명 뒤에 있는 환자 번호, 두 점 사이의 거리 공식 적용
                        result.drop('x', axis=1, inplace=True)
                        result.drop('y', axis=1, inplace=True)
                        result.drop('z', axis=1, inplace=True)  # x,y,z제거
                        result[self.lbl_name[-2]].loc[-1] = result[self.lbl_name[-2]].mean(axis=0)  # 평균 axis = 0 : 행방향, axis =1 : 열방향
                        list_land = result[self.lbl_name[-2]].tolist()  # 다음 df에 넣기 위해 list로 만듬

                        if len(df_sheet.columns) > 2:  # 엑셀에 id가 이미 존재
                            # df_sheet는 이미 정렬된 생태
                            df_sheet = df_sheet.drop(['Landmark_num', 'Landmark_name', 'Aver'], axis=1)  # id의 값만 추출
                            df_sheet = df_sheet.drop(df_sheet.index[len(df_sheet) - 1])  # 마지막 행 삭제
                            df_sheet.insert(len(df_sheet.columns), self.lbl_name[-2], list_land)  # 값 추가
                            df_sheet.loc[len(df_sheet)] = df_sheet.mean(axis=0)

                        # 엑셀에 id 처음 입력
                        else:
                            df_sheet = pd.DataFrame()
                            df_sheet.insert(0, self.lbl_name[-2],
                                            list_land)  # 새로운 데이터 프레임 첫번째에 추가됨. (0, 이름, 결과)
                            df_sheet.loc[len(df_sheet)] = df_sheet.mean(axis=0)

                        data_sum = df_sheet.sum()  # 각각 id의 sum
                        id_sum = data_sum.sum()  # data sum 의 sum
                        data_count = df_sheet.count()  # df_sheet 의 각각 id의 value 개수
                        data_count = data_count.sum()  # id의 value 개수 합
                        aver = id_sum / data_count  # 전체 평균

                        df_sheet['Aver'] = df_sheet.mean(axis=1)  # 마지막 열에 평균 추가

                        df_result = pd.concat([df_landmark, df_sheet], axis=1)  # 랜드마크, value 데이터 프레임 합치기
                        df_result.iat[-1, -1] = aver  # 마지막 행,열에 전체 aver 추가


                        self.df_result = df_result.fillna(-99999)  # 결측치에 -99999 입력 -> 엑셀에서 색상 변경시 숫자일 때만 가능 하기 때문
                        self.df_result.to_excel(writer, startcol=0, startrow=3,
                                           index=False, sheet_name='Sheet1')  # 0,3부터 엑셀로 저장, 인덱스 제거, Sheet1에 저장

                        # 시트 2
                        self.sheet2_value()

                        df_sheet2_name_aver = pd.DataFrame()
                        df_sheet2_name_aver['Name'] = self.df_result['Landmark_num'].astype(str) + '[' + self.df_result[
                            'Landmark_name'] + ']'  # 2[Sella] 형식으로 dataframe 만듬
                        df_sheet2_name_aver['Aver'] = self.df_result['Aver']
                        df_sheet2_name_aver = df_sheet2_name_aver.drop(df_sheet2_name_aver.index[len(df_sheet2_name_aver) - 1]) # 마지막 줄 제거

                        df_sheet2 = self.df_sheet2_name.merge(df_sheet2_name_aver, on='Name', how='left')   # 2[Sella] aver 형태로 합침
                        print(df_sheet2)                                                                    # 빈 칸 Nan 으로 합쳐짐
                        df_sheet2.to_excel(writer, startcol=0, startrow=3,
                                           index=False, sheet_name='Sheet2')

                        writer.save()   # Sheet1, Sheet2 저장
                        self.sheet_setting()
                        ############

                        # table, comment txt 에 저장 해서 다음에 불러올 때 그대로 사용 가능
                        txt = open("C:/woo_project/AI/comment.txt", 'w')

                        batch = self.table.item(0, 1).text()
                        rate = self.table.item(1, 1).text()
                        opti = self.table.item(2, 1).text()
                        aug = self.table.item(3, 1).text()
                        comment = self.edt.toPlainText()
                        pre = self.pre_name[:-2]
                        lbl = self.lbl_name[:-2]
                        error = self.edt_error.text()
                        pre_loc = "/".join(pre)
                        lbl_loc = "/".join(lbl)
                        info = batch, rate, opti, aug, comment, pre_loc, lbl_loc, error
                        txt.write(','.join(info))

                        txt.close()
                else:
                    pass
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
            self.landmark_name_value.append(str(self.landmark_value[i])+'['+str(self.landmark_key[i])+']')

    # 시트 색상, 테두리 스타일
    def sheet_style(self):
        self.yellow_color = PatternFill(start_color='ffffb3', end_color='ffffb3', fill_type='solid')
        self.red_color = PatternFill(start_color='ffcccc', end_color='ffcccc', fill_type='solid')
        self.green_color = PatternFill(start_color='c1f0c1', end_color='c1f0c1', fill_type='solid')
        self.blue_color = PatternFill(start_color='b3d9ff', end_color='b3d9ff', fill_type='solid')
        self.blue_color2 = PatternFill(start_color='ccf5ff', end_color='ccf5ff', fill_type='solid')
        self.gray_color = PatternFill(start_color='bfbfbf', end_color='bfbfbf', fill_type='solid')
        self.gray_color2 = PatternFill(start_color='e0e0eb', end_color='e0e0eb', fill_type='solid')

        self.thin_border = Border(left=borders.Side(style='thin'),
                                  right=borders.Side(style='thin'),
                                  top=borders.Side(style='thin'),
                                  bottom=borders.Side(style='thin'))

    # 시트 색상,테두리 설정
    def sheet_setting(self):
        wb = openpyxl.load_workbook(filename=self.file_name[0])
        ws = wb['Sheet1']

        # table에 default값 출력
        ws.cell(row=1,column=3).value = f'Hyperparameter Batch size = {self.table.item(0, 1).text()}' \
                                        f', Learning rate = {self.table.item(1, 1).text()}' \
                                        f', optimizer = {self.table.item(2, 1).text()}' \
                                        f', aug = {self.table.item(3, 1).text()} '

        # comment에 default값 출력
        ws.cell(row=2, column=3).value = f'comment : {self.edt.toPlainText()}'
        ws.cell(row=3, column=3).value = 'Patient_ID'
        ws.cell(3, 3).fill = self.blue_color
        ws.cell(4, 1).fill = self.blue_color
        ws.cell(4, 2).fill = self.blue_color

        for j in range(len(self.landmark_name)):
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
                if data > float(self.value[7]):  # 특정 수치 이상 이면 색상 변함
                    ws.cell(row=row, column=col).fill = self.red_color
                    ws.cell(row=row, column=col).border = self.thin_border
                elif data == -99999:
                    ws.cell(row=row, column=col).value = ''
                    ws.cell(row=row, column=col).fill = self.gray_color2
                    ws.cell(row=row, column=col).border = self.thin_border
                ws.cell(row=4, column=col).fill = self.gray_color

        for row in range(5, ws.max_row + 1):
            if ws.cell(row=row,column=2).value == 'Empty':
                ws.cell(row=row, column=2).fill = self.red_color


        ws.cell(4, ws.max_column).fill = self.blue_color
        ws.cell(4, ws.max_column).border = self.thin_border
        ws.cell(ws.max_row, 2).fill = self.blue_color
        ws.cell(ws.max_row, 2).border = self.thin_border

        ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=ws.max_column - 1)
        wb.save(filename=self.file_name[0])

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

        sheet2_group = []
        for i in range(len(group_key)):
            for j in range(len(group_value[i])):
                for k in range(len(self.landmark_name_value)):      # landmark_name_value = 2[Sella] 형태
                    name = self.landmark_name_value[k].split('[')
                    if str(group_value[i][j]) == name[0]:   # name[0] = 랜드마크 번호
                        group_value[i][j] = self.landmark_name_value[k] # value 즉 num 가 2[Sella] 형태가 됨

        for i in range(len(group_key)):
            sheet2_group.append(group_key[i])
            for j in range(len(group_value[i])):
                sheet2_group.append(group_value[i][j])
            sheet2_group.append('kkkk')     # group 변경 할 때 마다 빈칸 추가
        self.df_sheet2_name = pd.DataFrame()
        self.df_sheet2_name.insert(0, 'Name', sheet2_group)

    def dialog_close(self):
        self.dialog.close()
