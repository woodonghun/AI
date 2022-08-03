import numpy as np
from openpyxl.styles import Border, borders, Protection
import openpyxl
from PySide2.QtWidgets import QWidget, QPushButton, QFileDialog, QDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QMessageBox, QPlainTextEdit, QLineEdit
import pandas as pd
from openpyxl.styles import PatternFill


# 엑셀이 비었을 때, 셋팅이 되어 있을 때
# label, predict 번호 일치 하지 않을때
# 이미 존재하는 id 넣었을 때
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
        self.table.setSortingEnabled(False)  # 정렬기능
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()  # 이것만으로는 checkbox 열 은 잘 조절안됨.
        self.table.setColumnWidth(0, 80)  # checkbox 열 폭 강제 조절.
        self.table.setColumnWidth(1, 80)

        self.table.setItem(0, 0, QTableWidgetItem('Batch size'))
        self.table.setItem(0, 1, QTableWidgetItem(self.value[0]))
        self.table.setItem(1, 0, QTableWidgetItem('Learning rate'))
        self.table.setItem(1, 1, QTableWidgetItem(self.value[1]))
        self.table.setItem(2, 0, QTableWidgetItem('Optimizer'))
        self.table.setItem(2, 1, QTableWidgetItem(self.value[2]))
        self.table.setItem(3, 0, QTableWidgetItem('aug'))
        self.table.setItem(3, 1, QTableWidgetItem(self.value[3]))

        self.table.setHorizontalHeaderLabels(["name", "Value"])
        self.table.setGeometry(20, 180, 180, 145)

        btn_pre_path = QPushButton(self.dialog)
        btn_lbl_path = QPushButton(self.dialog)
        btn_export = QPushButton(self.dialog)

        btn_pre_path.setText('Predict path')
        btn_lbl_path.setText('Label path')
        btn_export.setText('Export excel')

        self.lbl_pre = QLabel(self.dialog)
        self.lbl_lbl = QLabel(self.dialog)
        lbl_error = QLabel(self.dialog)
        self.edt_error = QLineEdit(self.dialog)

        self.lbl_lbl.setGeometry(125, 16, 250, 30)
        self.lbl_pre.setGeometry(125, 41, 250, 30)
        lbl_error.setGeometry(220, 180, 100, 20)
        self.edt_error.setGeometry(220, 205, 50, 20)

        lbl_error.setText('허용 오차 범위')

        self.edt_error.setText(self.value[7])

        btn_lbl_path.setGeometry(20, 20, 100, 20)
        btn_pre_path.setGeometry(20, 45, 100, 20)
        btn_export.setGeometry(20, 335, 180, 30)

        btn_lbl_path.clicked.connect(self.btn_lbl_clicked)
        btn_pre_path.clicked.connect(self.btn_pre_clicked)
        btn_export.clicked.connect(self.btn_export_clicked)

        self.edt = QPlainTextEdit(self.dialog)
        self.edt.setPlainText(self.value[4])
        self.edt.setGeometry(20, 75, 300, 100)

        self.dialog.setWindowTitle('AI')
        self.dialog.setGeometry(500, 300, 370, 420)
        self.dialog.exec()

    def btn_lbl_clicked(self):
        self.landmark_name = []
        self.lbl_id = QFileDialog.getOpenFileName(self, self.tr("Openfile"), self.value[6])
        self.lbl_lbl.setText(self.lbl_id[0])
        name = self.lbl_id[0].replace('.', '/')
        self.lbl_name = name.split('/')

        label = open(self.lbl_id[0], "r", encoding="UTF-8")
        lines = label.read()
        lines = lines.replace("\n", ",")
        lines = lines.split(",")
        self.lines_chunk = [lines[i * 4:(i + 1) * 4] for i in
                            range((len(lines) + 4 - 1) // 4)]  # 4개 단위로 리스트 나눔 (id,x,y,z)
        # landmark 저장
        self.landmark()
        for i in range(len(self.lines_chunk)):
            for j in range(len(self.landmark_key)):
                if self.lines_chunk[i][0] == self.landmark_key[j]:  # landmark key 값 과 lbl에 있는
                    self.landmark_name.append(self.landmark_value[j])


    def btn_pre_clicked(self):
        self.pre_id = QFileDialog.getOpenFileName(self, self.tr("Open file"), self.value[5])  # 창 title, 처음위치
        self.lbl_pre.setText(self.pre_id[0])
        name = self.pre_id[0].replace('.', '/')
        self.pre_name = name.split('/')  # 파일명 뒤에 환자 id를 가져오기 위해 사용

        # predict open 하고 dataframe 에 저장
        predict = open(self.pre_id[0], "r", encoding="UTF-8")
        lines2 = predict.read()
        lines2 = lines2.replace("\n", ",")
        lines2 = lines2.split(",")
        self.lines_chunk2 = [lines2[i * 4:(i + 1) * 4] for i in
                        range((len(lines2) + 4 - 1) // 4)]

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

                    df_sheet = pd.read_excel(self.file_name[0], sheet_name=0, header=3,
                                             engine='openpyxl')  # result xlsx 가져 오면 df_sheet index 와 num 가 동일하게 됨

                    # 엑셀에 id 이미 존재함
                    if self.lbl_name[-2] in df_sheet.columns:
                        self.messagebox("이미 존재 하는 id 입니다")
                    # 엑셀에 id 없음
                    else:
                        self.sheet_style()
                        # label open 하고 dataframe 에 저장
                        df = pd.DataFrame(self.lines_chunk, columns=['landmark_num', 'x', 'y', 'z'])  # label 데이터 프레임

                        df['x'] = df['x'].astype(float)  # 타입 변경 안하면 연산 안됨
                        df['y'] = df['y'].astype(float)
                        df['z'] = df['z'].astype(float)
                        df['landmark_num'] = df['landmark_num'].astype(int)
                        df = df.replace(to_replace=-99999.00, value=pd.NA)      # float 라서 -99999.00 -> 결측치로 변경
                        df = df.sort_values(by='landmark_num')  # 데이터 정렬

                        df2 = pd.DataFrame(self.lines_chunk2, columns=['landmark_num', 'x', 'y', 'z'])  # predict 데이터 프래임

                        df2['x'] = df2['x'].astype(float)
                        df2['y'] = df2['y'].astype(float)
                        df2['z'] = df2['z'].astype(float)
                        df2 = df2.replace(to_replace=-99999.00, value=pd.NA)
                        df2['landmark_num'] = df2['landmark_num'].astype(int)
                        df2 = df2.sort_values(by='landmark_num')

                        result = df.sub(df2)  # 결과값 데이터 프레임 df-df2 ///정렬됨
                        result['landmark_num'] = df['landmark_num']     # 정렬된 df[landmark_num] 넣음

                        df_landmark = pd.DataFrame(self.lines_chunk2, columns=['landmark_num', 'x', 'y', 'z'])  # 랜드마크 번호, 이름 따로 dataframe 생성
                        df_landmark['landmark_num'] = df_landmark['landmark_num'].astype(int)
                        df_landmark.insert(1,'landmark_name', self.landmark_name)
                        df_landmark.drop('x', axis=1, inplace=True)
                        df_landmark.drop('y', axis=1, inplace=True)
                        df_landmark.drop('z', axis=1, inplace=True)
                        df_landmark = df_landmark.sort_values(by='landmark_num')  # 데이터 정렬
                        df_landmark = df_landmark.append({'landmark_num': 0, 'landmark_name':'aver'},ignore_index=True)

                        result[self.lbl_name[-2]] = (result['x'].pow(2) + result['y'].pow(2) + result['z'].pow(2)).pow(
                            1 / 2)  # name[-2] 파일명 뒤에 있는 환자 번호, 두 점 사이의 거리 공식 적용
                        result.drop('x', axis=1, inplace=True)
                        result.drop('y', axis=1, inplace=True)
                        result.drop('z', axis=1, inplace=True)  # x,y,z제거
                        result[self.lbl_name[-2]].loc[-1] = result[self.lbl_name[-2]].mean(axis=0)  # 평균
                        list_land = result[self.lbl_name[-2]].tolist()  # 다음 df에 넣기 위해 list로 만듬

                        if len(df_sheet.columns) > 2:  # 엑셀에 id가 이미 존재
                            # df_sheet는 이미 정렬된 생태
                            print(df_sheet)
                            df_sheet = df_sheet.drop(['landmark_num', 'landmark_name', 'aver'], axis=1)  # id의 값만 추출
                            df_sheet = df_sheet.drop(df_sheet.index[len(df_sheet) - 1])  # 마지막 행 삭제
                            df_sheet.insert(len(df_sheet.columns), self.lbl_name[-2], list_land)  # 값 추가
                            df_sheet.loc[len(df_sheet)] = df_sheet.mean(axis=0)
                            print(df_sheet)

                        # 엑셀에 id 처음 입력
                        else:
                            df_sheet = pd.DataFrame()
                            df_sheet.insert(0, self.lbl_name[-2],
                                            list_land)  # 새로운 데이터 프레임 첫번째에 추가됨. (0, 이름, 결과)
                            df_sheet.loc[len(df_sheet)] = df_sheet.mean(axis=0)

                        data_sum = df_sheet.sum()       # 각각 id의 sum
                        id_sum = data_sum.sum()         # data sum 의 sum
                        data_count = df_sheet.count()   # df_sheet 의 각각 id의 value 개수
                        data_count = data_count.sum()   # id의 value 개수 합
                        aver = id_sum / data_count      # 전체 평균

                        df_sheet['aver'] = df_sheet.mean(axis=1)    # 마지막 열에 평균 추가

                        df_result = pd.concat([df_landmark, df_sheet], axis=1)
                        df_result.iat[-1, -1] = aver    # 마지막 행,열에 전체 aver 추가

                        df_result = df_result.fillna(-99999)    # 결측치에 -99999 입력 -> 엑셀에서 색상 변경시 숫자일 때만 가능 하기 때문
                        df_result.to_excel(self.file_name[0], startcol=0, startrow=3, index=False)  # C:\woo_project\AI\root 예시 #

                        self.sheet_setting()

                        # 시트 2
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

        txt = open('C:/woo_project/AI/landmark.txt', 'r')
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
            landmark_dict[landmark_chunk[i][1]] = landmark_chunk[i][2]

        self.landmark_key = list(landmark_dict.keys())
        self.landmark_value = list(landmark_dict.values())

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
        ws = wb.active

        ws.cell(row=1,
                column=3).value = f'Hyperparameter Batch size = {self.table.item(0, 1).text()}, Learning rate = {self.table.item(1, 1).text()}, optimizer = {self.table.item(2, 1).text()}, aug = {self.table.item(3, 1).text()} '
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

        ws.cell(4, ws.max_column).fill = self.blue_color
        ws.cell(4, ws.max_column).border = self.thin_border
        ws.cell(ws.max_row, 2).fill = self.blue_color
        ws.cell(ws.max_row, 2).border = self.thin_border

        ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=ws.max_column - 1)
        wb.save(filename=self.file_name[0])

    def messagebox(self, i):
        signBox = QMessageBox()
        signBox.setWindowTitle("Warning")
        signBox.setText(i)

        signBox.setIcon(QMessageBox.Information)
        signBox.setStandardButtons(QMessageBox.Ok)
        signBox.exec_()

    def sheet2_value(self):
        Global_g = []
        Cranial_Base = []
        TMJ_1 = []
        TMJ_2 = []
        TMJ_3 = []
        Mx_S = []
        Mn_S_1 = []
        Mn_S_2 = []
        Dental_1 = []
        Dental_2 = []
        Dental_3 = []
        Dental_4 = []
        Dental_5 = []
        Dental_6 = []
        Dental_7 = []
        Dental_8 = []
        Dental_9 = []
        Soft_Tissue_2 = []
        Soft_Tissue_3 = []
        Soft_Tissue_4 = []
        for i in range(2):
            Global_g.append(
            [self.value_sheet2[i][0], self.value_sheet2[i][2], self.value_sheet2[i][3],
             self.value_sheet2[i][6]
                , self.value_sheet2[i][7], self.value_sheet2[i][16], self.value_sheet2[i][27],
             self.value_sheet2[i][28]])
            Cranial_Base.append(
            [self.value_sheet2[i][1], self.value_sheet2[i][4], self.value_sheet2[i][5],
             self.value_sheet2[i][8]
                , self.value_sheet2[i][9], self.value_sheet2[i][12], self.value_sheet2[i][13],
             self.value_sheet2[i][14]])
            TMJ_1.append(
            [self.value_sheet2[i][108], self.value_sheet2[i][109], self.value_sheet2[i][110],
             self.value_sheet2[i][110]
                , self.value_sheet2[i][120], self.value_sheet2[i][121], self.value_sheet2[i][122],
             self.value_sheet2[i][128]])
            TMJ_2.append(
            [self.value_sheet2[i][114], self.value_sheet2[i][115], self.value_sheet2[i][117],
             self.value_sheet2[i][118]
                , self.value_sheet2[i][126], self.value_sheet2[i][127], self.value_sheet2[i][129],
             self.value_sheet2[i][130]])
            TMJ_3.append(
            [self.value_sheet2[i][111], self.value_sheet2[i][112], self.value_sheet2[i][113],
             self.value_sheet2[i][119]
                , self.value_sheet2[i][123], self.value_sheet2[i][124], self.value_sheet2[i][125],
             self.value_sheet2[i][131]])

    def dialog_close(self):
        self.dialog.close()
