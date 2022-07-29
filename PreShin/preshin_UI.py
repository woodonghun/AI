from openpyxl.styles import Border, borders
import openpyxl
from PySide2.QtWidgets import QWidget, QPushButton, QLineEdit, QFileDialog, QDialog, QLabel, QTableWidget, \
    QTableWidgetItem, QMessageBox, QPlainTextEdit
import pandas as pd


# 엑셀이 비었을 때, 셋팅이 되어 있을 때
# label, predict 번호 일치 하지 않을때
# 이미 존재하는 id 넣었을 때
# 엑셀 위치 잘못 되었을 때
# label, predict 하나 라도 없을 때
# 색상, 셀 칸 범위
from openpyxl.styles import PatternFill


class Preshin_UI(QWidget):
    def __init__(self):
        super().__init__()
        self.dialog = QDialog()
        self.initUI()
        self.dialog_close()

    def initUI(self):

        self.table = QTableWidget(4, 2, self.dialog)
        self.table.setSortingEnabled(False)  # 정렬기능
        self.table.resizeRowsToContents()
        self.table.resizeColumnsToContents()  # 이것만으로는 checkbox 열 은 잘 조절안됨.
        self.table.setColumnWidth(0, 80)  # checkbox 열 폭 강제 조절.
        self.table.setColumnWidth(1, 80)

        self.table.setItem(0, 0, QTableWidgetItem('Batch size'))
        self.table.setItem(0, 1, QTableWidgetItem('4'))
        self.table.setItem(1, 0, QTableWidgetItem('Learning rate'))
        self.table.setItem(1, 1, QTableWidgetItem('2e-4'))
        self.table.setItem(2, 0, QTableWidgetItem('Optimizer'))
        self.table.setItem(2, 1, QTableWidgetItem('adam'))
        self.table.setItem(3, 0, QTableWidgetItem('aug'))
        self.table.setItem(3, 1, QTableWidgetItem('0'))

        self.table.setHorizontalHeaderLabels(["name", "Value"])
        self.table.setGeometry(10, 200, 530, 200)

        btn_pre_path = QPushButton(self.dialog)
        btn_lbl_path = QPushButton(self.dialog)
        btn_export = QPushButton(self.dialog)

        btn_pre_path.setText('Predict path')
        btn_lbl_path.setText('Label path')
        btn_export.setText('Export excel')

        self.lbl_pre = QLabel(self.dialog)
        self.lbl_lbl = QLabel(self.dialog)
        self.lbl_result = QLabel(self.dialog)

        self.lbl_lbl.setGeometry(150, 20, 400, 30)
        self.lbl_pre.setGeometry(150, 40, 400, 30)
        self.lbl_result.setGeometry(150, 60, 400, 30)

        btn_lbl_path.setGeometry(20, 20, 100, 20)
        btn_pre_path.setGeometry(20, 40, 100, 20)
        btn_export.setGeometry(20, 80, 100, 20)

        btn_lbl_path.clicked.connect(self.btn_lbl_clicked)
        btn_pre_path.clicked.connect(self.btn_pre_clicked)
        btn_export.clicked.connect(self.btn_export_clicked)

        self.comment = ''
        self.edt = QPlainTextEdit(self.dialog)
        self.edt.setPlainText(self.comment)
        self.edt.setGeometry(100, 100, 100, 100)

        self.dialog.setWindowTitle('AI')
        self.dialog.setGeometry(500, 300, 550, 450)
        self.dialog.exec()

    def btn_pre_clicked(self):
        self.pre_fname = QFileDialog.getOpenFileName(self, self.tr("Open file"), "C:/woo_project/AI/root/predict")
        self.lbl_pre.setText(self.pre_fname[0])
        name = self.pre_fname[0].replace('.', '/')
        self.pre_name = name.split('/')     # 환자 id 가지고 오려고 씀

    def btn_lbl_clicked(self):
        self.lbl_fname = QFileDialog.getOpenFileName(self, self.tr("Openfile"), "C:/woo_project/AI/root/label")
        self.lbl_lbl.setText(self.lbl_fname[0])
        name = self.lbl_fname[0].replace('.', '/')
        self.lbl_name = name.split('/')

    def btn_export_clicked(self):
        # lbl, pre 둘다 선택
        if self.lbl_lbl.text() != '' and self.lbl_pre.text() != '':
            # 환자 id 일치 x
            if self.lbl_name[-2] != self.pre_name[-2]:
                self.messagebox("label 과 predict의 환자 id가 일치하지 않습니다.")
            # 환자 id 일치 o
            else:
                self.file_name = QFileDialog.getSaveFileName(self, self.tr("Save Data file"), "C:/woo_project/AI/root",
                                                        self.tr("Data Files(*.xlsx)"))  # 저장창 뜨게함 창이름, 위치, 확장자
                # 저장 파일 선택 했을 때
                if self.file_name[0] != '':
                    self.sheet_style()
                    df_sheet = pd.read_excel(self.file_name[0], sheet_name=0, header=3,
                                             engine='openpyxl')  # result xlsx가져옴
                    # 엑셀에 id 이미 존재함
                    if self.lbl_name[-2] in df_sheet.columns:
                        self.messagebox("이미 존재 하는 id 입니다")
                    # 엑셀에 id 없음
                    else:
                        # 랜드마크 이름설정
                        self.landmark_name = ['N', 'Sella', 'R FZP', 'L FZP', 'R Or', 'L Or', 'R Po', 'L Po', 'R TFP', 'L TFP', 'R KRP',
                                         'L KRP', 'ANS', 'PNS', 'A', 'B', 'Pog', 'Gn', 'Me', 'R CP Point', 'L CP Point',
                                         'R Sigmoid notch', 'L Sigmoid notch', 'R Anterior ramal point', 'L Anterior ramal point',
                                         'R Post Go', 'L Post Go', 'R Go', 'L Go', 'R Inf Go', 'L Inf Go', 'U1MP', 'R U1CP',
                                         'R U1RP',
                                         'L U1CP', 'L U1RP', 'R U3CP', 'R U3RP', 'L U3CP', 'L U3RP', 'R U6CP', 'R U6RP', 'L U6CP',
                                         'L U6RP', 'L1MP', 'R L1CP', 'R L1RP', 'L L1CP', 'L L1RP', 'R L3CP', 'R L3RP', 'L '
                                                                                                                       'L3CP',
                                         'L L3RP', 'R L6CP', 'R L6RP', 'L L6CP', 'L L6RP', 'Stomion super', 'Columella	',
                                         'Subnasale	', 'Upper lip	', 'S Pogonion', 'R U2CP', 'R U2RP', 'L U2CP', 'L U2RP',
                                         'R U4CP', 'R U4RP', 'L U4CP', 'L U4RP', 'R U5CP', 'R U5RP', 'L U5CP', 'L U5RP', 'R U7CP',
                                         'R U7RP', 'L U7CP', 'L U7RP', 'R L2CP', 'R L2RP', 'L L2CP', 'L L2RP', 'R L4CP', 'R L4RP',
                                         'L L4CP', 'L L4RP', 'R L5CP', 'R L5RP', 'L L5CP', 'L L5RP', 'R L7CP', 'R L7RP', 'L L7CP',
                                         'L L7RP', 'R U8CP', 'R U8RP', 'L U8CP', 'L U8RP', 'R L8CP', 'R L8RP', 'L L8CP', 'L L8RP',
                                         'S Glabella', 'S Nasion', 'Pronasale', 'Lower lip', 'Mentolabial Sulc', 'Base of Epiglott',
                                         'R Cd-L', 'R Cd-M', 'R C Cd-S', 'R GF-L', 'R GF-M', 'R C GF-S', 'R Cd-A', 'R Cd-P',
                                         'R S Cd-S',
                                         'R GF-A', 'R GF-P', 'R S GF-S', 'L Cd-L', 'L Cd-M', 'L C Cd-S', 'L GF-L', 'L GF-M',
                                         'L C GF-S',
                                         'L Cd-A', 'L Cd-P', 'L S Cd-S', 'L GF-A', 'L GF-P', 'L S GF-S', 'R Ant. Zygoma',
                                         'L Ant. Zygoma', 'R Zygion', 'L Zygion', 'R Endocanthion', 'L Endocanthion',
                                         'R Exocanthion',
                                         'L Exocanthion', 'R S Ant. Zygoma', 'L S Ant. Zygoma', 'R S Zygion', 'L S Zygion',
                                         'R Alar',
                                         'L Alar', 'Labiale Superius', 'R Crista Philtri', 'L Crista Philtri', 'R Cheilion',
                                         'L Cheilion', 'Labiale Inferius', 'R S Go', 'L S Go', 'R U4PRP', 'R U5PRP', 'R U6PRP',
                                         'R U7PRP', 'L U4PRP', 'L U5PRP', 'L U6PRP', 'L U7PRP']

                        # label open 하고 dataframe 에 저장
                        label = open(self.lbl_fname[0], "r", encoding="UTF-8")
                        lines = label.read()
                        lines = lines.replace("\n", ",")
                        lines = lines.split(",")
                        lines_chunk = [lines[i * 4:(i + 1) * 4] for i in
                                       range((len(lines) + 4 - 1) // 4)]    # 4개 단위로 리스트 나눔

                        df = pd.DataFrame(lines_chunk, columns=['landmark_num', 'x', 'y', 'z'])  # label 데이터 프레임
                        df['x'] = df['x'].astype(float)
                        df['y'] = df['y'].astype(float)
                        df['z'] = df['z'].astype(float)
                        df['landmark_num'] = df['landmark_num'].astype(int)
                        df = df.sort_values(by=['landmark_num'])  # 랜드마크랑 일치하기 위해 정렬

                        # predict open 하고 dataframe 에 저장
                        predict = open(self.pre_fname[0], "r",encoding="UTF-8")
                        lines2 = predict.read()
                        lines2 = lines2.replace("\n", ",")
                        lines2 = lines2.split(",")
                        lines_chunk2 = [lines2[i * 4:(i + 1) * 4] for i in
                                        range((len(lines2) + 4 - 1) // 4)]

                        df2 = pd.DataFrame(lines_chunk2, columns=['landmark_num', 'x', 'y', 'z'])  # predict 데이터 프래임

                        df2['x'] = df2['x'].astype(float)
                        df2['y'] = df2['y'].astype(float)
                        df2['z'] = df2['z'].astype(float)
                        df2['landmark_num'] = df2['landmark_num'].astype(int)
                        df2 = df2.sort_values(by=['landmark_num'])  # 데이터 정렬

                        result = df.sub(df2)  # 결과값 데이터 프레임
                        result['landmark_num'] = df['landmark_num']
                        result[self.lbl_name[-2]] = (result['x'].pow(2) + result['x'].pow(2) + result['x'].pow(2)).pow(
                            1 / 2)  # name[-2] 파일명 뒤에 있는 환자 번호, 두 점 사이의 거리 공식 적용
                        result.drop('x', axis=1, inplace=True)  # 이름부분 뺌
                        result.drop('y', axis=1, inplace=True)
                        result.drop('z', axis=1, inplace=True) # x,y,z는 필요없어서 버림

                        if len(df_sheet.columns) > 2:  # 엑셀에 id가 이미 존재
                            self.comment += self.edt.toPlainText()

                            df_sheet = df_sheet.drop(['landmark_num', 'landmark_name', 'aver'], axis=1)  # 값만 추출
                            df_sheet = df_sheet.drop(df_sheet.index[len(df_sheet) - 1])  # 마지막 행 삭제
                            df_sheet.insert(len(df_sheet.columns), self.lbl_name[-2], result[self.lbl_name[-2]])  # 값 추가

                            # 행열의 마지막에 평균값 다시 해서 입력
                            sum = df_sheet.sum()
                            sum = sum.sum()
                            aver = sum / (len(df_sheet.columns) * (len(df_sheet)))

                            df_sheet['aver'] = df_sheet.mean(axis=1)
                            df.loc[len(df_sheet) - 1, 'aver'] = aver
                            df_sheet.loc[-1] = df_sheet.mean(axis=0)

                            # 맨 아래에 aver 추가
                            self.landmark_name.append('aver')

                            list_num = result['landmark_num'].tolist()
                            list_num.append('')

                            df_sheet.insert(0, 'landmark_num', list_num)
                            df_sheet.insert(1, 'landmark_name', self.landmark_name)
                            df_sheet.to_excel(self.file_name[0], startcol=0, startrow=3, index=False)

                            self.sheet_setting()    # 셀 시트 세팅

                        # 엑셀에 id 처음 입력
                        else:
                            df_first = pd.DataFrame()
                            df_first.insert(0, self.lbl_name[-2], result[self.lbl_name[-2]])  # 새로운 데이터 프레임 첫번째에 추가됨. (0, 이름, 결과)
                            df_first['aver'] = df_first.mean(axis=1)  # axis 1 열  마지막 열에 key: aver/ mean 평균 추가
                            df_first.loc[-1] = df_first.mean(axis=0)  # axis 0 행  마지막 행에 평균 추가

                            self.landmark_name.append('aver')  # 맨 아래에 aver 추가

                            list_num = result['landmark_num'].tolist()
                            list_num.append('')

                            df_first.insert(0, 'landmark_num', result['landmark_num'])
                            df_first.insert(1, 'landmark_name', self.landmark_name)

                            df_first.to_excel(self.file_name[0], startcol=0, startrow=3, index=False)  # C:\woo_project\AI\root 예시 #

                            self.sheet_setting()

                            self.comment += self.edt.toPlainText()

                        self.edt.setPlainText(self.comment)

                else:
                    pass
        elif self.lbl_lbl.text() == '' or self.lbl_pre.text() == '':
            self.messagebox("label 또는 predict 경로를 확인하세요.")

    def sheet_style(self):
        self.yellow_color = PatternFill(start_color='ffffb3', end_color='ffffb3', fill_type='solid')
        self.red_color = PatternFill(start_color='ffcccc', end_color='ffcccc', fill_type='solid')
        self.green_color = PatternFill(start_color='c1f0c1', end_color='c1f0c1', fill_type='solid')
        self.blue_color = PatternFill(start_color='b3d9ff', end_color='b3d9ff', fill_type='solid')
        self.blue_color2 = PatternFill(start_color='ccf5ff', end_color='ccf5ff', fill_type='solid')
        self.gray_color = PatternFill(start_color='bfbfbf', end_color='bfbfbf', fill_type='solid')

        self.thin_border = Border(left=borders.Side(style='thin'),
                             right=borders.Side(style='thin'),
                             top=borders.Side(style='thin'),
                             bottom=borders.Side(style='thin'))

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

        for j in range(len(self.landmark_name) - 1):
            ws.cell(5 + j, 1).border = self.thin_border
            ws.cell(5 + j, 1).fill = self.green_color
            ws.cell(5 + j, 2).border = self.thin_border
            ws.cell(5 + j, 2).fill = self.yellow_color

        for row in range(5, ws.max_row):
            ws.cell(row=row, column=ws.max_column).fill = self.blue_color2

        for col in range(3, ws.max_column+1):
            for row in range(5, ws.max_row):
                if float(ws.cell(row=row, column=col).value) > 10:  # 특정 수치 이상 이면 색상 변함
                    ws.cell(row=row, column=col).fill = self.red_color
                    ws.cell(row=row, column=col).border = self.thin_border
                    ws.cell(row=4, column=col).fill = self.gray_color

        ws.cell(4, ws.max_column).fill = self.blue_color

        ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=ws.max_column-1)
        wb.save(filename=self.file_name[0])

    def messagebox(self, i):
        signBox = QMessageBox()
        signBox.setWindowTitle("Warning")
        signBox.setText(i)

        signBox.setIcon(QMessageBox.Information)
        signBox.setStandardButtons(QMessageBox.Ok)
        signBox.exec_()

    def dialog_close(self):
        self.dialog.close()
