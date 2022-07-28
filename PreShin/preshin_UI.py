import numpy as np
from PySide2.QtWidgets import QWidget, QPushButton, QLineEdit, QFileDialog, QDialog, QLabel
import pandas as pd

# 엑셀이 비었을 때, 셋팅이 되어 있을 때


class Preshin_UI(QWidget):
    def __init__(self):
        super().__init__()
        self.dialog = QDialog()
        self.initUI()
        self.dialog_close()

    def initUI(self):
        btn_pre_path = QPushButton(self.dialog)
        btn_lbl_path = QPushButton(self.dialog)
        btn_result_path = QPushButton(self.dialog)
        btn_export = QPushButton(self.dialog)

        btn_pre_path.setText('Predict path')
        btn_lbl_path.setText('Label path')
        btn_result_path.setText('result path')
        btn_export.setText('Export excel')

        self.lbl_pre = QLabel(self.dialog)
        self.lbl_lbl = QLabel(self.dialog)
        self.lbl_result = QLabel(self.dialog)

        self.lbl_lbl.setGeometry(150, 20, 400, 30)
        self.lbl_pre.setGeometry(150, 40, 400, 30)
        self.lbl_result.setGeometry(150, 60, 400, 30)

        btn_lbl_path.setGeometry(20, 20, 100, 20)
        btn_pre_path.setGeometry(20, 40, 100, 20)
        btn_result_path.setGeometry(20, 60, 100, 20)
        btn_export.setGeometry(20, 80, 100, 20)

        btn_lbl_path.clicked.connect(self.btn_lbl_clicked)
        btn_result_path.clicked.connect(self.btn_result_clicked)
        btn_export.clicked.connect(self.btn_export_clicked)
        btn_pre_path.clicked.connect(self.btn_pre_clicked)

        edt = QLineEdit(self.dialog)
        edt.setGeometry(100, 100, 100, 100)

        self.dialog.setWindowTitle('AI')
        self.dialog.setGeometry(500, 300, 550, 450)
        self.dialog.exec()

    def btn_pre_clicked(self):
        self.pre_fname = QFileDialog.getOpenFileName(self)
        self.lbl_pre.setText(self.pre_fname[0])

    def btn_lbl_clicked(self):
        self.lbl_fname = QFileDialog.getOpenFileName(self)
        self.lbl_lbl.setText(self.lbl_fname[0])
        name = self.lbl_fname[0].replace('.', '/')
        self.name = name.split('/')

    def btn_result_clicked(self):
        self.result_fname = QFileDialog.getOpenFileName(self)
        self.lbl_result.setText(self.result_fname[0])

    def btn_export_clicked(self):
        # pd.set_option('mode.chained_assignment', None)      # 사본 사용 혼돈 경고 무시
        landmark_name = ['N', 'Sella', 'R FZP', 'L FZP', 'R Or', 'L Or', 'R Po', 'L Po', 'R TFP', 'L TFP', 'R KRP',
                         'L KRP', 'ANS', 'PNS', 'A', 'B', 'Pog', 'Gn', 'Me', 'R CP Point', 'L CP Point',
                         'R Sigmoid notch', 'L Sigmoid notch', 'R Anterior ramal point', 'L Anterior ramal point',
                         'R Post Go', 'L Post Go', 'R Go', 'L Go', 'R Inf Go', 'L Inf Go', 'U1MP', 'R U1CP', 'R U1RP',
                         'L U1CP', 'L U1RP', 'R U3CP', 'R U3RP', 'L U3CP', 'L U3RP', 'R U6CP', 'R U6RP', 'L U6CP',
                         'L U6RP', 'L1MP', 'R L1CP', 'R L1RP', 'L L1CP', 'L L1RP', 'R L3CP', 'R L3RP', 'L L3CP',
                         'L L3RP', 'R L6CP', 'R L6RP', 'L L6CP', 'L L6RP', 'Stomion super', 'Columella	',
                         'Subnasale	', 'Upper lip	', 'S Pogonion', 'R U2CP', 'R U2RP', 'L U2CP', 'L U2RP',
                         'R U4CP', 'R U4RP', 'L U4CP', 'L U4RP', 'R U5CP', 'R U5RP', 'L U5CP', 'L U5RP', 'R U7CP',
                         'R U7RP', 'L U7CP', 'L U7RP', 'R L2CP', 'R L2RP', 'L L2CP', 'L L2RP', 'R L4CP', 'R L4RP',
                         'L L4CP', 'L L4RP', 'R L5CP', 'R L5RP', 'L L5CP', 'L L5RP', 'R L7CP', 'R L7RP', 'L L7CP',
                         'L L7RP', 'R U8CP', 'R U8RP', 'L U8CP', 'L U8RP', 'R L8CP', 'R L8RP', 'L L8CP', 'L L8RP',
                         'S Glabella', 'S Nasion', 'Pronasale', 'Lower lip', 'Mentolabial Sulc', 'Base of Epiglott',
                         'R Cd-L', 'R Cd-M', 'R C Cd-S', 'R GF-L', 'R GF-M', 'R C GF-S', 'R Cd-A', 'R Cd-P', 'R S Cd-S',
                         'R GF-A', 'R GF-P', 'R S GF-S', 'L Cd-L', 'L Cd-M', 'L C Cd-S', 'L GF-L', 'L GF-M', 'L C GF-S',
                         'L Cd-A', 'L Cd-P', 'L S Cd-S', 'L GF-A', 'L GF-P', 'L S GF-S', 'R Ant. Zygoma',
                         'L Ant. Zygoma', 'R Zygion', 'L Zygion', 'R Endocanthion', 'L Endocanthion', 'R Exocanthion',
                         'L Exocanthion', 'R S Ant. Zygoma', 'L S Ant. Zygoma', 'R S Zygion', 'L S Zygion', 'R Alar',
                         'L Alar', 'Labiale Superius', 'R Crista Philtri', 'L Crista Philtri', 'R Cheilion',
                         'L Cheilion', 'Labiale Inferius', 'R S Go', 'L S Go', 'R U4PRP', 'R U5PRP', 'R U6PRP',
                         'R U7PRP', 'L U4PRP', 'L U5PRP', 'L U6PRP', 'L U7PRP']

        label = open(self.lbl_fname[0], "r", encoding="UTF-8")      # label 랜드마크 저장
        lines = label.read()
        lines = lines.replace("\n", ",")
        lines = lines.split(",")
        lines_chunk = [lines[i * 4:(i + 1) * 4] for i in
                       range((len(lines) + 4 - 1) // 4)]

        df = pd.DataFrame(lines_chunk, columns=['landmark_num', 'x', 'y', 'z'])     # label 데이터 프레임
        df['x'] = df['x'].astype(float)
        df['y'] = df['y'].astype(float)
        df['z'] = df['z'].astype(float)
        df['landmark_num'] = df['landmark_num'].astype(int)
        df = df.sort_values(by=['landmark_num'])  # 데이터 정렬

        predict = open(self.pre_fname[0], "r", encoding="UTF-8")        # predict 랜드마크 저장 . C:\woo_project\AI\predict\10.txt,
        lines2 = predict.read()
        lines2 = lines2.replace("\n", ",")
        lines2 = lines2.split(",")
        lines_chunk2 = [lines2[i * 4:(i + 1) * 4] for i in
                        range((len(lines2) + 4 - 1) // 4)]

        df2 = pd.DataFrame(lines_chunk2, columns=['landmark_num', 'x', 'y', 'z'])       # predict 데이터 프래임

        df2['x'] = df2['x'].astype(float)
        df2['y'] = df2['y'].astype(float)
        df2['z'] = df2['z'].astype(float)
        df2['landmark_num'] = df2['landmark_num'].astype(int)
        df2 = df2.sort_values(by=['landmark_num'])  # 데이터 정렬

        result = df.sub(df2)        # 결과값 데이터 프레임
        result['landmark_num'] = df['landmark_num']
        result[self.name[-2]] = (result['x'].pow(2) + result['x'].pow(2) + result['x'].pow(2)).pow(1 / 2)    # name[-2] 파일명 뒤에 있는 환자 번호
        result.drop('x', axis=1, inplace=True)  # 이름부분 뺌
        result.drop('y', axis=1, inplace=True)
        result.drop('z', axis=1, inplace=True)
        # https://trading-for-chicken.tistory.com/43 제곱 거듭 제곱 해야함


        df_sheet = pd.read_excel(self.result_fname[0], sheet_name=0, engine='openpyxl')     # result xlsx가져옴, *** header=2 추가해야됨 ***

        if len(df_sheet.columns) > 2:   # 이미 값이 들어있음
            df_sheet = df_sheet.drop(['landmark_num','landmark_name','aver'], axis=1)   # 값만 추출
            df_sheet = df_sheet.drop(df_sheet.index[len(df_sheet)-1])   # 마지막 행 삭제
            print(df_sheet)
            print(len(df_sheet))
            df_sheet.insert(len(df_sheet.columns), self.name[-2], result[self.name[-2]])    # 값 추가

            sum = df_sheet.sum()
            sum = sum.sum()
            aver = sum/(len(df_sheet.columns)*(len(df_sheet)))

            df_sheet['aver'] = df_sheet.mean(axis=1)
            df.loc[len(df_sheet)-1, 'aver'] = aver
            df_sheet.loc[-1] = df_sheet.mean(axis=0)

            landmark_name.append('aver')  # 맨 아래에 aver 추가
            df_sheet.insert(0, 'landmark_num', df['landmark_num'])
            df_sheet.insert(1, 'landmark_name', landmark_name)
            df_sheet.to_excel(self.result_fname[0], index=False)
            print(df_sheet)

        else:
            df_final = pd.DataFrame()       # 다음에는 새로 쓰지말고 덮어쓰기
            df_final.insert(0, self.name[-2], result[self.name[-2]])    # 새로운 데이터 프레임 첫번째에 추가됨. (0, 이름, 결과)
            df_final['aver'] = df_final.mean(axis=1)    # axis 1 열  마지막 열에 key: aver/ mean 평균 추가
            df_final.loc[-1] = df_final.mean(axis=0)     # axis 0 행  마지막 행에 평균 추가

            landmark_name.append('aver')  # 맨 아래에 aver 추가
            df_final.insert(0, 'landmark_num', df['landmark_num'])
            df_final.insert(1, 'landmark_name', landmark_name)

            df_final.to_excel(self.result_fname[0], index=False)    # C:\woo_project\AI\root 예시 #
            print(df_final)

    def dialog_close(self):
        self.dialog.close()
