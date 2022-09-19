import ast
import os

import openpyxl
import pandas as pd
from PySide2.QtWidgets import QFileDialog

import PreShin.preshin_UI
from PreShin.preshin_UI import average, messagebox
from PreShin.loggers import logger

class PreShin_UI_2d(PreShin.preshin_UI.PreShin_UI):
    def __init__(self):
        super().__init__()

    # [id, x, y] 형태 list로 만듬
    def landmark_id_format_change(self, loc, id_list):
        label = open(str(loc + '/' + id_list), "r", encoding="UTF-8")
        id_format = label.readlines()
        lines = []
        for line in id_format:
            line = line.replace("\n", "")
            line = line.split(",")
            if len(line) != 3:
                logger.error('id landmark format error : [id, x, y]')
                logger.error(line)
            else:
                lines.append(line)

        label.close()
        return lines

    def compare_landmark(self):
        # label 폴더의 제일 처음 환자 id를 읽음
        # landmark, x, y형태
        lines_chunk = self.landmark_id_format_change(self.lbl_id, self.lbl_list[0])

        # landmark 번호만 따로 저장
        lines_chunk_num = []
        for i in range(len(lines_chunk)):
            lines_chunk_num.append(lines_chunk[i][0])

        set_lines_chunk_num = set(lines_chunk_num)
        set_landmark_value = set(self.landmark_value)
        empty = set_lines_chunk_num - set_landmark_value
        empty_list = list(empty)  # 집합을 만들어 차집합 으로 landmark.dat 에 없는 num 를 찾음

        self.landmark_name = []  # 빈 리스트 생성

        # landmark 저장
        for i in range(len(lines_chunk)):
            for j in range(len(self.landmark_value)):
                if lines_chunk[i][0] == self.landmark_value[j]:  # 비교후 같은 값을 landmark_name 에 리스트로 추가
                    self.landmark_name.append(self.landmark_key[j])  # landmark key : id, value : number
                    continue
                if j > len(self.landmark_value) - 2:
                    for k in range(len(empty_list)):
                        if empty_list[k] == lines_chunk[i][0]:  # 없는 num 와 비교후 같으면 empty 저장
                            self.landmark_name.append('None')

    def id_dataframe(self, lines_chunk):
        df = pd.DataFrame(lines_chunk, columns=['Landmark_num', 'x', 'y'])  # label 데이터 프레임
        df['x'] = df['x'].astype(float)  # 타입 변경 안하면 연산 안됨
        df['y'] = df['y'].astype(float)
        df['Landmark_num'] = df['Landmark_num'].astype(int)
        df = df[df >= 0]

        df = df.sort_values(by='Landmark_num')  # 데이터 정렬
        return df

    # x,y 제거
    def drop_landmark(self,df):
        df.drop('x', axis=1, inplace=True)
        df.drop('y', axis=1, inplace=True)

    def btn_export_clicked(self):
        # lbl, pre 둘다 선택
        if self.lbl_lbl.text() != '' and self.lbl_pre.text() != '':

            if self.edt_xlsx_name.text() != '':  # 파일명 입력 했을때

                loc_xlsx = QFileDialog.getExistingDirectory(self, "Open file", os.getcwd(),
                                                            QFileDialog.ShowDirsOnly)
                if loc_xlsx != '':  # 폴더 선택 했을때
                    file = os.listdir(loc_xlsx)  # 엑셀 저장 위치에 있는 파일 읽기
                    if f'{self.edt_xlsx_name.text()}_folder' not in file:  # 동일한 파일명이 없을때

                        self.loc_xlsx = loc_xlsx + f'/{self.edt_xlsx_name.text()}_folder'
                        os.mkdir(self.loc_xlsx)

                        df_sheet = pd.DataFrame()
                        self.set_pre_lbl()  # id 정렬
                        wb = openpyxl.Workbook()
                        self.new_xlsx = self.loc_xlsx + f'/{self.edt_xlsx_name.text()}.xlsx'
                        self.new_xlsx_ouliter = self.loc_xlsx + f'/{self.edt_xlsx_name.text()}_outlier.xlsx'
                        wb.save(self.new_xlsx)

                        self.sheet2_value()  # sheet2 landmark name 설정

                        for i in range(len(self.id_list)):  # 환자 수 만큼 만들고 df합침
                            name = self.id_list[i].split('/')

                            lbl_chunk = self.landmark_id_format_change(self.lbl_id, self.id_list[i])  # 3개 단위로 리스트 나눔 (id,x,y)
                            pre_chunk = self.landmark_id_format_change(self.pre_id, self.id_list[i])

                            df_lbl = self.id_dataframe(lbl_chunk)
                            df_pre = self.id_dataframe(pre_chunk)

                            result = df_lbl.sub(df_pre)  # 결과값 데이터 프레임 df-df2
                            result['Landmark_num'] = df_lbl['Landmark_num']  # result[landmark_num] = 0이되서 정렬된 df[landmark_num] 넣음
                            df_landmark = pd.DataFrame(pre_chunk, columns=['Landmark_num', 'x', 'y'])  # 랜드마크 번호, 이름에 대한 dataframe 생성 2D
                            df_landmark['Landmark_num'] = df_landmark['Landmark_num'].astype(int)
                            df_landmark.insert(1, 'Landmark_name', self.landmark_name)
                            self.drop_landmark(df_landmark)
                            df_landmark = df_landmark.sort_values(by='Landmark_num')  # 데이터 정렬

                            new_df = pd.DataFrame({'Landmark_num': [0], 'Landmark_name': ['Aver']})

                            df_landmark = pd.concat([df_landmark, new_df], ignore_index=True)

                            result[name[0]] = (result['x'].pow(2) + result['y'].pow(2)).pow(
                                1 / 2)  # name[-2] 파일명 뒤에 있는 환자 번호, 두 점 사이의 거리 공식 적용 2D
                            self.drop_landmark(result)
                            result[name[0]].loc[-1] = result[name[0]].mean(axis=0)  # 평균 axis = 0 : 행방향, axis =1 : 열방향
                            list_land = result[name[0]].tolist()  # 다음 df에 넣기 위해 list로 만듬

                            # 엑셀에 id 처음 입력
                            patient_id = name[0].split('.')[0]

                            df_sheet.insert(0, patient_id, list_land)  # 새로운 데이터 프레임 첫번째에 추가됨. (0, 이름, 결과)
                        # outlier 의 수치 이하의 값만 출력 후 평균값 만들어서 landmark 와 합침
                        df_sheet_outlier = df_sheet[df_sheet < float(self.edt_outlier.text())]
                        aver_outlier = average(df_sheet_outlier)
                        df_sheet_outlier['Aver'] = df_sheet_outlier.mean(axis=1)
                        df_result_outlier = pd.concat([df_landmark, df_sheet_outlier], axis=1)

                        # 평균값 만들어서 landmark 와 합침
                        aver = average(df_sheet)
                        df_sheet['Aver'] = df_sheet.mean(axis=1)  # 마지막 열에 평균 추가
                        df_result = pd.concat([df_landmark, df_sheet], axis=1)  # 랜드마크, value 데이터 프레임 합치기

                        group = sum(self.group_num, [])
                        self.number = [int(i.split('[')[0]) for i in group]
                        self.number.append(0)

                        df_result = df_result.query(f'Landmark_num == {self.number}')
                        df_result_outlier = df_result_outlier.query(f'Landmark_num == {self.number}')

                        # 측정값과 num,name 나누어서 다시 평균 만들기 현재 평균은 group을 제거한 값도 같이 평균한 값임
                        df_result1 = df_result.iloc[:, 0:2]
                        df_result2 = df_result.iloc[:, 2: len(df_result.columns)]
                        df_result2.drop(df_result2.tail(1).index)
                        df_result2.loc[df_lbl.shape[0]] = df_result2.mean(axis=0)

                        # outlier 세팅
                        df_result2_outlier = df_result_outlier.iloc[:, 2: len(df_result_outlier.columns)]
                        df_result2_outlier.drop(df_result_outlier.tail(1).index)
                        df_result2_outlier.loc[df_lbl.shape[0]] = df_result2_outlier.mean(axis=0)

                        # 기본
                        self.df_result = pd.concat([df_result1, df_result2], axis=1)
                        self.df_result = self.df_result.fillna(-99999)  # 결측치에 -99999 입력 -> 엑셀에서 색상 변경시 숫자일 때만 가능 하기 때문
                        self.df_result.iat[-1, -1] = aver  # 마지막 행,열에 전체 aver 추가
                        self.df_result.reset_index(inplace=True, drop='index')

                        self.sheet_color()

                        # 엑셀
                        writer = pd.ExcelWriter(self.new_xlsx, engine='openpyxl')
                        self.df_result.to_excel(writer, startcol=0, startrow=3,
                                                index=False, sheet_name='Sheet1')  # 0,3부터 엑셀로 저장, 인덱스 제거, Sheet1에 저장

                        # outlier 세팅
                        self.df_result_outlier = pd.concat([df_result1, df_result2_outlier], axis=1)
                        self.df_result_outlier = self.df_result_outlier.fillna(-99999)
                        self.df_result_outlier.iat[-1, -1] = aver_outlier
                        self.df_result_outlier.reset_index(inplace=True, drop='index')

                        # outlier 엑셀
                        writer_outlier = pd.ExcelWriter(self.new_xlsx_ouliter, engine='openpyxl')
                        self.df_result_outlier.to_excel(writer_outlier, startcol=0, startrow=3,
                                                        index=False, sheet_name='Sheet1')  # 0,3부터 엑셀로 저장, 인덱스 제거, Sheet1에 저장

                        # 시트 2

                        self.sheet2(self.df_result, writer, aver)
                        self.sheet2(self.df_result_outlier, writer_outlier, aver_outlier)

                        self.sheet1_setting(self.new_xlsx)
                        self.sheet1_setting(self.new_xlsx_ouliter)
                        self.sheet2_setting(self.new_xlsx)
                        self.sheet2_setting(self.new_xlsx_ouliter)
                        ############
                        self.error_id(self.loc_xlsx, self.edt_xlsx_name.text())
                    else:
                        messagebox("동일한 파일명이 존재합니다. 다시 입력하세요")
                        logger.error("same file name exist")
                else:
                    pass
            else:
                messagebox("파일명을 입력하세요")
                logger.error("no file name")
        # label, predict 선택 되지 않았을 때
        elif self.lbl_lbl.text() == '' or self.lbl_pre.text() == '':
            messagebox("label 또는 predict 경로를 확인 하세요.")

        logger.info("btn_export out")

    def open_json(self):

        with open(f'{os.getcwd()}/group_points_preShin_2D.json', 'r') as inf:  # group : { landmark 번호, ...}
            group = ast.literal_eval(inf.read())  # 그룹 포인트 프리신을 dict 로 변환

        return group
