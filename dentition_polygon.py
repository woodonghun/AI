from PreShin.tooth import Tooth_UI, Tooth


def calculate(lbl: str, pre: str):
    """
        dice, diceloss, iou 계산
        :param lbl: label txt 주소
        :param pre: predict txt 주소
        :return: dice, diceloss, iou 값 => mode 에 따라서 return 값이 달라짐
        """

    lbl_count = 0  # lbl class 개수
    pre_count = 0  # pre class 개수
    intersection_count = 0  # 교집합
    union_count = 0  # 합집합
    # 두 개의 파일을 동시에 열기
    count_lines = 0
    with open(lbl, 'r') as lbl_f, open(pre, 'r') as pre_f:
        lbl_lines = lbl_f.readlines()  # 파일1의 모든 라인 읽기
        pre_lines = pre_f.readlines()  # 파일2의 모든 라인 읽기
        # 두 파일의 라인 개수 확인
        if len(lbl_lines) != len(pre_lines):
            print("경고: 두 파일의 라인 개수가 다릅니다.")
            print(f"label line : {len(lbl_lines)}, Predict line: {len(pre_lines)}")
            return None
        else:
            # 두 파일의 내용을 동시에 읽어 오기
            for lbl_line, pre_line in zip(lbl_lines, pre_lines):
                count_lines += 1
                # 각 파일의 각 줄에 대해 원하는 작업 수행
                lbl_line = lbl_line.strip()  # 공백 제거
                pre_line = pre_line.strip()
                print(f'파일1: {lbl_line}, 파일2: {pre_line}')
                print(f"line : {count_lines}")

                lbl_count += int(lbl_line)
                pre_count += int(pre_line)

                if int(lbl_line) + int(pre_line) > 0:
                    union_count += 1
                    if int(lbl_line) + int(pre_line) > 1:
                        intersection_count += 1

                print(lbl_count, pre_count, intersection_count, union_count)
                dice = intersection_count * 2 / (lbl_count + pre_count)
                iou = intersection_count / union_count

                dice_loss = 1 - dice

                print(dice, iou, dice_loss)

if __name__ == '__main__':
    label = r'C:\Users\3DONS\Desktop\temp\lbl\1#RU1.txt'
    predict = r'C:\Users\3DONS\Desktop\temp\pre\1#RU1.txt'
    calculate(label, predict)
