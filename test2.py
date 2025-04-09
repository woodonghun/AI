import os

import SimpleITK as sitk
import numpy as np

""" 하나의 nrrd 파일에 2개의 label 이 있을경우 분리하는 코드"""

original_folder_path = r'C:\Users\3DONS\Desktop\predict\20231219_predict'  # 분리할 nrrd 를 모아놓은 폴더
create_split_nrrd_path = r'C:\Users\3DONS\Desktop\predict\split_predict'  # 분리될 nrrd 를 생성할 폴더위치 - md, mx 폴더가 생성되며 아래에 파일이 생성된다
create_in_one_folder = r''  # 하나의 폴더에 분리하고 싶을 경우 경로를 지정한다. ex) id_1, id_2 파일이 생성됨


def split_nrrd(origin_folder_path: str, create_folder_path: str, all_in_one: str):
    """
    
    :param origin_folder_path:
    :param create_folder_path:
    :param all_in_one:
    :return:
    """
    file_list = os.listdir(origin_folder_path)
    if not os.path.exists(create_folder_path + '/' + "md"):
        os.makedirs(create_folder_path + '/' + "md")
    if not os.path.exists(create_folder_path + '/' + "mx"):
        os.makedirs(create_folder_path + '/' + "mx")
    for i in file_list:
        patient = i
        print(f'아이디 : {patient}')

        image = sitk.ReadImage(origin_folder_path + '/' + f'{patient}')

        # SimpleITK 이미지를 넘파이 배열로 변환
        img_array = sitk.GetArrayFromImage(image)

        # 각 레이블 값의 출현 빈도 계산
        label_0_count = np.sum(img_array == 0)
        label_1_count = np.sum(img_array == 1)
        label_2_count = np.sum(img_array == 2)

        print("레이블 0의 개수:", label_0_count)
        print("레이블 1의 개수:", label_1_count)
        print("레이블 2의 개수:", label_2_count)

        # 레이블 값 1과 2를 가진 부분을 따로 추출
        label_1_image = np.zeros_like(img_array)
        label_1_image[img_array == 1] = 1  # 레이블 값 1만 포함하는 이미지

        label_2_image = np.zeros_like(img_array)
        label_2_image[img_array == 2] = 2  # 레이블 값 2만 포함하는 이미지

        # 겹치는 부분에 대해 두 레이블 값을 모두 할당
        overlap_mask = np.logical_and(img_array == 1, img_array == 2)
        label_1_image[overlap_mask] = 1
        label_2_image[overlap_mask] = 2

        # 각각의 넘파이 배열을 SimpleITK 이미지로 변환
        label_1_sitk_image = sitk.GetImageFromArray(label_1_image)
        label_1_sitk_image.CopyInformation(image)

        label_2_sitk_image = sitk.GetImageFromArray(label_2_image)
        label_2_sitk_image.CopyInformation(image)

        sitk.WriteImage(label_1_sitk_image, create_folder_path + '/' + "md" + '/' + patient)
        sitk.WriteImage(label_2_sitk_image, create_folder_path + '/' + "mx" + '/' + patient)
        if create_in_one_folder != '':
            print('all_in_one')
            sitk.WriteImage(label_1_sitk_image, all_in_one + '/' + patient.split('.')[0] + '_1.nrrd')
            sitk.WriteImage(label_2_sitk_image, all_in_one + '/' + patient.split('.')[0] + '_2.nrrd')
        # 겹치는 부분에서 두 레이블 값이 동시에 등장하는 픽셀의 개수 카운트
        overlap_count = np.sum((img_array == 1) & (img_array == 2))

        print("두 레이블 값이 겹치는 픽셀 개수:", overlap_count)


if __name__ == "__main__":
    split_nrrd(original_folder_path, create_split_nrrd_path, create_in_one_folder)
    # nrrd 파일 읽기
    # image = sitk.ReadImage(r'C:\Users\3DONS\Desktop\MD_MX_TEST\MXMN\predict\10005_output.nrrd')

    # # 이미지의 크기 (size) 가져오기
    # size = image.GetSize()
    # print(size[0],size[1],size[2])
    # # 복셀 개수 계산 (가로 × 세로 × 높이)
    # voxel_count = size[0] * size[1] * size[2]
    # print("복셀 개수:", voxel_count)