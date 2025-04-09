import nrrd
import numpy as np

def calculate_iou(data1, data2):
    """
    두 세그멘테이션 데이터의 IoU (Intersection over Union)를 계산.
    
    Args:
        data1 (numpy.ndarray): 첫 번째 세그멘테이션 데이터.
        data2 (numpy.ndarray): 두 번째 세그멘테이션 데이터.
    
    Returns:
        float: IoU 값.
    """
    # 교집합 계산
    intersection = np.logical_and(data1, data2)
    
    # 합집합 계산
    union = np.logical_or(data1, data2)
    
    # IoU 계산
    iou = np.sum(intersection) / np.sum(union)
    return iou

# 파일 경로 설정
file1_path = r'C:\ai_project\onnx\MXMD_to_ONNX\MXMNseg\20250103_predicts\result/onnx\145.nrrd'  # 첫 번째 NRRD 파일 경로
file2_path = r'C:\ai_project\onnx\MXMD_to_ONNX\MXMNseg\20250103_predicts\result/145.nrrd'  # 두 번째 NRRD 파일 경로

# NRRD 파일 로드
data1, header1 = nrrd.read(file1_path)
data2, header2 = nrrd.read(file2_path)

# 데이터 크기 확인
if data1.shape != data2.shape:
    print("두 세그멘테이션 데이터의 크기가 다릅니다. IoU를 계산할 수 없습니다.")
else:
    # IoU 계산
    iou = calculate_iou(data1, data2)
    print(f"두 세그멘테이션의 IoU 값: {iou:.4f}")
