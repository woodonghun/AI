import json
import os
# 스크립트의 위치를 기준으로 상대 경로 설정
script_dir = os.path.dirname(__file__)  # 현재 스크립트의 디렉토리
config_path = os.path.join(script_dir, "config_preShin.json")
preShin_path = os.path.join(script_dir, "group_points_preShin.json")

try:
    with open(config_path, "r") as f:
        CONFIG: dict = json.load(f)

    with open(preShin_path, "r") as f:
        # with open(Preshin2_temp.args.path_group_points + "group_points.json", "r") as f:
        GROUP_POINTS: dict = json.load(f)
except Exception as e:
    print("Failed to read config.json:", e)
