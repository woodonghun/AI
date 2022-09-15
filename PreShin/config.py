import json

try:
    with open("config_preShin.json", "r") as f:
        CONFIG: dict = json.load(f)

    with open("group_points_preShin.json", "r") as f:
    # with open(Preshin2_temp.args.path_group_points + "group_points.json", "r") as f:
        GROUP_POINTS: dict = json.load(f)
except Exception as e:
    print("Failed to read config.json:", e)
    #exit(-1)