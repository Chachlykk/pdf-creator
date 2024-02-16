import json


def read_settings():
    """
    reads configuration settings

    :return: dictionary
    """
    with open("settings.json", 'r', encoding="utf-8") as f:
        data = json.load(f)
    return data
