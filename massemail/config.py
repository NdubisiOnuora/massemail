import json


class Config(object):
    def __init__(self, config_file_name):
        with open(config_file_name) as f:
            config = json.load(f)
            for config_key, config_value in config.items():
                setattr(self, config_key, config_value)
