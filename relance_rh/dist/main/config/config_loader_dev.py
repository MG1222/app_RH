import json
import os

def load_config(config_path='config/config_dev.json'):
    with open(config_path, 'r') as config_file:
        config = json.load(config_file)
    return config
