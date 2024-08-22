import json
import os
import time
import sys
from enum import Enum
import logging

class ConfigKeys(Enum):
    LOG_FILE_PATH = 'log_file_path'
    MAX_BYTES = 'max_bytes'
    BACKUP_COUNT = 'backup_count'
    WINDOW_WIDTH = 'window_width'
    WINDOW_HEIGHT = 'window_height'
    CLICK_WAIT_TIME = 'click_wait_time'
    TASKBAR_HEIGHT = 'taskbar_height'
    WINDOWS_PER_BATCH = 'windows_per_batch'

class ConfigManager:
    def __init__(self, config_file='config.json', reload_interval=360):
        self.config_file = config_file
        self.reload_interval = reload_interval
        self.last_reload_time = time.time()

        if getattr(sys, 'frozen', False):
            self.base_dir = sys._MEIPASS
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))

        self.config = self.load_and_resolve_config()

    def validate_config(self, config: dict):
        required_keys = {
            ConfigKeys.LOG_FILE_PATH: str,
            ConfigKeys.MAX_BYTES: int,
            ConfigKeys.BACKUP_COUNT: int,
            ConfigKeys.WINDOW_WIDTH: int,
            ConfigKeys.WINDOW_HEIGHT: int,
            ConfigKeys.CLICK_WAIT_TIME: int,
            ConfigKeys.TASKBAR_HEIGHT: int,
            ConfigKeys.WINDOWS_PER_BATCH: int
        }

        for key, expected_type in required_keys.items():
            value = config.get(key.value)
            if value is None:
                logging.error(f"Missing required configuration key: {key}")
                sys.exit(1)
            if not isinstance(value, expected_type):
                logging.error(f"Invalid type for {key}: expected {expected_type.__name__}, got {type(config[key]).__name__}")
                sys.exit(1)

            if key == ConfigKeys.LOG_FILE_PATH:
                if not os.path.isabs(value):
                    config[key.value] = os.path.abspath(value)
                if not value.startswith(self.base_dir):
                    logging.error(f"Invalid log file path: {value} is outside the allowed directory.")
                    sys.exit(1)
            elif key in {ConfigKeys.WINDOW_WIDTH, ConfigKeys.WINDOW_HEIGHT, ConfigKeys.CLICK_WAIT_TIME}:
                if value < 0:
                    logging.error(f"Invalid value for {key.value}: must be a positive integer.")
                    sys.exit(1)

    def load_and_resolve_config(self) -> dict:
        print(f"Loading configuration from: {self.config_file}")
        try:
            config_file_path = os.path.join(self.base_dir, self.config_file)
            with open(self.config_file, 'r') as file:
                config = json.load(file)
                self.resolve_paths(config)
                self.validate_config(config)
                return config
        except FileNotFoundError:
            print(f"Configuration file '{self.config_file}' not found.")
            sys.exit(1)
        except json.JSONDecodeError as e:
            print(f"Error decoding configuration: {e}")
            sys.exit(1)

    def resolve_paths(self, config):
        log_file_path = config.get(ConfigKeys.LOG_FILE_PATH.value, 'afk_script.log')
        if not os.path.isabs(log_file_path):
            config[ConfigKeys.LOG_FILE_PATH.value] = os.path.join(self.base_dir, log_file_path)

    def get(self, key: ConfigKeys, default=None):
        env_key = key.value.upper()
        if env_key in os.environ:
            return os.environ[env_key]
        return self.config.get(key.value, default)

    def reload_config(self):
        old_config = self.config.copy()
        self.config = self.load_and_resolve_config()

        changes = []
        for key, old_value in old_config.items():
            new_value = self.config.get(key)
            if new_value != old_value:
                changes.append((key, old_value, new_value))

        if changes:
            logging.info("Configuration reloaded with the following changes:")
            for key, old_value, new_value in changes:
                logging.info(f" - {key}: {old_value} -> {new_value}")
        else:
            logging.info("Configuration reloaded with no changes.")

    def dynamic_reload_config(self):
        current_time = time.time()
        if current_time - self.last_reload_time > self.reload_interval:
            config_mod_time = os.path.getmtime(self.config_file)
            if config_mod_time > self.last_reload_time:
                logging.info("Detected changes in the configuration file. Reloading configuration...")
                previous_config = self.config
                self.reload_config()
                self.last_reload_time = current_time

                if (self.config[ConfigKeys.WINDOW_WIDTH.value] != previous_config[ConfigKeys.WINDOW_WIDTH.value] or
                    self.config[ConfigKeys.WINDOW_HEIGHT.value] != previous_config[ConfigKeys.WINDOW_HEIGHT.value] or
                    self.config[ConfigKeys.WINDOWS_PER_BATCH.value] != previous_config[ConfigKeys.WINDOWS_PER_BATCH.value]):
                    logging.info("Window positions cache cleared due to configuration changes.")
            else:
                logging.info("No changes detected in the configuration file.")
        else:
            logging.info(f"Skipping config reload. Next check in {int(self.reload_interval - (current_time - self.last_reload_time))} seconds.")

