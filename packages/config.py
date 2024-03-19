"""
Loads a config file and gets config values
"""
import os
import configparser


class Config:
    """
Loads a config file and gets config values
    """

    def __init__(self, path):
        if not os.path.exists(path):
            raise Exception(f"Config file was not found: {os.path.abspath(path)}")
        self._config = configparser.ConfigParser()
        self._config.read(path)

    def get_entry(self, section, key):
        """
Reads a config value
        :param section: Section of the config file
        :param key: Name of the config entry to read
        :return: The config entry
        """
        if section in self._config:
            if key in self._config[section]:
                return self._config[section][key]
            else:
                raise Exception(f"Key '{key}' was not found in section '{section}' of the config file")
        else:
            raise Exception(f"Section '{section}' was not found in the config file")

    def get_float(self, section, key):
        """
Reads a config value and converts it to a float. Throws an error if the value cannot be interpreted as a float.
        :param section: Section of the config file
        :param key: Name of the config entry to be read as a float
        :return: The config entry converted to float
        """
        entry = self.get_entry(section, key)
        try:
            return float(entry)
        except ValueError:
            raise Exception(f"Invalid config entry: {entry} in section '{section}', key '{key}'. The type of the value has to be float")

    def get_int(self, section, key):
        """
Reads a config value and converts it to an int. Throws an error if the value cannot be interpreted as an int.
        :param section: Section of the config file
        :param key: Name of the config entry to read as an int
        :return: The config entry converted to int
        """
        entry = self.get_entry(section, key)
        try:
            return int(entry)
        except ValueError:
            raise Exception(f"Invalid config entry: '{entry}' for section '{section}', key '{key}'. The type of the value has to be an int")
