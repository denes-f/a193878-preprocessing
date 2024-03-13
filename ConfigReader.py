"""
Methods to read config values
"""
import os
import configparser

class ConfigReader:
    """
Reader to get any value, a float or an int from a config file
    """

    def __init__(self, path):
        if not os.path.exists(path):
            raise Exception('Config file was not found: {0}'.format(path))
        self._config_reader = configparser.ConfigParser()
        self._config_reader.read(path)

    def get_value(self, section, key):
        """
Reading a config value
        :param section: Section of the config file
        :param key: Name of the config entry to read
        :return: The config entry
        """
        if section in self._config_reader:
            if key in self._config_reader[section]:
                return self._config_reader[section][key]
            else:
                raise Exception('Key {0} was not found in section {1} of the config file'.format(key, section))
        else:
            raise Exception('Section {0} was not found in the config file'.format(section))

    def get_float(self, section, key):
        """
Reading a config value and converting it to a float. Throws an error if the value cannot be interpreted as a float.
        :param section: Section of the config file
        :param key: Name of the config entry to read as a float
        :return: The config entry converted to float
        """
        value = self.get_value(section, key)
        try:
            return float(value)
        except ValueError:
            raise Exception('Invalid config entry: {0} for section {1}, key {2}. The type of the value has to be float'.format(value, section, key))

    def get_int(self, section, key):
        """
Reading a config value and converting it to an int. Throws an error if the value cannot be interpreted as an int.
        :param section: Section of the config file
        :param key: Name of the config entry to read as an int
        :return: The config entry converted to int
        """
        value = self.get_value(section, key)
        try:
            return int(value)
        except ValueError:
            raise Exception('Invalid config entry: {0} for section {1}, key {2}. The type of the value has to be an int'.format(value, section, key))
