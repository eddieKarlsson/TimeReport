import json
import os


class Settings:
    """A class to store all settings, user-data in JSON format."""

    def __init__(self):
        """Initiliaze settings."""

        # Static setings
        self.debug_level = 0

        # internal var, used below in functions
        self.json_file = 'user_settings.json'
        self.indent = 1

    def _create_user_settings(self):
        """Create dict which contains all user data"""
        user_settings = {
            'in_time_folder': 'C:/Users/EddieKarlsson/Documents/M_Report/',
            'in_time_file': 'Månadsrapport 2020 Juli-Dec.xlsx',
            'output_folder': 'output',
            'projno': 200012,
            'month': 'Sept',
            'week': 36
        }

        return user_settings

    def load_user_settings(self):
        """Load stored user settings, otherwise create .json file"""

        # Check if settings file already exists, else create it.
        if os.path.isfile(self.json_file):
            with open(self.json_file, 'r') as f:
                user_settings = json.load(f)
        else:
            user_settings = self._create_user_settings()

            with open(self.json_file, 'w') as f:
                json.dump(user_settings, f, indent=self.indent)

        return user_settings

    def store_user_settings(self, user_settings):
        """Dump Dict to JSON file"""
        with open(self.json_file, 'w') as f:
            json.dump(user_settings, f, indent=self.indent)
