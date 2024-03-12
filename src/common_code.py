import json
import smtplib
from pathlib import Path

class Config:
    def __init__(self, defaults_filename, config_filename):
        self.defaults_file = Path(defaults_filename)
        self.config_file = Path(config_filename)

        self.extract_config()


    def extract_config(self):
        if not self.defaults_file.exists() or self.defaults_file.is_dir():
            raise Exception(f"Defaults file '{self.defaults_file}' not found")

        if self.config_file.is_dir():
            raise Exception(f"Config file '{self.defaults_file}' is a directory, not a readable file")

        # Write empty JSON object if the config doesn't already exist
        if not self.config_file.exists():
            with open(self.config_file, "w") as config:
                config.write("{}")

        with open(self.defaults_file) as defaults, open(self.config_file) as config:
            self.defaults = json.load(defaults)
            self.config = {**self.defaults, **json.load(config)}


    def get(self, key):
        return self.config[key]


    def set(self, key, value):
        self.config[key] = value


    def attach_var(self, var, key):
        var.set(self.get(key))

        var.trace_add("write", callback=lambda *_: self.set(key, var.get()))

    # Writes values that differ from the defaults to the config file
    def write(self):
        config = {key: value for key, value in self.config.items() if not key in self.defaults or self.defaults[key] != value}
        with open(self.config_file, "w") as config_file:
            json.dump(config, config_file)


def contains_excluded_phrase(category, defaults):
    for phrase in defaults['Excluded Category Phrases']:
        if phrase in category:
            return True

    return False


def calculate_percent_complete(grade_data, assignment_index):
    number_complete = 0
    for student in range(1, len(grade_data)):
        if not (grade_data[student][assignment_index] == '0' or grade_data[student][assignment_index] == '-'):
            number_complete = number_complete+1
    return round(number_complete/(len(grade_data)-1)*100, 2)  # The only non-student row is the header


def test_email_host(defaults, host, email, password):
    smtp_server = defaults[host+' SMTP']
    port = int(defaults[host+' Port'])
    server = smtplib.SMTP_SSL(smtp_server, port)
    server.connect(smtp_server, port)
    server.ehlo()
    server.starttls()
    server.ehlo()
    try:
        server.login(email, password)
    except:
        print('Email login failed!')
        if host == 'Gmail':
            print('You need an app password to use gmail with this program:' +
                  '\nhttps://support.google.com/accounts/answer/185833')
    server.quit()
