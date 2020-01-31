import smtplib


def extract_defaults():
    defaults_dict = {}
    with open('defaults.txt') as defaults:
        defaults_string = defaults.read()
        defaults_list = defaults_string.split('#')
        # The first item in the list is just formatting instructions, and can be discarded
        for option in defaults_list[1:]:
            # This extracts the option name and the corresponding value, excluding '#' and '\n'
            name_and_value = option.split('\\\\')
            # Name is in index 0, value is in index 1
            defaults_dict[name_and_value[0][:-1]] = name_and_value[1][:-1]
    return defaults_dict


def contains_excluded_phrase(category, defaults):
    excluded_category_phrases = defaults['Excluded Category Phrases'].split(',')
    for phrase in excluded_category_phrases:
        if phrase in category:
            return 1
    return 0


def calculate_percent_complete(grade_data, assignment_index):
    number_complete = 0
    for student in range(1, len(grade_data)):
        if not (grade_data[student][assignment_index] == '0' or grade_data[student][assignment_index] == '-'):
            number_complete = number_complete+1
    return round(number_complete/(len(grade_data)-1)*100, 2)  # The only non-student row is the header


def test_email_host(defaults, host, email, password):
    smtp_server = defaults[host+' SMTP']
    port = int(defaults[host+' Port'])
    server = smtplib.SMTP('smtp.mail.yahoo.com', 587)
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