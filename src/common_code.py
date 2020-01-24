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
