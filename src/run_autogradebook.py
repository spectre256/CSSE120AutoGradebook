# Made by Zach Kelly
# Last updated 1/24/20

from common_code import *
import tkinter
from tkinter import ttk, filedialog
import csv
import smtplib
import ssl
from email.message import EmailMessage

# TODO fix duped code between both scripts
# TODO standardize padding and such
# TODO remove unnecessary lambdas

# Notes to Dr. Fisher:
#   Create & use App Passwords
#      Go to your Google Account.
#      Select Security.
#      Under "Signing in to Google," select App Passwords. You may need to sign in. ...
#      At the bottom, choose Select app and choose the app you using Select device and choose the device you're using. ...
#      Follow the instructions to enter the App Password. ...
#      Tap Done.
# email --> fisherds@gmail.com
# Password --> 
# Server --> Gmail

# that app specific password was created for my Mac on 11/7/22, hopefully I can reuse it.

config = Config("defaults.json", "config.json")


def main():

    root = tkinter.Tk()
    root.title('Choose Autograde Settings')
    main_frame = ttk.Frame(root)
    main_frame.pack(fill='both', padx=25, pady=25)

    ##############
    # Misc. Options
    ##############

    # Misc. Options frame
    misc_options_frame = ttk.Frame(main_frame)
    misc_options_frame.grid_columnconfigure(0, weight=1)
    misc_options_frame.grid_columnconfigure(1, weight=1)
    misc_options_frame.pack(fill='both', pady=10, ipadx=100)

    # Test run checkbutton and label
    test_run_label = ttk.Label(misc_options_frame, text='Test Run?')
    test_run_label.grid(row=0, column=1)

    test_run_value = tkinter.BooleanVar()
    config.attach_var(test_run_value, "Test Run")
    test_run_list_checkbutton = ttk.Checkbutton(misc_options_frame, variable=test_run_value)
    test_run_list_checkbutton.grid(row=1, column=1)

    # Minimum missing assignments spinbox
    minimum_missing_label = ttk.Label(misc_options_frame, text='Minimum Missing Assignments')
    minimum_missing_label.grid(row=0, column=0)

    def minimum_missing_helper():
        try:
            int(minimum_missing_spinbox.get())
        except ValueError:
            if minimum_missing_spinbox.get() != '':
                minimum_missing_spinbox.set(config.get("Minimum Missing Assignments Default"))
            return
        if int(minimum_missing_spinbox.get()) > config.get("Minimum Missing Assignments Maximum"):
            minimum_missing_spinbox.set(config.get("Minimum Missing Assignments Maximum"))

    minimum_missing_value = tkinter.IntVar()
    config.attach_var(minimum_missing_value, "Minimum Missing Assignments Default")
    minimum_missing_spinbox = ttk.Spinbox(misc_options_frame, from_=0,
                                          to=config.get("Minimum Missing Assignments Maximum"),
                                          textvariable=minimum_missing_value)
    minimum_missing_spinbox.grid(row=1, column=0)
    minimum_missing_value.trace('w', lambda *args: minimum_missing_helper())

    ###############
    # Email options
    ###############
    # TODO: error handling (bad email, password, can't connect to service provider, etc.)
    # Email options frame
    email_options_frame = ttk.Frame(main_frame)
    email_options_frame.grid_columnconfigure(0, weight=0)
    email_options_frame.grid_columnconfigure(1, weight=0)
    email_options_frame.grid_columnconfigure(2, weight=1)
    email_options_frame.pack(fill='both', pady=10)

    # Sender login info
    sender_email_label = ttk.Label(email_options_frame, text='Sender Email:')
    sender_email_label.grid(row=0, column=0, sticky='w')
    sender_email_entry = ttk.Entry(email_options_frame)
    sender_email_entry.grid(row=0, column=1)

    sender_password_label = ttk.Label(email_options_frame, text='Sender Password:')
    sender_password_label.grid(row=1, column=0, sticky='w')
    sender_password_entry = ttk.Entry(email_options_frame, show='*')
    sender_password_entry.grid(row=1, column=1)

    # Email host
    email_host_label = ttk.Label(email_options_frame, text='Email Host:')
    email_host_label.grid(row=0, column=2)
    email_host_variable = tkinter.StringVar()
    config.attach_var(email_host_variable, "Email Host")
    email_host_value = ttk.Combobox(email_options_frame, values=config.get("Host List"), state="readonly", textvariable=email_host_variable)#, "Yahoo"], state="readonly")
    email_host_value.grid(row=1, column=2)

    # Email message format
    email_parameter_list = [config.get("Email Subject"), config.get("CC Emails"), config.get("Email Message")]
    email_format_button = ttk.Button(email_options_frame, text='Change Message Format',
                                     command=lambda: email_format_window(root, email_parameter_list))
    email_format_button.grid(row=2, columnspan=3, pady=10)

    ###############################
    # Gradebook file path selection
    ###############################

    # Gradebook path selection frame
    gradebook_path_selection_frame = ttk.Frame(main_frame)
    gradebook_path_selection_frame.grid_columnconfigure(0, weight=1)
    gradebook_path_selection_frame.grid_columnconfigure(1, weight=1)
    gradebook_path_selection_frame.pack(fill='both', pady=10)

    # Gradebook path selection button and labels
    gradebook_file = tkinter.StringVar()
    config.attach_var(gradebook_file, "Gradebook Path")
    gradebook_path_label1 = ttk.Label(gradebook_path_selection_frame, text='Current Gradebook Path:')
    gradebook_path_label1.grid(row=0, column=0, sticky='w')

    gradebook_path_label2 = ttk.Label(gradebook_path_selection_frame, textvariable=gradebook_file,
                                      font='Helvetica 8 bold')
    gradebook_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    gradebook_full_path = config.get("Gradebook Path")
    gradebook_path_button = ttk.Button(gradebook_path_selection_frame, text='Select File',
                                       command=lambda: gradebook_file.set(get_gradebook_path()))
    gradebook_path_button.grid(row=0, column=1)

    def get_gradebook_path():
        nonlocal gradebook_full_path
        max_path_characters = 65
        gradebook_path = tkinter.filedialog.askopenfilename(filetypes=[('csv', '.csv')], initialdir='.')
        if gradebook_path == '':
            gradebook_path = config.get("Gradebook Path")
        gradebook_full_path = gradebook_path
        if len(gradebook_path) > max_path_characters:
            gradebook_path = '...'+gradebook_path[len(gradebook_path)-max_path_characters+3:len(gradebook_path)]
        return gradebook_path

    ###############################
    # Student list file path selection
    ###############################

    # Student list path selection frame
    student_list_path_selection_frame = ttk.Frame(main_frame)
    student_list_path_selection_frame.grid_columnconfigure(0, weight=1)
    student_list_path_selection_frame.grid_columnconfigure(1, weight=1)
    student_list_path_selection_frame.pack(fill='both', pady=10)

    # Student list path selection button and labels
    student_list_file = tkinter.StringVar()
    config.attach_var(student_list_file, "Student List Path")
    student_list_path_label1 = ttk.Label(student_list_path_selection_frame, text='Current Student List Path:')
    student_list_path_label1.grid(row=0, column=0, sticky='w')

    student_list_path_label2 = ttk.Label(student_list_path_selection_frame, textvariable=student_list_file,
                                         font='Helvetica 8 bold')
    student_list_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    student_list_full_path = config.get("Student List Path")
    student_list_path_button = ttk.Button(student_list_path_selection_frame, text='Select File',
                                          command=lambda: student_list_file.set(get_student_list_path()))
    student_list_path_button.grid(row=0, column=1)

    def get_student_list_path():
        nonlocal student_list_full_path
        max_path_characters = 65
        student_list_path = tkinter.filedialog.askopenfilename(filetypes=[('txt', '.txt')], initialdir='.')
        if student_list_path == '':
            student_list_path = config.get("Student List Path")
        student_list_full_path = student_list_path
        if len(student_list_path) > max_path_characters:
            student_list_path = '...' + student_list_path[
                                     len(student_list_path) - max_path_characters + 3:len(student_list_path)]
        return student_list_path

    #################################
    # Assignments file path selection
    #################################

    # Assignments path selection frame
    assignments_path_selection_frame = ttk.Frame(main_frame)
    assignments_path_selection_frame.grid_columnconfigure(0, weight=1)
    assignments_path_selection_frame.grid_columnconfigure(1, weight=1)
    assignments_path_selection_frame.pack(fill='both', pady=10)

    # Assignment path selection button and labels
    assignments_file = tkinter.StringVar()
    config.attach_var(assignments_file, "Assignments Path")
    assignments_path_label1 = ttk.Label(assignments_path_selection_frame, text='Current Assignments Path:')
    assignments_path_label1.grid(row=0, column=0, sticky='w')

    assignments_path_label2 = ttk.Label(assignments_path_selection_frame, textvariable=assignments_file,
                                         font='Helvetica 8 bold')
    assignments_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    assignments_full_path = config.get("Assignments Path")
    assignments_path_button = ttk.Button(assignments_path_selection_frame, text='Select File',
                                          command=lambda: assignments_file.set(get_assignments_path()))
    assignments_path_button.grid(row=0, column=1)

    def get_assignments_path():
        nonlocal assignments_full_path
        max_path_characters = 65
        assignments_path = tkinter.filedialog.askopenfilename(filetypes=[('csv', '.csv')], initialdir='.')
        if assignments_path == '':
            assignments_path = config.get("Assignments Path")
        assignments_full_path = assignments_path
        if len(assignments_path) > max_path_characters:
            assignments_path = '...' + assignments_path[
                                        len(assignments_path) - max_path_characters + 3:len(assignments_path)]
        return assignments_path

    #############################
    # Percent threshold selection
    #############################

    # Percent threshold frame
    threshold_frame = ttk.Frame(main_frame)
    threshold_frame.grid_columnconfigure(0, weight=1)
    threshold_frame.grid_columnconfigure(1, weight=1)
    threshold_frame.pack(fill='both', pady=20)

    # Assignments file selection toggles the percentage threshold scale
    def check_percent_threshold_scale_active(*args):
        nonlocal percent_threshold_value_label
        if assignments_file.get() != config.get("Assignments Path"):
            percent_threshold_scale.set(config.get("Default Percentage Threshold"))
            percent_threshold_value_label.config(text=str(config.get("Default Percentage Threshold")) + '%')
            percent_threshold_value.set(config.get("Default Percentage Threshold"))

    assignments_file.trace('w', check_percent_threshold_scale_active)

    # Percent threshold scale and labels
    percent_threshold_value_label = ttk.Label(threshold_frame, font='Helvetica 18 bold', anchor=tkinter.CENTER, width=5)
    percent_threshold_value_label.grid(row=0, column=1, rowspan=2)

    percent_threshold_value = tkinter.IntVar()
    config.attach_var(percent_threshold_value, "Default Percentage Threshold")
    percent_threshold_scale = ttk.Scale(threshold_frame, from_=0, to=100,
                                        variable=percent_threshold_value, command=lambda value:
                                        percent_threshold_value_label.config(text=str(int(float(value))) + '%'))
    percent_threshold_scale.grid(row=1)
    percent_threshold_value.trace('w', check_percent_threshold_scale_active)

    percentage_threshold_text_label = ttk.Label(threshold_frame, text='Assignment Percentage Threshold:', anchor=tkinter.CENTER)
    percentage_threshold_text_label.grid(row=0, column=0)

    #######################
    # action buttons
    #######################

    # action buttons frame
    action_button_frame = ttk.Frame(main_frame)
    action_button_frame.grid_columnconfigure(0, weight=1)
    action_button_frame.grid_columnconfigure(1, weight=1)
    action_button_frame.pack(fill='both', pady=10)

    # action buttons
    ok_button = ttk.Button(action_button_frame, text='Ok', command=lambda: ok_helper())
    ok_button.grid(row=0, column=0)

    cancel_button = ttk.Button(action_button_frame, text='Cancel', command=lambda: root.quit())
    cancel_button.grid(row=0, column=1)

    def ok_helper():
        subject = 0
        cc_emails = 1
        message = 2

        root.quit()
        config.write()
        auto_grade(student_list_full_path, gradebook_full_path, assignments_full_path, test_run_value.get(),
                   percent_threshold_value.get(), int(minimum_missing_spinbox.get()), sender_email_entry.get(),
                   sender_password_entry.get(), email_host_value.get(), email_parameter_list[subject],
                   email_parameter_list[cc_emails], email_parameter_list[message])

    root.resizable(width=False, height=False)
    root.mainloop()


def email_format_window(root, email_parameter_list):
    subject = 0
    cc_emails = 1
    message = 2

    email_format_toplevel = tkinter.Toplevel(root)
    email_format_toplevel.title('Change Email Format')

    email_message_frame = ttk.Frame(email_format_toplevel)
    email_message_frame.grid_columnconfigure(0, weight=0)
    email_message_frame.grid_columnconfigure(1, weight=1)
    email_message_frame.pack()

    subject_label = ttk.Label(email_message_frame, text='Subject: ')
    subject_label.grid(row=0, column=0, sticky='w')

    subject_value = tkinter.StringVar()
    subject_entry = ttk.Entry(email_message_frame, textvariable=subject_value)
    subject_value.set(email_parameter_list[subject])
    subject_entry.grid(row=0, column=1, sticky='ew')

    cc_emails_label = ttk.Label(email_message_frame, text='CC: ')
    cc_emails_label.grid(row=1, column=0, sticky='w')

    cc_emails_value = tkinter.StringVar()
    cc_emails_entry = ttk.Entry(email_message_frame, textvariable=cc_emails_value)
    cc_emails_value.set(email_parameter_list[cc_emails])
    cc_emails_entry.grid(row=1, column=1, sticky='ew')

    email_label = ttk.Label(email_message_frame, text='\'[NAME]\' will insert a student\'s name,' +
                                                      'and \'[ASSIGNMENTS]\' will insert the missing assignment list')
    email_label.grid(row=2, columnspan=2)

    email = tkinter.Text(email_message_frame, wrap='word')
    email.insert(tkinter.INSERT, email_parameter_list[message])
    email.grid(row=3, column=0, columnspan=2)

    # action buttons
    def ok_helper():
        nonlocal email_parameter_list
        email_parameter_list[0:2] = [subject_value.get(), cc_emails_value.get(), email.get('0.0', tkinter.END)]
        email_format_toplevel.grab_release()
        email_format_toplevel.destroy()

    def quit_helper():
        email_format_toplevel.grab_release()
        email_format_toplevel.destroy()

    action_button_frame = ttk.Frame(email_format_toplevel)
    action_button_frame.grid_columnconfigure(0, weight=1)
    action_button_frame.grid_columnconfigure(1, weight=1)
    action_button_frame.pack(fill='both')

    ok_button = ttk.Button(action_button_frame, text='Ok', command=lambda: ok_helper())
    ok_button.grid(row=3, column=0, pady=10)

    cancel_button = ttk.Button(action_button_frame, text='Cancel', command=lambda: quit_helper())
    cancel_button.grid(row=3, column=1, pady=10)

    email_format_toplevel.resizable = False
    email_format_toplevel.update()

    email_format_toplevel.focus_set()
    email_format_toplevel.grab_set()


def auto_grade(students, grades, assignments, test_run, percent, min_missing, sender, password, host, subject, cc, msg):
    # print(students)         # Input
    # print(grades)           # Turn into CSV
    # print(assignments)      # Input
    # print(percent)          # Input
    # print(test_run)         # Input
    # print(min_missing)      # Input
    # print(sender)           # Input
    # print(password)         # Input
    # print(host)             # Get host settings (SMTP server/port)
    # print(subject)          # Input                 }
    # print(cc)               # Extract emails        } combine into html email?
    # print(msg)          # Figure out inserts    }
    # Extract list of assignments each student is missing
    #
    # Create final email object (mimetext?)
    with open(grades) as grade_file:
        grade_list = list(csv.reader(grade_file))
    assignments_list = extract_assignments_list(assignments, percent, grade_list)
    students_list = extract_students_list(students, grade_list)
    missing_assignments_list = extract_missing_assignments(students_list, grade_list, assignments_list)
    server = server_setup(sender, password, host)
    email_parameters = {'message': msg.translate({ord('\n'): None}), 'subject': subject, 'cc emails': cc.split(','),
                        'sender email': sender}
    send_emails(email_parameters, server, missing_assignments_list, test_run, min_missing)


def extract_assignments_list(assignments, percent, grade_list):
    assignments_list = []
    if assignments != config.get("Assignments Path"):
        with open(assignments) as assignments_file:
            assignments_data = list(csv.reader(assignments_file))
        for assignment in range(1,len(assignments_data)):
            if int(float(assignments_data[assignment][1])) >= percent: # Index 1 is percent submitted
                assignments_list.append(assignments_data[assignment][0]) # Index 0 is assignment name
    else:
        categories = grade_list[0]
        emails_column = categories.index('Email address')
        for category in range(emails_column + 1, len(categories)):
            if not contains_excluded_phrase(categories[category], defaults):
                percent_complete = calculate_percent_complete(grade_list, category)
                if percent_complete >= percent:
                    assignments_list.append(categories[category])
    return assignments_list


def extract_students_list(students, grade_list):
    if students == config.get("Student List Path"):
        students_list = []
        emails_index = grade_list[0].index('Email address')
        for student in range(1,len(grade_list)):
            email = grade_list[student][emails_index]
            at = email.index('@')
            students_list.append(email[:at])
    else:
        with open(students) as students_file:
            students_list = students_file.read().split('\n')
    return students_list


def extract_missing_assignments(students_list, grade_list, assignments_list): # TODO
    assignments_index = []
    categories = grade_list[0]
    for assignment in assignments_list:
        assignments_index.append(categories.index(assignment))
    email_index = categories.index('Email address')
    first_name_index = categories.index('First name')
    last_name_index = categories.index('Last name') # TODO: defaults
    student_assignment_list = []
    student_index = 1
    for student in students_list:
        while student not in grade_list[student_index][email_index]:
            student_index += 1
        missing_assignments = [grade_list[student_index][email_index], grade_list[student_index][first_name_index]+' '+
                               grade_list[student_index][last_name_index]]
        for assignment in assignments_index:
            grade = grade_list[student_index][assignment]
            if (grade == '-') or (grade == '0'):
                missing_assignments.append(categories[assignment])
        student_assignment_list.append(missing_assignments)
    return student_assignment_list


def server_setup(sender_email, sender_password, host):
    print(f"sender_email {sender_email}")
    print(f"len(sender_password) = {len(sender_password)}")
    print(f"host {host}")
    smtp_server = defaults[host + ' SMTP']
    port = int(defaults[host + ' Port'])
    context = ssl.SSLContext(ssl.PROTOCOL_TLS)
    server = smtplib.SMTP(smtp_server, port)
    # server.connect(smtp_server, port)
    server.ehlo()
    server.starttls(context=context)
    server.ehlo()
    try:
        server.login(sender_email, sender_password)
    except:
        print('Email login failed!')
        if host == 'Gmail':
            print('You need an app password to use gmail with this program:' +
                  '\nhttps://support.google.com/accounts/answer/185833')
    return server


def send_emails(email_parameters, server, missing_assignments, test_run, minimum_missing):
    for assignments_list in missing_assignments:
        if len(assignments_list)-2 >= minimum_missing:
            email = EmailMessage()
            email['Subject'] = email_parameters['subject']
            email['From'] = email_parameters['sender email']
            if test_run:
                email['To'] = email['From']
            else:
                email['To'] = assignments_list[0]  # Index 0 is the student email
            email['CC'] = email_parameters['cc emails']
            assignments = """
               
            """.join(assignments_list[2:])
            assignments = """
            
            {}
            
""".format(assignments)
            message = email_parameters['message']
            message = message.replace('[NAME]', assignments_list[1]).replace('[ASSIGNMENTS]', assignments)
            email.set_content(message)
            if test_run:
                server.send_message(email, email_parameters['sender email'], email_parameters['sender email'])
            else:
                server.send_message(email, email_parameters['sender email'], assignments_list[0])

    server.quit()


main()
