# Made by Zach Kelly
# Last updated 1/24/20

from src.common_code import *
import tkinter
from tkinter import ttk, filedialog

# TODO fix duped code between both scripts
# TODO standardize padding and such
# TODO remove unnecessary lambdas


def main():
    defaults = extract_defaults()

    root = tkinter.Tk()
    root.title('Prepare Input Files')
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
    test_run_value.set(int(defaults['Test Run']))
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
                minimum_missing_spinbox.set(int(defaults['Minimum Missing Assignments Default']))
            return
        if int(minimum_missing_spinbox.get()) > int(defaults['Minimum Missing Assignments Maximum']):
            minimum_missing_spinbox.set(int(defaults['Minimum Missing Assignments Maximum']))

    minimum_missing_value = tkinter.IntVar()
    minimum_missing_spinbox = ttk.Spinbox(misc_options_frame, from_=0,
                                          to=int(defaults['Minimum Missing Assignments Maximum']),
                                          textvariable=minimum_missing_value)
    minimum_missing_spinbox.set(int(defaults['Minimum Missing Assignments Default']))
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
    email_host_value = ttk.Combobox(email_options_frame, values=['Gmail', 'Outlook', 'Yahoo'], state='readonly')
    email_host_value.set(defaults['Email Host'])
    email_host_value.grid(row=1, column=2)

    # Email message format
    email_parameter_list = [defaults['Email Subject'], defaults['CC Emails'], str(defaults['Email Message'])]
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
    gradebook_file.set(defaults['Gradebook Path'])
    gradebook_path_label1 = ttk.Label(gradebook_path_selection_frame, text='Current Gradebook Path:')
    gradebook_path_label1.grid(row=0, column=0, sticky='w')

    gradebook_path_label2 = ttk.Label(gradebook_path_selection_frame, textvariable=gradebook_file,
                                      font='Helvetica 8 bold')
    gradebook_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    gradebook_full_path = defaults['Gradebook Path']
    gradebook_path_button = ttk.Button(gradebook_path_selection_frame, text='Select File',
                                       command=lambda: gradebook_file.set(get_gradebook_path()))
    gradebook_path_button.grid(row=0, column=1)

    def get_gradebook_path():
        nonlocal gradebook_full_path
        max_path_characters = 65
        gradebook_path = tkinter.filedialog.askopenfilename(filetypes=[('csv', '.csv')], initialdir='.')
        if gradebook_path == '':
            gradebook_path = defaults['Gradebook Path']
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
    student_list_file.set(defaults['Student List Path'])
    student_list_path_label1 = ttk.Label(student_list_path_selection_frame, text='Current Student List Path:')
    student_list_path_label1.grid(row=0, column=0, sticky='w')

    student_list_path_label2 = ttk.Label(student_list_path_selection_frame, textvariable=student_list_file,
                                         font='Helvetica 8 bold')
    student_list_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    student_list_full_path = defaults['Student List Path']
    student_list_path_button = ttk.Button(student_list_path_selection_frame, text='Select File',
                                          command=lambda: student_list_file.set(get_student_list_path()))
    student_list_path_button.grid(row=0, column=1)

    def get_student_list_path():
        nonlocal student_list_full_path
        max_path_characters = 65
        student_list_path = tkinter.filedialog.askopenfilename(filetypes=[('csv', '.csv')], initialdir='.')
        if student_list_path == '':
            student_list_path = defaults['Student List Path']
        student_list_full_path = student_list_path
        if len(student_list_path) > max_path_characters:
            student_list_path = '...' + student_list_path[
                                     len(student_list_path) - max_path_characters + 3:len(student_list_path)]
        return student_list_path

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
        auto_grade(student_list_full_path, gradebook_full_path, test_run_value.get(),
                   int(minimum_missing_spinbox.get()), sender_email_entry.get(), sender_password_entry.get(),
                   email_host_value.get(), email_parameter_list[subject], email_parameter_list[cc_emails],
                   email_parameter_list[message])

    root.resizable(width=False, height=False)
    root.mainloop()


def auto_grade(students, grades, test_run, min_missing, sender_email, sender_password, host, subject, cc, message):
    print(students)
    print(grades)
    print(test_run)
    print(min_missing)
    print(sender_email)
    print(sender_password)
    print(host)
    print(subject)
    print(subject)
    print(cc)
    print(message)
    pass


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

    email = tkinter.Text(email_message_frame, wrap="word")
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


main()
