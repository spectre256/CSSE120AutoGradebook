# Made by Zach Kelly
# Last updated 1/18/20

import tkinter
from tkinter import ttk, filedialog
import csv
import os

#TODO add a config file for default values
#TODO standardize padding and such
#TODO remove unecessary lambdas

def main():
    # TODO figure out number of assignments
    num_assignments = 100

    default_percentage_threshold = 75

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
    test_run_value.set(False)
    test_run_list_checkbutton = ttk.Checkbutton(misc_options_frame, variable=test_run_value)
    test_run_list_checkbutton.grid(row=1, column=1)

    # Minimum missing assignments spinbox
    minimum_missing_label = ttk.Label(misc_options_frame, text='Minimum Missing Assignments')
    minimum_missing_label.grid(row=0, column=0)

    def minimum_missing_helper(minimum_missing_spinbox, num_assignments):
        try:
            int(minimum_missing_spinbox.get())
        except ValueError:
            if minimum_missing_spinbox.get() != '':
                minimum_missing_spinbox.set(0)
            return
        if int(minimum_missing_spinbox.get()) > num_assignments:
            minimum_missing_spinbox.set(num_assignments)

    minimum_missing_value = tkinter.IntVar()
    minimum_missing_spinbox = ttk.Spinbox(misc_options_frame, from_=0, to=num_assignments, textvariable=minimum_missing_value)#command=lambda: minimum_missing_helper(minimum_missing_spinbox, num_assignments))
    minimum_missing_spinbox.set(0)
    minimum_missing_spinbox.grid(row=1, column=0)
    minimum_missing_value.trace('w', lambda *args: minimum_missing_helper(minimum_missing_spinbox, num_assignments))

    ###############
    # Email options
    ###############
    #TODO error handling (bad email, password, can't connect to service provider, etc.)
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

    # # CC Emails
    # cc_emails_label = ttk.Label(email_options_frame, text='CC Emails:')
    # cc_emails_label.grid(row=2, column=0, sticky='w')
    # cc_emails_value = tkinter.StringVar()
    # cc_emails_entry = ttk.Entry(email_options_frame, textvariable=cc_emails_value)
    # cc_emails_value.set('email 1, email 2, etc.')
    # cc_emails_entry.grid(row=2, column=1)

    # Email host
    email_host_label = ttk.Label(email_options_frame, text='Email Host:')
    email_host_label.grid(row=0, column=2)
    email_host_value = ttk.Combobox(email_options_frame, values=['Gmail','Outlook','Yahoo'], state='readonly')
    email_host_value.set('Gmail')
    email_host_value.grid(row=1, column=2)

    # Email message format
    #TODO
    email_parameter_list = []
    email_format_button = ttk.Button(email_options_frame, text='Change Message Format', command=lambda: email_format_window(root, email_parameter_list))
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
    gradebook_file.set('None Selected')
    gradebook_path_label1 = ttk.Label(gradebook_path_selection_frame, text='Current Gradebook Path:')
    gradebook_path_label1.grid(row=0, column=0, sticky='w')

    gradebook_path_label2 = ttk.Label(gradebook_path_selection_frame, text='None', textvariable=gradebook_file,
                                      font='Helvetica 8 bold')
    gradebook_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    gradebook_full_path = 'None Selected'
    gradebook_path_button = ttk.Button(gradebook_path_selection_frame, text='Select File',
                                       command=lambda: gradebook_file.set(get_gradebook_path()))
    gradebook_path_button.grid(row=0, column=1)

    def get_gradebook_path():
        nonlocal gradebook_full_path
        max_path_characters = 70
        gradebook_path = tkinter.filedialog.askopenfilename(filetypes=[('csv', '.csv')], initialdir='.')
        if gradebook_path == '':
            gradebook_path = 'None Selected'
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
    student_list_file.set('All Students')
    student_list_path_label1 = ttk.Label(student_list_path_selection_frame, text='Current Student List Path:')
    student_list_path_label1.grid(row=0, column=0, sticky='w')

    student_list_path_label2 = ttk.Label(student_list_path_selection_frame, text='None', textvariable=student_list_file,
                                      font='Helvetica 8 bold')
    student_list_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    student_list_full_path = 'All Students'
    student_list_path_button = ttk.Button(student_list_path_selection_frame, text='Select File',
                                       command=lambda: student_list_file.set(get_student_list_path()))
    student_list_path_button.grid(row=0, column=1)

    def get_student_list_path():
        nonlocal student_list_full_path
        max_path_characters = 70
        student_list_path = tkinter.filedialog.askopenfilename(filetypes=[('csv', '.csv')], initialdir='.')
        if student_list_path == '':
            student_list_path = 'All Students'
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
        #TODO: figure out variables
        #cc email input
        #email service provider selector?

        # sender email-same as login? (technically sender and login can be different I think, but prob not a good idea)

        #subject input
        #message format input
        #cc emails

        root.quit()
        #variables
        student_list_full_path
        gradebook_full_path
        test_run_value.get()
        minimum_missing_spinbox.get()
        sender_email_entry.get()
        sender_password_entry.get()
        email_host_value.get()
        # process_files('update_student_list_value.get()', 'update_assignments_value.get()', 'percent_threshold_value.get()',
        #               gradebook_full_path)

    root.resizable(width=False, height=False)
    root.mainloop()


def email_format_window(root, email_parameter_list):
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
    subject_value.set('Missing CSSE120 Assignments')
    subject_entry.grid(row=0, column=1, sticky='ew')

    cc_emails_label = ttk.Label(email_message_frame, text='CC: ')
    cc_emails_label.grid(row=1, column=0, sticky='w')

    cc_emails_value = tkinter.StringVar()
    cc_emails_entry = ttk.Entry(email_message_frame, textvariable=cc_emails_value)
    cc_emails_value.set('ex1@email.com, ex2@email.com, etc.')
    cc_emails_entry.grid(row=1, column=1, sticky='ew')

    email = tkinter.Text(email_message_frame, wrap="word")
    email.insert(tkinter.INSERT,"\nThis email was sent automatically, replies will be ignored.\n"#"\nThis is a test, please check the following list against the gradebook, and reply only if the list is incorrect.\n"
                   + "\nHello "+"[NAME]"+",\n\nYou are currently missing:\n\n"+"[MISSING ASSIGNMENT LIST]\n\n"
                   + "please submit these as soon as you can. Talk to your professor if you have any questions."
                   + " Remember, you must turn in ALL assignments to pass the class!\n\nThank you,\nZach Kelly\n"
                   + "CSSE120 Grader")
    email.grid(row=2, column=0, columnspan=2)

    # action buttons
    def ok_helper():
        #TODO
        email_parameter_list = [None, None]

    action_button_frame = ttk.Frame(email_format_toplevel)
    action_button_frame.grid_columnconfigure(0, weight=1)
    action_button_frame.grid_columnconfigure(1, weight=1)
    action_button_frame.pack(fill='both')

    ok_button = ttk.Button(action_button_frame, text='Ok', command=lambda: ok_helper())
    ok_button.grid(row=3, column=0, pady=10)

    cancel_button = ttk.Button(action_button_frame, text='Cancel', command=lambda: email_format_toplevel.quit())
    cancel_button.grid(row=3, column=1, pady=10)

    email_format_toplevel.resizable = False
    email_format_toplevel.update()

    email_format_toplevel.focus_set()
    email_format_toplevel.grab_set()
    email_format_toplevel.wait_window()
    email_format_toplevel.grab_release()

# def process_files(update_student_list, update_assignments, assignment_percentage_threshold, gradebook_path):
#     if gradebook_path == 'None Selected':
#         return
#
#     with open(gradebook_path) as gradebook:
#         grade_data = list(csv.reader(gradebook))
#
#     if not os.path.isdir('output'):
#         os.mkdir('output')
#
#     if update_student_list:
#         create_student_list(grade_data)
#
#     if update_assignments:
#         create_assignments_list(grade_data, assignment_percentage_threshold)
#
#
# # def create_student_list(grade_data):
# #     if os.path.exists('./output/studentlist.txt'):
# #         os.remove('./output/studentlist.txt')
# #     student_list = open('./output/studentlist.txt', 'w')
# #     emails_column = grade_data[0].index('Email address')
# #     for row in range(1, len(grade_data)):
# #         email = grade_data[row][emails_column]
# #         at = email.find('@')
# #         username = email[0:at]
# #         student_list.write(username)
# #         if row != len(grade_data)-1:
# #             student_list.write('\n')
# #     student_list.close()
# #
# #
# # def create_assignments_list(grade_data, assignment_percentage_threshold):
# #     if os.path.exists('./output/assignments.csv'):
# #         os.remove('./output/assignments.csv')
# #     assignments_list = open('./output/assignments.csv', 'w')
# #     assignments_list.write('assignment name,percent complete,will be processed\n')
# #     categories = grade_data[0]
# #     emails_column = categories.index('Email address')
# #     for category in range(emails_column+1, len(categories)):
# #         if not contains_excluded_phrase(categories[category]):
# #             percent_complete = calculate_precent_complete(grade_data, category)
# #             will_be_processed = percent_complete >= assignment_percentage_threshold
# #             assignments_list.write('\"'+categories[category]+'\"'+','+str(percent_complete)+','+str(will_be_processed))
# #             if not category == len(categories)-1:
# #                 assignments_list.write('\n')
# #     assignments_list.close()
# #
# #
# # def contains_excluded_phrase(category):
# #     excluded_category_phrases = ['total', 'Last downloaded from this course']
# #     for phrase in excluded_category_phrases:
# #         if phrase in category:
# #             return 1
# #     return 0
# #
# #
# # def calculate_precent_complete(grade_data, assignment_index):
# #     number_complete = 0
# #     for student in range(1, len(grade_data)):
# #         if not (grade_data[student][assignment_index] == 0 or grade_data[student][assignment_index] == '-'):
# #             number_complete = number_complete+1
# #     return round(number_complete/(len(grade_data)-1)*100, 2)  # The only non-student row is the header


main()
