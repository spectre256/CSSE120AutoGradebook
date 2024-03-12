# Made by Zach Kelly
# Last updated 1/24/20

from common_code import *
import tkinter
from tkinter import ttk, filedialog
import csv
import os

defaults = extract_defaults()


def main():

    root = tkinter.Tk()
    root.title('Prepare Input Files')
    main_frame = ttk.Frame(root)
    main_frame.pack(fill='both', padx=10, pady=25)

    ##############
    # Checkbuttons
    ##############

    # Checkbutton frame
    checkbutton_frame = ttk.Frame(main_frame)
    checkbutton_frame.grid_columnconfigure(0, weight=1)
    checkbutton_frame.grid_columnconfigure(1, weight=1)
    checkbutton_frame.pack(fill='both', pady=10)#, ipadx=100)

    # First checkbutton and label
    update_assignments_label = ttk.Label(checkbutton_frame, text='Update Assignments?', anchor=tkinter.CENTER)
    update_assignments_label.grid(row=0, column=1)

    update_assignments_value = tkinter.BooleanVar()
    # The update assignments checkbutton toggles the percentage threshold scale
    percent_threshold_scale_active = int(defaults['Update Assignments'])
    update_assignments_value.set(int(defaults['Update Assignments']))

    def toggle_percent_threshold_scale_active(*args):
        nonlocal percent_threshold_scale_active
        nonlocal percent_threshold_value_label
        percent_threshold_scale_active = not percent_threshold_scale_active
        if not percent_threshold_scale_active:
            percent_threshold_scale.set(int(defaults['Default Percentage Threshold']))
            percent_threshold_value_label.config(text=str(int(defaults['Default Percentage Threshold'])) + '%')

    update_assignments_value.trace('w', toggle_percent_threshold_scale_active)

    update_assignments_checkbutton = ttk.Checkbutton(checkbutton_frame, variable=update_assignments_value)
    update_assignments_checkbutton.grid(row=1, column=1)

    # Second checkbutton and label
    student_list_label = ttk.Label(checkbutton_frame, text='Update Student List?', anchor=tkinter.CENTER)
    student_list_label.grid(row=0, column=0)

    update_student_list_value = tkinter.BooleanVar()
    update_student_list_value.set(int(defaults['Update Student List']))
    update_student_list_checkbutton = ttk.Checkbutton(checkbutton_frame, variable=update_student_list_value)
    update_student_list_checkbutton.grid(row=1, column=0)

    #############################
    # Percent threshold selection
    #############################

    # Percent threshold frame
    threshold_frame = ttk.Frame(main_frame)
    threshold_frame.grid_columnconfigure(0, weight=1)
    threshold_frame.grid_columnconfigure(1, weight=1)
    threshold_frame.pack(fill='both', pady=20)

    # Percent threshold scale and labels
    percent_threshold_value_label = ttk.Label(threshold_frame, font='Helvetica 18 bold', anchor=tkinter.CENTER, width=5)
    percent_threshold_value_label.grid(row=0, column=1, rowspan=2)

    percent_threshold_value = tkinter.IntVar()
    percent_threshold_scale = ttk.Scale(threshold_frame, from_=0, to=100,
                                        variable=percent_threshold_value, command=lambda value:
                                        percent_threshold_value_label.config(text=str(int(float(value))) + '%'))
    percent_threshold_scale.set(int(defaults['Default Percentage Threshold']))
    percent_threshold_scale.grid(row=1)

    def percent_threshold_state_helper(*args):
        if not percent_threshold_scale_active:
            percent_threshold_value.set(int(defaults['Default Percentage Threshold']))

    percent_threshold_value.trace('w', percent_threshold_state_helper)

    percentage_threshold_text_label = ttk.Label(threshold_frame, text='Assignment Percentage Threshold:', anchor=tkinter.CENTER)
    percentage_threshold_text_label.grid(row=0, column=0)

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
    gradebook_path_label1 = ttk.Label(gradebook_path_selection_frame, text='Current Gradebook Path:', anchor=tkinter.CENTER)
    gradebook_path_label1.grid(row=0, column=0, sticky='ew')

    gradebook_path_label2 = ttk.Label(gradebook_path_selection_frame, textvariable=gradebook_file,
                                      font='Helvetica 8 bold', width=75, anchor=tkinter.CENTER)
    gradebook_path_label2.grid(row=1, columnspan=3, pady=20)

    # TODO: add error handling
    gradebook_full_path = defaults['Gradebook Path']
    gradebook_path_button = ttk.Button(gradebook_path_selection_frame, text='Select File',
                                       command=lambda: gradebook_file.set(get_gradebook_path()))
    gradebook_path_button.grid(row=0, column=1)

    def get_gradebook_path():
        nonlocal gradebook_full_path
        max_path_characters = 70
        gradebook_path = tkinter.filedialog.askopenfilename(filetypes=[('csv', '.csv')], initialdir='.')
        if gradebook_path == '':
            gradebook_path = defaults['Gradebook Path']
        gradebook_full_path = gradebook_path
        if len(gradebook_path) > max_path_characters:
            gradebook_path = '...'+gradebook_path[len(gradebook_path)-max_path_characters+3:len(gradebook_path)]
        return gradebook_path

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
        root.quit()
        process_files(update_student_list_value.get(), update_assignments_value.get(), percent_threshold_value.get(),
                      gradebook_full_path)

    root.resizable(width=False, height=False)
    root.mainloop()


def process_files(update_student_list, update_assignments, assignment_percentage_threshold, gradebook_path):
    if gradebook_path == defaults['Gradebook Path']:
        return

    with open(gradebook_path) as gradebook:
        grade_data = list(csv.reader(gradebook))

    if not os.path.isdir('output'):
        os.mkdir('output')

    if update_student_list:
        create_student_list(grade_data)

    if update_assignments:
        create_assignments_list(grade_data, assignment_percentage_threshold)


def create_student_list(grade_data):
    if os.path.exists('./output/studentlist.txt'):
        os.remove('./output/studentlist.txt')
    student_list = open('./output/studentlist.txt', 'w')
    emails_column = grade_data[0].index('Email address')
    for row in range(1, len(grade_data)):
        email = grade_data[row][emails_column]
        at = email.find('@')
        username = email[0:at]
        student_list.write(username)
        if row != len(grade_data)-1:
            student_list.write('\n')
    student_list.close()


def create_assignments_list(grade_data, assignment_percentage_threshold):
    if os.path.exists('./output/assignments.csv'):
        os.remove('./output/assignments.csv')
    assignments_list = open('./output/assignments.csv', 'w')
    assignments_list.write('assignment name,percent complete,will be processed\n')
    categories = grade_data[0]
    emails_column = categories.index('Email address')
    for category in range(emails_column+1, len(categories)):
        if not contains_excluded_phrase(categories[category], defaults):
            percent_complete = calculate_percent_complete(grade_data, category)
            will_be_processed = percent_complete >= assignment_percentage_threshold
            assignments_list.write('\"'+categories[category]+'\"'+','+str(percent_complete)+','+str(will_be_processed))
            if not category == len(categories)-1:
                assignments_list.write('\n')
    assignments_list.close()


main()
