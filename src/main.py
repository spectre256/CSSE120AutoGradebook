# Made by Zach Kelly
# Last updated: 1/9/20
#
# This script finds students
#

import smtplib
from email.mime.text import MIMEText
import csv
import getpass

def process_gradebook():
    # Percentage of students who have had the assignment graded needed for the assignment to be active
    gradebook_file = 'Gradebook.csv'
    with open(gradebook_file) as gradebook:
        grade_data = list(csv.reader(gradebook))
    assignments = find_active_assignments(grade_data)

    active_assignments = assignments[0]
    # for idx in range(len(active_assignments)):
    #     if active_assignments[idx] == 0:
    #         print(grade_data[0][idx])
    # print("\n\n")
    # for idx in range(len(active_assignments)):
    #     if active_assignments[idx] == 1:
    #         print(grade_data[0][idx])

    student_data = target_students(assignments, grade_data[0])
    targeted_students = student_data[0]
    missing_assignments_list = student_data[1]
    print(sum(targeted_students))
    # for idx in range(len(targeted_students)):
    #     if targeted_students[idx] == 1:
    #         print("\n"+(grade_data[idx][0])+" "+(grade_data[idx][1]))
    setup_email(targeted_students, missing_assignments_list, grade_data)


def setup_email(targeted_students, assignments, grade_data):
    server = setup()
    # name = "Zach"
    for student in range(len(targeted_students)):
        # target_email = "zzzz14767@gmail.com"
        # target_email = "kellyzc@rose-hulman.edu"
        if targeted_students[student] == 1:
            name = grade_data[student][0]+" "+grade_data[student][1]
            target_email = grade_data[student][5]
            missing_list = assignments[student-1]
            send_reminder(name, target_email, missing_list, server)
    server.quit()


def target_students(assignments, assignment_names):
    active_assignments = assignments[0]
    assignment_data = assignments[1]

    # with open("assignment_data.csv", "w") as outfile:
    #     for i in range(len(assignment_data)):
    #         for j in range(len(assignment_data[0])):
    #             outfile.write(str(assignment_data[i][j]))
    #             if j != len(assignment_data[0])-1:
    #                 outfile.write(",")
    #         if i != len(assignment_data)-1:
    #             outfile.write("\n")
    targeted_students = [0]
    max_num_missing = 5
    missing_assignments_list = []
    for student in range(len(assignment_data[0])):
        num_missing = 0
        missing_assignments = []
        for assignment in range(len(assignment_data)):
            if active_assignments[assignment] == 1:
                if assignment_data[assignment][student] == 0:
                    num_missing = num_missing + 1
                    missing_assignments.append(assignment_names[assignment])
        missing_assignments_list.append(missing_assignments)
        if num_missing >= max_num_missing:
            targeted_students.append(1)
        else:
            targeted_students.append(0)
    print(targeted_students)
    return targeted_students, missing_assignments_list


def find_active_assignments(grade_data):
    excluded_keywords = ['total', 'code', 'downloaded']
    assignments_first_column = 7
    students_first_row = 2
    active_threshold = 0.75
    active_assignments = []
    assignment_data = []
    num_students = len(grade_data)
    for i in range(assignments_first_column-1):
        active_assignments.append(0)
        submitted = []
        for student in range(students_first_row - 1, num_students):
            submitted.append(0)
        assignment_data.append(submitted)
    for assignment in range(assignments_first_column - 1, len(grade_data[0])):
        skip = 0
        for word in excluded_keywords:
            if word in grade_data[0][assignment]:
                active_assignments.append(0)
                skip = 1
                break
        submitted = []
        if skip == 0:
            for student in range(students_first_row - 1, num_students):
                grade = grade_data[student][assignment]
                # I really doubt anyone will get a zero on a submitted assignment
                if (grade == '-') or (grade == 0):
                    submitted.append(0)
                else:
                    submitted.append(1)
            if sum(submitted) >= num_students * active_threshold:
                active_assignments.append(1)
            else:
                active_assignments.append(0)
        else:
            for student in range(students_first_row - 1, num_students):
                submitted.append(0)
        assignment_data.append(submitted)
    return active_assignments, assignment_data


def setup():
    smtp_server = 'smtp.office365.com'
    port = 587
    my_user = username+"@rose-hulman.edu"
    # my_pass = "password"
    server = smtplib.SMTP(smtp_server, port)
    server.connect(smtp_server, port)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(my_user, my_pass)
    return server


def missing_string(missing_list):
    list_string = ''
    for assignment in missing_list:
        list_string = list_string+str(assignment)+'\n'
    list_string += '\n'
    return list_string


def send_reminder(name, target_email, missing_list, server):
    my_email = username + "@rose-hulman.edu"

    # Uncomment for testing purposes
    if dry_run:
        target_email = "fisherds@gmail.com"
        print("Dry run email sent to:" + target_email + "  ONLY! not the student")


    msg = MIMEText("\nThis email was sent automatically, replies will be ignored.\n"#"\nThis is a test, please check the following list against the gradebook, and reply only if the list is incorrect.\n"
                   + "\nHello "+name+",\n\nYou are currently missing:\n\n"+missing_string(missing_list)
                   + " please submit these as soon as you can. Talk to your professor if you have any questions."
                   + " Remember, you must turn in ALL assignments to pass the class!\n\nThank you,\nZach Kelly\n"
                   + "CSSE120 Grader")
    msg['From'] = my_email
    msg['To'] = target_email
    msg['Subject'] = "Missing CSSE120 Assignments"

    server.sendmail(my_email, target_email, msg.as_string())#target_email, msg.as_string())


dry_run = True
username = "fisherds"
my_pass = getpass.getpass("Enter your password: ")
process_gradebook()
