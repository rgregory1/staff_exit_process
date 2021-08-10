import pygsheets
import datetime
import pprint
import time
import yagmail
import credentials

# setup credentials for sending email
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP(gmail_user, gmail_password)

# get timestamp for log
temp_timestamp = str(datetime.datetime.now())
print(2 * "\n")
print(temp_timestamp)

# authorize pygsheets to work with sheets
gc = pygsheets.authorize()

# grab master sheet to get information from
master_sh = gc.open("Staff Exit Form")
master_list = master_sh.worksheet_by_title("Master")

# get all values from data entry form
values_mat = master_list.get_all_values(returnas="matrix")
# remove empty rows from list
exit_staff_list = [x for x in values_mat if x[0] != ""]


# add id's corrosponding to row number from spreadsheet
for x, line in enumerate(exit_staff_list):
    line.insert(0, x + 1)
# change initial row to id for zipping
exit_staff_list[0][0] = "id"

# create dictionary to work with from now on
keys = exit_staff_list[0]
data = [dict(zip(keys, values)) for values in exit_staff_list[1:]]

# initialize final strings
final_admin_todo = ""
final_office_todo = ""
final_admin_ass_todo = ""
final_tech_sup_todo = ""
final_tech_int_todo = ""

# being loop looking for incomplete staff memmbers
for staff in data:
    if staff["Status"] == "Not Done":

        # print name for log
        print(
            "{} information is incomplete, gathering info for emails".format(
                staff["Staff Name"]
            )
        )
        # get staff members info
        this_staff_sheet = master_sh.worksheet_by_title(staff["Staff Name"])
        this_staff_matrix = this_staff_sheet.get_all_values(returnas="matrix")

        # Initialize dict for recording status
        counter = 1
        this_staff_data = {}

        # feed this_staff_matrix into dict for processing, adding spreadsheet row numbers
        for line in this_staff_matrix:
            this_line_data = {}
            # this_line_data['row'] = counter
            this_line_data["a"] = line[0]
            this_line_data["b"] = line[1]
            this_staff_data[counter] = this_line_data
            counter = counter + 1
        # pprint.pprint(this_staff_data)

        # begin admin email notifications
        admin_list = [9, 10, 11, 12, 13, 14]
        admin_todo = ""
        for number in admin_list:
            # print(this_staff_data[number])
            if this_staff_data[number]["a"] == "":
                admin_todo = admin_todo + this_staff_data[number]["b"] + "\n"
        if admin_todo != "":
            final_admin_todo = (
                final_admin_todo + staff["Staff Name"] + "\n \n" + admin_todo + "\n\n"
            )
            print(staff["Staff Name"])
            print("\n" + admin_todo)

        # begin office manager notifications
        office_list = [17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27]
        office_todo = ""
        for number in office_list:
            # print(this_staff_data[number])
            if this_staff_data[number]["a"] == "":
                office_todo = office_todo + this_staff_data[number]["b"] + "\n"
        if office_todo != "":
            final_office_todo = (
                final_office_todo + staff["Staff Name"] + "\n \n" + office_todo + "\n\n"
            )
            print(staff["Staff Name"])
            print("\n" + office_todo)

        # begin Admin Assistant notifications
        admin_ass_list = [30, 31, 32]
        admin_ass_todo = ""
        for number in admin_ass_list:
            # print(this_staff_data[number])
            if this_staff_data[number]["a"] == "":
                admin_ass_todo = admin_ass_todo + this_staff_data[number]["b"] + "\n"
        if admin_ass_todo != "":
            final_admin_ass_todo = (
                final_admin_ass_todo
                + staff["Staff Name"]
                + "\n \n"
                + admin_ass_todo
                + "\n\n"
            )

        # begin tech support notifications
        tech_sup_list = [35, 36, 37, 38, 39, 40, 41]
        tech_sup_todo = ""
        for number in tech_sup_list:
            # print(this_staff_data[number])
            if this_staff_data[number]["a"] == "":
                tech_sup_todo = tech_sup_todo + this_staff_data[number]["b"] + "\n"
        if tech_sup_todo != "":
            final_tech_sup_todo = (
                final_tech_sup_todo
                + staff["Staff Name"]
                + "\n \n"
                + tech_sup_todo
                + "\n\n"
            )

        # begin tech int notifications
        tech_int_list = [44, 45]
        tech_int_todo = ""
        for number in tech_int_list:
            # print(this_staff_data[number])
            if this_staff_data[number]["a"] == "":
                tech_int_todo = (
                    tech_int_todo + "- " + this_staff_data[number]["b"] + "\n"
                )
        if tech_int_todo != "":
            final_tech_int_todo = (
                final_tech_int_todo
                + staff["Staff Name"]
                + "\n \n"
                + tech_int_todo
                + "\n\n"
            )

print("final admin todo follows")
print(final_admin_todo)

print("final office todo follows")
print(final_office_todo)

# begin email notifications

contents = "This is your friendly weekly reminder of things to do for exiting staff memmbers. \n \n \n"
contents2 = (
    "Due to your efficiency, there is actually nothing for you to do for exiting staff!"
)
html = '<a href="https://docs.google.com/spreadsheets/d/1kgLv2h_TWmb9FzBmDAe6dJDw5mkz-itbCrt69PdxO3c/edit#gid=0">Staff Exit Form spreadsheet</a>'

# Admin emails
if final_admin_todo != "":
    yag.send(
        ["Justina.Jennett@mvsdschools.org", "christopher.dodge@mvsdschools.org"],
        "Staff Exit Weekly Reminder",
        [contents, final_admin_todo, html],
    )
else:
    yag.send(
        ["Justina.Jennett@mvsdschools.org", "christopher.dodge@mvsdschools.org"],
        "Staff Exit Weekly Reminder",
        [contents, contents2, html],
    )
print("admin emails sent")

# office manager emails
if final_office_todo != "":
    yag.send(
        "Tanya.Racine@mvsdschools.org",
        "Staff Exit Weekly Reminder",
        [contents, final_office_todo, html],
    )
else:
    yag.send(
        "Tanya.Racine@mvsdschools.org",
        "Staff Exit Weekly Reminder",
        [contents, contents2, html],
    )
print("office manager emails sent")

# admin assistant emails
if final_admin_ass_todo != "":
    yag.send(
        ["dawn.tessier@mvsdschools.org", "Mary.Ellis@mvsdschools.org"],
        "Staff Exit Weekly Reminder",
        [contents, final_admin_ass_todo, html],
    )
else:
    yag.send(
        ["dawn.tessier@mvsdschools.org", "Mary.Ellis@mvsdschools.org"],
        "Staff Exit Weekly Reminder",
        [contents, contents2, html],
    )
print("admin assistant emails sent")

# tech support emails
if final_tech_sup_todo != "":
    yag.send(
        "josh.laroche@mvsdschools.org",
        "Staff Exit Weekly Reminder",
        [contents, final_tech_sup_todo, html],
    )
else:
    yag.send(
        "josh.laroche@mvsdschools.org",
        "Staff Exit Weekly Reminder",
        [contents, contents2, html],
    )
print("tech support emails sent")

# tech integration emails
if final_tech_int_todo != "":
    yag.send(
        "josh.laroche@mvsdschools.org",
        "Staff Exit Weekly Reminder",
        [contents, final_tech_int_todo, html],
    )
else:
    yag.send(
        "josh.laroche@mvsdschools.org",
        "Staff Exit Weekly Reminder",
        [contents, contents2, html],
    )
print("tech int emails sent")
