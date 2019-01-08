import pygsheets
import datetime
import pprint
import time
import yagmail
import credentials

# get timestamp for log
temp_timestamp = str(datetime.datetime.now())
print(2 * "\n")
print(temp_timestamp)

is_leaving_staff = False

# setup credentials for sending email
gmail_user = credentials.gmail_user
gmail_password = credentials.gmail_password
yag = yagmail.SMTP(gmail_user, gmail_password)

# authorize pygsheets to work with sheets
gc = pygsheets.authorize()

# open form responses
sh = gc.open("Staff Exit Form Entry (Responses)")
wks = sh.worksheet_by_title("Form Responses 1")

# # ONLY FOR TESTING PURPOSES!!!! --------------------------------------------
# wks.update_value('F2', '')


print("checking new staff form entries")
# get all values from data entry form
values_mat = wks.get_all_values(returnas="matrix")
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


for staff in data:

    if staff["Sheet Setup"] == "":

        # switch is_leaving_staff for email purposes
        is_leaving_staff = True
        print("beginning process for exiting staff member" + staff["Staff Member"])

        # grab master copy of new staff sheet
        orignial_exit_sheet = gc.open("Original Staff Exit Sheet")
        original_worksheet = orignial_exit_sheet.worksheet_by_title("Original")

        # grab master sheet to add to
        master_sh = gc.open("Staff Exit Form")
        master_list = master_sh.worksheet_by_title("Master")

        # find first empty line on master sheet
        master_mat = master_list.get_all_values(
            returnas="matrix"
        )  # get all values from master worksheet
        occupied_master_rows = [
            x for x in master_mat if x[0] != ""
        ]  # remove empty rows from list
        first_empty_master = len(occupied_master_rows) + 1

        # # ONLY FOR TESTING PURPOSES!!!! --------------------------------------------
        # disposable_sheet = master_sh[1]
        # master_sh.del_worksheet(disposable_sheet)
        # time.sleep(3)

        # add copy of original worksheets
        staff_member_wks = master_sh.add_worksheet(
            staff["Staff Member"], src_worksheet=original_worksheet
        )

        # move new worksheet to first positoin behind master list
        staff_member_wks.index = 1

        # fill in staff info on new worksheet
        staff_member_wks.update_value("C2", staff["Staff Member"])
        staff_member_wks.update_value("C3", staff["Exit Date"])
        staff_member_wks.update_value("C4", staff["Position"])

        # add staff info to master list to check completion on staff_member_wks
        link_sheets = "='" + staff["Staff Member"] + "'!"
        name_column = "A" + str(first_empty_master)
        status_column = "B" + str(first_empty_master)
        admin_column = "D" + str(first_empty_master)
        office_column = "E" + str(first_empty_master)
        admin_ass_column = "F" + str(first_empty_master)
        tech_spec_column = "G" + str(first_empty_master)
        tech_int_column = "H" + str(first_empty_master)
        master_list.update_value(name_column, staff["Staff Member"])
        master_list.update_value(status_column, link_sheets + "C5")
        master_list.update_value(admin_column, link_sheets + "E8")
        master_list.update_value(office_column, link_sheets + "E16")
        master_list.update_value(admin_ass_column, link_sheets + "E28")
        master_list.update_value(tech_spec_column, link_sheets + "E33")
        master_list.update_value(tech_int_column, link_sheets + "E42")

        # add x to 'sheet setup' column on form response sheed to mark staff as initiated
        sheet_setup_location = "F" + str(staff["id"])
        wks.update_value(sheet_setup_location, "X")
        print("Spreadsheets setup complete")

        # begin email notifications
        contents = "A new staff member, {}, was added to the Staff Exit Process spreadsheet, go and check it out. \n\n".format(
            staff["Staff Member"]
        )
        html = '<a href="https://docs.google.com/spreadsheets/d/1kgLv2h_TWmb9FzBmDAe6dJDw5mkz-itbCrt69PdxO3c/edit#gid=0">Staff Exit Process spreadsheet</a>'
        yag.send(
            [
                "rgregory@fnwsu.org",
                "jlaroche@fnwsu.org",
                "jjennett@fnwsu.org",
                "dstamour@fnwsu.org",
                "dtessier@fnwsu.org",
                "mellis@fnwsu.org",
                "clongway@fnwsu.org",
            ],
            "Notification of Employee Exiting",
            [contents, html],
        )
        print("sent exit notification emails for {}".format(staff["Staff Member"]))

if is_leaving_staff == False:
    print("No staff exits to report on")
else:
    print("Finished")
