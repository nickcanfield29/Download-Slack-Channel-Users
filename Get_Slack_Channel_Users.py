import requests
from xlsxwriter import Workbook
from tkinter import filedialog
import tkinter.messagebox
from tkinter import *
from pathvalidate import sanitize_filepath

def clean_filename(filename):

    cleaned_filename = sanitize_filepath(filename)
    return cleaned_filename

def get_directory():

    file_path = ''
    count_tries = 0

    root = Tk()
    root.withdraw()
    root.update()

    while file_path == '':
        if count_tries > 3:
            root.destroy()
            print("Sorry. Exiting program.")
            x = input("Press enter to exit program or exit out of this window.")
            exit()

        prompt = "Where do you want to save your file?"
        if count_tries > 0 and count_tries <= 3:
            prompt = "Must select folder"
        file_path = filedialog.askdirectory(title = prompt)
        count_tries += 1
    root.destroy()
    return file_path

def is_matching_channel(channel_to_check):

    if len(CHANNEL_NAME) == 0:
        return True
    for channel_name_options in CHANNEL_NAME:
        if channel_name_options in channel_to_check['name']:
            return True

def getUserMap():

    try:
        all_slack_users_addition_list = requests.get('https://slack.com/api/users.list?token=%s&limit=1000&pretty=1' % (SLACK_API_TOKEN)).json()
        addition_list_members = all_slack_users_addition_list['members']
        cursor_id = all_slack_users_addition_list['response_metadata']['next_cursor']
        number_of_additions = len(addition_list_members)
        print("Connection Successful!")
        print("Grabbing your Slack users. Please wait...")

        num = 0
        while num < number_of_additions:
            member_to_add = addition_list_members[num]
            all_slack_users_list.append(member_to_add)
            num += 1

        while number_of_additions >= 1000:

            try:
                all_slack_users_addition_list = requests.get('https://slack.com/api/users.list?token=%s&cursor=%s&limit=1000&pretty=1' % (SLACK_API_TOKEN, cursor_id)).json()
                addition_list_members = all_slack_users_addition_list['members']

                number_of_additions = len(addition_list_members)
                num = 0
                while num < number_of_additions:
                    member_to_add = addition_list_members[num]
                    all_slack_users_list.append(member_to_add)
                    num += 1

                cursor_id = all_slack_users_addition_list['response_metadata']['next_cursor']
                print("Still grabbing your users. Please wait...")
            except:
                break

    except:
        print("There was an error in grabbing your Slack Users. Sorry!")

    print("Done grabbing your users.")
    print("Found ", len(all_slack_users_list), " users")

def get_channel_users(all_channel_members):

    channel_users = []

    number_of_slack_users = len(all_slack_users_list)
    entry = 0

    print("Matching users with profile info...")
    while entry < number_of_slack_users:
        try:
            if all_slack_users_list[entry]['id'] in all_channel_members:
                channel_users.append(all_slack_users_list[entry])
        except:
            pass
        entry += 1
    return channel_users

######################################################
######################################################
######################################################
#Run Program

def show_output():
    path = get_directory()
    message_output = "ERROR"
    tkinter.messagebox.showinfo('ERROR', message_output)

def open_dialogue_box():

    global api_token
    global desired_filename

    main = Tk()

    Label(main, text="Enter Slack API Token:").grid(row=0)
    Label(main, text="Desired Filename:").grid(row=1)

    api_token = Entry(main)
    desired_filename = Entry(main)

    api_token.grid(row=0, column=1)
    desired_filename.grid(row=1, column=1)

    Button(main, text='Quit', command=main.destroy).grid(row=4, column=0, sticky=W, pady=4)
    Button(main, text='Get Channel Users', command=show_output).grid(row=4, column=1, sticky=W, pady=4)

    mainloop()

if __name__ == "__main__":

    open_dialogue_box()

    global all_slack_users_list
    all_slack_users_list = []

    print("TITLE: Download Slack Members From Specific Channel(s)")
    print("VERSION: 1.0")
    print("COPYRIGHT: Nick Canfield")
    print("SUPPORT: Please contact Nick @ nickcanfieldbiz@gmail.com")
    print("**********************")
    print("INSTRUCTIONS:\nMake sure you have set up a Slack App in your Slack account to get your API token.\nFor a detailed video on how to set up your app, get your API Token, and run this program,\nGo here: https://bit.ly/2XHzLn5")

    print("**********************")
    print("**********************")
    print("Please select the folder where you want to save your Slack Users Excel file")
    path = get_directory()
    print("**********************")
    print("**********************")

    print("Please type in your desired file name that you want to call your file.")
    print('File Name Example: my_slack_channel_users')
    name = input("Desired file name: ")
    if len(name) < 1:
        name = 'my_slack_channel_users'
    name = clean_filename(name)
    print("**********************")
    print("**********************")

    try:
        print('Trying to set up your Excel file')
        fileName = "{parent}/{file}.xlsx".format(parent=path, file=name)
        workbook = Workbook(fileName)
        worksheet = workbook.add_worksheet()
        cell_format_header = workbook.add_format()
        cell_format_header.set_border(1)
        cell_format_header.set_bg_color('green')
        cell_format_normal = workbook.add_format()
        cell_format_normal.set_border(1)
        row = 1
        worksheet.write('A' + str(row), 'CHANNEL_NAME', cell_format_header)
        worksheet.write('B' + str(row), 'USER_NAME', cell_format_header)
        worksheet.write('C' + str(row), 'EMAIL', cell_format_header)
        worksheet.write('D' + str(row), 'CHANNEL_TYPE', cell_format_header)
        worksheet.write('E' + str(row), 'USER_ID', cell_format_header)
        row += 1
        print("Great! We were able to set up an Excel file.")
    except:
        print('Something went wrong!')
        print('Please make sure you have Microsoft Excel or try running the program again with a different folder or file name.')
        input("Press enter to exit program.")
        exit()  # quit Python

    print("**********************")
    print("**********************")
    print("Let's grab your Slack Account App API.")
    slack_api_okay = False
    while slack_api_okay == False:
        SLACK_API_TOKEN = input("Please enter your Slack API Token: ")  # get one from https://api.slack.com/docs/oauth-test-tokens
        SLACK_API_TOKEN = SLACK_API_TOKEN.lstrip().rstrip()
        if len(SLACK_API_TOKEN) > 0:
            slack_api_okay = True
        else:
            print("You need to enter your API key. Please try again...")
            print("**********************")
    print("**********************")
    print("**********************")
    done_entering_channel_names = False
    CHANNEL_NAME = []
    print("Let's grab some channels to look for.")

    print("Please type in the exact channel name, or common text for multiple channel names.\nIf you want to grab channel users for all of your channels, type 'all'.\nWhen done adding channels, type done.")
    channel_additions = 0
    print("**********************")
    while done_entering_channel_names == False:
        if len(CHANNEL_NAME) > 0:
            print("Current Search List:")
            for channel in CHANNEL_NAME:
                print(channel)
        print("**********************")
        if channel_additions >= 1:
            print("Done adding channels? Type 'done'")
        addition = input("Channel or search term: ")
        if (addition.lower().lstrip().rstrip() == 'all' or addition.lower().lstrip().rstrip() == "'all'") and channel_additions == 0:
            print("Let's get all the channels!")
            done_entering_channel_names = True
            channel_additions += 1
        if (addition.lower() == 'done' or addition.lower() == "'done'")  and channel_additions > 0:
            print("Okay. We'll grab users from channels matching the following names: ", CHANNEL_NAME)
            done_entering_channel_names = True
        if addition.lower() == 'done' and channel_additions == 0:
            print("You need to enter a channel first before being done... Try again.")
        if len(addition) > 0 and addition.lower() != 'done' and done_entering_channel_names != True:
            channel_additions += 1
            CHANNEL_NAME.append(addition)
        if len(addition) == 0:
            print("Invalid entry. Try again.")
        print("**********************")

    print("**********************")

    print("Trying to connect to your Slack Account...")
    try:
        channel_list = requests.get('https://slack.com/api/conversations.list?token=%s' % SLACK_API_TOKEN, '&types=public_channel%2Cprivate_channel&pretty=1').json()['channels']
        getUserMap()
    except:
        print('Something went wrong!')
        print("Sorry. Your API Key was incorrect or some server error occurred.")
        print("Please try reconnecting to the internet and verify your Slack API key has not expired.")
        input("Press enter to exit program.")
        exit() # Quit Program

    print("**********************")
    print("List of Channels matching your search term(s): ", CHANNEL_NAME)
    matching_channel_count = 0

    for channel in channel_list:
        if is_matching_channel(channel):
            matching_channel_count += 1
            if channel['is_private'] == True:
                print('Channel Name: ', channel['name'], ' || Channel Type: Private')
            else:
                print('Channel Name: ', channel['name'], ' || Channel Type: Public')
    print("**********************")
    print("**********************")

    user_count = 0
    if matching_channel_count > 0:

        for channel in channel_list:

            if is_matching_channel(channel):
                print("Getting users for channel: ", channel['name'])
                is_private_channel_value = channel['is_private']
                channel_type = 'Public'
                if is_private_channel_value == True:
                    channel_type = 'Private'
                cursor_id = ''
                all_channel_members = []

                try:
                    channel_info_members = requests.get('https://slack.com/api/conversations.members?token=%s&channel=%s&limit=1000&pretty=1' % (SLACK_API_TOKEN, channel['id'])).json()
                    channel_members = channel_info_members['members']
                    cursor_id = channel_info_members['response_metadata']['next_cursor']
                    has_members = False
                    if len(channel_members) > 0:
                        has_members = True
                        for member in channel_members:
                            all_channel_members.append(member)

                        number_of_additions = len(channel_members)
                        print("Found " + str(number_of_additions) + " channel users...")

                    if has_members == True:
                        while number_of_additions >= 1000:
                            print("Grabbing all the channel members....")
                            try:
                                channel_info_members = requests.get('https://slack.com/api/conversations.members?token=%s&channel=%s&cursor=%s&limit=1000&pretty=1' % (SLACK_API_TOKEN, channel['id'], cursor_id)).json()
                                channel_members = channel_info_members['members']
                                cursor_id = channel_info_members['response_metadata']['next_cursor']
                                for member in channel_members:
                                    all_channel_members.append(member)
                                number_of_additions = len(channel_members)
                            except:
                                number_of_additions = 0
                                pass
                    try:
                        users = get_channel_users(all_channel_members)
                        for user in users:
                            user_count += 1
                            user_name = ''

                            try:
                                if user['real_name']:
                                    user_name = user['real_name']
                            except:
                                pass

                            try:
                                worksheet.write('A' + str(row), channel['name'], cell_format_normal)
                                worksheet.write('B' + str(row), user_name, cell_format_normal)
                                worksheet.write('C' + str(row), user['profile']['email'], cell_format_normal)
                                worksheet.write('D' + str(row), channel_type, cell_format_normal)
                                worksheet.write('E' + str(row), user['id'], cell_format_normal)

                            except:

                                try:
                                    worksheet.write('A' + str(row), channel['name'], cell_format_normal)
                                    worksheet.write('B' + str(row), user_name, cell_format_normal)
                                    worksheet.write('C' + str(row), 'NO EMAIL', cell_format_normal)
                                    worksheet.write('D' + str(row), channel_type, cell_format_normal)
                                    worksheet.write('E' + str(row), user['id'], cell_format_normal)

                                except:

                                    try:
                                        worksheet.write('A' + str(row), channel['name'], cell_format_normal)
                                        worksheet.write('B' + str(row), 'NAME CANNOT BE WRITTEN', cell_format_normal)
                                        worksheet.write('C' + str(row), user['profile']['email'], cell_format_normal)
                                        worksheet.write('D' + str(row), channel_type, cell_format_normal)
                                        worksheet.write('E' + str(row), user['id'], cell_format_normal)

                                    except:
                                        worksheet.write('A' + str(row), channel['name'], cell_format_normal)
                                        worksheet.write('B' + str(row), 'NAME CANNOT BE WRITTEN', cell_format_normal)
                                        worksheet.write('C' + str(row), 'EMAIL CANNOT BE WRITTEN', cell_format_normal)
                                        worksheet.write('D' + str(row), channel_type, cell_format_normal)
                                        worksheet.write('E' + str(row), user['id'], cell_format_normal)

                            row += 1
                        print("Wrote " + channel['name'] + "'s " + str(len(users)) + " members to your file.")
                        print("**********************")
                    except:
                        print("Couldn't write members to your file")
                        print("**********************")
                except:
                    print("Couldn't grab channel members.")

        # Close the file
        print("**********************")
        print("Program Ran Successfully!")
        print(user_count, " channel users added to file: " + fileName)
        workbook.close()
        print("**********************")
        print("**********************")
    else:
        print("No channels matched your search. Please rerun the program")
        workbook.close()

    print("Program Completed.")
    x = input("Press enter to exit program or exit out of this window.")
    exit()
    # END PROGRAM
