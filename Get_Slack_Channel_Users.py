import requests
import xlsxwriter

print("*******************")
print("Download Slack Members From Specific Channel(s)")
print("*******************")
print("\n")

print("What folder do you want to save your Slack Users in on your computer?")
print('Folder Path Example:', '"C:/Users/User/Downloads"')
path = input("Please paste the folder location: ") #Put the Path of where you want the file to be saved here
print("*******************")
print("What do you want the Excel file name to be called?")
print('File Name Example: "SlackUsers"')
name = input("Please write your desired new file name: ")  # Name of text file that you want your EO Slack Members to be saved in
print("*******************")
print("*******************")

try:
    print('Trying to create your new file')
    # Excel
    fileName = "{parent}/{file}.xlsx".format(parent=path, file=name)
    workbook = xlsxwriter.Workbook(fileName)
    worksheet = workbook.add_worksheet()
    row = 1
    worksheet.write('A' + str(row), 'CHANNEL_NAME')
    worksheet.write('B' + str(row), 'USER_NAME')
    worksheet.write('C' + str(row), 'EMAIL')
    worksheet.write('D' + str(row), 'CHANNEL_TYPE')
    row += 1
except:
    print('Something went wrong!')
    print('Try running the program again with a different folder or file name')
    exit()  # quit Python


def is_matching_channel(channel_to_check):
    if len(CHANNEL_NAME) == 0:
        return True
    for channel_name_options in CHANNEL_NAME:
        if channel_name_options in channel_to_check['name']:
            return True

def get_channel_users(all_channel_members, all_slack_users_list):
    channel_users = []
    num_of_sets = len(all_channel_members)
    for user in all_slack_users_list:
        set = 0
        while set < num_of_sets:
            if user['id'] in all_channel_members[set]:
                channel_users.append(user)
            set += 1
    return channel_users

print("Let's grab your Slack Account API")
print("*******************")
slack_api_okay = False
while slack_api_okay == False:
    SLACK_API_TOKEN = input("Please paste in your Slack API Token: ")  # get one from https://api.slack.com/docs/oauth-test-tokens
    if len(SLACK_API_TOKEN) > 0:
        slack_api_okay = True
    else:
        print("You need to enter your API key. Please try again...")
        print("*******************")
print("*******************")
done_entering_channel_names = False
CHANNEL_NAME = []
print("Please type in the exact channel name, or common text for multiple channel names, from which o grab users\nIf you want all channels, type all.\nIf done adding channels, type done.")
channel_additions = 0
print("******************")
while done_entering_channel_names == False:
    if channel_additions >= 1:
        print("Done adding channels? Type 'done'")
    addition = input("Channel or search term: ")
    if (addition.lower() == 'all' or addition.lower() == 'all ') and channel_additions == 0:
        print("Let's get all the channels!")
        done_entering_channel_names = True
        channel_additions += 1
    if (addition.lower() == 'done' or addition.lower() == "'done'")  and channel_additions > 0:
        print("Okay. We'll grab channels matching the following names: ", CHANNEL_NAME)
        done_entering_channel_names = True
    if addition.lower() == 'done' and channel_additions == 0:
        print("You need to enter a channel first before being done... Try again.")
    if len(addition) > 0 and addition.lower() != 'done' and done_entering_channel_names != True:
        channel_additions += 1
        CHANNEL_NAME.append(addition)
    if len(addition) == 0:
        print("Invalid entry. Try again.")
    print("******************")

print("******************")
print("*******************")
print("*******************")
print("Trying to connect to your Slack Account...")
try:
    channel_list = requests.get('https://slack.com/api/conversations.list?token=%s' % SLACK_API_TOKEN, '&types=public_channel%2Cprivate_channel&pretty=1').json()['channels']
    all_slack_users_list = requests.get('https://slack.com/api/users.list?token=%s' % SLACK_API_TOKEN).json()['members']
except:
    print("Sorry. Your API Key was incorrect or some server error occured.")
    print("Please try reconnecting to the internet and verify your Slack API key has not expired.")
    exit() # Quit Program

print("CONNECTION SUCCESSFUL!")
print("**********************")
print("List of Channels matching the search term(s): ", CHANNEL_NAME)
print("**********************")
for channel in channel_list:

    if is_matching_channel(channel):
        if channel['is_private'] == True:
            print('Channel Name: ', channel['name'], ' || Channel Type: Private', )
        else:
            print('Channel Name: ', channel['name'], ' || Channel Type: Public', )
print("**********************")
print("**********************")

user_count = 0

for channel in channel_list:

    if is_matching_channel(channel):
        print("Getting users for channel: ", channel['name'])
        is_private_channel_value = channel['is_private']
        channel_type = 'Public'
        if is_private_channel_value == True:
            channel_type = 'Private'
        cursor_id = ''
        all_channel_members = []
        channel_info_members = requests.get('https://slack.com/api/conversations.members?token=%s&channel=%s&limit=1000&pretty=1' % (SLACK_API_TOKEN, channel['id'])).json()


        try:
            channel_members = channel_info_members['members']
            all_channel_members.append(channel_members)
            cursor_id = channel_info_members['response_metadata']['next_cursor']


            while channel_info_members['response_metadata']['next_cursor'] != '':
                channel_info_members = requests.get('https://slack.com/api/conversations.members?token=%s&channel=%s&cursor=%s&limit=1000&pretty=1' % (SLACK_API_TOKEN, channel['id'], cursor_id)).json()
                channel_members = channel_info_members['members']
                all_channel_members.append(channel_members)
                cursor_id = channel_info_members['response_metadata']['next_cursor']

            users = get_channel_users(all_channel_members, all_slack_users_list)
            for user in users:
                user_count += 1
                user_name = ''

                try:
                    if user['real_name']:
                        user_name = user['real_name']
                except:
                    pass

                try:
                    worksheet.write('A' + str(row), channel['name'])
                    worksheet.write('B' + str(row), user_name)
                    worksheet.write('C' + str(row), user['profile']['email'])
                    worksheet.write('D' + str(row), channel_type)

                except:

                    try:
                        worksheet.write('A' + str(row), channel['name'])
                        worksheet.write('B' + str(row), user_name)
                        worksheet.write('C' + str(row), 'NO EMAIL')
                        worksheet.write('D' + str(row), channel_type)

                    except:

                        try:
                            worksheet.write('A' + str(row), channel['name'])
                            worksheet.write('B' + str(row), 'NAME CANNOT BE WRITTEN')
                            worksheet.write('C' + str(row), user['profile']['email'])
                            worksheet.write('D' + str(row), channel_type)

                        except:
                            worksheet.write('A' + str(row), channel['name'])
                            worksheet.write('B' + str(row), 'NAME CANNOT BE WRITTEN')
                            worksheet.write('C' + str(row), 'EMAIL CANNOT BE WRITTEN')
                            worksheet.write('D' + str(row), channel_type)

                row += 1
        except:
            pass

# Close the file
print("**********************")
print("**********************")
print("Program Ran Successfully!")
print(user_count, " channel users added to file: ", (path, (name + '.xlsx')))
workbook.close()
print("Program Completed")
