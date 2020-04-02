import requests
from os.path import join

print("*******************")
print("Download Slack Members From Specific Channel(s)")
print("*******************")
print("\n")

print("What folder do you want to save your Slack Users in on your computer?")
print('Folder Path Example:', '"C:/Users/User/Downloads"')
path = input("Please paste the folder location: ") #Put the Path of where you want the file to be saved here
print("*******************")
print("What do you want the file name to be called?")
print('File Name Example: "SlackUsers"')
name = input("Please write your desired new file name: ")  # Name of text file that you want your EO Slack Members to be saved in
print("*******************")
print("*******************")

try:
    file_object = open(join(path, (name + '.txt')), 'w')  # Trying to create a new file or open one
except:
    print('Something went wrong!')
    print('Try running the program again with a different folder or file name')
    exit()  # quit Python


def is_matching_channel(channel_to_check):
    if CHANNEL_NAME in channel_to_check['name']:
        return True

print("Let's connect to your Slack Account")
print("*******************")
SLACK_API_TOKEN = input("Please paste in your Slack API Token: ")  # get one from https://api.slack.com/docs/oauth-test-tokens
print("*******************")
CHANNEL_NAME = input("Please type in the exact channel name, or common text for multiple channel names, to search in: ")
print("******************")
print("*******************")
print("*******************")
print("Trying to connect to your Slack Acount...")
try:
    channel_list = requests.get('https://slack.com/api/channels.list?token=%s' % SLACK_API_TOKEN).json()['channels']
    users_list = requests.get('https://slack.com/api/users.list?token=%s' % SLACK_API_TOKEN).json()['members']
except:
    print("Sorry. Your API Key was incorrect or some server error occured.")
    print("Please try reconnecting to the internet and verify your Slack API key has not expired.")
    exit() # Quit Python

print("CONNECTION SUCCESSFUL!")
print("**********************")
print("LIST OF", CHANNEL_NAME, " CHANNELS")
print("**********************")

for channel in channel_list:

    if is_matching_channel(channel):
        print(channel['name'])
        
print("**********************")
print("**********************")

print('Creating your new file')
file_object.write("CHANNEL NAME"+","+"USER NAME"+","+"EMAIL")
file_object.write("\n")
user_count = 0

for channel in channel_list:
    #channel = find_channel(channel_list)

    if is_matching_channel(channel):
        channel_info = requests.get('https://slack.com/api/channels.info?token=%s&channel=%s' % (SLACK_API_TOKEN, channel['id'])).json()['channel']
        members = channel_info['members']
        users = filter(lambda u: u['id'] in members, users_list)
        for user in users:
            user_count += 1
            first_name, last_name = '', ''

            try:
                if user['real_name']:
                    first_name = user['real_name']
            except:
                pass

            try:
                file_object.write(channel['name']+","+first_name+","+user['profile']['email'])
            except:
                try:
                    file_object.write(channel['name'] + "," + first_name + "," + "NO EMAIL")
                except:
                    try:
                        file_object.write(channel['name'] + "," + "NAME CANNOT BE WRITTEN" + "," +user['profile']['email'])
                    except:
                        file_object.write(channel['name'] + "," + "NAME CANNOT BE WRITTEN" + "," +"EMAIL CANNOT BE WRITTEN")
            file_object.write("\n")

# Close the file
print("**********************")
print("**********************")
print("Program Ran Successfully!")
print(user_count, " channel users added to file: ", (path, (name + '.txt')))
file_object.close()
print("Program Completed")
