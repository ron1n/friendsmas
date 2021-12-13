import xlrd
import random
import boto3
import os
from botocore.exceptions import ClientError
from jinja2 import Environment, FileSystemLoader, select_autoescape

def render_template_email(**kwargs):
    #Load in the jinja2 things and point it to a template
    env = Environment(
        loader=FileSystemLoader(searchpath="./"),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template("email.html")
    #Send the actual email (jinja2 template) and all of the required variables
    return(template.render(**kwargs))

#Using AWS Simple Email Service to send the emails
def send_ses(to, subj, body):
    CHARSET = "UTF-8"
    SENDER = "nuzzo.salvatore@gmail.com"
    AWS_REGION = "us-east-1"
    client = boto3.client('ses',region_name=AWS_REGION)
    #Try to send the email.
    try:
        #Provide the contents of the email.
        response = client.send_email(
            Destination={
                'ToAddresses': [
                    to,
                ],
            },
            Message={
                'Body': {
                    'Html': {
                        'Charset': CHARSET,
                        'Data': body,
                    },
                    'Text': {
                        'Charset': CHARSET,
                        'Data': body,
                    },
                },
                'Subject': {
                    'Charset': CHARSET,
                    'Data': subj,
                },
            },
            Source=SENDER,
        )
    #Display an error if something goes wrong.	
    except ClientError as e:
        print(e.response['Error']['Message'])
    else:
        print("Email sent! Message ID:"),
        print(response['MessageId'])

#Build a class so we can make friend objects with the following properties
class Friend:
    def __init__(self, name, email, wishlist, blacklist, allergies, pair, pairdex):
        self.name = name
        self.email = email
        self.wishlist = wishlist
        self.blacklist = blacklist
        self.allergies = allergies
        self.pair = pair
        #pairdex is a made up word. It's just the list index of their pair
        #because I really don't want to implement a doubly linked list.
        #Simple and fast as it can be retrieved by index in just O(1) time complexity.
        self.pairdex = pairdex

#Set up a list of friend objects with all the data from the sheet
    def load_friends():
        sheet = load_sheet("friendsmas.xls")
        friend_list = []
        #iterate over this shit and we'll pair em up later
        for i in range(sheet.nrows):
            my_friend = Friend(sheet.cell_value(i, 0),
                    sheet.cell_value(i, 1),
                    sheet.cell_value(i, 2),
                    sheet.cell_value(i, 3),
                    sheet.cell_value(i, 4), 
                    "", 0)
            friend_list.append(my_friend)
        return(friend_list)

#This only needs to be used once. In case Amazon SES is being a BITCH
#about sending emails to unverified addresses
def send_verify(email_recipient):
    client = boto3.client('ses',region_name="us-east-1")
    
    response = client.verify_email_identity(
    EmailAddress = email_recipient
    )
    return(response)

#Load that sheeeeet
def load_sheet(path):
    wb = xlrd.open_workbook(path)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)
    return(sheet)



#Because I'm too lazy to implement a linked list
def match_friends(friend_list):
    #mix em up so its a surprise
    random.shuffle(friend_list)
    i = 0
    while i < len(friend_list) - 1:
        friend_list[i].pair = friend_list[i + 1].name
        #here I go making up words again
        friend_list[i].pairdex = i + 1
        i += 1

    #Quick hack to take the tail of the list and feed it the head as its pair
    friend_list[len(friend_list) - 1].pair = friend_list[0].name
    return friend_list
    

def main():
    #From inside out - load the friends into a list and pass it to the match_friends function so we have a list with pairs
    friend_list = match_friends(Friend.load_friends())
    i = 0  
    while i < len(friend_list):
        to      = friend_list[i].email
        SUBJECT = 'Secret Santa Invitation!'
        
        #Again, only send one verification email when its absolutely necessary. You'll know when you need this.
        #send_verify(to)

        #From inside out - send the vars to the Jinja2 template, render it, and then email it out to folks.
        send_ses(to, SUBJECT,
                render_template_email(
                name        =friend_list[i].name,
                pair        =friend_list[i].pair,
                wishlist    =friend_list[friend_list[i].pairdex].wishlist,
                blacklist   =friend_list[friend_list[i].pairdex].blacklist,
                allergies   =friend_list[friend_list[i].pairdex].allergies)
                )
        i += 1

if __name__ == "__main__":
    main()
