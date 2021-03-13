import win32com.client as client
import json

outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')
account = namespace.Folders['email@account']

inbox = account.Folders['Inbox']
messages = inbox.Items

# Direct Search
def search_mail(inbox):
    fileOpen = open("myJson.json", "r")
    details = json.loads(fileOpen.read())

    done = 0
    found = 0
    foundTotal = 0
    for item in inbox.Items:
        done += 1
        if done%100 == 0:
            print("Done" + str(done))
            print("Found" + str(found))
            print()
            foundTotal += found
            found = 0


        for i in range (0, len(details)):
            if details[i][1] == False:
                toSearch = details[i][0]
                if toSearch in item.body:
                    details[i][1] = True
                    found += 1

    fileOpen = open("mailDone.json", "w+")
    fileOpen.write(json.dumps(details))

    print(foundTotal)


# Collect all mail into a string (json file) and then compare. (fast)
def make_mail_body_string(inbox):
    done = 0
    myStr = ""
    for item in inbox.Items:
        done += 1
        if done%100 == 0:
            print("Done = " + str(done))

        myStr = myStr + item.body + " " + item.subject + " "

    myJson = {
        "all": myStr
    }

    fileOpen = open("mailString.json", "w+")
    fileOpen.write(json.dumps(myJson))


make_mail_body_string(inbox)