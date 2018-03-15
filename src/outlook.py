'''
    A very basic script to connect to my Outlook server and 
    iterate over my emails
'''

import win32com.client
import win_unicode_console
import sys

win_unicode_console.enable() # NOTE: There's a bug related to stdout streams that errors out the program

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

# Make this configurable through CLI options
LYFT_FOLDER_NAME = "Lyft Ride"
OUTLOOK_DOMAIN = "joakes@silasg.com"

class Oli():
    '''
        Helper class to iterate over Outlook objects
    '''

    def __init__(self, outlook_object):
        self._obj = outlook_object

    def items(self):
        array_size = self._obj.Count
        for item_index in range(1, array_size):
            yield (item_index, self._obj[item_index])

    def prop(self):
        return sorted( self._obj._prop_map_get_.keys() )

def getLyftSubFolder(lyftFolders, month = None):
    '''
        Iterate over Lyft's monthly subfolders
        to find the month we need to report from
    '''

    if month is None:
        month = getMonthToReport()

    for inx, subfolder in Oli(lyftFolders.Folders).items():
        print(subfolder.Name)
        if (subfolder.Name == month):
            return subfolder


def getMonthToReport():
    '''
        Find the month we need to report on
    '''
    import datetime
    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    return months[datetime.datetime.now().month - 1]

def readMessages(monthlySubFolder):
    '''
        Read the messages for the respective month to report
        and return an array of their string bodies
    '''

    messages = monthlySubFolder.Items
    message = messages.GetFirst()

    totalMessagesRead = 1 # Keep track of the number of messages read for verification

    messageBodies = []

    while message:
        messageBodies.append(message.Body)
        message = messages.GetNext()
        totalMessagesRead += 1
    
    print("Total messages found: %i" % totalMessagesRead)

    return messageBodies

def messageBodies():
    topLevelFolder = mapi.Folders[OUTLOOK_DOMAIN] # Retrieve the main outlook folder
    
    lyftRideInbox = topLevelFolder.Folders[LYFT_FOLDER_NAME] # Retrieve the folder corresponding to the 

    monthlySubFolder = getLyftSubFolder(lyftRideInbox)

    return readMessages(monthlySubFolder)

if __name__ == "__main__":
    messageBodies()


