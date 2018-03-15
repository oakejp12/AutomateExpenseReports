'''
    A very basic script to connect to my Outlook server and 
    iterate over my emails
'''

import win32com.client
import win_unicode_console

win_unicode_console.enable() # NOTE: There's a bug related to stdout streams that errors out the program

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

LYFT_FOLDER_NAME = "Lyft Ride"

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

def getLyftFolderIndex():
    '''
        Iterate over all Outlook folders (top level)
        and find the Lyft Ride index
    '''
    for inx, folder in Oli(mapi.Folders).items(): # Corresponds to user name in outlook (e.g. joakes@example.com)
        for inx, subfolder in Oli(folder.Folders).items(): # Grab actual top-level folders for user
            print("(%i) " % inx + "" + folder.Name + " => " + subfolder.Name)
            if (subfolder.Name == LYFT_FOLDER_NAME):
                print("Index (%i) " % inx + " for " + subfolder.Name)
                return inx

def getLyftSubFolder(lyftFolders):
    '''
        Iterate over Lyft's monthly subfolders
        to find the month we need to report from
    '''
    month = getMonthToReport()

    print("Retrieving " + str(lyftFolders) + " subfolder for %s" % month)

    for inx, subfolder in Oli(lyftFolders.Folders).items():
        if (subfolder.Name == month):
            print("Found (%s) " % month)
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
    '''

    messages = monthlySubFolder.Items
    message = messages.GetNext()
    
    totalMessagesRead = 1 # Keep track of the number of messages read for verification

    print(message.Body)

    # while message:
    #     print(message.Body)
    #     message = messages.GetNext()
    #     totalMessagesRead += 1
    
    print("Total messages found: %i" % totalMessagesRead)


def main():
    lyftFolderIndex = getLyftFolderIndex()

    topLevelFolder = mapi.Folders[1] # Retrieve the main outlook folder
    
    lyftRideInbox = topLevelFolder.Folders[lyftFolderIndex] # Retrieve the folder corresponding to the 

    monthlySubFolder = getLyftSubFolder(lyftRideInbox)

    readMessages(monthlySubFolder)


if __name__ == "__main__":
    main()


