'''
    A very basic script to connect to my Outlook server and 
    iterate over my emails
'''

import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

LYFT_FOLDER_NAME = "Lyft Ride"

class Oli():
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
    for inx, folder in Oli(mapi.Folders).items():
        for inx ,subfolder in Oli(folder.Folders).items():
            if (subfolder.Name == LYFT_FOLDER_NAME):
                return inx

lyftFolderIndex = getLyftFolderIndex()

lyftRideInbox = mapi.GetDefaultFolder(lyftFolderIndex)

print(lyftRideInbox.Items.GetLast().body.encode("utf-8"))

