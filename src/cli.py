'''
    Read from the command line and decide what to do...
'''

import argparse
import sys

parser = argparse.ArgumentParser("Retrieve Lyft expenses from Outlook receipts.")

# Establish possible arguments
parser.add_argument('-f', '--folder', help='Folder name where monthly subfolders reside.', dest="LYFT_FOLDER_NAME", default="Lyft Ride", type=str)
parser.add_argument('-d', '--domain', help='Outlook domain (e.g. joakes@silasg.com)', dest="OUTLOOK_DOMAIN", default="joakes@silasg.com", type=str)
parser.add_argument('-m', '--month', help='Month to retrieve receipts for.', dest="MONTH", type=str)

class Arguments():

    def __init__(self):
        args = self.__getArguments__()
        self.LYFT_FOLDER_NAME = args.LYFT_FOLDER_NAME
        self.OUTLOOK_DOMAIN = args.OUTLOOK_DOMAIN
        self.MONTH = args.MONTH

    def __getArguments__(self):
        return parser.parse_args()

    def getMonth(self):
        return self.MONTH
    
    def getFolderName(self):
        return self.LYFT_FOLDER_NAME

    def getOutlookDomain(self):
        return self.OUTLOOK_DOMAIN

if __name__ == "__main__":
    args = Arguments()
    print(args.getOutlookDomain())