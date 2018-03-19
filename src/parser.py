'''
    Parse the Outlook message bodies into
    a readable format.
'''

from outlook import messageBodies
import re
from decimal import Decimal

class Parse():
    
    def __init__(self, body):
        self.body = body
        self.visaSearch     = r'(Visa \*[0-9]+)(.*)\$([0-9]+.[0-9]+)'

    def __searchForFare__(self, searchString):
        '''
            Search for the fare string in the message body:
            Line fare
            Lyft Line Discount 
        '''
        fareObj = re.search(searchString, self.body)
        if fareObj:
            lineFare = fareObj.group(3) 
            return lineFare
        else:
            print("No fare found: " + searchString) # TODO: Throw an error here? Maybe not since discounts won't be found
            #print(self.body)
    
    def getFareTotal(self):
        '''
            Calculate the total fare owed at the end of a ride
        '''
        totalFare = self.__searchForFare__(self.visaSearch)
        if totalFare:
            return Decimal(totalFare)
        else:
            print("Couldn't retrieve values...")

def getTotalExpenses():
    '''
        Retrieve the total spent from each ride for a month
        This is stored in the body of the email
    '''

    bodies = messageBodies()

    totalsForEachMessage = [Parse(n).getFareTotal() for n in bodies]

    [print(str(total)) for total in totalsForEachMessage]

    print("Total " + str(sum(totalsForEachMessage)))

    return sum(totalsForEachMessage)

if __name__ == "__main__":
    getTotalExpenses()

    