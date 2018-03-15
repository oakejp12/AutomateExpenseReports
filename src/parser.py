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
        self.lineFareSearch = r'(fare \(.*\))(.*)\$([0-9]+.[0-9]+)'
        self.discountSearch = r'(Discount)(.*)\$([0-9]+.[0-9]+)'

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
    
    def getLineTotal(self):
        '''
            Calculate the total fare owed at the end of the ride
        '''
        lineFare = self.__searchForFare__(self.lineFareSearch)
        discount = self.__searchForFare__(self.discountSearch)
        if lineFare and discount is None:
            return Decimal(lineFare)
        elif lineFare and discount:
            return Decimal(lineFare) - Decimal(discount)
        else:
            print("Couldn't retrieve values...")


def getTotalExpenses():
    bodies = messageBodies()

    totalsForEachMessage = [Parse(n).getLineTotal() for n in bodies]

    [print(str(total)) for total in totalsForEachMessage]

    print("Total " + str(sum(totalsForEachMessage)))

    return sum(totalsForEachMessage)


if __name__ == "__main__":
    getTotalExpenses()

    