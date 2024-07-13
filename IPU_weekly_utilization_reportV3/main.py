# Ctrl + Shift + Esc

import win32com.client as win32
from index import index, getxlApp
import time

startTime = time.time()

for i in range(1,11):
    time.sleep(4)
    print('Retry Number - {}'.format(i))
    try:
        excel = getxlApp()
        index()
        excel.Quit()
    except Exception as error:
        print(error)
        print("Error PLP :",type(error).__name__)
        if(type(error).__name__=='com_error'):
            print('\nBug')
        if(type(error).__name__=='PermissionError'):
            time.sleep(2)
            excel = getxlApp()
            excel.Quit()
            print('\nPermission Issue (Missing Data)')
        if(type(error).__name__=='TypeError'):
            print('\nData Async (len Issue) - Closing the Excel Sheet')
            print('Retry Number - {}'.format(i))
    else:
        print('--------------Executed Successfully--------------')
        break
        
print('----- Completed in %s seconds -----' % (time.time()-startTime))
