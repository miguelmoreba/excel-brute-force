import win32com.client
import sys
from itertools import permutations
import time

start = time.time()

xlApp = win32com.client.Dispatch("Excel.Application")
filename = "C:\source\Excel.Pass\sample1.xlsx"
chars = 2

perms = [''.join(p) for p in permutations('abcdefghijklmnopqrstuwxyz', chars)]
print(f'Starting to try passwords with {chars} number of chars')

for random_pass in perms:
    try:
        xlApp.Workbooks.Open(filename, False, True, None, random_pass)
        print(f'The correct password is {random_pass}')
        break;
    except:
        if 'The password you supplied is not correct' in str(sys.exc_info()[1]):
            print(f'Password {random_pass} is incorrect')
        # pass

end = time.time()

print(f'Time elapsed {end-start}')

# result = xlApp.Workbooks.Open('C:\source\Excel.Pass\sample1.xlsx', False, True, None, 'add') 