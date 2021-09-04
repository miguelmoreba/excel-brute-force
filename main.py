import time
from itertools import permutations
import io
import msoffcrypto

start = time.time()

filename = "/Users/miguelmorenobaladron/code/brute-force-excel/sample1.xlsx"
chars = 2

perms = [''.join(p) for p in permutations('abcdefghijklmnopqrstuwxyz', chars)]

print(f'Starting to try passwords with {chars} number of chars')

for random_pass in perms:
    try:
        # xlApp.Workbooks.Open(filename, False, True, None, random_pass)
        decrypted_workbook = io.BytesIO()

        with open(filename, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=random_pass)
            office_file.decrypt(decrypted_workbook)
            print(f'The correct password is {random_pass}')
            break;
    except:
        print(f'Password {random_pass} is incorrect')

        pass

end = time.time()

print(f'Time elapsed {end-start}')