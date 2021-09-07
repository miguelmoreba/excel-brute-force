import time
from itertools import permutations
import io
import msoffcrypto
import openpyxl
import pandas
import os

# decrypted_workbook = io.BytesIO()
# filename = "/Users/miguelmorenobaladron/code/brute-force-excel/sample1.xlsx"
# print(decrypted_workbook.read())

# with open(filename, 'rb') as file:
#     office_file = msoffcrypto.OfficeFile(file)
#     office_file.load_key(password='cd')
#     office_file.decrypt(decrypted_workbook)
#     df = pandas.read_excel(decrypted_workbook, engine="openpyxl")
#     print(df)


with open(f'./maybe/yes', 'w'):
    pass