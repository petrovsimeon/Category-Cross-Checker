'''This is a script that checks if a company offers a category of services by cross checking if this company is already
present in another companies list'''
print('\nWelcome to my cross checking script')

# Imports
import pandas as pd
import xlsxwriter
import time
print('\nLibraries imported')

#Timer - shows how fast the script runs.
start = time.time()

# Output file
output_file = xlsxwriter.Workbook('RESULTS.xlsx')
worksheet = output_file.add_worksheet()
row = 0
col = 0

# Input files
# Companies
companies_file = pd.read_excel('Solicitors ALL CATEGORIES.xlsx')
companies = list(companies_file['Company'].values)

#Category files
cat_file1 = pd.read_csv('1. Buying_Selling_Properties.csv')
category_1 = list(cat_file1['Name'].values)

cat_file2 = pd.read_csv('2. Commercial Conveyancy.csv')
category_2 = list(cat_file2['Name'].values)

categories_list = [category_1, category_2]

for company in companies:
    row += 1
    col = 0

    worksheet.write (row, col, company)
    col +=1

    for category in categories_list:

        if company in category:
            worksheet.write(row, col, 'Yes')
            col += 1

        else:
            worksheet.write(row, col, 'No')
            col += 1

#End of script
output_file.close()
print('\nAll done!')
end = time.time()
print("\nTotal time for running:")
print(end - start)