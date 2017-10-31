'''This is a script that checks if a company offers a category of services by cross checking if this company is already
present in another companies list'''

# Imports
import pandas as pd
import xlsxwriter

# Output File
output_file = xlsxwriter.Workbook('Categories found.xlsx')
worksheet = output_file.add_worksheet()

def creating_output_file():
    global output_file
    global worksheet
    global col
    global row
    output_file = xlsxwriter.Workbook ('Results - '+str(search_term)+ ' in '+ str(location)+'.xlsx')
    worksheet = output_file.add_worksheet ()
    bold = output_file.add_format({'bold': True})

    worksheet.write (row, col, 'Company Name', bold)
    col = 1
    worksheet.write (row, col, 'Address', bold)
    col = 2
    worksheet.write (row, col, 'Phone number', bold)
    col = 3
    worksheet.write (row, col, 'Type of business', bold)
    col = 4
    worksheet.write (row, col, 'Number of reviews', bold)
    col = 5
    worksheet.write (row, col, 'Rating', bold)
    col = 6
    worksheet.write (row, col, 'Description', bold)
    col = 7
    worksheet.write (row, col, 'Price Range', bold)
    row += 1

# Input files
# Files with categories
all_companies = pd.read_excel('Solicitors ALL CATEGORIES.xlsx')
# category1 = pd.read_csv('category.csv')
# category2 = pd.read_csv('category.csv')
# category3 = pd.read_csv('category.csv')
# category4 = pd.read_csv('category.csv')
# category5 = pd.read_csv('category.csv')
# category6 = pd.read_csv('category.csv')
# category7 = pd.read_csv('category.csv')
# category8 = pd.read_csv('category.csv')
# category9 = pd.read_csv('category.csv')
# category10 = pd.read_csv('category.csv')
# category11 = pd.read_csv('category.csv')
# category12 = pd.read_csv('category.csv')
# category13 = pd.read_csv('category.csv')
# category14 = pd.read_csv('category.csv')
# category15 = pd.read_csv('category.csv')
# category16 = pd.read_csv('category.csv')
# category17 = pd.read_csv('category.csv')
# category18 = pd.read_csv('category.csv')
# category19 = pd.read_csv('category.csv')
# category20 = pd.read_csv('category.csv')
# category21 = pd.read_csv('category.csv')
# category22 = pd.read_csv('category.csv')
# category23 = pd.read_csv('category.csv')
# category24 = pd.read_csv('category.csv')
# category25 = pd.read_csv('category.csv')
# category26 = pd.read_csv('category.csv')
# category27 = pd.read_csv('category.csv')
# category28 = pd.read_csv('category.csv')
# category29 = pd.read_csv('category.csv')
# category30 = pd.read_csv('category.csv')
# category31 = pd.read_csv('category.csv')
# category32 = pd.read_csv('category.csv')
# category33 = pd.read_csv('category.csv')
# category34 = pd.read_csv('category.csv')
# category35 = pd.read_csv('category.csv')
# category36 = pd.read_csv('category.csv')
# category37 = pd.read_csv('category.csv')
# category38 = pd.read_csv('category.csv')
# category39 = pd.read_csv('category.csv')
# category40 = pd.read_csv('category.csv')
# category41 = pd.read_csv('category.csv')
# category42 = pd.read_csv('category.csv')
# category43 = pd.read_csv('category.csv')
# category44 = pd.read_csv('category.csv')
# category45 = pd.read_csv('category.csv')
# category46 = pd.read_csv('category.csv')
# category47 = pd.read_csv('category.csv')
# category48 = pd.read_csv('category.csv')
# category49 = pd.read_csv('category.csv')
# category50 = pd.read_csv('category.csv')
# category51 = pd.read_csv('category.csv')
# category52 = pd.read_csv('category.csv')
# category53 = pd.read_csv('category.csv')
# category54 = pd.read_csv('category.csv')
# category55 = pd.read_csv('category.csv')
# category56 = pd.read_csv('category.csv')
# category57 = pd.read_csv('category.csv')
# category58 = pd.read_csv('category.csv')
# category59 = pd.read_csv('category.csv')
# category60 = pd.read_csv('category.csv')
# category61 = pd.read_csv('category.csv')
# category62 = pd.read_csv('category.csv')
# category63 = pd.read_csv('category.csv')

# List of category files
# categories = [category1, category2, category3, category4, category5, category6, category7,
# category8, category9, category10, category11,category12, category13, category14, category15,
# category16, category17, category18, category19,category20, category21, category22, category23,
# category24, category25, category26, category27,category28, category29, category30, category31,
# category32, category33, category34, category35,category36, category37, category38, category39,
# category40, category41, category42, category43,category44, category45, category46, category47,
# category48, category49, category50, category51,category52, category53, category54, category55,
# category56, category57, category58, category59,category60, category61, category62, category63]

#Variables needed
row = 1
col = 0
# Getting companies list

companies = all_companies['Company'].values

for company in companies:
    row += 1

    for category in categories:
        if company in category:
