import serpscrap
import selenium
import xlrd
import xlsxwriter

a=0
index=0

## Abrir o excel

book = xlrd.open_workbook('a.xlsx')
sheet = book.sheet_by_index(0)

list_j = []

for k in range(1,sheet.nrows):
    list_j.append(str(sheet.row_values(k)[0]))
total =len(list_j)
print(total)

## Fazer o scrapping

keywords = list_j
url = []

for b in keywords:
    q=0
    b= b + ' contact'
    config = serpscrap.Config()
    config.set('scrape_urls', False)
    
    scrap = serpscrap.SerpScrap()
    scrap.init(config=config.get(), keywords=b)
    results = scrap.run()
    
    for result in results:
        if result['serp_rank']==1 and q==0:
            url.append([result['serp_url'], result['serp_snippet'], result['serp_title']])
            q=1
    index+=1
    print(str(index) +'/'+ str(total))

## Criar o excel e escrever

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('EmpresasScrap.xlsx')

worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
Dados = url

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for link, descricao, titulo in (Dados):
    worksheet.write(row, col, keywords[row])
    worksheet.write(row, col+1,     link)
    worksheet.write(row, col + 2, descricao)
    worksheet.write(row, col + 3, titulo)
    row += 1
    
workbook.close()
