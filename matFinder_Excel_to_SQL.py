import xlrd

name= 'amonia1'

## Abrir o excel

book = xlrd.open_workbook(name+'.xlsx')
sheet = book.sheet_by_index(0)

list_1 = []
list_2 = []


for c in range(sheet.nrows):
    for k in range(sheet.ncols):
        valorcelula=sheet.cell_value(c,k)
        list_2.append(valorcelula)
    list_1.append(list_2)
    list_2=[]


file = open(name+"111.sql","w")
file.write("INSERT INTO points"+"\n")
file.write("(material_id, reference_id, point_property, username, value, temperature, hx, hy, hz, ex, ey, ez, pressure)"+"\n")
file.write("VALUES"+"\n")


for ind in list_1:
    ind1="(" + str(ind[0]) + "," + str(ind[1]) + "," +"'"+ str(ind[2])+"'" + "," + "'"+str(ind[3]) +"'"+ "," + str(ind[4]) + "," + str(ind[5]) + "," + str(ind[6]) + "," + str(ind[7]) + "," + str(ind[8]) + "," + str(ind[9]) + "," + str(ind[10]) + "," + str(ind[11]) + "," + str(ind[12]) + ")"+","
    file.write(ind1+"\n")
