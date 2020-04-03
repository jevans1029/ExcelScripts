import xlrd
import xlwt
import os

#remove extension from filename
originalfilename = ""
xlsx = ".xlsx"
original = originalfilename + xlsx
newfile = "New " + originalfilename + ".xls"
try:
    os.remove(newfile)
except:
    print("no new file")
col1=0
col2=1
newcol1=0
newcol2=1


"""This function takes an excel spreadsheet and loops through two columns and matches each column entry to 
    every entry in the other column. The original file must be of file type .xlsx and the output file has extension .xls. 
    It first deletes any file with the target file name. 
    The table in the original file must start at row 1 with the column titles as the first row. The value to group by should 
    be in column 1. 
    """
#use 0 index values for column
def forEach(col1, col2, newcol1, newcol2):

    wb = xlrd.open_workbook(original)
    sheet = wb.sheet_by_index(0)
    rows = sheet.nrows
    list1 = []
    list2 = []
    col1_title = sheet.cell_value(0, col1)
    col2_title = sheet.cell_value(0, col2)
    if col1_title=='' or col2_title=='':
        raise ValueError("No values in row 1. The Table must start in row 1 with the column titles")

    #generate lists of each column
    for i in range(rows):
        list1.append(sheet.cell_value(i, col1))
        list2.append(sheet.cell_value(i, col2))
    list1 = list1[1:]
    list2 = list2[1:]
    wb.release_resources()

    nwb = xlwt.Workbook()
    nws = nwb.add_sheet('Sheet 1')
    nws.write(0, newcol1, col1_title)
    nws.write(0, newcol2, col2_title)

    newcells = []
    #generate a list of the new cells
    for i in range(len(list1)):
        for l in range(len(list2)):
            tuple = (list1[i], list2[l])
            newcells.append(tuple)
    #write the new cells
    for row in range(len(newcells)):
        nws.write(row+1, newcol1, newcells[row][0])
        nws.write(row+1, newcol2, newcells[row][1])

    nwb.save(newfile)

if __name__ == "__main__":
    forEach(col1, col2, newcol1, newcol2)
