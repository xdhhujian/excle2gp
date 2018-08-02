################################################读取excel文件及数据###################################
#########读取excel文件
import xlrd
def open_excel(file='file.xlsx'):
    data = xlrd.open_workbook(file)
    return data
########根据索引获取Excel表格中的数据   参数:  file：Excel文件路径     by_index：表的索引
import datetime
from matplotlib.dates import num2date
def excel_read_byindex(file='file.xlsx', by_index=2):
    data = open_excel(file)
    table = data.sheets()[by_index]
    rows = table.nrows  # 行数
    cols = table.ncols  # 列数
    all_content = []
    for i in range(rows):
        row_content=[]
        for j in range(cols):
            cell = table.cell_value(i, j)
            ctype = table.cell(i, j).ctype
            if ctype == 3:
                startDt = datetime.date(1899, 12, 31).toordinal() - 1
                cell = num2date(startDt + cell).strftime('%Y-%m-%d %H:%M:%S')
            row_content.append(cell)
        all_content.append(row_content)
    return all_content
print(excel_read_byindex(file='E:/gmf/BJ/cs/5gzqh_gzetf.xlsx', by_index=2))
############根据文件名获取excel表中的数据  参数：file：文件路径   by_name:表名
import datetime
from matplotlib.dates import num2date
def excel_read_byname(file='file.xlsx', by_name='cs1'):
    data = open_excel(file)
    table = data.sheet_by_name(by_name)
    rows = table.nrows  # 行数
    cols = table.ncols  # 列数
    all_content = []
    for i in range(rows):
        row_content=[]
        for j in range(cols):
            cell = table.cell_value(i, j)
            ctype = table.cell(i, j).ctype
            if ctype == 3:
                startDt = datetime.date(1899, 12, 31).toordinal() - 1
                cell = num2date(startDt + cell).strftime('%Y-%m-%d %H:%M:%S')
            row_content.append(cell)
        all_content.append(row_content)
    return all_content
print(excel_read_byname(file='E:/gmf/BJ/cs/5gzqh_gzetf.xlsx', by_name='cs1'))
def main():
    tables = excel_read_byindex('E:/gmf/BJ/CS/5gzqh_gzetf.xlsx')
    for row in tables:
        print(row)
if __name__=='__main__':
    main()
