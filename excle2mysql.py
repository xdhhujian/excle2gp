#-*-coding:gbk -*-

'''
Usage:
   excle2mysql.py -fcsnm <file> <column> <sheet> <table> <start>
   excle2mysql.py -ft <file> <type>
   excle2mysql.py -dt <dir> <type>
   excle2mysql.py -ftn <file> <type> <table>
   excle2mysql.py -dtn <dir> <type> <table>
   excle2mysql.py -fcs <file> <column> <sheet>
Options:
   -h --help  查看帮助
   -f         文件
   -d         目录
   -t         类型
   -n         表名
   -c         列名
   -s         excle sheet 名称
   -m         start data

Example:

    excle2mysql.py -fcsnm d:/asd/q 0 0 szidc.tablename 3

"""
'''

import sqlite3
import pymysql
import xlrd
import datetime
from matplotlib.dates import num2date
from docopt import docopt


#获取excle中的数据,file传入文件名,第一列为字段名，by_name为子sheet名称
#这里固定了子sheet的名称
def excle_table_byindex(file='D:\\test1\\dir_a.xlsx',colnameindex=0,sheet_index=0,start_row=0):
    excle_handle=xlrd.open_workbook(file)
    #print(len(excle_handle.sheet_names()))
    #for i in range(0,len(excle_handle.sheet_names())):
    data = excle_handle.sheet_by_index(sheet_index)#获取sheet内容
    nrows = data.nrows#数据行数
    colnames = data.row_values(colnameindex)
    list = []
    for rownum in range(start_row,nrows):
        row = data.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                cell = data.cell_value(rownum,i)
                cell_type = data.cell(rownum,i).ctype
                if cell_type == 3:
                    startDt = datetime.date(1899,12,31).toordinal()-1
                    cell = num2date(startDt+cell).strftime('%Y-%m-%d %H:%M:%S')
                    app[colnames[i].replace('/','')] = cell
                else:
                    app[colnames[i].replace('/','')] = row[i]
            list.append(app)
    return list


def open_mysql(host = '192.168.3.16',user = 'root',password = 'hujian0605',db = 'test_db',port = 3306,charset = 'utf8'):
    try:
        conn = pymysql.connect(host = '192.168.3.16',user = 'root',password = 'hujian0605',db = 'test_db',port = 3306,charset = 'utf8')
    except Exception as e:
        print("connect error")
    return conn


def get_input():
    arguments = docopt(__doc__)
    print(arguments)
    try:
        file= arguments['<file>']
    except Exception as e:
        file=""
    try:
        type =arguments['<type>']
    except Exception as e:
        type=""
    try:
        dir = arguments['<dir>']
    except Exception as e:
        dir=""
    try:
        table_name=arguments['<table>']
    except Exception as e:
        table_name=""
    print(file,type)
    try:
        column_id=arguments['<column>']
    except Exception as e:
        column_id=""
    try:
        sheet_id=arguments['<sheet>']
    except Exception as e:
        sheet_id=""
    try:
        start_row=arguments['<start>']
    except Exception as e:
        start_row=""

    return dir,file,type,table_name,column_id,sheet_id,start_row

def table_exists(cur,table_name):
    try:
        sql_text = "select 1 from " + table_name +" limit 1"
        cur.execute(sql_text)
        exists = True
    except Exception as e:
       if str(e).endswith("doesn't exist\")"):
           exists = False
       else:
           exists = 'other'
    return exists



def main():
     dir,file,type,table_name,column_id,sheet_id,start_row=get_input()
     if file:
         try:
            datas = excle_table_byindex(file,int(column_id),int(sheet_id),int(start_row))
            print(datas)
         except Exception as e:
             e
         conn = open_mysql()
         cur = conn.cursor()
         hava_table = table_exists(cur,table_name)
         if hava_table == True:
             for data in datas:
                 COLstr = ''
                 ROWstr = ''
                 for key in data.keys():
                     COLstr = COLstr + key + ','
                     ROWstr = (ROWstr.replace("%","")+'"%s"'+',')%(data[key])
                 COLstr = COLstr[:-1]
                 ROWstr = ROWstr[:-1]
                 insert_sql = "insert into " + table_name +" values(" + ROWstr +");"
                 print(insert_sql)
                 cur.execute(insert_sql)
             cur.close()
             conn.commit()
             conn.close()
         elif hava_table == False:
             print("表不存在,请先创建表")
         else:
             print("其他原因导致失败")

if __name__=="__main__":
    main()