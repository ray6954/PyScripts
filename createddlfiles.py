# coding: utf-8

import sys
import xlrd  # xls:rw,xlsx:r
#import re

# -- インデックスを設定する。ゼロオリジン。-- #

SHEET_INDEX = 2      # テーブル定義シート

COLUMN_INDEX = {
'COL_TBLNAME'    : 3      # テーブル名
,'COL_COLNAME'   : 9      # カラム名
,'COL_COLTYPE'   : 14     # カラムデータ型
,'COL_COLDIGIT'  : 15     # カラム桁
,'COL_PRECISION' : 16     # カラム精度
,'COL_PK'        : 10     # PK
,'COL_NOTNULL'   : 11     # 必須
,'COL_DEFAULT'   : 19     # デフォルト値
}

ROW_FROM = 9      # 読込開始行番号

# -------------------------------------------- #



def write_ddl(rows):

    SPACE = '        '  # 半角空白8文字
    
    tablename = rows[0]['COL_TBLNAME'].value
    output = []
    pk = ',CONSTRAINT PK_%s PRIMARY KEY(' % tablename
    pkcols = ''
    
    output.append('CREATE TABLE %s (' % tablename)
    
    line = ''
    for i,row in enumerate(rows):
        if i != 0:
            line = SPACE + ','
        else:
            line = SPACE
        
        line = line + row['COL_COLNAME'].value
        line = line + '  ' + row['COL_COLTYPE'].value
        
        if 0 == row['COL_COLDIGIT'].ctype:
            pass
        else:
            line = line + '(%s' % int(row['COL_COLDIGIT'].value)
            if 0 == row['COL_PRECISION'].ctype:
                line = line + ')'
            else:
                line = line + ',%s)' % int(row['COL_PRECISION'].value)

        if 'Yes' == row['COL_PK'].value:
            print  row['COL_PK'].value
            if 0 == len(pkcols):
                pkcols = row['COL_COLNAME'].value
            else:
                pkcols = pkcols + ',' + row['COL_COLNAME'].value
            print row['COL_COLNAME'].value
            
        if 0 == row['COL_DEFAULT'].ctype:
            pass
        elif 1 == row['COL_DEFAULT'].ctype:
            line = line + ' DEFAULT ' + "'%s'" % str(row['COL_DEFAULT'].value)
        elif 2 == row['COL_DEFAULT'].ctype:
            line = line + ' DEFAULT ' + str(int(row['COL_DEFAULT'].value))

        if 'Yes' == row['COL_NOTNULL'].value:
            line = line + ' ' + 'NOT NULL'
        
        output.append(line)
            
    if 0 != len(pkcols):
        pk = SPACE + pk + pkcols + ')'
        output.append(pk)

    output.append(')')

    file = open("%s.ddl" % tablename,"w")
    for l in output:
        file.write(l + '\n')
    file.close;
        


def create_files(filename):    

    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(SHEET_INDEX)

    row_to = sheet.nrows  # 最終行番号+1
    
    extable = ''
    cols = {} 
    rows = []
    for n in range(ROW_FROM,row_to):
        if ROW_FROM == n:
            extable = sheet.cell_value(n,COLUMN_INDEX['COL_TBLNAME'])
        else:
            if extable != sheet.cell_value(n,COLUMN_INDEX['COL_TBLNAME']):
                # テーブルが変わったので前テーブルのddl作成
                write_ddl(rows)
                rows[:] = []

        # 1行分の列を取得
        cols = {}
        for colname, index in COLUMN_INDEX.iteritems():
            cols[colname] = sheet.cell(n,index)

        rows.append(cols)
        extable = sheet.cell_value(n,COLUMN_INDEX['COL_TBLNAME'])
    
    # 最終テーブルの処理
    write_ddl(rows)



if __name__ == '__main__':

    argvs = sys.argv
    argc = len(argvs)

    if (argc != 2):
        print u'起動パラメータを設定してください。。。'
        print 'ex. > python %s filename' % argvs[0]
        quit()
    
    print '////// START //////'
    create_files(argvs[1])
    print '////// END //////'