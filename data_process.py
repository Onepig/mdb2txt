# -*- coding: cp936 -*-

import pypyodbc,os
import re
from multiprocessing import Pool
import datetime

def printPath(level, mdbdir):
    '''
    打印一个目录下的所有文件夹和文件
    '''
    # 所有文件夹，第一个字段是次目录的级别
    dirList = []
    # 所有文件
    fileList = []
    # 返回一个列表，其中包含在目录条目的名称(google翻译)
    files = os.listdir(mdbdir)
    # 先添加目录级别
    dirList.append(str(level))
    for f in files:
        if(os.path.isdir(mdbdir + '\\' + f)):
            # 排除隐藏文件夹。因为隐藏文件夹过多
            if(f[0] == '.'):
                pass
            else:
                # 添加非隐藏文件夹
                dirList.append(f)
        if(os.path.isfile(mdbdir + '\\' + f)):
            # 添加文件
            
            pat=re.compile(r'(.*)\.mdb')
            m=pat.match(f)
            if m:
                print '-' * (int(dirList[0])), f
                fileList.append(f)

    # 当一个标志使用，文件夹列表第一个级别不打印
##    for fl in fileList:
##        # 打印文件
##        pat=re.compile(r'(.*)\.mdb')
##        m=pat.match(fl)
##        if m:
##            print '-' * (int(dirList[0])), fl
##            mdb2txt(mdbdir, txtdir, fl)

    i_dl = 0
    for dl in dirList:
        if(i_dl == 0):
            i_dl = i_dl + 1
        else:
            # 打印至控制台，不是第一个的目录
            print'\n' 
            print '-' * (int(dirList[0])), dl
            # 打印目录下的所有文件夹和文件，目录级别+1
            printPath((int(dirList[0]) + 1), mdbdir + '\\' + dl)
    return fileList


        


def mdb2txt(filename):
    mdbdir=r'E:\\Study\\Hadoop\\data\\card\\'
    txtdir=r'E:\\Study\\Hadoop\\data\\card\\txt\\'
    tablename=filename[0:-4]
    print 'tablename:\t%s'%tablename

    txtname=txtdir+filename[0:-3]+r'txt'
    print 'txtname:\t%s'%txtname

    if os.path.exists(txtname):
        os.remove(txtname)
    f=open(txtname,'a')
    
    connection_string = 'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s'%(mdbdir+filename)
##    print connection_string
    conn = pypyodbc.connect(connection_string)
    cur = conn.cursor()
    '''
##    s0="SELECT NAME FROM MSYSOBJECTS WHERE TYPE=1 AND FLAGS=0"
##    s0="SELECT table_name FROM %s..Sysobjects Where XType='s' ORDER BY Name"%(filename[0:-4])
##    s0="select * from tableName"
##    s0="select * from MSysObjects"
##    s0="select * from information_schema.tables"
##    s0="select name from %s..Sysobjects where type='table' order by name"
##    print s0
##    cur.execute(s0)
##    for line in cur:
##        print line'''

    
    s1="select * from %s"%tablename
####    print s1
    cur.execute(s1)
    for row in cur:
##        f.write(str(row))
        for i in range(len(row)-1):
            f.write(str(row[i])+'\t')
        f.writetr(row[len(row)-1])+'\n')

        
    cur.close()
    conn.close()
    f.close()
    print '\n'

if __name__=='__main__':
    start=datetime.datetime.now()

    global mdbdir
    mdbdir=r'E:\\Study\\Hadoop\\data\\card\\'
    global txtdir
    txtdir=r'E:\\Study\\Hadoop\\data\\card\\txt\\'
    if os.path.exists(txtdir):
        pass
    else:
        os.mkdir(txtdir)
    
    mdbfile=printPath(1, mdbdir)
    
    pool=Pool()
    pool.map(mdb2txt, mdbfile)
    pool.close()
    pool.join()
    
    end=datetime.datetime.now()
    print 'Total time: %s'%(end-start)

    
    

    
    


