# -*- coding: cp936 -*-

import pypyodbc,os
import re, datetime

def printPath(level, path):
    '''
    打印一个目录下的所有文件夹和文件
    '''
    # 所有文件夹，第一个字段是次目录的级别
    dirList = []
    # 所有文件
    fileList = []
    # 返回一个列表，其中包含在目录条目的名称(google翻译)
    files = os.listdir(path)
    # 先添加目录级别
    dirList.append(str(level))
    for f in files:
        if(os.path.isdir(path + '\\' + f)):
            # 排除隐藏文件夹。因为隐藏文件夹过多
            if(f[0] == '.'):
                pass
            else:
                # 添加非隐藏文件夹
                dirList.append(f)
        if(os.path.isfile(path + '\\' + f)):
            # 添加文件
            fileList.append(f)
    # 当一个标志使用，文件夹列表第一个级别不打印
    for fl in fileList:
        # 打印文件
        pat=re.compile(r'(.*)\.mdb')
        m=pat.match(fl)
        if m:
            print '-' * (int(dirList[0])), fl
            mdb2txt(path,fl)

    i_dl = 0
    for dl in dirList:
        if(i_dl == 0):
            i_dl = i_dl + 1
        else:
            # 打印至控制台，不是第一个的目录
            print'\n' 
            print '-' * (int(dirList[0])), dl
            # 打印目录下的所有文件夹和文件，目录级别+1
            printPath((int(dirList[0]) + 1), path + '\\' + dl)


        


def mdb2txt(path,filename):


    tablename=filename[0:-4]
    print 'tablename:\t%s'%tablename
    
    txtname=path+'txt\\'+filename[0:-3]+r'txt'
    print 'txtname:\t%s'%txtname
    
##    if os.path.exists(txtname):
##        os.remove(txtname)
##    f=open(txtname,'a')
    
    connection_string = 'Driver={Microsoft Access Driver (*.mdb)};DBQ=%s'%(path+filename)
    print connection_string
    conn = pypyodbc.connect(connection_string)
    cur = conn.cursor()
    s1="select * from %s"%tablename
    print s1
    cur.execute(s1)
    for row in cur:
        l=len(row)
        print l
##        for field in row:
##            f.write(str(field)+'\t')
##        f.write('\n')
        
    cur.close()
    conn.close()
##    f.close()

if __name__=='__main__':
    printPath(1, r'E:\Study\Hadoop\data\card')
    

    
    


