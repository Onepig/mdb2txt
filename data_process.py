# -*- coding: cp936 -*-

import pypyodbc,os
import re, datetime

def printPath(level, path):
    '''
    ��ӡһ��Ŀ¼�µ������ļ��к��ļ�
    '''
    # �����ļ��У���һ���ֶ��Ǵ�Ŀ¼�ļ���
    dirList = []
    # �����ļ�
    fileList = []
    # ����һ���б������а�����Ŀ¼��Ŀ������(google����)
    files = os.listdir(path)
    # ������Ŀ¼����
    dirList.append(str(level))
    for f in files:
        if(os.path.isdir(path + '\\' + f)):
            # �ų������ļ��С���Ϊ�����ļ��й���
            if(f[0] == '.'):
                pass
            else:
                # ���ӷ������ļ���
                dirList.append(f)
        if(os.path.isfile(path + '\\' + f)):
            # �����ļ�
            fileList.append(f)
    # ��һ����־ʹ�ã��ļ����б���һ�����𲻴�ӡ
    for fl in fileList:
        # ��ӡ�ļ�
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
            # ��ӡ������̨�����ǵ�һ����Ŀ¼
            print'\n' 
            print '-' * (int(dirList[0])), dl
            # ��ӡĿ¼�µ������ļ��к��ļ���Ŀ¼����+1
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
    

    
    

