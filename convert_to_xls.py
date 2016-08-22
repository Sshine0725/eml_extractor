
# -*- coding: utf-8 -*-
#!/usr/bin/env python
import sys, re, os, email , optparse
from xlrd import open_workbook
from xlwt import Workbook, easyxf
from os.path import join
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

'''
# 确定运行环境的encoding
__g_codeset = sys.getdefaultencoding()
if "ascii"==__g_codeset:
    __g_codeset = locale.getdefaultlocale()[1]
#

def utf8_to_mbs(s):
    return s.decode("utf-8").encode(__g_codeset)
#

def mbs_to_utf8(s):
    return s.decode(__g_codeset).encode("utf-8")
'''

'''
设置单元格样式
'''

def set_style(name,height,bold=False):
  style = xlwt.XFStyle() # 初始化样式
 
  font = xlwt.Font() # 为样式创建字体
  font.name = name # 'Times New Roman'
  font.bold = bold
  font.color_index = 4
  font.height = height
 
  # borders= xlwt.Borders()
  # borders.left= 6
  # borders.right= 6
  # borders.top= 6
  # borders.bottom= 6
 
  style.font = font
  # style.borders = borders
 
  return style
 
 
#写excel
def write_excel():
    f = Workbook() #创建工作簿
 
    '''
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet
    row0 = [u'发件方',u'收件方',u'抄送方',u'密送方',u'是否有附件']
 
    default = easyxf('font: name Times New Roman;')
    #生成第一行
    for i in range(0,len(row0)):
        sheet1.write(0,i,row0[i],default)
    
    count = 0
    for item in os.listdir(dest):
        path = dest + os.sep + item
        if os.path.isfile(path) and item.endswith('.eml'):
            fp = open(path,'r')
            global msg
            msg = email.message_from_file(fp)
            count = count + 1
            try:
                sheet1.write(count,0,get_from(),default)
                sheet1.write(count,1,get_to(),default)
                sheet1.write(count,2,get_cc(),default)
                sheet1.write(count,3,get_bcc(),default)
                sheet1.write(count,4,get_attachment(),default)
            except:
                print count
                raise
        if (count%100)==0 :
            f.save(savedest) #保存文件
    f.save(savedest) #保存文件

'''
#获取脚本文件的当前路径
def cur_file_dir():
    #获取脚本路径
    path = sys.path[0]
    s = os.sep #根据unix或win，s为\或/
    root = path + s + "all" + s #要遍历的目录
    return root
'''

#获取发件人邮箱   
def get_from():
    if msg.get('from') != None:
        address = msg.get('from').split('>,')
        i = 0
        addresses = ''
        if(len(address)==1):
            return email.utils.parseaddr(address[0])[1]
        else:
            while((i<len(address))):
                addresses = addresses + email.utils.parseaddr(address[i])[1] + ';'
                i = i+1
            return addresses
    else:
        return ''

#获取收件人邮箱
def get_to():
    if msg.get('to') != None:
        address = msg.get('to').split('>,')
        i = 0
        addresses = ''
        if(len(address)==1):
            return email.utils.parseaddr(address[0])[1]
        else:
            while((i<len(address))):
                addresses = addresses + email.utils.parseaddr(address[i])[1] + ';'
                i = i+1
            return addresses
    else:
        return ''

#获取cc邮箱
def get_cc():
    if msg.get('cc') != None:
        address = msg.get('cc').split('>,')
        i = 0
        addresses = ''
        if(len(address)==1):
            return email.utils.parseaddr(address[0])[1]
        else:
            while((i<len(address))):
                addresses = addresses + email.utils.parseaddr(address[i])[1] + ';'
                i = i+1
            return addresses
    else:
        return ''

#获取bcc邮箱
def get_bcc():
    if msg.get('bcc') != None:
        address = msg.get('bcc').split('>,')
        i = 0
        addresses = ''
        if(len(address)==1):
            return email.utils.parseaddr(address[0])[1]
        else:
            while((i<len(address))):
                addresses = addresses + email.utils.parseaddr(address[i])[1] + ';'
                i = i+1
            return addresses
    else:
        return ''

#attatchment have or not
def get_attachment():
    if msg.get('Content-Type') != None:
        if msg.get('Content-Type').split(';')[0] == 'multipart/mixed' :
            return 1
        else :
            return 0
    else:
        empty_obj()

def main() :
    parser = optparse.OptionParser('usage: %prog [options]')
    parser.add_option('-p', '--filepath', dest='dir_path', default=None, type='string', help='you need a way to go')
    (options, args) = parser.parse_args()
    if options.dir_path == None:
        sys.exit(0)
    else:
        global dest , savedest
        dest = options.dir_path
        savedest = dest + os.sep + 'eml_contact.xls'
        print(savedest)
        write_excel()
                    
 
if __name__ == "__main__" :
    main()

