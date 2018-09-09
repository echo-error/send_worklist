
#-*- coding: utf-8 -*-

import os
import smtplib
#from   email import encoders
from   email.header import Header
from   email.mime.text import MIMEText
from   email.utils import parseaddr, formataddr
from   email.mime.multipart import MIMEMultipart
from   email.mime.application import MIMEApplication


#工单所在路径

#subdir_name=raw_input(u'请输入目录的月份,如201712：'.encode('gbk'))
subdir_name=raw_input(u'请输入目录的月份,如201712：')
                                                                  #raw_input函数不支持unicode,如果需要在windows环境cmd下运行，
                                                                  #需要将raw_input()中文转码为gbk,否则乱码'''
#subject=raw_input(u"请输入邮件标题，如  '2017.12月工单及11月得分':".encode('gbk'))
subject=raw_input(u"请输入邮件标题，如  '2017.12月工单及11月得分':")
#work_path=u'H:\\work\\work_list\\'+subdir_name+u'教师工单\\'
work_path=u'/root/work/'+subdir_name+u'教师工单/'
print work_path

workfile_list=os.listdir(work_path)                               #获得工单文件名list

workfile_list.sort()                                              #按文件名排序

print len(workfile_list)
i=0

while i<len(workfile_list):
    print workfile_list[i]
    i=i+1


def _format_addr(s):                                              #定义函数，格式化收发件人的地址
    name, addr = parseaddr(s)
    return formataddr((
        Header(name, 'utf-8').encode(),
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))

from_addr = 'z_zming@xx.com'
password = raw_input('Password: ')


#to_addr = raw_input('To: ')

to_addr_list=['z_zming@xX.com',                                  #获得收件人地址list
              'zzming@xx.com']

smtp_server = 'smtp.xx.com'

#msg = MIMEText('hello, send by Python...', 'plain', 'utf-8')
#to_addr_failed_list=[]
#attach_failed_list=[]


def feeds():                                                      #产生一个目标地址以及应该发送给该地址的两份附件信息
    to_addr = to_addr_list.pop(0)                                 #左弹出收件地址列表并赋值于收件人
    attach1_name = workfile_list.pop(0)                           #左弹出工单列表并赋值于附件1
    attach2_name = workfile_list.pop(0)                           #继续左弹出工单列表并赋值于附件2
    attach1 = os.path.join(work_path,attach1_name)
    attach2 = os.path.join(work_path,attach2_name)
    return to_addr,attach1_name,attach2_name,attach1,attach2


def assemble(to_addr,attach1_name,attach2_name,attach1,attach2):  #组装邮件正文及附件
    msg = MIMEMultipart()
    msg['From'] = _format_addr(u'发件人 <%s>' % from_addr)
    msg['To'] = _format_addr(u'学术部 <%s>' % to_addr)
    msg['Subject'] = Header(subject, 'utf-8').encode()
    part_text = MIMEText(u'请看附件','plain','utf-8')               #这是email正文部分
    msg.attach(part_text)
    #---这是附件部分---
    #xlsx类型附件1
    part1 = MIMEApplication(open(attach1,'rb').read())
    part1.add_header('Content-Disposition', 'attachment', filename=attach1_name.encode('gb18030'))
    msg.attach(part1)
    #xlsx类型附件2
    part2 = MIMEApplication(open(attach2,'rb').read())
    part2.add_header('Content-Disposition', 'attachment', filename=attach2_name.encode('gb18030'))
    msg.attach(part2)
    return msg

count=0                                                             #初始化计数器


#循环发送邮件
while len(to_addr_list)>0:
    to_addr,attach1_name,attach2_name,attach1,attach2=feeds()
    msg=assemble(to_addr,attach1_name,attach2_name,attach1,attach2)
    try:                                                            #发送邮件
        server = smtplib.SMTP(smtp_server, 25)
        #server.set_debuglevel(1)
        server.login(from_addr, password)
        server.sendmail(from_addr, [to_addr], msg.as_string())
        server.quit()
    except  smtplib.SMTPException,e:
        #print "email sent failed! %s" %to_addr
        print u"发送给%s的邮件失败,原因：%s" %(to_addr,e)
        flag=raw_input(u'是否重新投递Y/n ?')
        if flag=="Y" or flag=="y" or flag=="":                      #用户如果重发，则"故障目标邮件地址"与两个"附件"重新入列
                to_addr_list.append(to_addr)
                workfile_list.append(attach1)
                workfile_list.append(attach2)
               #to_addr_failed_list.append(to_addr)
               #attach_failed_list.append(attach1)
               #attach_failed_list.append(attach2)
    else:
        print "email sent successfully! %s" %to_addr
        count = count + 1

print u'成功发送%d封邮件！' %count
#print u'发送失败的对象 %s' %to_addr_failed_list
