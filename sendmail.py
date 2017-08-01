
#-*- coding: utf-8 -*-

import os
import smtplib
#from   email import encoders
from   email.header import Header
from   email.mime.text import MIMEText
from   email.utils import parseaddr, formataddr
from   email.mime.multipart import MIMEMultipart
from   email.mime.application import MIMEApplication


#获得工单文件名list
workfile_list=os.listdir(u'H:\\work\\work_list\\201707教师工单\\')
#按文件名排序
workfile_list.sort()
i=0
while i<22:
    print workfile_list[i]
    i=i+1

#定义函数，格式化收发件人的地址
def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((
        Header(name, 'utf-8').encode(),
        addr.encode('utf-8') if isinstance(addr, unicode) else addr))

from_addr = 'z_zming@126.com'
password = raw_input('Password: ')


#to_addr = raw_input('To: ')
#获得收件人地址list
to_addr_list=['z_zming@126.com','zzming2@hljqn.com','z_zming3@126.com','zzming4@hljqn.com',
              'z_zming5@126.com','zzming@hljqn.com','z_zming@126.com','zzming8@hljqn.com',
              'z_zming9@126.com','zzming@hljqn.com','z_zming11@126.com']
smtp_server = 'smtp.126.com'
#工单所在路径
path=u'H:\\work\\work_list\\201707教师工单\\'
#msg = MIMEText('hello, send by Python...', 'plain', 'utf-8')
count=0 #初始化计数器
#to_addr_failed_list=[]
#attach_failed_list=[]

#循环发送邮件
while len(to_addr_list)>0:
    to_addr = to_addr_list.pop(0) #左弹出收件地址列表并赋值于收件人
    attach1_name = workfile_list.pop(0)#左弹出工单列表并赋值于附件1
    attach2_name = workfile_list.pop(0)  # 继续左弹出工单列表并赋值于附件2
    attach1 = os.path.join(path,attach1_name)
    attach2 = os.path.join(path,attach2_name)

    msg = MIMEMultipart()
    msg['From'] = _format_addr(u'小明 <%s>' % from_addr)
    msg['To'] = _format_addr(u'亲爱的同事 <%s>' % to_addr)
    msg['Subject'] = Header(u'2017.08月工单……', 'utf-8').encode()
    #---这是email正文部分---
    part_text = MIMEText(u'请看附件','plain','utf-8')
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

    #发送邮件
    try:
        server = smtplib.SMTP(smtp_server, 25)
        #server.set_debuglevel(1)
        server.login(from_addr, password)
        server.sendmail(from_addr, [to_addr], msg.as_string())
        server.quit()

    except smtplib.SMTPException,e:
        #print "email sent failed! %s" %to_addr
        print "发送给%s的邮件失败,原因：%s" %(to_addr,e)
        flag=raw_input("是否重新投递Y/n ?")
        if flag=="Y" or flag=="y" or flag=="":     #用户如果重发，则"故障目标邮件地址"与两个"附件"重新入列
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
