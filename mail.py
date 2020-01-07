import  smtplib
import email
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.header import Header

#sender是发送人邮箱，passWord是服务器授权码，mail_host是服务器地址
sender = '905536460@qq.com'
passWord = 'roxnduicqxqnbefc'
mail_host = 'smtp.qq.com'
#邮件接送人
receivers = ['2100567390@qq.com','651231513@qq.com']

#新建MIMEMultipart对象
msg = MIMEMultipart('related')

#设置邮件头部内容
subject_content = 'Python 邮件测试' #主题
msg['From'] = 'sender_name<905536460@qq.com>' #发送人
msg['TO'] = 'receivers_1_name<2100567390@qq.com>,receivers_2_name<651231513@qq.com>' #接收人
msg['Subject'] = Header(subject_content,'utf-8')

#邮件正文
body_content = '这是一个Python测试邮件'
message_text = MIMEText(body_content,'plain','utf-8') #构造文本，参数：正文内容，文本格式，编码方式
msg.attach(message_text) #向MIMEMultipart对象添加文本对象

#添加图片
with open('D:\py\ciyun.PNG','rb') as f: #二进制读取图片
    message_image = MIMEImage(f.read()) #读取二进制数据
msg.attach(message_image) #添加图片文件到邮件信息当中去

#发送邮件
stp = smtplib.SMTP()  #创建SMTP对象
stp.connect(mail_host,25) #设置发件人邮箱的域名和端口
stp.set_debuglevel(1) # set_debuglevel(1)可以打印出和SMTP服务器交互的所有信息
stp.login(sender,passWord)
stp.sendmail(sender,receivers,msg.as_string())
stp.quit()