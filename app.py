from flask import *
from wtforms import *
from flask_wtf import CSRFProtect,FlaskForm
from xlrd import open_workbook
from xlutils.copy import copy
from email.header import Header
from email.mime.text import MIMEText
import smtplib

app = Flask(__name__)
app.config['SECRET_KEY'] = 'abgiueb2389406@#%^%DFWE'
CSRFProtect(app)
from_addr = 'zhouzihao2008@126.com'
password = 'zhouzihao626078'
smtp_server = 'smtp.126.com'
fp = open('mailmsg.html', 'r', encoding='utf-8')
text = fp.read()
fp.close()


class MyFrom(FlaskForm):
    name = StringField('姓名',[validators.Length(min=2,max=4,message='我想你的姓名应该在2个字到4个字之间')])
    number = IntegerField('胸卡号',[validators.NumberRange(min=20000000,max=210000000,message='通中学生胸卡号应该是9位的，一中的是8位的')])
    qqmail = StringField('QQ邮箱',[validators.Email(message='请填写你的邮箱'),validators.DataRequired(message='请填写你的邮箱')])
    school = RadioField('你的学校',[validators.DataRequired(message='请选择你的学校')],choices=[('通中','通中'),('一中','一中')])
    submit = SubmitField('提交')


def check(name, number, qqmail, school):
    workbook = open_workbook('number.xls')
    sheet = workbook.sheet_by_index(0)
    i = int(sheet.cell_value(0, 1))  # 通中
    if i >= 150:
        return 'fail', 3, ''
    ii = int(sheet.cell_value(0, 2))  # 一中
    if ii >= 150:
        return 'fail', 3, ''
    iii = int(sheet.cell_value(0, 0))  # 总数
    for x in range(1, i):
        if number == sheet.cell_value(x, 1):
            if name == sheet.cell_value(x, 0).encode('utf-8') and qqmail == sheet.cell_value(x, 2).encode(
                    'utf-8') and school == sheet.cell_value(x, 3).encode('utf-8'):
                return 'success', 2, sheet.cell_value(x,4)
            return 'fail', 1, ''
        if qqmail == sheet.cell_value(x, 2).encode('utf-8'):
            return 'fail', 2, ''
    newbook = copy(workbook)
    sheet1 = newbook.get_sheet(0)
    sheet1.write(i,0,name)
    sheet1.write(i,1,number)
    sheet1.write(i,2,qqmail)
    sheet1.write(i,3,school)
    if school == '通中':
        i = i+1
    else:
        ii = ii+1
    iii = iii+1
    sheet1.write(0,0,iii)
    sheet1.write(0,1,i)
    sheet1.write(0,2,ii)
    newbook.save('number.xls')
    return 'success', 1, sheet.cell_value(i-1,4)


@app.route('/',methods=['GET','POST'])
def home():
    form = MyFrom()
    if form.validate_on_submit():
        name = form.name.data
        number = form.number.data
        qqmail = form.qqmail.data
        school = form.school.data
        print(name,number,school,qqmail)
        status,code,key = check(name,number,qqmail,school)
        if status == 'success':
            server = smtplib.SMTP(smtp_server, 25)
            server.set_debuglevel(1)
            server.login(from_addr, password)
            msg = MIMEText(text % (name, int(key)), 'html', 'utf-8')
            msg['From'] = from_addr
            msg['To'] = qqmail
            msg['Subject'] = Header('欢迎观赏辩论友谊赛', 'utf-8').encode()
            server.sendmail(from_addr, [qqmail], msg.as_string())
            server.quit()
        return redirect(url_for(status,id=code))
    return render_template('home.html',form=form)


@app.route('/success/<id>',methods=['GET'])
def success(id):
    if id=='1':
        message='请打开邮箱查看入场券，如果没有收到可以查看一下垃圾箱'
    if id=='2':
        message='通行证已经重新发到你的邮箱，请注意查收'
    return render_template('success.html',message=message)


@app.route('/fail/<id>',methods=['GET'])
def fail(id):
    if id=='1':
        message='该胸卡号已领取入场券，如果没有在邮箱收到通行证，请重新填写与申请时一样的信息，我们将重新发送通行证'
    if id=='2':
        message='该邮箱已领取入场券，如果没有在邮箱收到通行证，请重新填写与申请时一样的信息，我们将重新发送通行证'
    if id=='3':
        message='该学校的入场券已申请完'
    return render_template('fail.html',message=message)


if __name__ == '__main__':
    app.run(host='0.0.0.0',port=14250)
