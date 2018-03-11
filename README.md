# 随机门票码获取网页(为秋云辩社设计的)

## 1、相关支持

- flask（需安装）
- xlutils（需安装）
- wtforms（需安装）
- flask_wtf（需安装）
- random
- email
- smtplib

## 2、使用方法

首先，运行keys.py创建含有随机数的number.xls表格<br>
**注意：这会清除原有所有数据！！！**<br>
将app.py中的address，password和smtp更改为你自己的<br><br>
使用cmd进入app.py所在的目录，输入：

	python app.py

即可在本机运行服务器，局域网内其他用户访问：

	(hostip):14250

（hostip为你的本机ip）便能进入网页。<br>

### 更改邮件内容

直接更改mailmsg.html即可，注意同步修改在app.py中的占位符！

### 更改表单内容

自学flask-wtf相关后可以轻易修改。

## 3、注意事项

可能会由于网络原因而导致发送邮件失败，可能会引起服务器堵塞，我是没遇到过。

## 4、数据保存

number.xls保存着所有的表单数据。