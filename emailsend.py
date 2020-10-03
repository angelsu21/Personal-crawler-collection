from email.mime.text import MIMEText
from email.header import Header
from smtplib import SMTP_SSL
import json
import time
import pandas as pd


class Mail:
    def __init__(self, json_data, mail_subject, mail_content, mail_type):
        self.host_server = json_data["host_server"]
        self.sender_qq = json_data["sender_qq"]
        self.pwd = json_data["pwd"]
        self.sender = json_data["sender"]
        self.mail_subject = mail_subject
        self.receivers = json_data["receivers"]
        self.mail_content = mail_content
        self.mail_type = mail_type


    def send(self):
        try:
            # ssl登录
            smtp = SMTP_SSL(self.host_server)
            smtp.ehlo(self.host_server)
            smtp.login(self.sender_qq, self.pwd)
            mail_content = self.mail_content
            # mail_content = mail_content + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

            if self.mail_type == 'html':
                msg = MIMEText(mail_content, "html", "utf-8")    # plain 是以普通文本发送
            else:
                msg = MIMEText(mail_content, "plain", "utf-8")
            msg["Subject"] = Header(self.mail_subject, "utf-8")   # 邮件标题
            msg["From"] = self.sender
            msg["To"] = self.receivers

            smtp.sendmail(self.sender, self.receivers.split(","), msg.as_string())  # 发送邮件
            smtp.quit()   # 关闭SMTP对话
            smtp.close()  # 关闭SMTPu服务器连接
            print("邮件发送成功")
        except Exception as e:
            print("邮件发送失败")
            print(e)

def htmlcreat(filter_merge_data, outline):
    pd.set_option('display.max_colwidth', None)  # 设置表格数据完全显示（不出现省略号）
    df_html = filter_merge_data.to_html(escape=False)  # DataFrame数据转化为HTML表格形式
    head = \
        """
        <head>
            <meta charset="utf-8">
            <STYLE TYPE="text/css" MEDIA=screen>

                table.dataframe {
                    border-collapse: collapse;
                    border: 2px solid #a19da2;
                    /*居中显示整个表格*/
                    margin: auto;
                }

                table.dataframe thead {
                    border: 2px solid #91c6e1;
                    background: #f1f1f1;
                    padding: 10px 10px 10px 10px;
                    color: #333333;
                }

                table.dataframe tbody {
                    border: 2px solid #91c6e1;
                    padding: 10px 10px 10px 10px;
                }

                table.dataframe tr {

                }

                table.dataframe th {
                    vertical-align: top;
                    font-size: 14px;
                    padding: 10px 10px 10px 10px;
                    color: #105de3;
                    font-family: arial;
                    text-align: center;
                }

                table.dataframe td {
                    text-align: center;
                    padding: 10px 10px 10px 10px;
                }

                body {
                    font-family: 宋体;
                }

                h1 {
                    color: #5db446
                }

                div.header h2 {
                    color: #0002e3;
                    font-family: 黑体;
                }

                div.content h2 {
                    text-align: center;
                    font-size: 28px;
                    text-shadow: 2px 2px 1px #de4040;
                    color: #fff;
                    font-weight: bold;
                    background-color: #008eb7;
                    line-height: 1.5;
                    margin: 20px 0;
                    box-shadow: 10px 10px 5px #888888;
                    border-radius: 5px;
                }

                h3 {
                    font-size: 22px;
                    background-color: rgba(0, 2, 227, 0.71);
                    text-shadow: 2px 2px 1px #de4040;
                    color: rgba(239, 241, 234, 0.99);
                    line-height: 1.5;
                }

                h4 {
                    color: #e10092;
                    font-family: 楷体;
                    font-size: 20px;
                    text-align: center;
                }

                td img {
                    /*width: 60px;*/
                    max-width: 300px;
                    max-height: 300px;
                }

            </STYLE>
        </head>
        """

    body = \
        """
        <body>

        <div align="center" class="header">
            <!--标题部分的信息-->
            <h1 align="center">{outline}</h1>
            <h2 align="center">{daytime}</h2>
        </div>

        <hr>

        <div class="content">
            <!--正文内容-->
            <h2> </h2>

            <div>
                <h4></h4>
                {df_html}

            </div>
            <hr>

            <p style="text-align: center">

            </p>
        </div>
        </body>
        """.format(outline= outline,daytime=time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), df_html=df_html)
    html_msg = "<html>" + head + body + "</html>"
    html_msg = html_msg.replace('\n', '').encode("utf-8")  # 拼接后的HTML可能含有换行符\n，需要将其去掉
    return html_msg
