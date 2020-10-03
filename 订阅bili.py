from selenium import webdriver
from pyquery import PyQuery as pq
import time
import json
import pandas as pd
import emailsend

# 获取cookie
# driver = webdriver.Chrome()
# driver.maximize_window()
# driver.implicitly_wait(5)
# driver.get('https://t.bilibili.com/')
# driver.delete_all_cookies()
# time.sleep(30)
# dictcookies = driver.get_cookies()
# jsoncookies = json.dumps(dictcookies)
#
# with open('bilicookie.txt', 'w') as f:
#     f.write(jsoncookies)
# print('cookie:{}'.format(jsoncookies))
# driver.quit()

# 让滚动条滚动到昨天看到的位置，以加载所需网页
def scroll_bar_we_need():
    scrollTop = 10000
    countnum = 0
    while countnum == 0:
        down = "var q=document.documentElement.scrollTop=" + str(scrollTop)
        driver.execute_script(down)
        time.sleep(3)
        doc = pq(driver.page_source)
        pubtime = doc('[class="detail-link tc-slate"]').items()
        for i in pubtime:
            if i.text()[0:2] == '昨天':
                countnum = countnum + 1
                break
        scrollTop = scrollTop + 2000   # 避免更新太多最终下拉框固定到10000位置陷入死循环
    return driver.page_source

# 从获取的网页中提取要通过邮件发送的部分
def email_content(page_source):
    doc = pq(page_source)
    totalpostnum = 0
    for i in doc('[class="c-pointer"]').items():
        totalpostnum = totalpostnum + 1
    print("there are {} posts waiting for your reading".format(totalpostnum))
    postlists = doc('[class="card"]').items()
    c_pointers = []
    pubtimes = []
    pub_types = []
    titles = []
    reposts = []
    contentfulls = []
    contents =[]
    for idx, postlist in enumerate(postlists):
        if idx == totalpostnum:
            break
        c_pointer = postlist('[class="c-pointer"]').text()
        c_pointers.append(c_pointer)
        pubtime = postlist('[class="detail-link tc-slate"]').text()
        pubtimes.append(pubtime)
        if postlist('[class="post-content repost"]').text() == '':
            pub_types.append('原创')
            reposts.append(' ')
            if postlist('[class="content-ellipsis"]').text() == '':
                contentfulls.append(postlist('[class="content-full"]').text()+' ')
            else:
                contentfulls.append(postlist('[class="content-ellipsis"]').text())
        else:
            pub_types.append('转发')
            reposts.append(postlist('[class="content-full"]').eq(0).text())
            contentfulls.append(postlist('[class="content-full"]').eq(1).text()+' ')

        title = postlist('[class="title"]').text() + postlist('[class="title fs-16 tc-black"]').text()
        if title == '':
            titles.append(' ')
        else:
            titles.append(title)
        textarea = postlist('[class="text-area"]')
        contents.append(textarea('[class="content"]').text()+' ')
    dict1 = {
        '发布者':c_pointers,
        '发布时间':pubtimes,
        '发布类型':pub_types,
        '视频标题':titles,
        '转发评论':reposts,
        '发布内容':contentfulls,
        '视频内容':contents
    }
    dfdata = pd.DataFrame(dict1)
    return dfdata


if __name__ == '__main__':
    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.implicitly_wait(5)
    driver.get('https://t.bilibili.com/')
    f = open('bilicookie.txt', 'r')
    listcookie = json.loads(f.read())
    for cookie in listcookie:
        driver.add_cookie(cookie)
    driver.refresh()
    driver.implicitly_wait(5)
    time.sleep(5)
    page_source = scroll_bar_we_need()
    driver.quit()
    dfdata = email_content(page_source)

    fp = open('./edata.json', 'r', encoding='utf8')
    json_data = json.load(fp)
    mail_subject = 'bilibili订阅消息'
    outline = 'bili订阅消息'
    mail_content = emailsend.htmlcreat(dfdata, outline=outline)
    mail = emailsend.Mail(json_data, mail_subject=mail_subject, mail_content=mail_content, mail_type='html')
    mail.send()


