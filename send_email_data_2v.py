import requests
import re
from bs4 import BeautifulSoup
import lxml
import pandas as pd

#登录获取cookies
def login(user,pwd):
    url='https://m.exmail.qq.com/cgi-bin/login'
    data={
        'device':'',
        'f':'xhtml',
        'tfcont':'',
        'uin':user,
        'pwd':pwd,
        'btlogin':'登录'
    }
    res=requests.post(url=url,data=data)
    # print(res.cookies.get_dict())
    cookies=res.cookies.get_dict()
    return cookies

#破解sid参数
def get_sid():
    msid=cookies['msid']
    mystr1=msid.split('&')[1].split('@')[0]
    mystr2=msid.split('&')[1].split('@')[1]
    mystr3=cookies['qm_sk'].split('&')[1]
    mystr4=cookies['qm_sid'].split(',')[1]
    # print(mystr1,mystr2,mystr3,mystr4)
    sid=mystr1 + mystr3 + ',' + str(mystr2) + ',' + mystr4
    # print(sid)
    return sid


def get_send_page():
    url='https://m.exmail.qq.com/cgi-bin/mail_list'
    data={
        'fromsidebar':'1',
        'sid':sid,
        'folderid':'3',
        'page':'0',
        'pagesize':'10',
        'sorttype':'time',
        't':'mail_list',
        'loc':'folderlist,,xhtml,1'
    }
    res=requests.get(url=url,params=data,cookies=cookies)
    html=res.text
    # print(html)
    page=re.findall(r'qm_page_item qm_page_item_Mid">1 / (\d{1,5})',html)[0]
    # print(page)
    return page

def get_send_mail():
    page = int(get_send_page())
    for p in range(page):
        try:
            url = 'https://m.exmail.qq.com/cgi-bin/mail_list'
            data = {
                'fromsidebar':'1',
                'sid':sid,
                'folderid':'3',
                'page':str(p),
                'pagesize':'10',
                'sorttype':'time',
                't':'mail_list',
                'loc':'folderlist,,xhtml,1'
            }
            res = requests.get(url=url, params=data, cookies=cookies)
            html = res.text
            # print(html)
            soup=BeautifulSoup(html,'lxml')
            # print(soup)
            readmail_list=soup.find(class_='readmail_list')
            # print(readmail_list)
            divs=readmail_list.find_all('div',class_='maillist_listItem')
            for div in divs:
                # print(div)
                name = div.find('nobr').text.strip()  #接收人
                # send_time=div.find('span',class_='maillist_listItem_time').text.strip()
                subject=div.find('div',class_='maillist_listItemLineSecond func_ellipsis').text.strip()  #主题
                # body=div.find('div',class_='maillist_listItem_abstract func_ellipsis').text.strip().replace('&nbsp;','')

                href = 'https://m.exmail.qq.com' + div.find('a')['href']
                # print(href)
                res = requests.get(url=href, cookies=cookies)
                send_time = BeautifulSoup(res.text, 'lxml').find('span',class_='readmail_item_date').text.strip()   #发送时间
                body = BeautifulSoup(res.text, 'lxml').find('div',class_='readmail_item_contentNormal qmbox').text.strip()  # 正文

                print(name,send_time,subject,body)
                receive_data['接收人'].append(name)
                receive_data['发送时间'].append(send_time)
                receive_data['主题'].append(subject)
                receive_data['正文'].append(body)
        except:
            pass


def save_receive_mail():
    get_send_mail()
    pd.DataFrame(receive_data).to_excel(user.split('@')[0]+'发件箱数据.xlsx',index=False)


if __name__=='__main__':
    user=input('请输入邮箱地址：').strip()
    pwd=input('请输入邮箱密码：').strip()
    cookies=login(user,pwd)
    sid=get_sid()
    receive_data = {
        '接收人': [],
        '发送时间': [],
        '主题': [],
        '正文': []
    }
    save_receive_mail()