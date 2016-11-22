# -*- coding: utf-8 -*-
##__author__ =='liam'
import re
import requests
import random
import os
import json
import ctypes
import win32ui
import sys
import math
import time
import datetime
from lxml import etree
import subprocess
import xlsxwriter as wx
import ctypes
from PIL import Image

STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE= -11
STD_ERROR_HANDLE = -12

FOREGROUND_BLACK = 0x0
FOREGROUND_BLUE = 0x01 # text color contains blue.
FOREGROUND_GREEN= 0x02 # text color contains green.
FOREGROUND_RED = 0x04 # text color contains red.
FOREGROUND_INTENSITY = 0x08 # text color is intensified.

BACKGROUND_BLUE = 0x10 # background color contains blue.
BACKGROUND_GREEN= 0x20 # background color contains green.
BACKGROUND_RED = 0x40 # background color contains red.
BACKGROUND_INTENSITY = 0x80 # background color is intensified.

class Color:
    ''''' See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winprog/winprog/windows_api_reference.asp
    for information on Windows APIs.'''
    std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)

    def set_cmd_color(self, color, handle=std_out_handle):
        """(color) -> bit
        Example: set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE | FOREGROUND_INTENSITY)
        """
        bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
        return bool

    def reset_color(self):
        self.set_cmd_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)

    def print_red_text(self, print_text):
        self.set_cmd_color(FOREGROUND_RED | FOREGROUND_INTENSITY)
        print (print_text)
        self.reset_color()

    def print_green_text(self, print_text):
        self.set_cmd_color(FOREGROUND_GREEN | FOREGROUND_INTENSITY)
        print (print_text)
        self.reset_color()

    def print_blue_text(self, print_text):
        self.set_cmd_color(FOREGROUND_BLUE | FOREGROUND_INTENSITY)
        print (print_text)
        self.reset_color()

    def print_red_text_with_blue_bg(self, print_text):
        self.set_cmd_color(FOREGROUND_RED | FOREGROUND_INTENSITY| BACKGROUND_BLUE | BACKGROUND_INTENSITY)
        print (print_text)
        self.reset_color()

def parseDate(datestr):
##    print(datestr)
    if re.search('(\d+).*天[之|以]?前',datestr):
        tmp=re.search('(\d+).*天[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(days = int(tmp)))
    elif re.search('(\d+).*日[之|以]?前',datestr):
        tmp=re.search('(\d+).*日[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(days = int(tmp)))
    elif re.search('(\d+).*周[之|以]?前',datestr):
        tmp=re.search('(\d+).*周[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(weeks = int(tmp)))
    elif re.search('(\d+).*秒[钟]?[之|以]?前',datestr):
        tmp=re.search('(\d+).*秒[钟]?[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(seconds = int(tmp)))
    elif re.search('(\d+).*分钟[之|以]?前',datestr):
        tmp=re.search('(\d+).*分钟[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(minutes = int(tmp)))
    elif re.search('(\d+)个?.*星期[之|以]?前',datestr):
        tmp=re.search('(\d+)个?.*星期[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(weeks = int(tmp)))
    elif re.search('(\d+)个?.*礼拜[之|以]?前',datestr):
        tmp=re.search('(\d+)个?.*礼拜[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(weeks = int(tmp)))
    elif re.search('(\d+)个?.*小时[之|以]?前',datestr):
        tmp=re.search('(\d+)个?.*小时[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(hours = int(tmp)))
    elif re.search('(\d+)个?.*钟头[之|以]?前',datestr):
        tmp=re.search('(\d+)个?.*钟头[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(hours = int(tmp)))
    elif re.search('(\d+)个?.*钟点[之|以]?前',datestr):
        tmp=re.search('(\d+)个?.*钟点[之|以]?前',datestr).group(1)
        date_pa = (datetime.datetime.now() - datetime.timedelta(hours = int(tmp)))
    elif re.search('(\d+)个?.*月[之|以]?前',datestr):
        tmp=re.search('(\d+)个?.*月[之|以]?前',datestr).group(1)
        date_pa = datetime.datetime.now() - relativedelta.relativedelta(months = int(tmp)) 
    elif re.search('(\d+).*年[之|以]?前',datestr):
        tmp=re.search('(\d+).*年[之|以]?前',datestr).group(1)
        date_pa = datetime.datetime.now() - relativedelta.relativedelta(years = int(tmp))       
    elif re.search('\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}',datestr):
        tmp=re.search('\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}',datestr).group()
        date_pa=time.strptime(tmp, "%Y-%m-%d %H:%M:%S")
    elif re.search('\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}',datestr):
        tmp=re.search('\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}',datestr).group()
        date_pa=time.strptime(tmp, "%Y-%m-%d %H:%M")
    elif  re.search('\d{4}-\d{1,2}-\d{1,2}',datestr): 
        tmp=re.search('\d{4}-\d{1,2}-\d{1,2}',datestr).group()
        date_pa=time.strptime(tmp, "%Y-%m-%d")
    elif  re.match('.*今.*天.*',datestr):
        today = datetime.date.today()
        if re.search('\d{1,2}:\d{1,2}:\d{1,2}',datestr):
            tmp=re.search('\d{1,2}:\d{1,2}:\d{1,2}',datestr).group()
            date_pa=time.strptime(str(today)+' '+tmp, "%Y-%m-%d %H:%M:%S")
        else:
            date_pa=time.strptime(str(today), "%Y-%m-%d")
    elif re.match('.*昨.*天.*',datestr):
        day = datetime.date.today()- datetime.timedelta(days=1) 
        if re.search('\d{1,2}:\d{1,2}:\d{1,2}',datestr):
            tmp=re.search('\d{1,2}:\d{1,2}:\d{1,2}',datestr).group()
            date_pa=time.strptime(str(day)+' '+tmp, "%Y-%m-%d %H:%M:%S")
        else:
            date_pa=time.strptime(str(day), "%Y-%m-%d")
    elif re.match('.*前.*天.*',datestr):
        day = datetime.date.today()- datetime.timedelta(days=2) 
        if re.search('\d{1,2}:\d{1,2}:\d{1,2}',datestr):
            tmp=re.search('\d{1,2}:\d{1,2}:\d{1,2}',datestr).group()
            date_pa=time.strptime(str(day)+' '+tmp, "%Y-%m-%d %H:%M:%S")
        else:
            date_pa=time.strptime(str(day), "%Y-%m-%d")
    return date_pa

def parseDateStr(date_pa):
    return time.strftime("%Y-%m-%d %H:%M:%S", date_pa)

def parseDateStrToStamp(datestr):
       return time.mktime(time.strptime(datestr,'%Y-%m-%d %H:%M:%S'))

def checkThreadPage(xmldata):
    if(len(getThreadNodes(xmldata))>0):
        return False
    else:
        return True

def checkPostPage(xmldata):
    if(len(getRowNodes(xmldata))>0):
        return False
    else:
        return True

def getRowNodes(xmldata):
    data = xmldata
    rownodes=data.xpath('.//div[@class="mod-bd"]//div[@class = "comment-item"]')

    # contains:.//a[contains(@class,'btnX') and .//text()='Sign in']
	# starts-with：.//a[starts-with(@class,'btnSelectedBG')]

    if len(rownodes)==0:
        raise NameError('Can not parse post RowNodes!')
    return rownodes


    
def parsePosterName(rownode):
    node=rownode.xpath('.//span[@class="comment-info"]/a')
    if len(node)==0:
        raise NameError('Can not parse PosterName!')
    node = node[0].xpath('string(.)').strip()
    return node

def parseContent(rownode):
    node=rownode.xpath('.//div[@class="comment"]/p//text()')

    if len(node)==0:
        raise NameError('Can not parse Content!')
    content = ' '.join(node)
    return content

def parsePosterURL(rownode):
    node=rownode.xpath('.//span[@class="comment-info"]/a/@href')

    if len(node)==0:
        return None

    return node[0]


def parseNumofUseful(rownode):
    node=rownode.xpath('.//span[@class ="comment-vote"]/span[@class = "votes pr5"]')
    if len(node)==0:
        raise NameError('Can not parse NumofUseful!')
    elif len(node) ==1:
        useful = node[0].xpath('string(.)')

    return useful

def parsePosterID(url):
    if url ==None:
        return None
        
    if re.search('com/people/(.*?)/',url):
        return re.search('com/people/(.*?)/',url).group(1)
    else:
        return 0


def parseRating(rownode):

        rating = rownode.xpath('.//span[@class = "comment-info"]/span')
        if len(rating) == 2:
            return  rating[0].xpath('./@class')[0].replace('allstar','').replace(' rating','').replace('0','')
        else:
            return 0

def parseDateOfPost(rownode):

    node=rownode.xpath('.//span[@class="comment-info"]/span')

    if len(node)==0:
        raise NameError('Can not parse DateOfPost!')
    else:
        for idx,oneRow  in  enumerate(node):
            oneRow = oneRow.xpath('string(.)').replace('\n', '').replace(' ', '')
            if  re.search('(\d+)-(\d+)-(\d+)',oneRow):
                    oneRow  = re.search('(\d+-\d+-\d+)', oneRow).group(1)
                    node = parseDateStr(parseDate(oneRow))
                    return  node

        return node
   
def parseSinglePostRow(rownode):
    global Subject,threadurl
    subject = Subject
    posterName=parsePosterName(rownode)
    dateOfPost=parseDateOfPost(rownode)
    content=parseContent(rownode)
    posterURL=parsePosterURL(rownode)
    NumofUseful=parseNumofUseful(rownode)
    rating  = parseRating(rownode)
    posterID=parsePosterID(posterURL)
    threadURL=threadurl


    node = [subject,content,dateOfPost,NumofUseful,posterName,posterURL,rating,posterID,threadURL]
    return node

def parseSinglePostPageAndNeedTurnToNext(xmldata):
    global  postDateTime,postData
    if checkPostPage(xmldata):
        raise NameError('This Page is not a Post Page!')
    nodes = getRowNodes(xmldata)

    for node in nodes:
        post=parseSinglePostRow(node)

        if parseDateStrToStamp(post[2]) >= parseDateStrToStamp(parseDateStr(parseDate(postDateTime))):
            postData.append(post)
            print('save a record successfully !')
        else:
            return False

    return True if getNextPostPageNode(xmldata) != None else False


def getNextPostPageNode(xmldata):

    node=xmldata.xpath('.//a[@class="next"]')
    if len(node) == 0:
        return None
    node = node[0].xpath('@href')[0]
    return node

def turnToPage(url):
    global  content_headers
    waitTime=random.uniform(1, 2)
    time.sleep(waitTime)
    print(" Turn to next  Page : "+url)
    res = DoubanClient().getSession().get(url, headers  =content_headers,cookies = DoubanClient().loadCookie(),timeout=20,allow_redirects = False)
    # print(res)
    return res.text

def parseSubject(xmldata):

    subject = xmldata.xpath('.//td[@class = "ptm pbn"]/div[@class = "ts z h1"]')
    subject  =subject[0].xpath('string(.)').strip().replace('[复制链接]','')
    return subject


def doCapture(url):

    global threadurl,postDateTime,Subject,clr,content_headers


    content_headers  ={
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, sdch, br",
        "Accept-Language": "en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4",
        "Cache-Control": "max-age=0",
        "Connection": "keep-alive",
        "Host":"movie.douban.com",
        "Referer":"https://www.douban.com/accounts/login?source=movie",
        "Upgrade-Insecure-Requests":"1",
        "User-Agent":"Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36"
    }

    threadurl = url
    if re.search('.*?comments',threadurl):
        partOfUrl = re.search('(.*?comments)',threadurl).group(1)
    else:
        partOfUrl = None

    try:

        # print(DoubanClient().loadCookie())
        res = DoubanClient().getSession().get(threadurl, headers = content_headers,cookies = DoubanClient().loadCookie(),timeout=20,allow_redirects = False)
        # print(res)
        xmldata = etree.HTML(res.text)

        # 抓取标题
        Subject = xmldata.xpath('.//div[@id = "content"]/h1')
        if  Subject:
            Subject = Subject[0].xpath('string(.)').strip()
        else:
            Subject = 'None!'

        while (parseSinglePostPageAndNeedTurnToNext(xmldata)):

            pageNode = partOfUrl+getNextPostPageNode(xmldata)
            if not pageNode:
                break

            xml = turnToPage(pageNode)
            if not xml:
                break
            xmldata = etree.HTML(xml)
            
    except Exception as err:
        print ('has an error while spidering')
        print(err)
    finally:
        print('Finish Spidering')


class DoubanClient(object):

    loginURL = r"https://www.douban.com/login"
    homeURL = r"http://www.douban.com"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, sdch, br",
        "Accept-Language":"en-US,en;q=0.8,zh-CN;q=0.6,zh;q=0.4",
        "Host": "www.douban.com",
        "Connection":"keep-alive",
        "Cache-Control":"max-age=0",
        # "Referer":"https://www.baidu.com/link?url=BCG5AfxNkL6vibSGQS5bypnPCS1cASjrzeiUkBCODrG&wd=&eqid=e2bd91340002c19400000004582a86f3",
        "Upgrade-Insecure-Requests": "1"
    }

    # captchaFile = os.path.join(sys.path[0], "captcha.gif") 
    # cookieFile = os.path.join(sys.path[0], "cookie")
    captchaFile  = os.path.join(os.getcwd(), 'captcha.gif')
    cookieFile = os.path.join(os.getcwd(), 'cookie')


    def __init__(self):
        # os.chdir(sys.path[0])  # 设置脚本所在目录为当前工作目录
        os.chdir(os.path.dirname(os.path.abspath(__file__))) #绝对路径
        self.__session = requests.Session()
        self.__session.headers = self.headers
        # 若已经有 cookie 则直接登录
        # self.__cookie = self.__loadCookie()
        # if self.__cookie:
        #     print("检测到cookie文件，直接使用cookie登录")
        #     self.__session.cookies.update(self.__cookie)
        #     print("cookie 已保存")
        # else:
        #     print("没有找到cookie文件，请调用login方法登录一次！")

    def open(self, url, delay=0, timeout=10):
        """打开网页，返回Response对象"""
        if delay:
            time.sleep(delay)
        return self.__session.get(url, timeout=timeout)

    def login(self, username, password):
        # 登录

        self.__username = username
        self.__password = password
        self.__loginURL = self.loginURL

        while True:
            captchaid = None
            captcha = None

            html = self.open(self.__loginURL).text
            formatHtml = etree.HTML(html)
            if  len(formatHtml.xpath('.//img[@class = "captcha_image"]/@src')) > 0:      #判断是否有验证码
                captchaURL = formatHtml.xpath('.//img[@class = "captcha_image"]/@src')[0]
                captchaid = re.search('id=(.*?)&size',captchaURL).group(1)
            else:
                captchaURL = None

            if captchaURL:
                captcha = self.open(captchaURL).content
                # print(self.captchaFile+"---------"+os.getcwd())
                with open(self.captchaFile, "wb") as output:
                    output.write(captcha)
                print("=" * 40)
                print("已打开验证码图片，请识别！")
                # subprocess.call(self.captchaFile, shell=True)
                im = Image.open(self.captchaFile)
                im.show()
                captcha = input("请输入验证码：")
                os.remove(self.captchaFile)

            # 发送POST请求
            data = {
                "source": None,
                "form_password": self.__password,
                "form_email":self.__username,
                "captcha-solution":captcha,
                "captcha-id":captchaid,
                "login": "登录"
            }
            res = self.__session.post(self.__loginURL, data=data)
            print("=" * 40)
            print(res)
            if res.text.find('nav-user-account')>0:
                print("登录成功")
                self.__saveCookie()
                return True
                break
            else:
                print("登录失败")
                return False


    def __saveCookie(self):
        """cookies 序列化到文件
        即把dict对象转化成字符串保存
        """
        with open(self.cookieFile, "w") as output:
            cookies = self.__session.cookies.get_dict()
            json.dump(cookies, output)
            # print("=" * 50)
            print("已在同目录下生成cookie文件：", self.cookieFile)

    def __loadCookie(self):
        """读取cookie文件，返回反序列化后的dict对象，没有则返回None"""
        if os.path.exists(self.cookieFile):
            with open(self.cookieFile, "r") as f:
                cookie = json.load(f)
                return cookie
        return None

    def  loadCookie(self):
        if os.path.exists(self.cookieFile):
            with open(self.cookieFile, "r") as f:
                cookie = json.load(f)
                print('成功加载cookies')
                return cookie
        return None


    def getSession(self):
        return self.__session

def main():

    global clr,postDateTime,postData
    clr = Color()
    clr.print_green_text('*'*40)
    clr.print_green_text('##  Python  3.4')
    clr.print_green_text('##  Author  Liam')
    clr.print_green_text('##  Date    11/13/2016')
    clr.print_green_text('##  Crawl   DoubanCrawl')
    clr.print_green_text('*'*40)

    clr.print_green_text('Enter to Open File')
    dlg = win32ui.CreateFileDialog(1)   # 表示打开文件对话框
    dlg.SetOFNInitialDir('C:/')   # 设置打开文件对话框中的初始显示目录
    dlg.DoModal()
    filename = dlg.GetPathName()
    clr.print_green_text('Open File : '+filename)

    if filename is None or filename == '':
       sys.exit(0)

    postDateTime = input('请输入抓取截止日期 (格式：2016-1-1):')
    # postDateTime = '2016-11-1'
    postData = []

    f = open(filename,'rb')
    task_lines = [i for i in f.readlines()]
    f.close()

    client = DoubanClient()
        # 测试账号
    # client.login("liam.bao@cicdata.com", "cicdata123456")

    while True:    	
	    userAccount = input('请输入你的用户名\n>  ')
	    userPas = input("请输入你的密码\n>  ")
	    flat = client.login(userAccount,userPas)
	    if flat:
	    	break

    count = 0
    try:

        for line in task_lines:
            try:
                count += 1
                line = str(line, encoding='utf-8')
                line = line.strip()
                
                if not line:
                    continue
                clr.print_green_text('Start Parsing Url : '+str(line))
                if len(line):
                    doCapture(line)
                    clr.print_green_text('Url: '+str(line)+ ' parsing Done!')
                # for i in data:
                #     allthread.append(i)

                if len(postData) > 20000:
                    clr.print_green_text('Counts ' + str(len(postData)) + '  posts')
                    getExcel(postData)
                    postData = []
                    waitTime = random.uniform(3, 5)
                    clr.print_green_text("  Wait for "+str(int(waitTime))+" Seconds!")
                    time.sleep(waitTime)
            except Exception as err:
                clr.print_red_text (err)
        if postData:
            clr.print_green_text('Counts ' + str(len(postData)) + '  posts')
            getExcel(postData)

    except Exception as err:
        clr.print_red_text(err)
    finally:
        os.remove(DoubanClient.cookieFile)



def getExcel(data):
    clr = Color()
    try:
        title = ['subject','content','dateOfPost','NumofUseful','posterName','posterURL','rating','posterID','threadURL']

        file_name = '%s%s' % ('Output_',("%d" % (time.time() * 1000)))
        
        workbook = wx.Workbook(file_name+'.xlsx')
        worksheet = workbook.add_worksheet('post')
        for i in range(len(data)):
            for j in range(len(title)):
                if i==0:
                    worksheet.write(i, j, title[j])
                worksheet.write(i+1, j, data[i][j])

        workbook.close()
        clr.print_green_text('\n File '+file_name+' Done!')
    except Exception as err:
        clr.print_red_text(err)


if __name__ == '__main__':
    main()


####
####version test pull
####  git pull
####version test push
