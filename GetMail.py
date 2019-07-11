# -*- coding: utf-8 -*-
#!/usr/bin/python
#coding=utf-8
#

# 取邮件及其附件保存到指定目录（WORK_FOLDER）
# panhwa@hotmail.com
# 2018.1.11
#
#
# 注意：修改了python包中的文件：
# 1.因为邮件长度超过了poplib中 _MAXLINE的定义，直接修改了pyton2.7的poplib.py _MAXLINE=65536
# 2.python2.7 mail模块处理特殊媒体类型时，会简单返回 text/plain，修改了 email/message.py, 改为 text/unknown 方便捕获
#
# todo：同时读多个服务器
import sys
import inspect
import poplib
import cStringIO
import email
import base64,os,sys,re
import time
import datetime
import ConfigParser
import codecs
import socket
from HTMLParser import HTMLParser

ENV_INFO = ""
DEBUG = False # read from ini
JUST_LOCAL_EML = False
EML_PATH_NAME = "Eml/"

def Usage():
    print "Please edit GetMail.ini first."
    # print ""

#保留差错日志
def ErrorLog(ErrInfo):
    """write error log to error.log in current path"""
    global ENV_INFO
    global DEBUG
    callerframerecord = inspect.stack()[1]
    frame = callerframerecord[0]
    info = inspect.getframeinfo(frame)
    WriteStr = time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time()))
    #WriteStr = WriteStr + " [" + info.filename + "/" + info.function + "/" + str(info.lineno) + "][" + ENV_INFO + "] "
    WriteStr = WriteStr + " ["                        + info.function + "/" + str(info.lineno) + "][" + ENV_INFO + "] "
    WriteStr = WriteStr + ErrInfo
    ErrorFile = open('error.log','ab')
    ErrorFile.writelines(WriteStr)
    ErrorFile.writelines("\n")
    ErrorFile.close()

def MyUnicode(Data,ChrCode):
    """Do a little more things than Unicode(). """
    global ENV_INFO
    global DEBUG
    if DEBUG:
        print "Data type is ",type(Data).__name__
    if isinstance(Data, unicode):#已经是Unicode
        return Data
    if DEBUG:
        print "ChrCode=[",ChrCode,"]",
    if not ChrCode or ChrCode == "":# or ChrCode.lower().find("utf8") >= 0
        ChrCode = "utf-8"
    try:
        unic = unicode(Data,ChrCode,'replace')
    except Exception,e:
        unic = Data
        ErrorLog("MyUnicode error, "+ str(e)+", source data:")
        ErrorLog(Data)
    return unic

def MyMailDecode(HeaderString):
    """Decode something like '=?gb18030?=...',
    python 2.7 email model seems can't decode Chinese correctly in headers(from/to/cc/bcc etc.)"""
    global ENV_INFO
    global DEBUG
    if not HeaderString:
        HeaderString = "<N/A>"
        # return None
    #找到编码的子串
    p = re.compile('=\?.*?\?[qQbB]\?.*?\?=')
    rtStr = ""
    StrB = 0
    #缺省utf-8，试图修复qq邮箱"撤回邮件"没有编码头=?utf-8?的bug
    ChrCode = "utf-8"
    #for 每个子串
    for match in p.finditer(HeaderString):
        #保持不变的串的结束位置
        StrE=match.span()[0]
        #含有编码的子串
        sub = HeaderString[match.span()[0]:match.span()[1]]
        h = email.Header.Header(sub)
        dh = email.Header.decode_header(h)
        ChrCode = dh[0][1]
        SourceStr = dh[0][0]
        # ChrCode = re.search('=\?(.*?)\?',sub).group(1)
        # SourceStr = re.search('=\?.*?\?(.*?)\?=',sub).group(1)
        #上一个处理完成的串 + 之间不用处理的部分 + 解码后的子串
        #rtStr = rtStr + HeaderString[StrB:StrE] + MyUnicode(SourceStr,ChrCode)
        rtStr = rtStr + HeaderString[StrB:StrE] + SourceStr
        #下一个保持不变的串的开始位置
        StrB=match.span()[1]
        if DEBUG:
            print "sub=",sub
            print "SourceStr=",SourceStr,"ChrCode=",ChrCode
    rtStr = rtStr + HeaderString[StrB:]
    # 如果有调用email.Header.decode_header解析过，用最后一个的ChrCode编码（通常是统一的）
    return MyUnicode(rtStr,ChrCode)
    # return rtStr

# 用到的全部头部标记名字
HEADS=["Received","From","To","Cc","Bcc","Date","Message-Id","Subject",
    "DKIM-Signature","MIME-Version","Content-Transfer-Encoding","Content-type"]

def GetHeader(MailMsg,HeaderName):
    """Get header by name form email message. """
    global ENV_INFO
    global DEBUG
    if DEBUG:
        print 'Header [',HeaderName,']=', MailMsg.get(HeaderName)
    return "[ "+HeaderName+" = "+MyMailDecode(MailMsg.get(HeaderName))+"]\n"

def GetMsg(Pop,EmlFolder,MailNo,FromSvr):
    """get email.message object from local file or pop3 server"""
    # 本地有就读本地，没有就pop上取
    EmlFileName = EmlFolder + str(MailNo)+".eml"
    LocalEmlExist = os.path.isfile(EmlFileName)
    if not FromSvr and LocalEmlExist:
        print "from local"
        try:
            f=open(EmlFileName,'rb')
            msg = email.message_from_file(f)
        except:
            ErrorLog("Read eml file error:"+str(e))
        finally:
            f.close()
    else:
        print "from server"
        mail = Pop.retr(MailNo)

        buf = cStringIO.StringIO()
        for j in mail[1]:
            print >>buf, j
        buf.seek(0)
        msg = email.message_from_file(buf)
    return msg

def SaveEmlFile(Msg,EmlFolder,MailNo,Rewrite):
    """Save eml file to the EmlFolder. """
    try:
        if not os.path.exists(EmlFolder):
            os.makedirs(EmlFolder)
        EmlFileName = EmlFolder + str(MailNo)+".eml"
        # 文件不存在或者强制重写
        if not os.path.isfile(EmlFileName) or Rewrite:
            EmlFile = open(EmlFileName,'wb')
            EmlFile.writelines(Msg.as_string())
            EmlFile.close()
    except Exception,e:
        ErrorLog("Write eml file error:"+str(e))

class Html2Text(HTMLParser):
    """Rewrite html in mail into plain text. """
    #todo 更精细的tag控制
    Text = ""
    def handle_starttag(self, tag, attrs):
        donothing=tag
        # print "Encountered a start tag:", tag

    def handle_endtag(self, tag):
        # donothing=tag
        # print "Encountered an end tag :", tag
        if tag == 'p':
            self.Text += ("\n")
        if tag == 'br':
            self.Text += ("\n")
    def handle_data(self, data):
        if re.match("\S",data):#非空才打印
            try:
                self.Text += (data)
                # self.Text += ('\n')
            except Exception,e:
                self.Text += ("[- write text/html error -]")
                ErrorLog("Mail body write error:")
                ErrorLog(str(e))
                # ErrorLog(body)
                # todo throw sth

# ---------------------------------------------------------------------------
# 
#
#
def main(argv):
    global ENV_INFO
    global DEBUG
    global JUST_LOCAL_EML
    ENV_INFO ="Main thread"
    CurrentPath = os.path.split(os.path.realpath(__file__))[0] + "/"
    print "CurrentPath=",CurrentPath
    # 1、读配置文件-------------------------------------
    config = ConfigParser.ConfigParser()
    config.read(CurrentPath+"GetMail.ini")
    PopServer = config.get("Mail","PopServer")
    MailUser = config.get("Mail","MailUser")
    MailPW = config.get("Mail","MailPW") #todo：密码加密
    StopAfter = int(config.get("Mail","StopAfter"))
    SaveAttachment = bool(config.get("Mail","SaveAttachment"))
    LastMailNo = int(config.get("Mail","LastMailNo"))
    if PopServer =="pop.mail":
        Usage()
        raw_input("Press enter to exit > ")
        return
    if int(config.get("Control","DEBUGLevel"))>0:
        DEBUG = True
    else:
        DEBUG = False

    LocalMailPath = config.get("Mail","LocalMailPath")
    LocalMailPath = LocalMailPath + "/" + PopServer + "/" + MailUser + "/"
    EmlFolder = LocalMailPath + EML_PATH_NAME
    if(LastMailNo<0):
        LastMailNo = 0

    print "DEBUG =",str(DEBUG)
    print "StopAfter =",str(StopAfter)
    print "SaveAttachment =",str(SaveAttachment)
    print "LocalMailPath =",str(LocalMailPath)
    print "PopServer =",str(PopServer)
    print "MailUser =",str(MailUser)
    print "LastMailNo =",LastMailNo

    # 开始连接POP3 Server
    try:
        Pop3 = poplib.POP3( PopServer )   #邮件下载服务器
        Pop3.user( MailUser )    #邮箱地址
        Pop3.pass_(MailPW)   #密码
        #有多少封信（最后一封信的编号）
        num = len(Pop3.list()[1])
        print "Mail count =",num
        config.set("Mail","MailCount",num)
        config.write(open(CurrentPath+"GetMail.ini", "w"))
    except socket.error as e:
        ErrorLog("socket error:"+str(e))
        print "连接服务器失败，仅本地缓存"
        JUST_LOCAL_EML = True #todo
    except Exception,e:
        ErrorLog("connect pop3 server error:"+str(e))
        #todo

    #检索邮件范围
    if StopAfter > 0:
        rge = range(LastMailNo+1,LastMailNo+StopAfter+1,1)
    else:
        # 从LastMailNo到最近邮件num
        rge = range(LastMailNo+1,num+1,1)
    if DEBUG:
        print "Begin deal with mail No."+str(LastMailNo+1)+" to "+str(num)
    EveryThingOK = True
    for MailNo in rge:
        ENV_INFO = "Mail No."+str(MailNo)
        ii = 1 #文件名非法时用的临时序号
        #解析信件内容
        try:
            Msg = GetMsg(Pop3,EmlFolder,MailNo,False)
        except socket.error as e:
            ErrorLog("socket error:"+str(e))
            break
        except Exception,e:
            ErrorLog("retr mail error:"+str(e))
            continue
        #保存邮件本地缓存
        SaveEmlFile(Msg,EmlFolder,MailNo,False)#todo:配置Rewrite

        # 保存邮件体
        MailTxtName = LocalMailPath + str(MailNo) +"-mail.txt"
        ftxt = codecs.open(MailTxtName, 'wb', encoding='utf-8')
        # ftxt = open(MailTxtName, 'wb')
        ftxt.write('['+ '#'*26+ ' Mail No.'+ str(MailNo) +  '#'*26+ ']\n')
        # if DEBUG:
            # print "_Mail source code begin_"+'_'*30
            # print Msg
            # print "_Mail source code end_"+'_'*32
        # Mail header
        # Msg.get("Auto-Submitted"),Msg.get("X-QQ-MAIL-TYPE"),Msg.get("X-QQ-STYLE") # ,Msg.get("")
        for h in HEADS:
            try:
                ftxt.write(GetHeader(Msg,h))
            except Exception,e:
                EveryThingOK = False
                ErrorLog("Mail header["+ h +"] write error - : ")
                ErrorLog(str(e))
                ftxt.write("[-- Header write error, error.log for detail. -- ]\n")

        #处理消息体，包括附件
        ThisMailHasHtml = False
        # print "[for each part of body]"
        # 循环信件中的每一个mime的数据块
        for Part in Msg.walk():
            # print Part
            # print Part.get_content_type()
            if Part.is_multipart(): # 这里要判断是否是multipart，是的话，里面的数据是无用的，至于为什么可以了解mime相关知识。
                # print "[it's multipart,do nothing.]"
                continue
            if DEBUG:
                print '[', '-'*26, 'Mail part', '-'*26, ']'
            ftxt.write('['+ '-'*61+ ']\n')
            name = Part.get_param("name") #如果是附件，这里就会取出附件的文件名
            # todo：这样判断不够严谨，暂时没有发现例外
            if name:
                #有附件
                # fname = MyMailDecode(name)
                fname = re.sub("[\\\/\\\:\*\?\"\<\>\|\t\r\n]",'',MyMailDecode(name))
                # print "sub",re.sub("[\\\/\\\:\*\?\"\<\>\|\t\r\n]",'',MyMailDecode(name))
                # print "no sub",fname
                #加上邮件编号，避免重名，便于检索
                fname = LocalMailPath + str(MailNo) +"A-" + fname
                if SaveAttachment:
                    data = Part.get_payload(decode=True) #　解码出附件数据，然后存储到文件中
                    try:
                        fAttch = open(fname, 'wb') #注意一定要用wb来打开文件，因为附件一般都是二进制文件
                        # todo: 重名处理
                    except Exception,e:
                        # EveryThingOK = False
                        # todo: 1、保留原文件后缀名
                        #       2、修改原文件名为合法，保留原名主要内容
                        ErrorLog(str(e))
                        ftxt.write("[-- Name of attachment ("+fname+") is invalid, use 'mailno-1' instead. -- ]\n")
                        fname = LocalMailPath + str(MailNo) +"A-"+str(ii)
                        fAttch = open(fname, 'wb')
                        ii+=1
                    fAttch.write(data)
                    fAttch.close()
                    ftxt.write("[ -- Attachment '"+ fname+ "' has been saved to "+ LocalMailPath+ " -- ]\n")
                else:
                    ftxt.write("[ -- Attachment '"+ fname+ "' -- ]\n")
            else:
            #不是附件，是文本内容
            #todo:其他媒体类型还不能正确处理
                body = Part.get_payload(decode=True) # 解码出文本内容，直接输出来就可以了。
                bcode = Part.get_charsets()[0]
                body = MyUnicode(body,bcode)
                ContentType = Part.get_content_type()
                if DEBUG:
                    print "Boundary=",Part.get_boundary()
                    print "ContentType=",ContentType
                    print "-name=",Part.get_param("name")
                    print "-charset=",Part.get_param("charset")
                    print "Items=",Part.items()
                    print "__Part content_____"
                    print body
                    print "__End of part content_____"
                if ContentType == "text/html":
                    #超文本只简单的读出各tag的文本部分及主要的回车
                    #todo 更精确解析或直接另存为html文件
                    ThisMailHasHtml = True # 无效，见下
                    ftxt.write("[ --"+ContentType+"-- ]\n")
                    # html 转 text
                    parser = Html2Text()
                    try:
                        parser.feed(body)
                    except Exception,e:
                        EveryThingOK = False
                        ErrorLog("parse html error:")
                        ErrorLog(str(e))
                    ftxt.write(parser.Text)
                elif ContentType == "text/plain":#文本
                    if re.match("\S",body) and not ThisMailHasHtml:#非空才打印
                        #todo: 通常text/plain 在先，ThisMailHasHtml不起作用
                        ftxt.write("[ --"+ContentType+"-- ]\n")
                        try:
                            ftxt.write(body)
                            ftxt.write('\n')
                        except Exception,e:
                            EveryThingOK = False
                            ftxt.write("[- write text/plain error -]")
                            ErrorLog("Mail body write error:")
                            ErrorLog(str(e))
                            # ErrorLog(body)
                elif ContentType == "text/unknown":
                    #已知bug，python缺省时使用text/plain类型，导致qqmail的部分类型识别错误，暂时修改message.py为text/unknown以捕捉此问题
                    ErrorLog("text/unknown QQMail ContentType, marked, fix late, maybe. ")
                else:#不认识的内容
                    ftxt.write("[ --"+ContentType+"-- ]")
                    # 如果能如预期找出文件名的附件，当做附件保存处理
                    # if
                    # 不能识别，报错
                    # else:
                    ErrorLog("Unknown ContentType:"+ContentType)
                    # ErrorLog(body)
                ftxt.write("\n")
                # print unicode(body,Part.get_charsets()[0])
            # print '+'*60 # 用来区别各个部分的输出
        print '\n',
        ftxt.write('['+ '#'*23+ ' End of Mail No.'+ str(MailNo) +  '#'*23+ ']\n')
        ftxt.close()
        #rename ftxt
        NewTxtName = str(MailNo) + "M-[" + MyMailDecode(Msg.get("Subject")) + "]-[" + MyMailDecode(Msg.get("Date")) +"].txt"
        NewTxtName = LocalMailPath + re.sub("[\\\/\\\:\*\?\"\<\>\|\t\r\n]",'',NewTxtName)
        try:
            os.rename(MailTxtName,NewTxtName)
        except Exception,e:
            ErrorLog("Rename Mail text file error:")
            ErrorLog("from["+MailTxtName+"]to["+NewTxtName+"]")
            ErrorLog(str(e))

        if EveryThingOK:
            print "All OK, write lastmailno=",MailNo
            config.set("Mail","LastMailNo",MailNo)
            config.write(open(CurrentPath+"GetMail.ini", "w"))
        else:
            print "Something wrong in mail No.=", MailNo
            break
    Pop3.quit()
    if EveryThingOK:
        config.remove_option("Mail","SomethingWrongLastTime")
    else:
        config.set("Mail","SomethingWrongLastTime",MailNo)
    config.set("Mail","LastTime",time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time())))
    config.write(open(CurrentPath+"GetMail.ini", "w"))

    # raw_input("Press enter to exit > ")

if __name__ == "__main__":
    main(sys.argv)

