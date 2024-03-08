# -*- coding: utf-8 -*-
#!/usr/bin/python
# coding=utf-8
# python 3

# 取邮件及其附件保存到指定目录（WORK_FOLDER）
# panhwa@hotmail.com
# 2018.1.11
#
# 注意：修改了python包中的文件：
# 1.因为邮件长度超过了poplib中 _MAXLINE的定义，直接修改了pyton2.7的poplib.py _MAXLINE=65536
# 2.python2.7 mail模块处理特殊媒体类型时，会简单返回 text/plain，修改了 email/message.py, 改为 text/unknown 方便捕获
#
# rewrite for python 3
# 去掉了全部的无类型判断 except Exception, 改为抛出未捕获的异常
# 日常采用调试状态运行，以便发现问题
# 1._MAXLINE=65536 已改为本地值覆盖，见下Line 41
# 2.mail模块媒体类型修改已取消
# 2020.4.8
#

# TODO 密码加密
# TODO 更多容错处理，如 -EOF重连，noop等
# TODO 顺序检查本地文件是否齐全 1-N
# TODO 检查本地文件是否格式正确
# TODO 同时读多个服务器
from html.parser import HTMLParser
import socket
import codecs
import configparser
import datetime
import time
import io
import re
import os
import base64
import email.parser
import sys
import inspect
import poplib
from tkinter import Tk,Label,Entry #,messagebox

# 覆盖 poplib.py _MAXLINE=2048
poplib._MAXLINE = 65536


ENV_INFO = "GetMail.py"  # print errorlog
DEBUG_LEVEL = 99  # if not read from ini file
JUST_LOCAL_EML = False
EML_PATH_NAME = "Eml/"
BREAK_A_SECOND = 60
ConfigFileName = "./Config.ini"
ErrorLogName = "./error.log"
MailPW = ""


def Usage():
    print("Please edit GetMail.ini first.")
    # print ""


def getInput(title, message):
    def return_callback(event):
        # print('quit...')
        root.quit()

    def close_callback():
        # messagebox.showinfo('message', 'no click...')
        global rt
        rt = ""
        root.quit()
    root = Tk(className=title)
    root.wm_attributes('-topmost', 1)
    screenwidth, screenheight = root.maxsize()
    width = 300
    height = 100
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width)/2, (screenheight - height)/2)
    root.geometry(size)
    root.resizable(0, 0)
    lable = Label(root, height=2)
    lable['text'] = message
    lable.pack()
    entry = Entry(root)
    entry.bind('<Return>', return_callback)
    entry.pack()
    entry.focus_set()
    root.protocol("WM_DELETE_WINDOW", close_callback)
    root.mainloop()
    rt = entry.get()
    root.destroy()
    return rt

# 保留差错日志
def ErrorLog(ErrInfo):
    """write error log to error.log in current path"""
    global ENV_INFO, BREAK_A_SECOND, ErrorLogName
    callerframerecord = inspect.stack()[1]
    frame = callerframerecord[0]
    info = inspect.getframeinfo(frame)
    WriteStr = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
    # WriteStr = WriteStr + " [" + info.filename + "/" + info.function + "/" + str(info.lineno) + "][" + ENV_INFO + "] "
    WriteStr = WriteStr + " [" + info.function + \
        "/" + str(info.lineno) + "][" + ENV_INFO + "] "
    WriteStr = WriteStr + ErrInfo
    ErrorFile = open(ErrorLogName, 'at')
    ErrorFile.writelines(WriteStr)
    ErrorFile.writelines("\n")
    ErrorFile.close()

def Debug_Print(Info, Print_Level):
    global DEBUG_LEVEL
    if DEBUG_LEVEL >= Print_Level:
        print(Info)

# CODES = ['UTF-8', 'GB18030', 'GBK', 'hz', 'BIG5', 'GB2312', 'iso2022_jp_2', 'big5hkscs', 'cp950', ]
# All python codecs built-in
CODES = ['utf_8', 'ascii', 'gb18030', 'hz', 'gbk', 'big5', 'big5hkscs', 'cp037', 'cp424', 'cp437', 'cp500', 'cp737', 'cp775', 'cp850', 'cp852', 'cp855', 'cp856',
         'cp857', 'cp860',
         'cp861', 'cp862', 'cp863', 'cp864', 'cp865', 'cp866', 'cp869', 'cp874', 'cp875', 'cp932', 'cp949', 'cp950', 'cp1006', 'cp1026', 'cp1140', 'cp1250',
         'cp1251', 'cp1252', 'cp1253', 'cp1254', 'cp1255', 'cp1256', 'cp1257', 'cp1258', 'euc_jp', 'euc_jis_2004', 'euc_jisx0213', 'euc_kr', 'gb2312',
         'iso2022_jp', 'iso2022_jp_1', 'iso2022_jp_2', 'iso2022_jp_2004', 'iso2022_jp_3', 'iso2022_jp_ext', 'iso2022_kr', 'latin_1', 'iso8859_2',
         'iso8859_3', 'iso8859_4', 'iso8859_5', 'iso8859_6', 'iso8859_7', 'iso8859_8', 'iso8859_9', 'iso8859_10', 'iso8859_13', 'iso8859_14', 'iso8859_15',
         'johab', 'koi8_r', 'koi8_u', 'mac_cyrillic', 'mac_greek', 'mac_iceland', 'mac_latin2', 'mac_roman', 'mac_turkish', 'ptcp154', 'shift_jis',
         'shift_jis_2004', 'shift_jisx0213', 'utf_16', 'utf_16_be', 'utf_16_le', 'utf_7']

def AutoDecode(data, charset=None):
    # TODO search mail file, ’cause there is usualy only one charset in one mail.
    # if (data.__class__.__name__ == "str" and (charset is None or data=="" )) or data is None: # 无需解码
    global CODES
    if data is None or data.__class__.__name__ == "str":  # 无需解码
        return data
    if charset is not None:
        CODES = [charset.upper()]+CODES
    # 遍历编码类型
    for code in CODES:
        try:
            rt = data.decode(encoding=code)
            # Debug_Print("Codec is "+code,2)
            return rt
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(code, data, 0, -1, "全部编码均解析失败")
    # return data

def AutoOpen(file, mode):
    global CODES
    for code in CODES:
        try:
            f = open(file, mode, encoding=code)
            f.read()
            f.seek(0)
            # Debug_Print("File Codec is "+code,2)
            return f
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(code, None, 0, -1, "用全部编码均打开失败")

def MyMailDecode(HeaderString):
    try:
        """Decode something like '=?gb18030?=...',
        python 2.7 email model seems can't decode Chinese correctly in headers(from/to/cc/bcc etc.)"""
        # if not HeaderString:
        if HeaderString is None or HeaderString == "":
            HeaderString = ""
            # return None
        # 找到编码的子串
        p = re.compile(r'=\?.*?\?[qQbB]\?.*?\?=')
        rtStr = ""
        StrB = 0
        # 缺省utf-8，试图修复qq邮箱"撤回邮件"没有编码头=?utf-8?的bug
        # ChrCode = "utf-8"
        # for 每个子串
        for match in p.finditer(HeaderString):
            # 保持不变的串的结束位置
            StrE = match.span()[0]
            # # 含有编码的子串
            # sub = HeaderString[match.span()[0]:match.span()[1]]
            # h = email.header.Header(sub)
            # dh = email.header.decode_header(h)
            # dh = email.header.decode_header(match.group())
            # ChrCode = dh[0][1]
            # SourceStr = dh[0][0]
            # ChrCode = re.search('=\?(.*?)\?',sub).group(1)
            # SourceStr = re.search('=\?.*?\?(.*?)\?=',sub).group(1)
            # 上一个处理完成的串 + 之间不用处理的部分 + 解码后的子串
            # rtStr = rtStr + HeaderString[StrB:StrE] + MyUnicode(SourceStr,ChrCode)
            value, charset = email.header.decode_header(match.group())[0]
            if charset:
                value = AutoDecode(value, charset)

            rtStr = rtStr + HeaderString[StrB:StrE] + value
            # 下一个保持不变的串的开始位置
            StrB = match.span()[1]
            # Debug_Print("sub="+sub, 2)
            # Debug_Print( "SourceStr="+SourceStr+"ChrCode="+ChrCode,2 )
        rtStr = rtStr + HeaderString[StrB:]
        # 如果有调用email.Header.decode_header解析过，用最后一个的ChrCode编码（通常是统一的）
        return str(rtStr)
    except UnicodeDecodeError:
        # 有个奇怪的bug，直接解中文有问题，上面的处理会报中文编码错误 AutoDecode() 失败
        if HeaderString is None or HeaderString == "":
            return ""
        value, charset = email.header.decode_header(HeaderString)[0]
        return AutoDecode(value, charset)
    # return rtStr


# 用到的全部头部标记名字
HEADS = ["Received", "From", "To", "Cc", "Bcc", "Date", "Message-Id", "Subject",
         "DKIM-Signature", "MIME-Version", "Content-Transfer-Encoding", "Content-type"]


def GetHeader(MailMsg, HeaderName):
    """Get header by name form email message. """
    strHeader = MyMailDecode(MailMsg.get(HeaderName))
    # Debug_Print('Header ['+HeaderName+']=' + strHeader, 2)
    if strHeader == "":
        rt = ""
    else:
        rt = "[ "+HeaderName+" = "+strHeader+"]\n"
    return rt


def GetAllHeader(Msg):
    rt = ""
    for h in HEADS:
        rt += GetHeader(Msg, h)
    return rt


def GetMsg(Pop, EmlFolder, MailNo, FromSvr):
    """get email.message object from local file or pop3 server"""
    # 本地有就读本地，没有就pop上取
    EmlFileName = EmlFolder + str(MailNo)+".eml"
    LocalEmlExist = os.path.isfile(EmlFileName)
    octets = 0
    if not FromSvr and LocalEmlExist:
        Debug_Print("Get Mail from local cache file(*.eml)", 2)
        try:
            f = AutoOpen(EmlFileName, 'rt')
            msg = email.message_from_file(f)
        except BaseException as e:
            ErrorLog("Read eml file error:"+str(e))
        finally:
            f.close()
    else:
        Debug_Print("Get Mail {} from server(LocalEmlExist={})".format(
            str(MailNo), str(LocalEmlExist)), 2)
        resp, lines, octets = Pop.retr(MailNo)
        Debug_Print((resp, octets), 1)
        if b"+OK" not in resp:
            ErrorLog("Pop.retr:"+resp.decode())
        content = AutoDecode(b'\r\n'.join(lines))
        emailparser = email.parser.Parser()
        msg = emailparser.parsestr(content)
    return msg,octets


def SaveEmlFile(Msg, EmlFolder, MailNo, Rewrite):
    """Save eml file to the EmlFolder. """
    # try:
    if not os.path.exists(EmlFolder):
        os.makedirs(EmlFolder)
    EmlFileName = EmlFolder + str(MailNo)+".eml"
    # 文件不存在或者强制重写
    if not os.path.isfile(EmlFileName) or Rewrite:
        EmlFile = open(EmlFileName, 'wt')
        EmlFile.writelines(Msg.as_string())
        EmlFile.close()
    # except Exception as e:
    #     raise(e)
    #     ErrorLog("Write eml file error:"+str(e))


class Html2Text(HTMLParser):
    """Rewrite html in mail into plain text. """
    # todo 更精细的tag控制
    Text = ""
    CurrentTag = ""

    # def handle_starttag(self, tag, attrs):
    # if tag == 'head':
    # self.CurrentTag = "head"
    # donothing=tag
    # print "Encountered a start tag:", tag

    def handle_endtag(self, tag):
        # print "Encountered an end tag :", tag
        if tag == 'p':
            # tag大小写处理
            self.Text += ("\n")
        if tag == 'br':
            self.Text += ("\n")

    def handle_data(self, data):
        # if self.CurrentTag == "head":
        #     # donothing
        #     pass
        # else:
        if re.match(r"\S", data):  # 非空才打印
            try:
                self.Text += (data)
                # self.Text += ('\n')
            except Exception as e:
                raise(e)
            #     self.Text += ("[- write text/html error -]")
            #     ErrorLog("Mail body write error:")
            #     ErrorLog(str(e))
                # ErrorLog(body)
                # todo throw sth


def GetCharset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        m = re.match(r".*charset=['\"]*([^'\"]+)['\"]*",
                     content_type, re.M | re.S)
        # m = re.match(r".*",content_type,re.M|re.S)
        if m is not None:
            charset = m.group(1)
        # pos = content_type.find('charset=')
        # if pos >= 0:
        #     charset = content_type[pos+8:].strip()
    return charset


def Msg2TxtFile(Mail, MailNo, ftxt, AttPath):
    """
    取消息内容，写入f
    消息体是递归调用
    Args:
        f ([type]): [description]
        Msg ([type]): [description]
    """
    # 循环信件中的每一个mime的数据块
    # 处理消息体，包括附件
    for Msg in Mail.walk():  # 遍历整个树，深度优先
        # 消息头
        # 头部解读写入文件
        ftxt.write('[ Head ' + '-'*50 + ']\n')
        ftxt.write(GetAllHeader(Msg))
        ftxt.write('[' + '-'*55 + ']\n')
        # 提取用于控制后续解码相关的头部信息
        ContentType = Msg.get_content_type().lower()
        ContentID = Msg.get_all("Content-ID")
        Charset = GetCharset(Msg)
        ContentTransferEncoding = Msg.get_all("Content-Transfer-Encoding")
        FileName = Msg.get_filename()
        if Msg.is_multipart():
            # 除了头部，其他没有
            if ContentType == "message/rfc822" and FileName:
                continue
            else:
                continue

            #     # 模块将message/rfc822视作多部份part，继续向内处理
            #     # TODO 考虑选择另一种方案，跳过go into，保存为附件
        # print(Msg.as_string())
        ####################################################################################
        # GUESS 似乎QQMail服务有bug
        # FIXME 效果有问题，暂时屏蔽了修改
        if ContentTransferEncoding and ContentType and "base64" in ContentTransferEncoding and ContentType not in ("text/plain", "text/html"):
            Changed = False
            newlines = []
            lines = Msg.get_payload().split("\n")
            LineLength = len(lines[0])
            for line in lines:
                if len(line) > LineLength:
                    # Changed = True
                    ErrorLog(
                        "Find a line with wrong length in base64 string, maybe 'QQMail-xxx' in line.")
                    ErrorLog(line)
                elif "QQMail-" in line:
                    raise(
                        KeyError("Find 'QQMail-xxx' in base64 string, maybe QQMail bug"))
                newlines.append(line[:LineLength])
            if Changed:
                newpayload = "\n".join(newlines)
                Msg.set_payload(payload=newpayload)
        # end GUESS
        ####################################################################################
        # if HeadContentTransferEncoding or ContentTransferEncoding:
        if ContentTransferEncoding:
            data = Msg.get_payload(decode=True)
        else:
            data = Msg.get_payload()
        if ContentType in ("text/plain", "text/html", "message/rfc822"):
            # 文本要解码
            # data = data.decode(Charset)
            # 有编码错误的情况，所以用遍历的方法
            data = AutoDecode(data, Charset)
        ftxt.write('[' + '-'*61 + ']\n')
        # 准备保存附件、嵌入文件
        if FileName:
            # 有附件
            SaveFile = True
            fname = re.sub(r"[\\\/\\\:\*\?\"\<\>\|\t\r\n]",
                           '', MyMailDecode(FileName))
            # 加上邮件编号，避免重名，便于检索
            fname = AttPath + str(MailNo) + "A-" + fname
        elif ContentType in ("application/png",) and ContentID:
            # 嵌入式文件
            SaveFile = True
            fname = AttPath + str(MailNo) + "A-CID-"+str(ContentID) + ".png"
        else:
            SaveFile = False
        if SaveFile:
            if AttPath:  # AttPath参数控制是否保存附件
                try:
                    fAttch = open(fname, 'wb')
                    # todo: 重名处理
                except Exception as e:
                    raise(e)
                #     # EveryThingOK = False
                #     ErrorLog(str(e))
                #     ftxt.write("[-- Name of attachment ("+fname+") is invalid, use 'mailno-1' instead. -- ]\n")
                #     fname = LocalMailPath + str(MailNo) +"A-"+str(ii)
                #     fAttch = open(fname, 'wb')
                #     ii+=1

                if data.__class__.__name__ == "str":
                    data = data.encode()
                fAttch.write(data)
                fAttch.close()
                ftxt.write("[ -- Attachment '" + fname +
                           "' has been saved to " + AttPath + " -- ]\n")
            else:
                ftxt.write("[ -- Attachment '" + fname + "', not save -- ]\n")
        else:
            if ContentType == "text/html":
                # 超文本只简单的读出各tag的文本部分及主要的回车
                # todo 更精确解析或直接另存为html文件
                # ThisMailHasHtml = True  # 无效，见下
                # ftxt.write("[ --"+ContentType+"-- ]\n")
                # html 转 text
                parser = Html2Text()
                try:
                    parser.feed(data)
                except Exception as e:
                    raise(e)
                #     EveryThingOK = False
                #     ErrorLog("parse html error:")
                #     ErrorLog(str(e))
                ftxt.write(parser.Text)
            elif ContentType == "text/plain":  # 文本
                # if re.match(r"\S", data): # and not ThisMailHasHtml:  # 非空才打印
                # 上面正则有误，空字符开始的data不行
                if True:
                    # FIXME: 通常text/plain 在先，ThisMailHasHtml不起作用
                    ftxt.write("[ --"+ContentType+"-- ]\n")
                    try:
                        ftxt.write(data)
                        ftxt.write('\n')
                    except Exception as e:
                        raise(e)
                    #     EveryThingOK = False
                    #     ftxt.write("[- write text/plain error -]")
                    #     ErrorLog("Mail body write error:")
                    #     ErrorLog(str(e))
                        # ErrorLog(data)
            else:  # 不认识的内容
                ftxt.write("[ --"+ContentType+"-- ]")
                # 如果能如预期找出文件名的附件，当做附件保存处理
                # if
                # 不能识别，报错
                # else:
                ErrorLog("Unknown ContentType:"+ContentType)
                # ErrorLog(data)
            # ftxt.write("\n")
            # ftxt.write('[' + '-'*26 + ' \\Body ' + '-'*26 + ']\n')
            # print unicode(data,Msg.get_charsets()[0])
        # print '+'*60 # 用来区别各个部分的输出


class MailCfg(object):
    def __init__(self):
        self.ConfigObj = None
        # self.DEBUG_LEVEL = 0
        # self.BREAK_A_SECOND = ''
        self.PopServer = ''
        self.MailUser = ''
        self.MailPW = ''
        self.StopAfter = 0
        self.SaveAttachment = True
        self.MailCount = 0
        self.LastMailNo = 0
        self.RedoList = ''
        self.RedoListLocal = ''
        self.LocalMailPath = ''
        self.EmlFolder = ''
# ---------------------------------------------------------------------------
#
#
#


def ReadCfg():
    global DEBUG_LEVEL, BREAK_A_SECOND, MailPW
    cfg = MailCfg()
    CurrentPath = os.path.split(os.path.realpath(__file__))[0] + "/"
    Debug_Print("CurrentPath="+CurrentPath, 1)
    # 1、读配置文件-------------------------------------
    if not ConfigFileName:
        ErrorLog(r"Run like this: GetMail c:\Configfile.ini")
        Debug_Print(r"Run like this: GetMail c:\Configfile.ini", 0)
        exit(0)
    config = configparser.ConfigParser()
    if os.path.exists(ConfigFileName):
        config.read(ConfigFileName)
    else:
        ErrorLog("Config file not exist:"+ConfigFileName)
        Debug_Print("Config file not exist:"+ConfigFileName, 1)
        exit(0)
    cfg.ConfigObj = config
    # Control段
    DEBUG_LEVEL = int(config.get("Control", "DebugLevel"))
    BREAK_A_SECOND = int(config.get("Control", "BreakASecond"))
    if BREAK_A_SECOND <= 0 or BREAK_A_SECOND > 10000:
        BREAK_A_SECOND = 60
    # Mail段
    cfg.PopServer = config.get("MailServer", "PopServer")
    cfg.MailUser = config.get("MailServer", "MailUser")
    cfg.MailPW = config.get("MailServer", "MailPW")  # todo：密码加密
    if MailPW == '':
        if cfg.MailPW != "":
            MailPW = cfg.MailPW
        else:
            MailPW = getInput(
                "GetMail", "Please enter Password for\n"+cfg.MailUser+"@"+cfg.PopServer+">")
            if MailPW == "":
                ErrorLog("No Password input.")
                Debug_Print("No Password input.", 1)
                exit(0)
    cfg.StopAfter = int(config.get("MailServer", "StopAfter"))
    cfg.SaveAttachment = bool(config.get("MailServer", "SaveAttachment"))
    cfg.LastMailNo = int(config.get("WorkState", "LastMailNo"))

    cfgRedoList = config.get("WorkState", "RedoList")
    if cfgRedoList.strip() != "":
        cfg.RedoList = [int(x) for x in cfgRedoList.split(",")]
    else:
        cfg.RedoList = []
    cfgRedoListLocal = config.get("WorkState", "RedoListLocal")
    if cfgRedoListLocal.strip() != "":
        cfg.RedoListLocal = [int(x) for x in cfgRedoListLocal.split(",")]
    else:
        cfg.RedoListLocal = []
    if cfg.PopServer == "pop.mail":
        Usage()
        input("Press enter to exit > ")
        return
    cfg.LocalMailPath = config.get("MailServer", "LocalMailPathRoot")
    cfg.LocalMailPath = cfg.LocalMailPath + "/" + \
        cfg.PopServer + "/" + cfg.MailUser + "/"
    cfg.EmlFolder = cfg.LocalMailPath + EML_PATH_NAME
    if(cfg.LastMailNo < 0):
        cfg.LastMailNo = 0

    Debug_Print("DEBUG_LEVEL ="+str(DEBUG_LEVEL), 1)
    Debug_Print("BREAK_A_SECOND ="+str(BREAK_A_SECOND), 1)
    Debug_Print("StopAfter ="+str(cfg.StopAfter), 1)
    Debug_Print("SaveAttachment ="+str(cfg.SaveAttachment), 1)
    Debug_Print("LocalMailPath ="+str(cfg.LocalMailPath), 1)
    Debug_Print("PopServer ="+str(cfg.PopServer), 1)
    Debug_Print("MailUser ="+str(cfg.MailUser), 1)
    Debug_Print("LastMailNo ="+str(cfg.LastMailNo), 1)
    return cfg


def GetMail():
    global ENV_INFO
    global DEBUG_LEVEL
    global JUST_LOCAL_EML
    global BREAK_A_SECOND
    global ConfigFileName
    global MailPW
    ENV_INFO = "Main thread"
    # CurrentPath = os.path.split(os.path.realpath(__file__))[0] + "/"
    # Debug_Print("CurrentPath="+CurrentPath, 1)
    cfg = ReadCfg()
    LastMailNo = cfg.LastMailNo  # 以本地为准
    # 开始连接POP3 Server
    try:
        # Pop3 = poplib.POP3(PopServer)
        Pop3 = poplib.POP3_SSL(cfg.PopServer)
        Pop3.set_debuglevel(0)
        Pop3.user(cfg.MailUser)  # 邮箱地址
        Pop3.pass_(MailPW)  # 密码
        # 有多少封信（最后一封信的编号）
        resp, mails, octets = Pop3.list()
        Debug_Print((resp, octets), 1)
        if b"+OK" not in resp:
            ErrorLog("Pop3.list:"+resp.decode())
        MailNumber = len(mails)
        Debug_Print("Mail count ="+str(MailNumber), 1)
        cfg.MailCount = MailNumber  # 先不写文件，后面每处理一封邮件都会保存
    except (poplib.error_proto, socket.error) as e:
        Debug_Print("Connection error:"+str(e), 1)
        ErrorLog("Connection error:"+str(e))
        # Debug_Print("连接服务器失败，仅本地缓存", 1)
        # JUST_LOCAL_EML = True  # todo
        return
    # except Exception as e:
    #     ErrorLog("connect pop3 server error:"+str(e))
        # todo
    # if MailNumber <= LastMailNo:
    #     Debug_Print("Nothing to do. Quit.", 1)
    #     Pop3.quit()
    #     return
    # 检索邮件范围
    if cfg.StopAfter > 0:
        rge = range(LastMailNo+1, min(LastMailNo +
                                      cfg.StopAfter+1, MailNumber+1), 1)
    else:
        # 从LastMailNo到最近邮件MailNumber
        rge = range(LastMailNo+1, MailNumber+1, 1)
    rge = list(rge)
    if rge is None:
        rge = []
    # TODO redo完成后及时清理
    rge = rge+cfg.RedoList+cfg.RedoListLocal
    rge = list(set(rge))
    dealmsg = "Redo:" + str(cfg.RedoList)
    dealmsg += "Redo Local:" + str(cfg.RedoListLocal)
    if MailNumber > LastMailNo:
        dealmsg += " and Mail No." + \
            str(LastMailNo+1)+" to No."+str(MailNumber)
    Debug_Print("Begin deal with:" + dealmsg, 2)
    if not rge:
        Debug_Print("Nothing to do. Quit.", 1)
        Pop3.quit()
        return
    for MailNo in rge:
        ENV_INFO = "Mail No."+str(MailNo)
        # ii = 1 #文件名非法时用的临时序号
        # 解析信件内容
        ForceGetFromServer = (MailNo in cfg.RedoList)
        try:
            Msg = GetMsg(Pop3, cfg.EmlFolder, MailNo, ForceGetFromServer)[0]
        except (poplib.error_proto, socket.error) as e:
            ErrorLog("Connection error:"+str(e))
            return
        # except Exception as e:
        #     ErrorLog("retr mail error:"+str(e))
        #     continue
        # 保存邮件本地缓存
        SaveEmlFile(Msg, cfg.EmlFolder, MailNo, False)  # todo:配置Rewrite

        # 保存邮件体
        MailTxtName = cfg.LocalMailPath + str(MailNo) + "-mail.txt"
        ftxt = codecs.open(MailTxtName, 'wb', encoding='utf-8')
        ftxt.write('[' + '#'*26 + ' Mail No.' + str(MailNo) + '#'*26 + ']\n')
        # 开始遍历全部MSG
        Msg2TxtFile(Msg, MailNo, ftxt, cfg.LocalMailPath)
        # print '\n',
        ftxt.write('[' + '#'*23 + ' End of Mail No.' +
                   str(MailNo) + '#'*23 + ']\n')
        ftxt.close()

        # 根据邮件主题等重命名一次
        NewTxtName = str(MailNo) + "M-[" + MyMailDecode(
            Msg.get("Subject")) + "]-[" + MyMailDecode(Msg.get("Date")) + "].txt"
        Debug_Print(NewTxtName, 4)
        NewTxtName = cfg.LocalMailPath + \
            re.sub(r"[\\\/\\\:\*\?\"\<\>\|\t\r\n]", '', NewTxtName)
        if os.path.exists(NewTxtName):
            os.remove(NewTxtName)
        os.rename(MailTxtName, NewTxtName)
        #     ErrorLog("Rename Mail text file error:")
        #     ErrorLog("from["+MailTxtName+"]to["+NewTxtName+"]")
        #     ErrorLog(str(e))

        Debug_Print("All OK, write lastmailno="+str(MailNo), 2)
        config = ReadCfg().ConfigObj
        config.set('WorkState', 'MailCount', str(cfg.MailCount))
        if MailNo in cfg.RedoList:
            cfg.RedoList.remove(MailNo)
            config.set("WorkState", "RedoList", ",".join(str(x)
                                                         for x in cfg.RedoList))
        elif MailNo in cfg.RedoListLocal:
            cfg.RedoListLocal.remove(MailNo)
            config.set("WorkState", "RedoListLocal", ",".join(str(x)
                                                              for x in cfg.RedoListLocal))
        else:
            config.set("WorkState", "LastMailNo", str(MailNo))
        config.write(open(ConfigFileName, "w"))

    # 全部处理完成，退出
    Pop3.quit()
    # 写之前再读一次配置，避免覆盖
    config = ReadCfg().ConfigObj
    # if EveryThingOK:
    #     config.remove_option("Mail", "SomethingWrongLastTime")
    # else:
    #     config.set("Mail", "SomethingWrongLastTime", str(MailNo))
    config.set("WorkState", "LastTime", time.strftime(
        "%Y-%m-%d %H:%M:%S", time.localtime(time.time())))
    config.write(open(ConfigFileName, "w"))
    # raw_input("Press enter to exit > ")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        ConfigFileName = sys.argv[1]
        ErrorLogName = ConfigFileName + "-err.log"
    # ErrorLog(str(sys.argv))
    while True:
        GetMail()
        Debug_Print("Sleep "+str(BREAK_A_SECOND) + " secs.", 1)
        time.sleep(BREAK_A_SECOND)
        Debug_Print("Reconnecting.", 1)
    # r=range(1,10,2)
    # r.stop=100
    # r=list(r)
    # print(r.__class__.__name__)
    # print(r)
