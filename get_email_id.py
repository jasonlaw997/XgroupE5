import poplib
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import pyperclip as clip
import win32api
import win32con
import win32gui
from time import sleep
from pynput import keyboard

def Ctrl_X(key):
    input_win_title = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    if input_win_title =="PingID" :
        win32api.keybd_event(9, 0, 0, 0)       #tab
        win32api.keybd_event(9, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
        sleep(0.3)
        if key == "v":
            win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
            win32api.keybd_event(86, 0, 0, 0)  # v键位码是86
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
            win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    elif input_win_title == "微信":
        if key == "v":
            win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
            win32api.keybd_event(86, 0, 0, 0)  # v键位码是86
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
            win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False

# 输入邮件地址, 口令和POP3服务器地址:
# email = input('Email: ')
# password = input('Password: ')
# pop3_server = input('POP3 server: ')

# 连接到POP3服务器:
server = poplib.POP3("pop.qq.com")
# 可以打开或关闭调试信息:
server.set_debuglevel(1)
# 可选:打印POP3服务器的欢迎文字:
# print(server.getwelcome().decode('utf-8'))

# 身份认证:
server.user("1025927107@qq.com")
server.pass_("eysuavnqpawobeeh")

# stat()返回邮件数量和占用空间:
# print('Messages: %s. Size: %s' % server.stat())
# list()返回所有邮件的编号:
resp, mails, octets = server.list()
# 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
# print(mails)

# 获取最新一封邮件, 注意索引号从1开始:
index = len(mails)
resp, lines, octets = server.retr(index)

# lines存储了邮件的原始文本的每一行,
# 可以获得整个邮件的原始文本:
msg_content = b'\r\n'.join(lines).decode('utf-8')
# 稍后解析出邮件:
msg = Parser().parsestr(msg_content)

# 可以根据邮件索引号直接从服务器删除邮件:
# server.dele(index)
# 关闭连接:
server.quit()

def print_info(msg, indent=0):
    if indent == 0:
        for header in ['From', 'To', 'Subject']:
            value = msg.get(header, '')
            if value:
                if header=='Subject':
                    value = decode_str(value)
                    if is_number(value):
                        clip.copy(value)

                else:
                    hdr, addr = parseaddr(value)
                    name = decode_str(hdr)
                    value = u'%s <%s>' % (name, addr)
            print(11111,'%s%s: %s' % ('  ' * indent, header, value))
            if header == "Form":       #想剪切板发送文本
                clip.copy(value)
                sleep(0.3)
                Ctrl_X("v")

    if (msg.is_multipart()):
        parts = msg.get_payload()
        for n, part in enumerate(parts):
            print(22,'%spart %s' % ('  ' * indent, n))
            print(3333333,'%s--------------------' % ('  ' * indent))
            print_info(part, indent + 1)
    else:
        content_type = msg.get_content_type()
        if content_type=='text/plain' or content_type=='text/html':
            content = msg.get_payload(decode=True)
            charset = guess_charset(msg)
            if charset:
                content = content.decode(charset)
            print(4444444,'%sText: %s' % ('  ' * indent, content + '...'))
        else:
            print(555555,'%sAttachment: %s' % ('  ' * indent, content_type))
def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def guess_charset(msg):
    charset = msg.get_charset()
    if charset is None:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset


def  funs_combkey_listener():
    global shortcut_key_listener
    with keyboard.GlobalHotKeys({
    '<ctrl>+<caps_lock>': print(222222)
    }) as shortcut_key_listener:
     shortcut_key_listener.join()

if __name__=='__main__':
    funs_combkey_listener()
    # print_info(msg)