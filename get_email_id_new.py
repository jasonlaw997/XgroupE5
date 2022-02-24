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
    # print(input_win_title)
    if input_win_title == "微信" or input_win_title =="PingID":
        if key == "v":
            win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
            win32api.keybd_event(86, 0, 0, 0)  # v键位码是86
            win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
            win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    sleep(0.5)
    win32api.keybd_event(13, 0, 0, 0)  # enter
    win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    # sleep(1)
    # caps_status=win32api.GetKeyState(win32con.VK_CAPITAL)
    # if caps_status == 1:
    #     win32api.keybd_event(20, 0, 0, 0)  # caps lock
    #     win32api.keybd_event(20, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键

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


def print_info(indent=0):
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
            print('%s%s: %s' % ('  ' * indent, header, value))
            if header == "Subject":       #想剪切板发送文本
                print(value)
                clip.copy(value)
                sleep(0.6)
                Ctrl_X("v")

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
    '<ctrl>+<alt>': print_info
    }) as shortcut_key_listener:
     shortcut_key_listener.join()

if __name__=='__main__':
    funs_combkey_listener()