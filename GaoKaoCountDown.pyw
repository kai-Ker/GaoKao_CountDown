import os
import time
from threading import Timer

import psutil
import pywintypes
import win32api
import win32con
import win32gui
import win32ui
from dateutil import parser

__author__ = 'kai_Ker (kai_Ker@buaa.edu.cn)'

FPS = 75


class TimerRunner(Timer):
    def run(self):
        while not self.finished.is_set():
            self.function(*self.args, **self.kwargs)
            self.finished.wait(self.interval)


class TimerApplier:
    def __init__(self, interval, functionName, *args, **kwargs):
        self.timer = TimerRunner(interval, functionName, *args, **kwargs)

    def timer_start(self):
        self.timer.start()

    def timer_cancel(self):
        self.timer.cancel()


def loger(text, type='info'):
    with open('./log', 'a') as f:
        f.write(f'{time.strftime("%Y-%m-%d %H:%M:%S")}|{type}|{text}\n')


def getScreenBitmap(hScrDC, rect: tuple):
    hMemDC = win32gui.CreateCompatibleDC(hScrDC)
    hBitmap = win32gui.CreateCompatibleBitmap(
        hScrDC, rect[2], rect[3]-rect[1])
    hOldBitmap = win32gui.SelectObject(hMemDC, hBitmap)

    win32gui.BitBlt(hMemDC, rect[0], rect[1], rect[2]-rect[0],
                    rect[3]-rect[1], hScrDC, rect[0], rect[1], win32con.SRCCOPY)
    win32gui.SelectObject(hMemDC, hOldBitmap)

    win32gui.DeleteDC(hMemDC)
    return hBitmap


def getDeskHandle():
    hWndDesktop = 0
    hProgMan = win32gui.FindWindow("ProgMan", None)
    if hProgMan:
        hShellDefView = win32gui.FindWindowEx(
            hProgMan, None, 'SHELLDLL_DefView', None)
        if hShellDefView:
            hWndDesktop = win32gui.FindWindowEx(
                hShellDefView, None, 'SysListView32', None)
    return hWndDesktop


def drawText(hWndDesktop, hScrBitmap, text='', hFont=None, color=(255, 75, 75)):
    # get text
    text = getTargetString() if text == '' else text

    # get hDC
    # win32gui.InvalidateRect(hWndDesktop, rect, True)
    # win32gui.UpdateWindow(hWndDesktop)
    hDeskDC = win32gui.GetWindowDC(hWndDesktop)

    # elimate the twinkle
    hMemDC = win32gui.CreateCompatibleDC(hDeskDC)  # memery hdc
    hBitDC = win32gui.CreateCompatibleDC(hDeskDC)
    hBitmap = win32gui.CreateCompatibleBitmap(
        hDeskDC, rect[2], rect[3])  # bitmap hdc
    hOldMemSel = win32gui.SelectObject(hMemDC, hBitmap)
    hOldScrSel = win32gui.SelectObject(hBitDC, hScrBitmap)

    win32gui.SetBkMode(hMemDC, win32con.TRANSPARENT)
    win32gui.SetBkMode(hBitDC, win32con.TRANSPARENT)
    win32gui.SelectObject(hMemDC, hFont) if not hFont == None else ...
    oldColor = win32gui.SetTextColor(hMemDC, win32api.RGB(*color))

    win32gui.BitBlt(hMemDC, rect[0], rect[1], rect[2]-rect[0],
                    rect[3]-rect[1], hBitDC, rect[0], rect[1], win32con.SRCCOPY)
    win32gui.DrawText(hMemDC, text, len(text), rect, win32con.DT_CENTER)

    win32gui.BitBlt(hDeskDC, rect[0], rect[1], rect[2]-rect[0],
                    rect[3]-rect[1], hMemDC, rect[0], rect[1], win32con.SRCCOPY)

    # # text color, etc
    # win32gui.SelectObject(hMemDC, hFont) if not hFont == None else ...
    # win32gui.SetBkMode(hDeskDC, win32con.TRANSPARENT)
    # win32gui.SetTextColor(hDeskDC, win32api.RGB(*color))

    # # draw
    # win32gui.DrawText(hDeskDC, text, len(text), rect, win32con.DT_CENTER)

    win32gui.SetTextColor(hMemDC, oldColor)
    win32gui.SelectObject(hMemDC, hOldMemSel)
    win32gui.SelectObject(hBitDC, hOldScrSel)
    win32gui.DeleteObject(hBitmap)
    win32gui.DeleteDC(hMemDC)
    win32gui.DeleteDC(hBitDC)
    win32gui.ReleaseDC(hWndDesktop, hDeskDC)
    win32gui.UpdateWindow(hWndDesktop)


def getTargetString():
    nowTime = parser.parse(time.strftime('%Y-%m-%d %H:%M:%S'))
    tgtTime = parser.parse('2022-06-07 09:00:00')
    deltaTime = tgtTime - nowTime
    if deltaTime.total_seconds() <= 0:
        return ''
    dayStr = str(deltaTime.days)
    secs = deltaTime.seconds
    horStr = str(secs // 3600).zfill(2)
    minStr = str(secs % 3600 // 60).zfill(2)
    secStr = str(secs % 60).zfill(2)
    return f'距高考还有{dayStr}天{horStr}时{minStr}分{secStr}秒'


def getScale() -> float:
    scale = 1.0
    try:
        with open('config.cfg', 'r', encoding='utf-8') as f:
            line=f.readlines()[0].replace('\r', '').replace('\n', '')
            scale = float(line.split('=')[1])
    except:
        loger('cannot get the scale', type='warning')
        try:
            with open('config.cfg', 'w', encoding='utf-8') as f:
                f.writelines('scale=1.0')
            loger('rewrite the config file')
        except:
            loger('cannot rewrite the config file', type='error')
    return scale


def getProcName(pid: int) -> str:
    name = None
    for proc in psutil.process_iter():
        if proc.pid == pid:
            name = proc.name()
    return name


def killProc():
    try:
        with open('proc', 'r') as f:
            pid, name = f.read().split('|')
        proc = psutil.Process(pid)
        if proc.name().lower() == name.lower():
            proc.terminate()
    except:
        pass


if __name__ == '__main__':
    killProc()
    pid = os.getpid()
    procName = getProcName(os.getpid())
    loger(f'program started, pid: {pid}, name: {procName}', type='init')

    scale = getScale()
    loger(f'get the scale: {scale}')
    fontH = int(50*scale)
    fontW = int(24*scale)
    halfW = int(333*scale)

    hWndDesktop = getDeskHandle()
    loger(f'get the hWndDesktop: {hWndDesktop}')
    # hScrDC = win32gui.CreateDC('DISPLAY', None, None)
    hScrDC = win32gui.GetWindowDC(hWndDesktop)
    loger(f'get the hScrDC: {hScrDC}')

    # refresh the rect width
    w = int(win32ui.GetDeviceCaps(hScrDC, win32con.HORZRES))
    rect = ((w-2*halfW)//2, 0, w//2+halfW, fontH)
    loger(f'rect: {rect}')

    hScrBitmap = getScreenBitmap(hScrDC, rect)

    font = win32ui.CreateFont(
        {'height': fontH, 'width': fontW, 'weight': win32con.FW_BOLD})
    hFont = pywintypes.HANDLE(font.GetSafeHandle())
    loger(f'get the hFont: {hFont}')

    if not (hWndDesktop and hScrDC and hScrBitmap and hFont):
        loger(f'program pop out', type='error')
        exit(-1)

    t = TimerApplier(0.02, drawText, [
                     hWndDesktop, hScrBitmap], {'hFont': hFont})
    t.timer_start()
    loger('timer started')

    with open('proc', 'w') as f:
        f.writelines(f'{pid}|{procName}')
    # t.timer_cancel()
    loger('main program to quit\n--------------------')
