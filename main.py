import ctypes
import difflib
import os
from typing import List, Union

import easyocr
import numpy
import win32con
import win32gui
import win32ui
from loguru import logger

from question_and_anwser import tiku

possibilities = tiku.keys()

workpath = os.getcwd()
model_storage_directory = os.path.join(workpath, 'model')
reader = easyocr.Reader(['ch_sim'], model_storage_directory=model_storage_directory)


def get_text(imsrc):
    result = reader.readtext(imsrc, detail=0)
    return result


def get_question(word, n=2, cutoff=0.8):
    return difflib.get_close_matches(word, possibilities, n=n, cutoff=cutoff)


class Window:
    def __init__(self, hwnd: Union[int, None], name: Union[str, None] = None, cls=None):
        if hwnd is not None:
            self.hwnd = hwnd
        else:
            assert name
            self.hwnd = win32gui.FindWindow(cls, name)
        assert self.hwnd
        logger.info('find hwnd {}'.format(hwnd))
        self.hWndDC = win32gui.GetDC(self.hwnd)
        self.hMfcDc = win32ui.CreateDCFromHandle(self.hWndDC)
        self.hMemDc = self.hMfcDc.CreateCompatibleDC()
        self.y_slice = slice(41, 140)
        self.x_slice = slice(742, 1080)
        self.resize()

    @staticmethod
    def get_all_windows(label: str) -> List[int]:
        def callback(hwnd, hwnds):
            if win32gui.GetWindowText(hwnd) == label:
                hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds

    @staticmethod
    def is_admin() -> bool:
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def reloadimg(self):
        self.width, self.height = win32gui.GetClientRect(self.hwnd)[2:]
        hBmp = win32ui.CreateBitmap()
        hBmp.CreateCompatibleBitmap(self.hMfcDc, self.width, self.height)
        self.hMemDc.SelectObject(hBmp)
        self.hMemDc.BitBlt((0, 0), (self.width, self.height), self.hMfcDc, (0, 0), win32con.SRCCOPY)
        result = numpy.frombuffer(hBmp.GetBitmapBits(True), dtype=numpy.uint8).reshape(self.height, self.width, 4)
        win32gui.DeleteObject(hBmp.GetHandle())
        self.imsrc = result

    def __del__(self):
        self.hMemDc.DeleteDC()
        self.hMfcDc.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, self.hWndDC)

    def resize(self, wight=1152, height=679):
        self.width, self.height = win32gui.GetClientRect(self.hwnd)[2:]
        self.x_slice = slice(int(self.width * 742 / 1136), int(self.width * 1080 / 1136))
        self.y_slice = slice(int(self.height * 41 / 640), int(self.width * 140 / 640))

    def run(self):
        self.reloadimg()
        need = self.imsrc[(self.y_slice, self.x_slice)]
        result = get_text(need)
        word = ''.join(result)
        logger.info("识别:{}".format(word))
        questions = get_question(word)
        for q in questions:
            logger.info("问题:{}".format(q))
            logger.info("答案:{}".format(tiku[q]))


def test():
    import cv2

    input_file = 'onmyoji_GJ3G4ZS3pU.png'
    imsrc = cv2.imread(input_file)
    y_slice = slice(41, 140)
    x_slice = slice(742, 1080)
    need = imsrc[(y_slice, x_slice)]
    result = get_text(need)
    word = ''.join(result)
    logger.info("识别:{}".format(word))
    questions = get_question(word, cutoff=0.8)
    for q in questions:
        logger.info("问题:{}".format(q))
        logger.info("答案:{}".format(tiku[q]))


if __name__ == '__main__':
    if not Window.is_admin():
        import sys

        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    else:
        import time

        label = '阴阳师-网易游戏'
        wins = Window.get_all_windows(label)
        if wins:
            game = Window(wins[0])
            while True:
                game.run()
                time.sleep(1.5)
        else:
            logger.info("no game found!")