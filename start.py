import cv2
import numpy as np
import win32gui as wg
import win32ui, win32con, win32api
import os
import keyboard
import pyautogui
import time
import configparser
import threading
from tkinter import *
from tkinter.font import *



class Settings:

    MarkerImage = 'resources/1.png'
    AnalysisMethod = 0
    Threshhold = 4.0
    BufferSize = 4
    PauseAfterClick = 1.5

    UiDraw = True
    UiDraw_ScanBox = True
    UiDraw_MarkerBox = True
    UiDraw_MarkerBox_Margin = 3
    UiKey_Pause = 'pause'
    UiKey_Throw = 'del'
    UiKey_MoveMouseToTarget = 'f12'
    UiKey_IncThreshhold = '+'
    UiKey_DecThreshhold = '-'
    UiKey_EXIT = 'space'

    @staticmethod
    def __init__():
        os.system('cls')
        print("Reading settings.ini...")
        cfg = configparser.ConfigParser()
        if cfg.read('settings.ini'):
            Settings.MarkerImage = cfg['DEFAULT']['MarkerImage']
            Settings.AnalysisMethod = int(cfg['DEFAULT']['AnalysisMethod'])
            Settings.Threshhold = float(cfg['DEFAULT']['Threshhold'])
            Settings.BufferSize = int(cfg['DEFAULT']['BufferSize'])
            Settings.PauseAfterClick = float(cfg['DEFAULT']['PauseAfterClick'])

            Settings.UiDraw = cfg['UserInterface'].getboolean('UiDraw')
            Settings.UiDraw_ScanBox = cfg['UserInterface'].getboolean('UiDraw_ScanBox')
            Settings.UiDraw_MarkerBox = cfg['UserInterface'].getboolean('UiDraw_MarkerBox')
            Settings.UiDraw_MarkerBox_Margin = int(cfg['UserInterface']['UiDraw_MarkerBox_Margin'])
            Settings.UiKey_Pause = cfg['UserInterface']['UiKey_Pause']
            Settings.UiKey_Throw = cfg['UserInterface']['UiKey_Throw']
            Settings.UiKey_MoveMouseToTarget = cfg['UserInterface']['UiKey_MoveMouseToTarget']
            Settings.UiKey_IncThreshhold = cfg['UserInterface']['UiKey_IncThreshhold']
            Settings.UiKey_DecThreshhold = cfg['UserInterface']['UiKey_DecThreshhold']
            Settings.UiKey_EXIT = cfg['UserInterface']['UiKey_EXIT']
            print('Reading success!')
            pass
        else:
            print("Settings file not found! Create new...")
            Settings.savecfg(cfg)

    @staticmethod
    def savecfg(config=None):

        cfg = config if config else configparser.ConfigParser()
        print("Settings saved!")
        cfg['DEFAULT'] = {
            'MarkerImage': Settings.MarkerImage,
            'AnalysisMethod': str(Settings.AnalysisMethod),
            'Threshhold': str(Settings.Threshhold),
            'BufferSize': str(Settings.BufferSize),
            'PauseAfterClick': str(Settings.PauseAfterClick),
        }
        cfg['UserInterface'] = {
            'UiDraw': str(Settings.UiDraw),
            'UiDraw_ScanBox': str(Settings.UiDraw_ScanBox),
            'UiDraw_MarkerBox': str(Settings.UiDraw_MarkerBox),
            'UiDraw_MarkerBox_Margin': str(Settings.UiDraw_MarkerBox_Margin),
            'UiKey_Pause': Settings.UiKey_Pause,
            'UiKey_Throw': Settings.UiKey_Throw,
            'UiKey_MoveMouseToTarget': Settings.UiKey_MoveMouseToTarget,
            'UiKey_IncThreshhold': Settings.UiKey_IncThreshhold,
            'UiKey_DecThreshhold': Settings.UiKey_DecThreshhold,
            'UiKey_EXIT': Settings.UiKey_EXIT
        }
        cfg['Explaining'] = {
            'AnalysisMethod': '0 -cv2.TM_CCOEFF, 1 -cv2.TM_CCORR_NORMED, 2 -cv2.TM_SQDIFF',
            'Threshhold': 'Vertical axis oscillations. Usually between 3.0 and 4.5',
            'UiKey_MoveMouseToTarget': 'For testing. Moves the mouse cursor to the found marker on the screen.',
            'UiDraw_MarkerBox_Margin': 'Margin of marker-image. Pixels integer. Default is 3'
        }
        with open('settings.ini', 'w') as configfile:
            cfg.write(configfile)

class Window:

    name: ""
    windows: []
    hwnd: None
    rect: [] # x rect[0] y rect[1] w rect[2] h rect[3]

    def __init__(self, name='1'):

        self.name = name
        self.windows = []
        self.rect = []
        wg.EnumWindows(self.__windowenumerationhandler, self.windows)

    def __windowenumerationhandler(self, hwnd, top_windows):
        top_windows.append((hwnd, wg.GetWindowText(hwnd)))


    def find(self):
        for i in self.windows:
            if self.name in i[1].lower():
                print(f'Целевое окно найдено! {i}')
                self.hwnd = i[0]
                self.rect = wg.GetWindowRect(self.hwnd)
                if self.rect[0] == 0:
                    break

    def focus(self):
        wg.ShowWindow(self.hwnd, 1)
        wg.SetForegroundWindow(self.hwnd)
        self.rect = wg.GetWindowRect(self.hwnd)


    def grab_screen(self, region=None):

        hwin = wg.GetDesktopWindow()

        if region:
            left,top,x2,y2 = region
            width = x2 #- left + 1
            height = y2 #- top + 1
        else:
            width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
            height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
            left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
            top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

        hwindc = wg.GetWindowDC(hwin)
        srcdc = win32ui.CreateDCFromHandle(hwindc)
        memdc = srcdc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(srcdc, width, height)
        memdc.SelectObject(bmp)
        memdc.BitBlt((0, 0), (width, height), srcdc, (left, top), win32con.SRCCOPY)

        signedIntsArray = bmp.GetBitmapBits(True)
        #img = np.fromstring(signedIntsArray, dtype='uint8')
        img = np.frombuffer(signedIntsArray, dtype='uint8')
        img.shape = (height,width,4)

        srcdc.DeleteDC()
        memdc.DeleteDC()
        wg.ReleaseDC(hwin, hwindc)
        wg.DeleteObject(bmp.GetHandle())

        return cv2.cvtColor(img, cv2.COLOR_BGRA2GRAY)

class Finder:
    __area: None
    method: str
    methods: []
    marker: None
    markercoord: []
    markercenter: []

    # method id's [ 0 - 2 ] #cv2.TM_CCOEFF_NORMED
    def __init__(self, marker=Settings.MarkerImage, method_id=Settings.AnalysisMethod):
        self.marker = cv2.imread(marker, 0)
        self.markercoord = [0, 0]
        self.markercoord[0], self.markercoord[1] = self.marker.shape;
        self.markercenter = [self.markercoord[0] / 2, self.markercoord[1] /2 ]
        self.methods = ['cv2.TM_CCOEFF', 'cv2.TM_CCORR_NORMED', 'cv2.TM_SQDIFF']
        self.method = self.methods[method_id]


    def frame_in(self, screenshot='resources/testarea.jpg'):
        self.__area = cv2.imread(screenshot, 3)

        buffered_frame = self.__area.copy()
        method = eval(self.method)

        res = cv2.matchTemplate(buffered_frame,self.marker, method)

        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

        if method is cv2.TM_SQDIFF:
            top_left = min_loc
        else:
            top_left = max_loc

        bottom_right = (top_left[0] + self.markercoord[0], top_left[1] + self.markercoord[1])

        return (top_left, bottom_right)

    def frame_in(self, frame):
        self.__area = frame

        buffered_frame = self.__area.copy()
        method = eval(self.method)

        res = cv2.matchTemplate(buffered_frame,self.marker, method)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

        if method is cv2.TM_SQDIFF:
            top_left = min_loc
        else:
            top_left = max_loc

        bottom_right = (top_left[0] + self.markercoord[0], top_left[1] + self.markercoord[1])

        return (top_left, bottom_right)

    def get_area(self):
        return self.__area.copy()

class TkOverlay(threading.Thread):

    transparent_color = '#%02x%02x%02x' % (255, 0, 128)
    draw_color = '#%02x%02x%02x' % (82, 227, 149)
    Window_coord: '0x0+0+0'
    scan_area_rect = [0, 0, 0, 0]
    scan_box_rect = [0, 0, 0, 0]
    __timer_started = False
    __queue:[]
    __Exit = False

    def __init__(self, rect=[0, 0, 0, 0] ):
        threading.Thread.__init__(self)
        self.daemon = True
        x, y, w, h = rect
        self.Window_coord = str(w) + 'x' + str(h) + '+' + str(x) + '+' + str(y)

    def Exit(self):
        self.__Exit = True

    def run(self):

        self.root = Tk()

        self.root.title('Solitaire')
        self.root.geometry(self.Window_coord)
        self.root.configure(bg=self.transparent_color)
        self.root.overrideredirect(True)
        self.root.lift()
        self.root.wm_attributes('-topmost', True)
        self.root.wm_attributes("-disabled", True)
        self.root.wm_attributes('-transparentcolor', self.transparent_color )

        self.root.iconbitmap('gameico.ico')

        self.scan_area = Frame(self.root,
                               bg = self.transparent_color,
                               highlightbackground = self.draw_color,
                               highlightcolor = self.draw_color,
                               highlightthickness=2,
                               bd=0,
                               height = self.scan_area_rect[3],
                               width = self.scan_area_rect[2]
                               )
        self.scan_box = Frame(self.root,
                              bg = self.transparent_color,
                              highlightbackground = self.draw_color,
                              highlightcolor = self.draw_color,
                              highlightthickness=2,
                              bd=0,
                              height = self.scan_box_rect[3],
                              width = self.scan_box_rect[2]
                              )
        tfont = Font(family = 'fixedsys', size='16')

        self.label = Label(self.root,
                           text='',
                           bg = self.transparent_color,
                           fg='light green',
                           font=tfont)

        Controls.text_start = True
        while (not self.__Exit):

            if Settings.UiDraw_ScanBox:
                self.DrawScanArea()
            if Settings.UiDraw_MarkerBox:
                self.DrawScanBox()

            if self.timer(2):
                self.DrawText()
            else:
                self.DrawText(clear=True)

            self.root.update()

        self.root.destroy()

    def DrawScanArea(self):
        self.scan_area.place(x = self.scan_area_rect[0], y = self.scan_area_rect[1], bordermode = 'outside')
        pass

    def DrawScanBox(self):
        self.scan_box.configure(height = self.scan_box_rect[3], width = self.scan_box_rect[2])
        self.scan_box.place(x = self.scan_box_rect[0], y = self.scan_box_rect[1], bordermode = 'outside')
        pass

    def DrawText(self, clear = False):
        text = f'{Settings.Threshhold:.1f}' if not clear else ''
        self.label.configure(text=text)
        cursor = pyautogui.position()
        self.label.place(x = cursor[0], y = cursor[1]-20)

    def timer(self, sec=1 ):
        if Controls.text_start:
            if not self.__timer_started:
                self.__start = time.time()
                time.process_time()
                self.__timer_started = True
                return True
            else:
                if time.time() - self.__start > sec:
                    self.__timer_started = False
                    Controls.text_start = False
                    return False
                else:
                    return True

class Filter:

    __queue=[]

    def __init__(self):
        self.__queue = [0] * Settings.BufferSize

    def pushToBuffer(self, y: int, sens=2.5)->bool:
        self.__queue.insert(0, y)
        self.__queue.pop()
        sum = 0
        for p in self.__queue:
            sum = sum + p
        diff = abs(sum/self.__queue.__len__() - y)

        return True if diff > sens else False

class Controls:
    Pause = True
    Display = None
    writingtext = False
    EXIT = False
    MouseMoveX = 0
    MouseMoveY = 0
    text_start = False
    text_show_start = 0
    text_show_elaps = 99

    @staticmethod
    def __init__():
        keyboard.on_release(Controls.on_release)

    @staticmethod
    def on_release(key):

        #print(key)
        if Settings.UiKey_EXIT in key.name:
            Controls.EXIT = True

        elif Settings.UiKey_Pause in key.name:
            Controls.Pause = not Controls.Pause

        elif Settings.UiKey_IncThreshhold in key.name:
            if Settings.Threshhold < 50.0:
                Settings.Threshhold = Settings.Threshhold + 0.1

            lock = threading.Lock()
            lock.acquire()
            Controls.text_start = True
            lock.release()

        elif Settings.UiKey_DecThreshhold in key.name:
            if Settings.Threshhold > 1.0:
                Settings.Threshhold = Settings.Threshhold - 0.1

            lock = threading.Lock()
            lock.acquire()
            Controls.text_start = True
            lock.release()

        elif Settings.UiKey_MoveMouseToTarget in key.name:
            pyautogui.moveTo(Controls.mousemoveX,
                             Controls.mousemoveY, 0)
            pyautogui.click()

if __name__ == '__main__':

    Settings()
    Controls()

    try:
        targetWindow = Window()
        targetWindow.find()
        targetWindow.focus()
    except:
        print('Target window not found!')

    RelativeScanRect = [int(targetWindow.rect[2]/3),
                        int(targetWindow.rect[3]/2.5 -30),
                        int(targetWindow.rect[2]/3),
                        int(targetWindow.rect[3]/2.5 - 150)
                        ]

    if Settings.UiDraw:
        GUI = TkOverlay(targetWindow.rect)
        GUI.start()
        GUI.scan_area_rect = RelativeScanRect

    print(targetWindow.rect)

    ScanRect = [int(targetWindow.rect[0]+targetWindow.rect[2]/3 ),
                int(targetWindow.rect[1]+targetWindow.rect[3]/2.5 -30),
                int(targetWindow.rect[2]/3),
                int(targetWindow.rect[3]/2.5 - 150)
                ]

    FIN = Finder(Settings.MarkerImage, Settings.AnalysisMethod)

    start = time.time()
    time.process_time()
    elapsed = 0
    lock = threading.Lock()
    f = Filter()
    while True:
        elapsed = time.time() - start

        if not Controls.Pause:
            frame = targetWindow.grab_screen(ScanRect)
            top_left, bottom_right = FIN.frame_in(frame)

            if f.pushToBuffer(top_left[0], Settings.Threshhold) and elapsed > 3:
                lastpos = pyautogui.position()
                pyautogui.moveTo(ScanRect[0]+top_left[0]-3+FIN.markercenter[0],
                                 ScanRect[1]+top_left[1]-3+FIN.markercenter[1], 0)
                pyautogui.click(pause=Settings.PauseAfterClick)
                pyautogui.moveTo(lastpos)
                start = time.time()

                pyautogui.press(Settings.UiKey_Throw)

            if Settings.UiDraw and Settings.UiDraw_MarkerBox:
                lock.acquire()
                GUI.scan_box_rect = [
                                     RelativeScanRect[0] + top_left[0] - Settings.UiDraw_MarkerBox_Margin - 5,
                                     RelativeScanRect[1] + top_left[1] - Settings.UiDraw_MarkerBox_Margin - 5,
                                     FIN.markercoord[0] + Settings.UiDraw_MarkerBox_Margin,
                                     FIN.markercoord[1] + Settings.UiDraw_MarkerBox_Margin
                                     ]
                lock.release()

            Controls.mousemoveX = ScanRect[0]+top_left[0]-3+FIN.markercenter[0]
            Controls.mousemoveY = ScanRect[1]+top_left[1]-3+FIN.markercenter[1]

        if Controls.EXIT:
            break

    GUI.Exit()
    Settings.savecfg()



    pass
