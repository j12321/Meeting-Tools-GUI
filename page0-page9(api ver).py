import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
import tkinter.messagebox as tkmessage
import pygsheets
import pandas as pd
import seaborn as sns
import calendar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
plt.rcParams['font.sans-serif'] = ['GenSenRounded TW']

gc = pygsheets.authorize(service_file='python final project-8cd68ae55bf9.json')

colors = ['red4', 'firebrick4', 'firebrick3', 'red2', 'red', 'firebrick1', 'OrangeRed2', 'tomato2', 'tomato',
          'chocolate1', 'dark orange', 'orange', 'goldenrod1', 'gold', 'yellow', 'DarkOliveGreen1', 'OliveDrab1',
          'green yellow', 'lawn green', 'chartreuse2', 'lime green', 'green3', 'SpringGreen3', 'SeaGreen3',
          'medium sea green', 'springGreen4', 'sea green', 'forest green', 'green4', 'dark green']

root = tk.Tk()


class Page0:

    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        self.root = master
        self.page0 = tk.Frame(self.root, width=1000, height=700)
        self.page0.configure(bg=self._from_rgb((68, 84, 106)))
        self.page0.master.title("開會小助手")
        self.page0.grid()

        f1 = tkfont.Font(size=45, family="源泉圓體 B")
        f2 = tkfont.Font(size=20, family="源泉圓體 M")

        self.lblCaption = tk.Label(self.page0, text='開 會 小 助 手', bg=self._from_rgb((68, 84, 106)), fg='white',
                                   font=f1, width=30, height=2, anchor='center')
        self.btnStart = tk.Button(self.page0, text='開始使用', bg=self._from_rgb((255, 217, 102)), font=f2, width=10,
                                  height=1, command=self.click_btnStart)

        self.lblCaption.place(relx=0.5, rely=0.5, anchor='center')
        self.btnStart.place(relx=0.5, y=450, anchor='center')

    def click_btnStart(self):
        self.page0.destroy()
        Page1()


class Page1:

    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        self.root = master
        self.page1 = tk.Frame(self.root, width=1000, height=700, bg=self._from_rgb((208, 224, 227)))
        self.page1.master.title("會議")
        self.page1.grid()

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=20, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘
        color_4 = self._from_rgb((255, 230, 153))  # 淡黃

        # 加scrollbar
        self.canvas1 = tk.Canvas(self.page1, width=1000, height=700, bg=color_2)
        self.canvas1.place(x=0, y=0)
        self.slb1 = tk.Scrollbar(self.page1, orient='vertical')
        self.slb1.place(x=980, width=20, height=700)
        self.canvas1.configure(yscrollcommand=self.slb1.set)
        self.slb1.configure(command=self.canvas1.yview)
        self.frame_context1 = tk.Frame(self.canvas1, width=1000, height=10000, bg=color_2)
        self.canvas1.create_window((-2, -2), window=self.frame_context1, anchor='nw')

        self.canvas_height_p1 = 200

        self.lblTitle_A = tk.Label(self.frame_context1, text=" 會議", height=1, width=15, font=f1,
                                   bg=color_1, fg='white', anchor='w')
        self.btnCreate_New = tk.Button(self.frame_context1, text="創建新會議", height=1, width=10, font=f2,
                                       bg=color_3, fg='black', command=self.click_btnCreate_New)
        self.btnCreate_folder = tk.Button(self.frame_context1, text="創建會議群組", height=1, width=12, font=f2,
                                          bg=color_3, fg='black', command=self.click_btnCreate_folder)
        self.lblSearch = tk.Label(self.frame_context1, text="關鍵字：", font=f3, bg=color_2)
        self.btnSearch = tk.Button(self.frame_context1, text='搜尋', command=self.click_btnSearch, height=1, width=3,
                                   font=f3, bg=color_4)

        global keywords
        keywords = tk.StringVar()
        self.inputKey = tk.Entry(self.frame_context1, textvariable=keywords, width=22, font=f3)

        self.lblTitle_A.place(x=0, y=50)
        self.btnCreate_New.place(x=780, y=50)
        self.btnCreate_folder.place(x=550, y=50)
        self.lblSearch.place(x=550, y=126)
        self.btnSearch.place(x=904, y=120)
        self.inputKey.place(x=640, y=126)

        global wb_names, sheet_names, df_sheet_names
        wb_names = gc.open('會議名稱')
        sheet_names = wb_names.worksheet_by_title('會議')

        df_sheet_names = sheet_names.get_as_df(has_header=False, include_tailing_empty=False)
        df_sheet_names.rename(columns={0: 'names', 1: 'status', 2: 'kind'}, inplace=True)

        global meeting_names, finish_meeting, meeting_or_folder
        try:
            meeting_names = df_sheet_names['names'].tolist()
            finish_meeting = df_sheet_names['status'].tolist()
            meeting_or_folder = df_sheet_names['kind'].tolist()
        except KeyError:
            df_sheet_names = pd.DataFrame({'names': [], 'status': []})
            meeting_names = []
            finish_meeting = []
            meeting_or_folder = []

        yaxis = 180
        k = 0
        m = 0
        self.pixel = tk.PhotoImage(height=2, width=10)

        for i in range(len(meeting_names)):
            if meeting_or_folder[i] == 'folder':
                length = len(meeting_names[i])
                font = tkfont.Font(size=34 - length, family="源泉圓體 B")
                self.btn_names = tk.Button(self.frame_context1, text=meeting_names[i], image=self.pixel, relief='solid',
                                           font=font, height=120, width=252, compound="center", wraplength=200,
                                           justify="left", bg='light grey',
                                           command=lambda b=i: self.click_btn_folder(b))
                self.btn_names.place(x=44 + 325 * (k % 3), y=180 + 150 * (k // 3))
                yaxis = 180 + 150 * (k // 3 + 1)

                if k % 3 == 0:
                    self.canvas_height_p1 += 150

                if finish_meeting[i] == 'finished':
                    self.btn_names.config(fg='light grey')

                k += 1

        for i in range(len(meeting_names)):
            if meeting_or_folder[i] == 'meeting':
                length = len(meeting_names[i])
                font = tkfont.Font(size=34 - length, family="源泉圓體 B")
                self.btn_names = tk.Button(self.frame_context1, text=meeting_names[i], image=self.pixel, relief='solid',
                                           font=font, height=120, width=252, compound="center", wraplength=200,
                                           justify="left", bg='white', command=lambda a=i: self.click_btn_meetings(a))
                self.btn_names.place(x=44 + 325 * (m % 3), y=yaxis + 150 * (m // 3))

                if m % 3 == 0:
                    self.canvas_height_p1 += 150

                if finish_meeting[i] == 'finished':
                    self.btn_names.config(fg='light grey')

                m += 1

        if self.canvas_height_p1 > 700:
            self.canvas1.configure(scrollregion=(0, 0, 1000, self.canvas_height_p1))
        else:
            self.canvas1.configure(scrollregion=(0, 0, 1000, 700))

    def click_btnSearch(self):
        f2 = tkfont.Font(size=20, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_4 = self._from_rgb((255, 230, 153))  # 淡黃

        self.canvasS = tk.Canvas(self.page1, width=1000, height=700, bg=color_2)
        self.canvasS.place(relx=0, rely=0.15)
        self.slbS = tk.Scrollbar(self.page1, orient='vertical')
        self.slbS.place(relx=0.98, width=20, height=700)
        self.canvasS.configure(yscrollcommand=self.slbS.set)
        self.slbS.configure(command=self.canvasS.yview)
        self.frame_contextS = tk.Frame(self.canvasS, width=1000, height=10000, bg=color_2)
        self.canvasS.create_window((-2, -2), window=self.frame_contextS, anchor='nw')

        self.canvas_height_pS = 250

        global keywords, meeting_names, finish_meeting, meeting_or_folder
        fit_meetings = []
        fit_meetings_location = []

        for i in range(len(meeting_names)):
            if keywords.get() in meeting_names[i]:
                fit_meetings.append(meeting_names[i])
                fit_meetings_location.append(i)

        yaxis = 80
        k = 0
        m = 0
        self.pixel = tk.PhotoImage(height=2, width=10)

        for i in range(len(fit_meetings)):
            if meeting_or_folder[fit_meetings_location[i]] == 'folder':
                length = len(fit_meetings[i])
                font = tkfont.Font(size=34 - length, family="源泉圓體 B")
                self.btn_names = tk.Button(self.frame_contextS, text=fit_meetings[i], image=self.pixel, relief='solid',
                                           font=font, height=120, width=252, compound="center", wraplength=200,
                                           justify="left", bg='light grey',
                                           command=lambda b=fit_meetings_location[i]: self.click_btn_folder(b))
                self.btn_names.place(x=44 + 325 * (k % 3), y=80 + 150 * (k // 3))
                yaxis = 80 + 150 * (k // 3 + 1)

                if k % 3 == 0:
                    self.canvas_height_pS += 150

                if finish_meeting[fit_meetings_location[i]] == 'finished':
                    self.btn_names.config(fg='light grey')

                k += 1

        for i in range(len(fit_meetings)):
            if meeting_or_folder[fit_meetings_location[i]] == 'meeting':
                length = len(fit_meetings[i])
                font = tkfont.Font(size=34 - length, family="源泉圓體 B")
                self.btn_names = tk.Button(self.frame_contextS, text=fit_meetings[i], image=self.pixel, relief='solid',
                                           font=font, height=120, width=252, compound="center", wraplength=200,
                                           justify="left", bg='white',
                                           command=lambda a=fit_meetings_location[i]: self.click_btn_meetings(a))
                self.btn_names.place(x=44 + 325 * (m % 3), y=yaxis + 150 * (m // 3))

                if m % 3 == 0:
                    self.canvas_height_pS += 150

                if finish_meeting[fit_meetings_location[i]] == 'finished':
                    self.btn_names.config(fg='light grey')

                m += 1

        counts = len(fit_meetings)

        self.lblText = tk.Label(self.frame_contextS, text="符合\"" + keywords.get() + "\"的會議共有" + str(counts) + "個：",
                                font=f2, bg=color_2)
        self.btn_back = tk.Button(self.frame_contextS, text="返回", height=1, font=f3, command=self.click_btn_back,
                                  bg=color_4)

        self.lblText.place(x=50, y=20)
        self.btn_back.place(x=890, y=25)

        if self.canvas_height_pS > 700:
            self.canvasS.configure(scrollregion=(0, 0, 1000, self.canvas_height_pS))
        else:
            self.canvasS.configure(scrollregion=(0, 0, 1000, 700))

    def click_btn_back(self):
        self.page1.destroy()
        Page1()

    def click_btnCreate_New(self):
        self.create_window()

    def click_btnCreate_folder(self):
        f1 = tkfont.Font(size=20, family="源泉圓體 B")
        f2 = tkfont.Font(size=15, family="源泉圓體 M")
        f3 = tkfont.Font(size=10, family="源泉圓體 M")

        self.window = tk.Toplevel()
        self.window.geometry('600x420')
        self.window.resizable(0, 0)
        self.window.title('群組名稱')
        self.window.configure(bg=self._from_rgb((208, 224, 227)))

        self.lblTitle_B = tk.Label(self.window, text=" 創建會議群組", height=1, width=15, font=f1,
                                   bg=self._from_rgb((68, 84, 106)), fg='white', anchor='w')
        self.lblname = tk.Label(self.window, text="群組名稱：", bg=self._from_rgb((208, 224, 227)), height=1, width=10,
                                font=f2)
        self.btnYes = tk.Button(self.window, text="確認", height=1, width=5, bg=self._from_rgb((255, 217, 102)), font=f2,
                                command=self.click_btnYes_1)

        global folder_name

        folder_name = tk.StringVar()
        self.inputname = tk.Entry(self.window, textvariable=folder_name, width=30, font=f2)

        self.lblTitle_B.place(x=0, y=25)
        self.lblname.place(x=120, y=150)
        self.inputname.place(relx=0.5, y=200, anchor='center')
        self.btnYes.place(relx=0.5, y=380, anchor='center')

    def click_btnYes_1(self):
        if self.inputname.get() == "":
            self.window.lower(belowThis=None)
            tkmessage.showerror(title="輸入未完整", message="您尚未輸入群組名稱")
            self.window.wm_attributes('-topmost', 1)
        else:
            global meeting_names
            if self.inputname.get() in meeting_names:
                self.window.lower(belowThis=None)
                tkmessage.showerror(title="名稱錯誤", message="此名稱已存在")
                self.window.wm_attributes('-topmost', 1)
            else:
                self.window.destroy()
                wb_names.add_worksheet(folder_name.get())
                sheet_names.append_table([folder_name.get(), 'unfinished', 'folder'], dimension='ROWS', overwrite=False)

                self.page1.destroy()
                Page1()

    def create_window(self):
        global date_list
        date_list = []

        f1 = tkfont.Font(size=20, family="源泉圓體 B")
        f2 = tkfont.Font(size=15, family="源泉圓體 M")
        f3 = tkfont.Font(size=10, family="源泉圓體 M")

        self.window = tk.Toplevel()
        self.window.geometry('600x420')
        self.window.resizable(0, 0)
        self.window.title('會議日期')
        self.window.configure(bg=self._from_rgb((208, 224, 227)))

        self.lblTitle_B = tk.Label(self.window, text=" 創建新會議", height=1, width=15, font=f1,
                                   bg=self._from_rgb((68, 84, 106)), fg='white', anchor='w')
        self.lblname = tk.Label(self.window, text="會議名稱：", bg=self._from_rgb((208, 224, 227)), height=1, width=10,
                                font=f2)
        self.lblchoose = tk.Label(self.window, text="你已選擇：", bg=self._from_rgb((208, 224, 227)), height=1, width=10,
                                  font=f2)

        global meeting_name

        meeting_name = tk.StringVar()
        self.inputname = tk.Entry(self.window, textvariable=meeting_name, width=30, font=f2)

        self.enydate = tk.Listbox(self.window, height=7, width=15, font=f2, selectmode=tk.MULTIPLE)

        width = root.winfo_reqwidth() + 50
        height = 100  # 窗口大小
        x, y = (root.winfo_screenwidth() - width) / 2, (root.winfo_screenheight() - height) / 2

        self.btnYes = tk.Button(self.window, text="確認", height=1, width=5, bg=self._from_rgb((255, 217, 102)), font=f2,
                                command=self.click_btnYes)
        self.btn_delete = tk.Button(self.window, text="刪除日期", font=f3, command=self.click_btn_delete)
        self.scroll_dates = tk.Scrollbar(self.window, command=self.enydate.yview)

        self.lblTitle_B.place(x=0, y=25)
        self.lblname.place(x=70, y=75)
        self.lblchoose.place(x=70, y=105)
        self.enydate.place(x=80, y=140)
        self.inputname.place(x=190, y=75)
        self.btnYes.place(relx=0.5, y=380, anchor='center')
        self.btn_delete.place(x=130, y=320)
        self.scroll_dates.place(x=232, y=142, relheight=0.403)

        self.enydate.config(yscrollcommand=self.scroll_dates.set)

        datetime = calendar.datetime.datetime  # 日期和時間結合(從這邊複製)
        timedelta = calendar.datetime.timedelta  # 時間差

        class Calendar:
            def __init__(s, point=None, position=None):
                # point    窗口位置
                # position 窗口在點的位置 'ur'-右上, 'ul'-左上, 'll'-左下, 'lr'-右下
                fwday = calendar.SUNDAY
                year = datetime.now().year  # 為使打開頁面時為當下年份
                month = datetime.now().month  # 為使打開頁面時為當下月份
                locale = None
                sel_bg = '#ecffc4'  # 設定點擊日期後的框顏色
                sel_fg = '#05640e'  # 設定點擊日期後的字底色
                s._date = datetime(year, month, 1)  # 該月份第一天
                s._selection = None  # 設置未選中的日期
                s.G_Frame = ttk.Frame(self.window)
                s._cal = s.__get_calendar(locale, fwday)  # 實例化適當的日曆類
                s.__setup_styles()  # 創建自定義樣式
                s.__place_widgets()  # pack/grid 小部件
                s.__config_calendar()  # 調整日曆列和安裝標記
                # 配置畫布和正確的绑定，以選擇日期。
                s.__setup_selection(sel_bg, sel_fg)
                # 存儲項ID，用於稍後插入。
                s._items = [s._calendar.insert('', 'end', values='') for _ in range(6)]
                # 在當前空日曆中插入日期
                s._update()
                s.G_Frame.place(x=290, y=120)
                self.window.update_idletasks()  # 刷新頁面

                self.window.deiconify()  # 還原視窗
                self.window.focus_set()  # 焦點設置在所需小部件上
                self.window.wait_window()  # 直到按確定

            def __get_calendar(s, locale, fwday):  # 日曆文字化
                if locale is None:
                    return calendar.TextCalendar(fwday)
                else:
                    return calendar.LocaleTextCalendar(fwday, locale)

            def __setup_styles(s):  # 自定義TTK風格
                style = ttk.Style(self.window)
                arrow_layout = lambda dir: (
                    [('Button.focus', {'children': [('Button.%sarrow' % dir, None)]})])  # 返回參數性質
                style.layout('L.TButton', arrow_layout('left'))  # 製作點選上個月的箭頭
                style.layout('R.TButton', arrow_layout('right'))  # 製作點選下個月的箭頭

            def __place_widgets(s):  # 標題框架及其小部件
                Input_judgment_num = self.window.register(s.Input_judgment)  # 需要将函数包装一下，必要的
                hframe = ttk.Frame(s.G_Frame)
                gframe = ttk.Frame(s.G_Frame)
                bframe = ttk.Frame(s.G_Frame)
                hframe.pack(in_=s.G_Frame, side='top', pady=5, anchor='center')  # 月曆的上視窗
                gframe.pack(in_=s.G_Frame, fill=tk.X, pady=5)
                bframe.pack(in_=s.G_Frame, side='bottom', pady=5)
                lbtn = ttk.Button(hframe, style='L.TButton',
                                  command=s._prev_month)  # 月曆上方左箭頭，點選後月曆切換至前個月
                lbtn.grid(in_=hframe, column=0, row=0, padx=12)
                rbtn = ttk.Button(hframe, style='R.TButton',
                                  command=s._next_month)  # 月曆上方右箭頭，點選後月曆切換至下個月
                rbtn.grid(in_=hframe, column=5, row=0, padx=12)

                s.CB_year = ttk.Combobox(hframe, width=5, values=[str(year) for year in
                                                                  range(datetime.now().year, datetime.now().year + 11,
                                                                        1)], validate='key',
                                         validatecommand=(Input_judgment_num, '%P'))  # 製作月曆年份的下拉選單
                s.CB_year.current(0)  # 月曆一開始呈現當下年分
                s.CB_year.grid(in_=hframe, column=1, row=0)
                s.CB_year.bind("<<ComboboxSelected>>", s._update)  # 選取(年)下拉式選單後，月曆更新
                tk.Label(hframe, text='年', justify='left').grid(in_=hframe, column=2, row=0,
                                                                padx=(0, 5))  # 下拉式選單後面的單位(年)
                s.CB_month = ttk.Combobox(hframe, width=3, values=['%02d' % month for month in range(1, 13)],
                                          state='readonly')  # 製作月曆年份的下拉選單
                s.CB_month.current(datetime.now().month - 1)  # 月曆一開始呈現當下月份
                s.CB_month.grid(in_=hframe, column=3, row=0)
                s.CB_month.bind("<<ComboboxSelected>>", s._update)  # 選取(月)下拉式選單後，月曆更新
                tk.Label(hframe, text='月', justify='left').grid(in_=hframe, column=4, row=0)  # 下拉式選單後面的單位(年)
                # 日曆部件
                s._calendar = ttk.Treeview(gframe, show='', selectmode='none', height=7)  # 建立放上日期的頁面
                s._calendar.pack(expand=1, fill='both', side='bottom', padx=5)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=0, relwidth=1, relheigh=2 / 200)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=198 / 200, relwidth=1, relheigh=2 / 200)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=0, relwidth=2 / 200, relheigh=1)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=198 / 200, rely=0, relwidth=2 / 200, relheigh=1)

            def __config_calendar(s):  # 設計日曆架構
                cols = ['日', '一', '二', '三', '四', '五', '六']  # 日曆上的星期幾
                s._calendar['columns'] = cols  # 設定日曆欄
                s._calendar.tag_configure('header', background='grey90')
                s._calendar.insert('', 'end', values=cols, tag='header')  # 調整其列寬
                font = tkfont.Font()
                maxwidth = max(font.measure(col) for col in cols)
                for col in cols:
                    s._calendar.column(col, width=maxwidth, minwidth=maxwidth, anchor='center')

            def __setup_selection(s, sel_bg, sel_fg):
                def __canvas_forget(evt):
                    canvas.place_forget()
                    s._selection = None

                s._font = tkfont.Font()
                s._canvas = canvas = tk.Canvas(s._calendar, background=sel_bg, borderwidth=0, highlightthickness=0)
                canvas.text = canvas.create_text(0, 0, fill=sel_fg, anchor='w')

                s._calendar.bind('<Button-1>', s._pressed)  # 點出要的日期

            def _build_calendar(s):  # 建立月曆-日期
                year = s._date.year  # s._date = datetime(year, month, 1)， 因此設定變數year為操作程式當下年分
                month = s._date.month  # 設定變數month為操作程式當下月份
                header = s._cal.formatmonthname(year, month, 0)  # 更新日曆顯示的日期
                cal = s._cal.monthdayscalendar(year, month)  # 此函數會建立指定年月份的周列表，以供屆時放入treeview
                for indx, item in enumerate(s._items):  # s._items為在treeview中先在每格插入""，一排放7個(先建立月曆的架構)
                    week = cal[indx] if indx < len(cal) else []  # 因為一個月頂多四周多，因此若indx>cal長度，將week設為空list
                    fmt_week = [('%02d' % day) if day else '' for day in week]
                    s._calendar.item(item, values=fmt_week)  # 將該年月份的日期分配放置treeview

            def _show_select(s, text, bbox):  # 秀出挑選的日子
                x, y, width, height = bbox
                textw = s._font.measure(text)
                canvas = s._canvas
                canvas.configure(width=width, height=height)
                canvas.coords(canvas.text, (width - textw) / 2, height / 2 - 1)
                canvas.itemconfigure(canvas.text, text=text)
                canvas.place(in_=s._calendar, x=x, y=y)

            def _pressed(s, evt=None, item=None, column=None, widget=None, confirm=False):  # 在日曆的某個地方點擊。

                if not item:
                    x, y, widget = evt.x, evt.y, evt.widget
                    item = widget.identify_row(y)
                    column = widget.identify_column(x)
                if not column or item not in s._items:  # 在工作日行中單擊或僅在列外單擊。
                    return
                item_values = widget.item(item)['values']  # 點選的日期該周
                if not len(item_values):  # 該行是空的。
                    return
                text = item_values[int(column[1]) - 1]  # text 為選擇日期
                if not text:  # 日期為空
                    return
                bbox = widget.bbox(item, column)
                if not bbox:  # 日曆尚未出現
                    self.window.after(20, lambda: s._pressed(item=item, column=column, widget=widget))
                    return
                text = '%02d' % text
                if confirm:
                    pass
                else:
                    s._selection = (text, item, column)
                    year = s._date.year  # 使用者選取日期中的年份
                    month = s._date.month  # 使用者選取日期中的月份
                    choose_date = s._selection[0]  # 使用者選取日期中的"日"
                    today_year = datetime.now().year
                    today_month = datetime.now().month
                    today_day = datetime.now().day
                    date = str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0]))
                    if int(str(year)) > int(str(today_year)):
                        if len(date_list) != 0:
                            if date in date_list:  # 若選擇日期出現在date_list，代表重複日期，必須跳出錯誤訊息
                                self.window.lower(belowThis=self.page1)
                                tkmessage.showerror(title="日期重複", message="此日期已選擇")
                                self.window.wm_attributes('-topmost', 1)
                            else:
                                s._show_select(text, bbox)
                                date_list.append(str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0])))
                                self.enydate.insert("end", (str(year) + "/" + str(month) + "/" +
                                                            ("%02d" % int(s._selection[0]))))
                        else:
                            s._show_select(text, bbox)
                            date_list.append("1")
                    elif int(str(year)) == int(str(today_year)) and int(str(month)) > int(str(today_month)):
                        if len(date_list) != 0:
                            if date in date_list:
                                self.window.lower(belowThis=self.page1)
                                tkmessage.showerror(title="日期重複", message="此日期已選擇")
                                self.window.wm_attributes('-topmost', 1)
                            else:
                                s._show_select(text, bbox)
                                date_list.append(str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0])))
                                self.enydate.insert("end", (str(year) + "/" + str(month) + "/" +
                                                            ("%02d" % int(s._selection[0]))))
                        else:
                            s._show_select(text, bbox)
                            date_list.append("1")
                    elif int(str(year)) == int(str(today_year)) and int(str(month)) == int(str(today_month)) and int(
                            str(choose_date)) >= int(str(today_day)):
                        if len(date_list) != 0:
                            if date in date_list:
                                self.window.lower(belowThis=self.page1)
                                tkmessage.showerror(title="日期重複", message="此日期已選擇")
                                self.window.wm_attributes('-topmost', 1)
                            else:
                                s._show_select(text, bbox)
                                date_list.append(str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0])))
                                self.enydate.insert("end", (str(year) + "/" + str(month) + "/" +
                                                            ("%02d" % int(s._selection[0]))))
                        else:
                            s._show_select(text, bbox)
                            date_list.append("1")
                    else:  # 若選取的日期為今天之前的日期，跳出錯誤訊息
                        self.window.lower(belowThis=self.page1)
                        tkmessage.showerror(title="日期無效", message="請選擇有效日期")
                        self.window.wm_attributes('-topmost', 1)

            def _prev_month(s):  # 按下左箭頭後，月曆頁面需進行切換
                s._canvas.place_forget()
                s._selection = None
                s._date = s._date - timedelta(days=1)
                s._date = datetime(s._date.year, s._date.month, 1)
                s.CB_year.set(s._date.year)
                s.CB_month.set(s._date.month)
                s._update()

            def _next_month(s):  # 按下右箭頭後，月曆頁面需進行切換
                s._canvas.place_forget()
                s._selection = None
                year, month = s._date.year, s._date.month
                s._date = s._date + timedelta(
                    days=calendar.monthrange(year, month)[1] + 1)
                s._date = datetime(s._date.year, s._date.month, 1)
                s.CB_year.set(s._date.year)
                s.CB_month.set(s._date.month)
                s._update()

            def _update(s, event=None, key=None):
                if key and event.keysym != 'Return':
                    return

                year = int(s.CB_year.get())
                month = int(s.CB_month.get())

                if year == 0 or year > 9999:
                    return

                s._canvas.place_forget()
                s._date = datetime(year, month, 1)
                s._build_calendar()  # 重建日曆

                if year == datetime.now().year and month == datetime.now().month:
                    day = datetime.now().day
                    for _item, day_list in enumerate(s._cal.monthdayscalendar(year, month)):
                        if day in day_list:
                            item = 'I00' + str(_item + 2)
                            column = '#' + str(day_list.index(day) + 1)
                            self.window.after(100, lambda: s._pressed(item=item, column=column, widget=s._calendar))

            def _main_judge(s):
                try:
                    if self.window.focus_displayof() is None or 'toplevel' not in str(self.window.focus_displayof()):
                        s._pressed()
                    else:
                        self.window.after(10, s._main_judge)
                except:
                    self.window.after(10, s._main_judge)

            def selection(s):
                if not s._selection:
                    return None
                year = s._date.year
                month = s._date.month

                return str(datetime(year, month, int(s._selection[0])))[:10]

            def Input_judgment(s, content):
                if content.isdigit() or content == "":
                    return True
                else:
                    return False

        cal = Calendar()

    def click_btn_delete(self):
        choose = []
        choose_date = []
        choose_tuple = self.enydate.curselection()
        global date_list

        for i in choose_tuple:
            choose.append(i)
            choose_date.append(str(self.enydate.get(i)))

        for k in range(len(choose)):
            self.enydate.delete(choose[k] - k)

        for k in choose_date:
            date_list.remove(k)

    def click_btnYes(self):
        if self.enydate.size() != 0 and self.inputname.get() == "":
            self.window.lower(belowThis=None)
            tkmessage.showerror(title="輸入未完整", message="您尚未輸入會議名稱")
            self.window.wm_attributes('-topmost', 1)
        elif self.enydate.size() == 0 and self.inputname.get() != "":
            self.window.lower(belowThis=None)
            tkmessage.showerror(title="輸入未完整", message="您尚未輸入會議日期")
            self.window.wm_attributes('-topmost', 1)
        elif self.enydate.size() == 0 and self.inputname.get() == "":
            self.window.lower(belowThis=None)
            tkmessage.showerror(title="輸入未完整", message="您尚未輸入會議名稱及日期")
            self.window.wm_attributes('-topmost', 1)
        else:
            global meeting_names
            if self.inputname.get() in meeting_names:
                self.window.lower(belowThis=None)
                tkmessage.showerror(title="會議名稱錯誤", message="此會議名稱已存在")
                self.window.wm_attributes('-topmost', 1)
            else:
                self.window.destroy()
                name = meeting_name.get()
                gc.create(name)

                global wb_record, sheet_time
                wb_record = gc.open(name)
                sheet_time = wb_record[0]
                sheet_time.title = '時間統計'
                ws = wb_record.add_worksheet('出缺勤')
                wb_record.add_worksheet('Meeting record')

                wb_record.share('arial5623@gmail.com', role='writer', type='user')
                wb_record.share('alice0911496698@gmail.com', role='writer', type='user')
                wb_record.share('sunnyliu891114@gmail.com', role='writer', type='user')
                wb_record.share('jenny900707@gmail.com', role='writer', type='user')
                wb_record.share('teresa890101@gmail.com', role='writer', type='user')
                wb_record.share('l1849261495@gmail.com', role='writer', type='user')

                global date_list
                date_list.remove('1')
                date_list = sorted(date_list)

                columns = []
                rows = []

                for i in range(len(date_list)):
                    columns.append(date_list[i])

                for i in range(16):
                    rows.append(str(7 + i) + ':00-' + str(8 + i) + ':00')

                df = pd.DataFrame(columns=columns, index=rows)
                sheet_time.set_dataframe(df, start='A1', copy_index=True, copy_head=True, nan='')

                global sheet_names
                sheet_names.append_table([meeting_name.get(), 'unfinished', 'meeting'], dimension='ROWS',
                                         overwrite=False)

                ws.update_row(index=1, values=['name', 'absence', 'mission'], col_offset=0)

                self.page1.destroy()
                Page1()

    def click_btn_meetings(self, a):
        global wb_record, name, location, sheet_time, df_sheet_time, dates
        name = meeting_names[a]
        finish = finish_meeting[a]
        location = a
        wb_record = gc.open(str(name))
        sheet_time = wb_record[0]

        df_sheet_time = sheet_time.get_as_df(has_header=True, include_tailing_empty=False)
        df_sheet_time.drop(columns=[''], inplace=True)
        dates = df_sheet_time.columns.tolist()

        if finish == 'finished':
            self.page1.destroy()
            Page9()
        else:
            self.page1.destroy()
            Page3()

    def click_btn_folder(self, b):
        global folder_location
        folder_location = b

        self.page1.destroy()
        Page2()


class Page2:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘
        color_4 = self._from_rgb((255, 230, 153))  # 淡黃

        global meeting_names, folder_location
        self.root = master
        self.page2 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page2.master.title(str(meeting_names[folder_location]))
        self.page2.grid()

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=20, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        # 加scrollbar
        self.canvas1 = tk.Canvas(self.page2, width=1000, height=700, bg=color_2)
        self.canvas1.place(relx=0, rely=0)
        self.slb1 = tk.Scrollbar(self.page2, orient='vertical')
        self.slb1.place(relx=0.98, width=20, height=700)
        self.canvas1.configure(yscrollcommand=self.slb1.set)
        self.slb1.configure(command=self.canvas1.yview)
        self.frame_context1 = tk.Frame(self.canvas1, width=1000, height=10000, bg=color_2)
        self.canvas1.create_window((-2, -2), window=self.frame_context1, anchor='nw')

        self.canvas_height_p1 = 200

        self.lblTitle_A = tk.Label(self.frame_context1, text=" " + str(meeting_names[folder_location]), height=1,
                                   width=15, font=f1, bg=color_1, fg='white', anchor='w')
        self.btnCreate_New = tk.Button(self.frame_context1, text="創建新會議", height=1, width=10, font=f2,
                                       bg=color_3, fg='black', command=self.click_btnCreate_New)
        self.btn_back = tk.Button(self.page2, text="返回", height=1, font=f3, command=self.click_btn_back, bg=color_3)
        self.lblSearch = tk.Label(self.frame_context1, text="關鍵字：", font=f3, bg=color_2)
        self.btnSearch = tk.Button(self.frame_context1, text='搜尋', command=self.click_btnSearch, height=1, width=3,
                                   font=f3, bg=color_4)

        global keywords
        keywords = tk.StringVar()
        self.inputKey = tk.Entry(self.frame_context1, textvariable=keywords, width=22, font=f3)

        self.lblTitle_A.place(relx=0, rely=0.005, anchor='nw')
        self.btnCreate_New.place(relx=0.78, rely=0.005, anchor='nw')
        self.btn_back.place(x=700, y=65)
        self.lblSearch.place(x=550, y=126)
        self.btnSearch.place(x=904, y=120)
        self.inputKey.place(x=640, y=126)

        global wb_names, sheet_names, df_sheet_names, folder_meeting_names, folder_finish_meeting
        sheet_names = wb_names.worksheet_by_title(str(meeting_names[folder_location]))

        df_sheet_names = sheet_names.get_as_df(has_header=False, include_tailing_empty=False)
        df_sheet_names.rename(columns={0: 'names', 1: 'status'}, inplace=True)

        try:
            folder_meeting_names = df_sheet_names['names'].tolist()
            folder_finish_meeting = df_sheet_names['status'].tolist()
        except KeyError:
            df_sheet_names = pd.DataFrame({'names': [], 'status': []})
            folder_meeting_names = []
            folder_finish_meeting = []

        self.pixel = tk.PhotoImage(height=2, width=10)
        for i in range(len(folder_meeting_names)):
            length = len(folder_meeting_names[i])
            font = tkfont.Font(size=34 - length, family="源泉圓體 B")
            self.btn_names = tk.Button(self.frame_context1, text=folder_meeting_names[i], image=self.pixel,
                                       relief='solid', font=font, height=120, width=252, compound="center",
                                       wraplength=200, justify="left", bg='white',
                                       command=lambda a=i: self.click_btn_meetings(a))
            self.btn_names.place(x=44 + 325 * (i % 3), y=180 + 150 * (i // 3))

            if i % 3 == 0:
                self.canvas_height_p1 += 150

            if folder_finish_meeting[i] == 'finished':
                self.btn_names.config(fg='light grey')

        if self.canvas_height_p1 > 700:
            self.canvas1.configure(scrollregion=(0, 0, 1000, self.canvas_height_p1))
        else:
            self.canvas1.configure(scrollregion=(0, 0, 1000, 700))

    def click_btnSearch(self):
        f2 = tkfont.Font(size=20, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_4 = self._from_rgb((255, 230, 153))  # 淡黃

        self.canvasS = tk.Canvas(self.page2, width=1000, height=700, bg=color_2)
        self.canvasS.place(relx=0, rely=0.15)
        self.slbS = tk.Scrollbar(self.page2, orient='vertical')
        self.slbS.place(relx=0.98, width=20, height=700)
        self.canvasS.configure(yscrollcommand=self.slbS.set)
        self.slbS.configure(command=self.canvasS.yview)
        self.frame_contextS = tk.Frame(self.canvasS, width=1000, height=10000, bg=color_2)
        self.canvasS.create_window((-2, -2), window=self.frame_contextS, anchor='nw')

        self.canvas_height_pS = 250

        global keywords, folder_meeting_names, folder_finish_meeting, meeting_or_folder
        fit_meetings = []
        fit_meetings_location = []

        for i in range(len(folder_meeting_names)):
            if keywords.get() in folder_meeting_names[i]:
                fit_meetings.append(folder_meeting_names[i])
                fit_meetings_location.append(i)

        self.pixel = tk.PhotoImage(height=2, width=10)

        for i in range(len(fit_meetings)):
            length = len(fit_meetings[i])
            font = tkfont.Font(size=34 - length, family="源泉圓體 B")
            self.btn_names = tk.Button(self.frame_contextS, text=fit_meetings[i], image=self.pixel, relief='solid',
                                       font=font, height=120, width=252, compound="center", wraplength=200,
                                       justify="left", bg='white',
                                       command=lambda a=fit_meetings_location[i]: self.click_btn_meetings(a))
            self.btn_names.place(x=44 + 325 * (i % 3), y=80 + 150 * (i // 3))

            if i % 3 == 0:
                self.canvas_height_pS += 150

            if folder_finish_meeting[fit_meetings_location[i]] == 'finished':
                self.btn_names.config(fg='light grey')

        counts = len(fit_meetings)

        self.lblText = tk.Label(self.frame_contextS, text="符合\"" + keywords.get() + "\"的會議共有" + str(counts) + "個：",
                                font=f2, bg=color_2)
        self.btn_back = tk.Button(self.frame_contextS, text="返回", height=1, font=f3, command=self.click_btn_back_1,
                                  bg=color_4)

        self.lblText.place(x=50, y=20)
        self.btn_back.place(x=890, y=25)

        if self.canvas_height_pS > 700:
            self.canvasS.configure(scrollregion=(0, 0, 1000, self.canvas_height_pS))
        else:
            self.canvasS.configure(scrollregion=(0, 0, 1000, 700))

    def click_btn_back_1(self):
        self.page2.destroy()
        Page2()

    def click_btn_back(self):
        self.page2.destroy()
        Page1()

    def click_btnCreate_New(self):
        self.create_window()

    def create_window(self):
        global date_list
        date_list = []

        f1 = tkfont.Font(size=20, family="源泉圓體 B")
        f2 = tkfont.Font(size=15, family="源泉圓體 M")
        f3 = tkfont.Font(size=10, family="源泉圓體 M")

        self.window = tk.Toplevel()
        self.window.geometry('600x420')
        self.window.resizable(0, 0)
        self.window.title('會議日期')
        self.window.configure(bg=self._from_rgb((208, 224, 227)))

        self.lblTitle_B = tk.Label(self.window, text=" 創建新會議", height=1, width=15, font=f1,
                                   bg=self._from_rgb((68, 84, 106)), fg='white', anchor='w')
        self.lblname = tk.Label(self.window, text="會議名稱：", bg=self._from_rgb((208, 224, 227)), height=1, width=10,
                                font=f2)
        self.lblchoose = tk.Label(self.window, text="你已選擇：", bg=self._from_rgb((208, 224, 227)), height=1, width=10,
                                  font=f2)

        global meeting_name

        meeting_name = tk.StringVar()
        self.inputname = tk.Entry(self.window, textvariable=meeting_name, width=30, font=f2)

        self.enydate = tk.Listbox(self.window, height=7, width=15, font=f2, selectmode=tk.MULTIPLE)

        width = root.winfo_reqwidth() + 50
        height = 100  # 窗口大小
        x, y = (root.winfo_screenwidth() - width) / 2, (root.winfo_screenheight() - height) / 2

        self.btnYes = tk.Button(self.window, text="確認", height=1, width=5, bg=self._from_rgb((255, 217, 102)), font=f2,
                                command=self.click_btnYes)
        self.btn_delete = tk.Button(self.window, text="刪除日期", font=f3, command=self.click_btn_delete)
        self.scroll_dates = tk.Scrollbar(self.window, command=self.enydate.yview)

        self.lblTitle_B.place(x=0, y=25)
        self.lblname.place(x=70, y=75)
        self.lblchoose.place(x=70, y=105)
        self.enydate.place(x=80, y=140)
        self.inputname.place(x=190, y=75)
        self.btnYes.place(relx=0.5, y=380, anchor='center')
        self.btn_delete.place(x=130, y=320)
        self.scroll_dates.place(x=232, y=142, relheight=0.403)

        self.enydate.config(yscrollcommand=self.scroll_dates.set)

        datetime = calendar.datetime.datetime  # 日期和時間結合(從這邊複製)
        timedelta = calendar.datetime.timedelta  # 時間差

        class Calendar:
            def __init__(s, point=None, position=None):
                # point    窗口位置
                # position 窗口在點的位置 'ur'-右上, 'ul'-左上, 'll'-左下, 'lr'-右下
                fwday = calendar.SUNDAY
                year = datetime.now().year  # 為使打開頁面時為當下年份
                month = datetime.now().month  # 為使打開頁面時為當下月份
                locale = None
                sel_bg = '#ecffc4'  # 設定點擊日期後的框顏色
                sel_fg = '#05640e'  # 設定點擊日期後的字底色
                s._date = datetime(year, month, 1)  # 該月份第一天
                s._selection = None  # 設置未選中的日期
                s.G_Frame = ttk.Frame(self.window)
                s._cal = s.__get_calendar(locale, fwday)  # 實例化適當的日曆類
                s.__setup_styles()  # 創建自定義樣式
                s.__place_widgets()  # pack/grid 小部件
                s.__config_calendar()  # 調整日曆列和安裝標記
                # 配置畫布和正確的绑定，以選擇日期。
                s.__setup_selection(sel_bg, sel_fg)
                # 存儲項ID，用於稍後插入。
                s._items = [s._calendar.insert('', 'end', values='') for _ in range(6)]
                # 在當前空日曆中插入日期
                s._update()
                s.G_Frame.place(x=290, y=120)
                self.window.update_idletasks()  # 刷新頁面

                self.window.deiconify()  # 還原視窗
                self.window.focus_set()  # 焦點設置在所需小部件上
                self.window.wait_window()  # 直到按確定

            def __get_calendar(s, locale, fwday):  # 日曆文字化
                if locale is None:
                    return calendar.TextCalendar(fwday)
                else:
                    return calendar.LocaleTextCalendar(fwday, locale)

            def __setup_styles(s):  # 自定義TTK風格
                style = ttk.Style(self.window)
                arrow_layout = lambda dir: (
                    [('Button.focus', {'children': [('Button.%sarrow' % dir, None)]})])  # 返回參數性質
                style.layout('L.TButton', arrow_layout('left'))  # 製作點選上個月的箭頭
                style.layout('R.TButton', arrow_layout('right'))  # 製作點選下個月的箭頭

            def __place_widgets(s):  # 標題框架及其小部件
                Input_judgment_num = self.window.register(s.Input_judgment)  # 需要将函数包装一下，必要的
                hframe = ttk.Frame(s.G_Frame)
                gframe = ttk.Frame(s.G_Frame)
                bframe = ttk.Frame(s.G_Frame)
                hframe.pack(in_=s.G_Frame, side='top', pady=5, anchor='center')  # 月曆的上視窗
                gframe.pack(in_=s.G_Frame, fill=tk.X, pady=5)
                bframe.pack(in_=s.G_Frame, side='bottom', pady=5)
                lbtn = ttk.Button(hframe, style='L.TButton',
                                  command=s._prev_month)  # 月曆上方左箭頭，點選後月曆切換至前個月
                lbtn.grid(in_=hframe, column=0, row=0, padx=12)
                rbtn = ttk.Button(hframe, style='R.TButton',
                                  command=s._next_month)  # 月曆上方右箭頭，點選後月曆切換至下個月
                rbtn.grid(in_=hframe, column=5, row=0, padx=12)

                s.CB_year = ttk.Combobox(hframe, width=5, values=[str(year) for year in
                                                                  range(datetime.now().year, datetime.now().year + 11,
                                                                        1)], validate='key',
                                         validatecommand=(Input_judgment_num, '%P'))  # 製作月曆年份的下拉選單
                s.CB_year.current(0)  # 月曆一開始呈現當下年分
                s.CB_year.grid(in_=hframe, column=1, row=0)
                s.CB_year.bind("<<ComboboxSelected>>", s._update)  # 選取(年)下拉式選單後，月曆更新
                tk.Label(hframe, text='年', justify='left').grid(in_=hframe, column=2, row=0,
                                                                padx=(0, 5))  # 下拉式選單後面的單位(年)
                s.CB_month = ttk.Combobox(hframe, width=3, values=['%02d' % month for month in range(1, 13)],
                                          state='readonly')  # 製作月曆年份的下拉選單
                s.CB_month.current(datetime.now().month - 1)  # 月曆一開始呈現當下月份
                s.CB_month.grid(in_=hframe, column=3, row=0)
                s.CB_month.bind("<<ComboboxSelected>>", s._update)  # 選取(月)下拉式選單後，月曆更新
                tk.Label(hframe, text='月', justify='left').grid(in_=hframe, column=4, row=0)  # 下拉式選單後面的單位(年)
                # 日曆部件
                s._calendar = ttk.Treeview(gframe, show='', selectmode='none', height=7)  # 建立放上日期的頁面
                s._calendar.pack(expand=1, fill='both', side='bottom', padx=5)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=0, relwidth=1, relheigh=2 / 200)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=198 / 200, relwidth=1, relheigh=2 / 200)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=0, rely=0, relwidth=2 / 200, relheigh=1)
                tk.Frame(s.G_Frame, bg='#565656').place(x=0, y=0, relx=198 / 200, rely=0, relwidth=2 / 200, relheigh=1)

            def __config_calendar(s):  # 設計日曆架構
                cols = ['日', '一', '二', '三', '四', '五', '六']  # 日曆上的星期幾
                s._calendar['columns'] = cols  # 設定日曆欄
                s._calendar.tag_configure('header', background='grey90')
                s._calendar.insert('', 'end', values=cols, tag='header')  # 調整其列寬
                font = tkfont.Font()
                maxwidth = max(font.measure(col) for col in cols)
                for col in cols:
                    s._calendar.column(col, width=maxwidth, minwidth=maxwidth, anchor='center')

            def __setup_selection(s, sel_bg, sel_fg):
                def __canvas_forget(evt):
                    canvas.place_forget()
                    s._selection = None

                s._font = tkfont.Font()
                s._canvas = canvas = tk.Canvas(s._calendar, background=sel_bg, borderwidth=0, highlightthickness=0)
                canvas.text = canvas.create_text(0, 0, fill=sel_fg, anchor='w')

                s._calendar.bind('<Button-1>', s._pressed)  # 點出要的日期

            def _build_calendar(s):  # 建立月曆-日期
                year = s._date.year  # s._date = datetime(year, month, 1)， 因此設定變數year為操作程式當下年分
                month = s._date.month  # 設定變數month為操作程式當下月份
                header = s._cal.formatmonthname(year, month, 0)  # 更新日曆顯示的日期
                cal = s._cal.monthdayscalendar(year, month)  # 此函數會建立指定年月份的周列表，以供屆時放入treeview
                for indx, item in enumerate(s._items):  # s._items為在treeview中先在每格插入""，一排放7個(先建立月曆的架構)
                    week = cal[indx] if indx < len(cal) else []  # 因為一個月頂多四周多，因此若indx>cal長度，將week設為空list
                    fmt_week = [('%02d' % day) if day else '' for day in week]
                    s._calendar.item(item, values=fmt_week)  # 將該年月份的日期分配放置treeview

            def _show_select(s, text, bbox):  # 秀出挑選的日子
                x, y, width, height = bbox
                textw = s._font.measure(text)
                canvas = s._canvas
                canvas.configure(width=width, height=height)
                canvas.coords(canvas.text, (width - textw) / 2, height / 2 - 1)
                canvas.itemconfigure(canvas.text, text=text)
                canvas.place(in_=s._calendar, x=x, y=y)

            def _pressed(s, evt=None, item=None, column=None, widget=None, confirm=False):  # 在日曆的某個地方點擊。

                if not item:
                    x, y, widget = evt.x, evt.y, evt.widget
                    item = widget.identify_row(y)
                    column = widget.identify_column(x)
                if not column or item not in s._items:  # 在工作日行中單擊或僅在列外單擊。
                    return
                item_values = widget.item(item)['values']  # 點選的日期該周
                if not len(item_values):  # 該行是空的。
                    return
                text = item_values[int(column[1]) - 1]  # text 為選擇日期
                if not text:  # 日期為空
                    return
                bbox = widget.bbox(item, column)
                if not bbox:  # 日曆尚未出現
                    self.window.after(20, lambda: s._pressed(item=item, column=column, widget=widget))
                    return
                text = '%02d' % text
                if confirm:
                    pass
                else:
                    s._selection = (text, item, column)
                    year = s._date.year  # 使用者選取日期中的年份
                    month = s._date.month  # 使用者選取日期中的月份
                    choose_date = s._selection[0]  # 使用者選取日期中的"日"
                    today_year = datetime.now().year
                    today_month = datetime.now().month
                    today_day = datetime.now().day
                    date = str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0]))
                    if int(str(year)) > int(str(today_year)):
                        if len(date_list) != 0:
                            if date in date_list:  # 若選擇日期出現在date_list，代表重複日期，必須跳出錯誤訊息
                                self.window.lower(belowThis=self.page2)
                                tkmessage.showerror(title="日期重複", message="此日期已選擇")
                                self.window.wm_attributes('-topmost', 1)
                            else:
                                s._show_select(text, bbox)
                                date_list.append(str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0])))
                                self.enydate.insert("end", (str(year) + "/" + str(month) + "/" +
                                                            ("%02d" % int(s._selection[0]))))
                        else:
                            s._show_select(text, bbox)
                            date_list.append("1")
                    elif int(str(year)) == int(str(today_year)) and int(str(month)) > int(str(today_month)):
                        if len(date_list) != 0:
                            if date in date_list:
                                self.window.lower(belowThis=self.page2)
                                tkmessage.showerror(title="日期重複", message="此日期已選擇")
                                self.window.wm_attributes('-topmost', 1)
                            else:
                                s._show_select(text, bbox)
                                date_list.append(str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0])))
                                self.enydate.insert("end", (str(year) + "/" + str(month) + "/" +
                                                            ("%02d" % int(s._selection[0]))))
                        else:
                            s._show_select(text, bbox)
                            date_list.append("1")
                    elif int(str(year)) == int(str(today_year)) and int(str(month)) == int(str(today_month)) and int(
                            str(choose_date)) >= int(str(today_day)):
                        if len(date_list) != 0:
                            if date in date_list:
                                self.window.lower(belowThis=self.page2)
                                tkmessage.showerror(title="日期重複", message="此日期已選擇")
                                self.window.wm_attributes('-topmost', 1)
                            else:
                                s._show_select(text, bbox)
                                date_list.append(str(year) + "/" + str(month) + "/" + ("%02d" % int(s._selection[0])))
                                self.enydate.insert("end", (str(year) + "/" + str(month) + "/" +
                                                            ("%02d" % int(s._selection[0]))))
                        else:
                            s._show_select(text, bbox)
                            date_list.append("1")
                    else:  # 若選取的日期為今天之前的日期，跳出錯誤訊息
                        self.window.lower(belowThis=self.page2)
                        tkmessage.showerror(title="日期無效", message="請選擇有效日期")
                        self.window.wm_attributes('-topmost', 1)

            def _prev_month(s):  # 按下左箭頭後，月曆頁面需進行切換
                s._canvas.place_forget()
                s._selection = None
                s._date = s._date - timedelta(days=1)
                s._date = datetime(s._date.year, s._date.month, 1)
                s.CB_year.set(s._date.year)
                s.CB_month.set(s._date.month)
                s._update()

            def _next_month(s):  # 按下右箭頭後，月曆頁面需進行切換
                s._canvas.place_forget()
                s._selection = None
                year, month = s._date.year, s._date.month
                s._date = s._date + timedelta(
                    days=calendar.monthrange(year, month)[1] + 1)
                s._date = datetime(s._date.year, s._date.month, 1)
                s.CB_year.set(s._date.year)
                s.CB_month.set(s._date.month)
                s._update()

            def _update(s, event=None, key=None):
                if key and event.keysym != 'Return':
                    return

                year = int(s.CB_year.get())
                month = int(s.CB_month.get())

                if year == 0 or year > 9999:
                    return

                s._canvas.place_forget()
                s._date = datetime(year, month, 1)
                s._build_calendar()  # 重建日曆

                if year == datetime.now().year and month == datetime.now().month:
                    day = datetime.now().day
                    for _item, day_list in enumerate(s._cal.monthdayscalendar(year, month)):
                        if day in day_list:
                            item = 'I00' + str(_item + 2)
                            column = '#' + str(day_list.index(day) + 1)
                            self.window.after(100, lambda: s._pressed(item=item, column=column, widget=s._calendar))

            def _main_judge(s):
                try:
                    if self.window.focus_displayof() is None or 'toplevel' not in str(self.window.focus_displayof()):
                        s._pressed()
                    else:
                        self.window.after(10, s._main_judge)
                except:
                    self.window.after(10, s._main_judge)

            def selection(s):
                if not s._selection:
                    return None
                year = s._date.year
                month = s._date.month

                return str(datetime(year, month, int(s._selection[0])))[:10]

            def Input_judgment(s, content):
                if content.isdigit() or content == "":
                    return True
                else:
                    return False

        cal = Calendar()

    def click_btn_delete(self):
        choose = []
        choose_date = []
        choose_tuple = self.enydate.curselection()
        global date_list

        for i in choose_tuple:
            choose.append(i)
            choose_date.append(str(self.enydate.get(i)))

        for k in range(len(choose)):
            self.enydate.delete(choose[k] - k)

        for k in choose_date:
            date_list.remove(k)

    def click_btnYes(self):
        if self.enydate.size() != 0 and self.inputname.get() == "":
            self.window.lower(belowThis=None)
            tkmessage.showerror(title="輸入未完整", message="您尚未輸入會議名稱")
            self.window.wm_attributes('-topmost', 1)
        elif self.enydate.size() == 0 and self.inputname.get() != "":
            self.window.lower(belowThis=None)
            tkmessage.showerror(title="輸入未完整", message="您尚未輸入會議日期")
            self.window.wm_attributes('-topmost', 1)
        elif self.enydate.size() == 0 and self.inputname.get() == "":
            self.window.lower(belowThis=None)
            tkmessage.showerror(title="輸入未完整", message="您尚未輸入會議名稱及日期")
            self.window.wm_attributes('-topmost', 1)
        else:
            global folder_meeting_names
            if self.inputname.get() in folder_meeting_names:
                self.window.lower(belowThis=None)
                tkmessage.showerror(title="名稱錯誤", message="此名稱已存在")
                self.window.wm_attributes('-topmost', 1)
            else:
                self.window.destroy()
                name = meeting_name.get()
                gc.create(str(name) + " in " + str(meeting_names[folder_location]))

                global wb_record, sheet_time
                wb_record = gc.open(str(name) + " in " + str(meeting_names[folder_location]))
                sheet_time = wb_record[0]
                sheet_time.title = '時間統計'
                ws = wb_record.add_worksheet('出缺勤')
                wb_record.add_worksheet('Meeting record')

                wb_record.share('arial5623@gmail.com', role='writer', type='user')
                wb_record.share('alice0911496698@gmail.com', role='writer', type='user')
                wb_record.share('sunnyliu891114@gmail.com', role='writer', type='user')
                wb_record.share('jenny900707@gmail.com', role='writer', type='user')
                wb_record.share('teresa890101@gmail.com', role='writer', type='user')
                wb_record.share('l1849261495@gmail.com', role='writer', type='user')

                global date_list
                date_list.remove('1')
                date_list = sorted(date_list)

                columns = []
                rows = []

                for i in range(len(date_list)):
                    columns.append(date_list[i])

                for i in range(16):
                    rows.append(str(7 + i) + ':00-' + str(8 + i) + ':00')

                df = pd.DataFrame(columns=columns, index=rows)
                sheet_time.set_dataframe(df, start='A1', copy_index=True, copy_head=True, nan='')

                global sheet_names
                sheet_names.append_table([name, 'unfinished'], dimension='ROWS', overwrite=False)

                ws.update_row(index=1, values=['name', 'absence', 'mission'], col_offset=0)

                self.page2.destroy()
                Page2()

    def click_btn_meetings(self, a):
        global name, location, wb_record, sheet_time, df_sheet_time, dates
        name = folder_meeting_names[a]
        finish = folder_finish_meeting[a]
        location = a
        wb_record = gc.open(str(name) + " in " + str(meeting_names[folder_location]))
        sheet_time = wb_record[0]

        df_sheet_time = sheet_time.get_as_df(has_header=True, include_tailing_empty=False)
        df_sheet_time.drop(columns=[''], inplace=True)
        dates = df_sheet_time.columns.tolist()

        if finish == 'finished':
            self.page2.destroy()
            Page9()
        else:
            self.page2.destroy()
            Page3()


class Page3:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘

        self.root = master
        self.page3 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page3.master.title(name)
        self.page3.grid()

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=20, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        self.lab_title_3 = tk.Label(self.page3, text=' ' + str(name), height=1, width=15, font=f1, bg=color_1,
                                    fg='white', anchor='w')
        self.btn_createtime = tk.Button(self.page3, text="新增你的時間", height=1, width=18, font=f2, fg=color_1,
                                        relief='solid', command=self.click_btn_createtime)
        self.btn_times = tk.Button(self.page3, text="查看所有人的時間", height=1, width=18, font=f2, fg=color_1,
                                   relief='solid', command=self.click_btn_times)
        self.btn_meetingrecord = tk.Button(self.page3, text="紀錄會議", height=1, width=18, font=f2, fg=color_1,
                                           relief='solid', command=self.click_btn_meetingrecord)
        self.btn_back = tk.Button(self.page3, text="返回", height=1, font=f3, command=self.click_btn_back, bg=color_3)

        self.lab_title_3.place(x=0, y=50)
        self.btn_createtime.place(relx=0.5, y=200, anchor='center')
        self.btn_times.place(relx=0.5, y=300, anchor='center')
        self.btn_meetingrecord.place(relx=0.5, y=400, anchor='center')
        self.btn_back.place(relx=0.5, y=600, anchor='center')

    def click_btn_createtime(self):
        self.page3.destroy()
        Page4()

    def click_btn_times(self):
        self.page3.destroy()
        Page5()

    def click_btn_meetingrecord(self):
        self.page3.destroy()
        Page6()

    def click_btn_back(self):
        self.page3.destroy()
        try:
            wb = gc.open(str(name) + " in " + str(meeting_names[folder_location]))
            Page2()
        except (pygsheets.exceptions.SpreadsheetNotFound, NameError):
            Page1()


class Page4:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘

        self.root = master
        self.page4 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page4.master.title("新增你的時間")
        self.page4.grid()

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=20, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        global var_name, var_selectall

        self.lab_title = tk.Label(self.page4, text=' 新增你的時間', height=1, width=15, font=f1,
                                  bg=color_1, fg='white', anchor='w').place(x=0, y=50)
        self.lab_name = tk.Label(self.page4, text='姓名：', font=f2, bg=color_2)

        var_name = tk.StringVar()
        self.int_name = tk.Entry(self.page4, textvariable=var_name, width=18, font=f2)
        self.btn_yes = tk.Button(self.page4, text='確定', bg=color_3, height=1, font=f3, command=self.click_btn_yes)
        self.btn_back = tk.Button(self.page4, text='返回', bg=color_3, height=1, font=f3, command=self.click_btn_back)

        var_selectall = tk.IntVar()
        self.btn_selectall = tk.Checkbutton(self.page4, onvalue=1, offvalue=0, variable=var_selectall, font=f1,
                                            bg=color_2, command=self.click_selectall)
        self.lab_selectall = tk.Label(self.page4, text='全選', font=f3, bg=color_2)

        self.lab_name.place(x=100, y=170)
        self.int_name.place(x=100, y=210, height=30)
        self.btn_back.place(x=440, y=620)
        self.btn_yes.place(x=515, y=620)
        self.btn_selectall.place(x=855, y=35)
        self.lab_selectall.place(x=878, y=50)

        global chk_btns
        chk_btns = []

        self.canvas = tk.Canvas(self.page4, width=414, height=510, bg=color_2)
        self.canvas.place(x=498, y=80)
        self.slb = tk.Scrollbar(self.page4, orient='horizontal')
        self.canvas.configure(xscrollcommand=self.slb.set)
        self.slb.configure(command=self.canvas.xview)
        self.frame_context = tk.Frame(self.canvas, width=10000, height=1000, bg=color_2)
        self.canvas.create_window((0, 0), window=self.frame_context, anchor='nw')

        self.canvas_width = 2

        if len(dates) > 7:
            self.slb.place(x=555, y=595, relwidth=0.362, height=10)

        for i in range(len(dates) + 1):
            if i != 0:
                chk_btns.append([])
            self.canvas_width += 52
            for j in range(17):
                if i == 0:
                    if j == 0:
                        tk.Label(self.page4, relief='solid', borderwidth=1, width=10, height=2, bg=color_2).place(x=480,
                                                                                                                  y=80)
                    else:
                        tk.Label(self.page4, text=str(6 + j) + ':00-' + str(7 + j) + ':00', relief='solid',
                                 borderwidth=1, width=10, height=2, bg=color_2).place(x=480, y=80 + 30 * j)
                else:
                    if j == 0:
                        tk.Label(self.frame_context, text=dates[i - 1][5:], borderwidth=1, relief='solid', width=7,
                                 height=2, bg=color_2, anchor='center').place(x=52 * i, y=0)
                    else:
                        var_i = tk.IntVar()
                        tk.Label(self.frame_context, borderwidth=1, relief='solid', width=7, height=2,
                                 bg=color_2).place(x=52 * i, y=30 * j)
                        tk.Checkbutton(self.frame_context, onvalue=1, offvalue=0, variable=var_i,
                                       bg=color_2).place(x=17 + 52 * i, y=5 + 30 * j)
                        chk_btns[i - 1].append(var_i)

        self.canvas.configure(scrollregion=(0, 0, self.canvas_width, 530))

    def click_btn_yes(self):
        abs_sheet = wb_record.worksheet_by_title('出缺勤')

        if var_name.get() == "":
            tkmessage.showerror(title="請輸入姓名", message="請輸入姓名！")
        else:
            member = sheet_time.cell((18, 1)).value
            if member == '':
                list_member = [var_name.get()]
                abs_sheet.update_col(index=1, values=list_member, row_offset=1)
                for i in range(len(dates)):
                    for j in range(16):
                        if chk_btns[i][j].get() == 1:
                            available = df_sheet_time.iat[j, i]
                            if available == '':
                                list_available = [var_name.get()]
                            else:
                                list_available = available.split(',')
                                list_available.append(var_name.get())
                            df_sheet_time.iloc[j, i] = ",".join(list_available)
            else:
                list_member = member.split(',')
                if var_name.get() not in list_member:
                    list_member.append(var_name.get())
                    abs_sheet.update_col(index=1, values=list_member, row_offset=1)
                    for i in range(len(dates)):
                        for j in range(16):
                            if chk_btns[i][j].get() == 1:
                                available = df_sheet_time.iat[j, i]
                                if available == '':
                                    list_available = [var_name.get()]
                                else:
                                    list_available = available.split(',')
                                    list_available.append(var_name.get())
                                df_sheet_time.iloc[j, i] = ",".join(list_available)
                else:
                    for i in range(len(dates)):
                        for j in range(16):
                            available = df_sheet_time.iat[j, i]
                            if chk_btns[i][j].get() == 1:
                                if available == '':
                                    list_available = [var_name.get()]
                                else:
                                    list_available = available.split(',')
                                    if var_name.get() not in available:
                                        list_available.append(var_name.get())
                                df_sheet_time.iloc[j, i] = ",".join(list_available)
                            else:
                                if var_name.get() in available:
                                    available = available.replace(var_name.get(), "").replace(",,", ",").strip(",")
                                    df_sheet_time.iloc[j, i] = available

            sheet_time.update_value((18, 1),  (",".join(list_member)))
            sheet_time.set_dataframe(df_sheet_time, start='B2', copy_index=False, copy_head=False)

            self.page4.destroy()
            Page5()

    def click_btn_back(self):
        self.page4.destroy()
        Page3()

    def click_selectall(self):
        if var_selectall.get() == 1:
            for i in range(len(dates)):
                for j in range(16):
                    chk_btns[i][j].set(1)
        else:
            for i in range(len(dates)):
                for j in range(16):
                    chk_btns[i][j].set(0)


class Page5:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘

        self.root = master
        self.page5 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page5.master.title("查看所有人的時間")
        self.page5.grid()

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=12, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")
        f4 = tkfont.Font(size=10, family="源泉圓體 M")

        self.lab_title = tk.Label(self.page5, text=" 查看所有人的時間", height=1, width=15, font=f1, bg=color_1,
                                  fg='white', anchor='w').place(x=0, y=50)

        self.btn_yes = tk.Button(self.page5, text='確定', bg=color_3, height=1, font=f3, command=self.click_btn_yes)
        self.lab_allmembers = tk.Label(self.page5, text='all members', font=f2, bg=color_2)
        self.lab_select = tk.Label(self.page5, text='select', font=f2, bg=color_2)
        self.lab_available = tk.Label(self.page5, text='available', font=f2, bg=color_2)
        self.lab_unavailable = tk.Label(self.page5, text='unavailable', font=f2, bg=color_2)

        self.btn_yes.place(x=475, y=620)
        self.lab_allmembers.place(x=100, y=400)
        self.lab_select.place(x=330, y=400)
        self.lab_available.place(x=320, y=125)
        self.lab_unavailable.place(x=102, y=125)

        self.scroll_available = tk.Scrollbar(self.page5)
        self.scroll_unavailable = tk.Scrollbar(self.page5)
        self.scroll_allmembers = tk.Scrollbar(self.page5)
        self.scroll_color = tk.Scrollbar(self.page5)

        self.scroll_available.place(x=413, y=152, relheight=0.33)
        self.scroll_unavailable.place(x=208, y=152, relheight=0.33)
        self.scroll_allmembers.place(x=208, y=426, relheight=0.194)
        self.scroll_color.place(x=413, y=426, relheight=0.194)

        var_available = tk.StringVar()
        self.lst_available = tk.Listbox(self.page5, listvariable=var_available, font=f2, width=14, height=12,
                                        yscrollcommand=self.scroll_available.set)
        self.lst_available.place(x=285, y=151)

        var_unavailable = tk.StringVar()
        self.lst_unavailable = tk.Listbox(self.page5, listvariable=var_unavailable, font=f2, width=14, height=12,
                                          yscrollcommand=self.scroll_unavailable.set)
        self.lst_unavailable.place(x=80, y=151)

        var_allmembers = tk.StringVar()
        self.lst_allmembers = tk.Listbox(self.page5, listvariable=var_allmembers, font=f2, width=14, height=7,
                                         yscrollcommand=self.scroll_allmembers.set)
        self.lst_allmembers.place(x=80, y=425)

        all_members = sheet_time.cell((18, 1)).value.split(',')
        how_many_people = len(all_members)

        for member in all_members:
            self.lst_allmembers.insert("end", member)

        self.btn_try = tk.Button(self.page5, text='try', font=f4, width=5, command=self.click_btn_try)
        self.btn_reset = tk.Button(self.page5, text='reset', font=f4, width=5, command=self.click_btn_reset)

        self.btn_try.place(x=298, y=570)
        self.btn_reset.place(x=363, y=570)

        self.lst_color = tk.Listbox(self.page5, width=14, height=7, font=f2, selectmode=tk.MULTIPLE,
                                    yscrollcommand=self.scroll_color.set)
        self.lst_color.place(x=285, y=425)
        self.color = str()

        global color_list, people_list, btn_list
        color_list = []
        people_list = []
        btn_list = []

        for i in range(how_many_people):
            self.lst_color.insert('end', i + 1)
            self.color = colors[int(30 / int(how_many_people)) * i]
            color_list.append(self.color)
            self.lst_color.itemconfig('end', bg=self.color, selectbackground=self.color)

        # 顏色漸層
        """hls_color_list = sns.color_palette("Reds", n_colors=how_many_people)
        for color in hls_color_list:
            color_list.append(self._from_rgb((int(250 * color[0]), int(250 * color[1]), int(250 * color[2]))))

        for i in range(how_many_people):
            self.lst_color.insert('end', i + 1)
            self.color = color_list[i]
            self.lst_color.itemconfig('end', bg=self.color, selectbackground=self.color)"""

        self.canvas = tk.Canvas(self.page5, width=414, height=510, bg=color_2)
        self.canvas.place(x=498, y=80)
        self.slb = tk.Scrollbar(self.page5, orient='horizontal')
        if len(dates) > 7:
            self.slb.place(x=555, y=595, relwidth=0.362, height=10)
        self.canvas.configure(xscrollcommand=self.slb.set)
        self.slb.configure(command=self.canvas.xview)
        self.frame_context = tk.Frame(self.canvas, width=10000, height=1000, bg=color_2)
        self.canvas.create_window((0, 0), window=self.frame_context, anchor='nw')

        self.canvas_width = 2

        for i in range(len(dates) + 1):
            if i != 0:
                people_list.append([])
                btn_list.append([])
            self.canvas_width += 52
            for j in range(17):
                if i == 0:
                    if j == 0:
                        tk.Label(self.page5, relief='solid', borderwidth=1, width=10, height=2, bg=color_2).place(x=480,
                                                                                                                  y=80)
                    else:
                        tk.Label(self.page5, text=str(6 + j) + ':00-' + str(7 + j) + ':00', relief='solid',
                                 borderwidth=1, width=10, height=2, bg=color_2).place(x=480, y=80 + 30 * j)
                else:
                    if j == 0:
                        tk.Label(self.frame_context, text=dates[i - 1][5:], borderwidth=1, relief='solid', width=7,
                                 height=2, bg=color_2, anchor='center').place(x=52 * i, y=0)
                    else:
                        tk.Label(self.frame_context, relief='solid', borderwidth=1, width=7, height=2,
                                 bg=color_2).place(x=52 * i, y=30 * j)
                        try:
                            how_many_available = len(df_sheet_time.iat[j - 1, i - 1].split(','))
                            if df_sheet_time.iat[j - 1, i - 1] == '':
                                how_many_available = 0
                        except IndexError:
                            how_many_available = 0

                        if how_many_available != 0:
                            self.btn = tk.Button(self.frame_context, bg=str(color_list[how_many_available - 1]),
                                                 width=5, height=1, command=lambda a=i, b=j: self.click_btn(a, b))
                        else:
                            self.btn = tk.Button(self.frame_context, bg='white', width=5, height=1,
                                                 command=lambda a=i, b=j: self.click_btn(a, b))
                        self.btn.place(x=4 + 52 * i, y=3 + 30 * j)
                        people_list[i - 1].append(how_many_available)
                        btn_list[i - 1].append(self.btn)

        self.canvas.configure(scrollregion=(0, 0, self.canvas_width, 530))

        self.scroll_available.config(command=self.lst_available.yview)
        self.scroll_unavailable.config(command=self.lst_unavailable.yview)
        self.scroll_allmembers.config(command=self.lst_allmembers.yview)
        self.scroll_color.config(command=self.lst_color.yview)

    def click_btn_try(self):
        number_choose = list()
        number_choose_tuple = self.lst_color.curselection()
        for i in number_choose_tuple:
            number_choose.append(self.lst_color.get(i))

        for i in range(len(dates)):
            for j in range(16):
                for k in range(len(number_choose)):
                    if k == 0:
                        if int(people_list[i][j]) == number_choose[k]:
                            btn_list[i][j].config(bg=str(color_list[int(number_choose[k]) - 1]))
                        else:
                            btn_list[i][j].config(bg='white')
                    else:
                        if int(people_list[i][j]) == number_choose[k]:
                            btn_list[i][j].config(bg=str(color_list[int(number_choose[k]) - 1]))

    def click_btn_reset(self):
        self.lst_color.select_clear(first=0, last='end')

        for i in range(1, len(dates) + 1):
            for j in range(1, 17):
                try:
                    how_many_available = len(df_sheet_time.iat[j - 1, i - 1].split(','))
                    if df_sheet_time.iat[j - 1, i - 1] == '':
                        how_many_available = 0
                except IndexError:
                    how_many_available = 0

                if how_many_available != 0:
                    btn_list[i - 1][j - 1].config(bg=str(color_list[how_many_available - 1]))
                    people_list[i - 1].append(how_many_available)
                else:
                    btn_list[i - 1][j - 1].config(bg='white')
                    people_list[i - 1].append(how_many_available)

    def click_btn_yes(self):
        self.page5.destroy()
        Page3()

    def click_btn(self, a, b):
        self.lst_available.delete(0, "end")
        self.lst_unavailable.delete(0, "end")

        all_members = str(sheet_time.cell((18, 1)).value).split(',')
        try:
            available_member = df_sheet_time.iat[b - 1, a - 1].split(',')
        except IndexError:
            available_member = []

        for member in available_member:
            self.lst_available.insert("end", member)

        for member in all_members:
            if member not in available_member:
                self.lst_unavailable.insert("end", member)


class Page6:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘

        self.root = master
        self.page6 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page6.master.title(name + '-紀錄')
        self.page6.grid()

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=20, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        self.lblTitle_6 = tk.Label(self.page6, text=' ' + name + '-紀錄', height=1, width=15, font=f1, bg=color_1,
                                   fg='white', anchor='w').place(x=0, y=50)
        self.btn6_1 = tk.Button(self.page6, text="出缺勤", height=1, width=18, font=f2, relief='solid', fg=color_1,
                                command=self.click_btn6_1)
        self.btn6_2 = tk.Button(self.page6, text="會議記錄", height=1, width=18, font=f2, relief='solid', fg=color_1,
                                command=self.click_btn6_2)
        self.btn6_3 = tk.Button(self.page6, text="結束會議", height=1, width=18, font=f2, relief='solid', fg=color_1,
                                command=self.click_btn6_3)
        self.btn6_4 = tk.Button(self.page6, text="返回", font=f3, bg=color_3, command=self.click_btn6_4)

        self.btn6_1.place(relx=0.5, y=200, anchor='center')
        self.btn6_2.place(relx=0.5, y=300, anchor='center')
        self.btn6_3.place(relx=0.5, y=400, anchor='center')
        self.btn6_4.place(relx=0.5, y=600, anchor='center')

    def click_btn6_1(self):
        self.page6.destroy()
        Page7()

    def click_btn6_2(self):
        self.page6.destroy()
        Page8()

    def click_btn6_3(self):
        message = tkmessage.askokcancel(title="確定結束會議？", message="結束會議後，您將無法作任何更動")
        if message:
            sheet_names.update_value((location + 1, 2), 'finished')

            self.page6.destroy()
            Page9()

    def click_btn6_4(self):
        self.page6.destroy()
        Page3()


class Page7:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘

        self.root = master
        self.page7 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page7.master.title("出缺勤")
        self.page7.grid()
        # 因為scrollbar不能直接應用在frame上，所以就創了一個canvas，在canvas上加scrollbar，然後再把canvas放進frame
        self.canvas = tk.Canvas(self.page7, width=1000, height=700, bg=color_2)
        self.canvas.grid()
        self.slb = tk.Scrollbar(self.page7, orient='vertical')
        self.slb.place(x=980, width=20, height=700)
        self.canvas.configure(yscrollcommand=self.slb.set)
        self.slb.configure(command=self.canvas.yview)
        self.frame_context = tk.Frame(self.canvas, width=1000, height=1000000, bg=color_2)
        self.canvas.create_window((-2, -2), window=self.frame_context, anchor='nw')

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=12, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        self.lab_title = tk.Label(self.frame_context, width=15, height=1, text=' 出缺勤', font=f1, fg='white', bg=color_1,
                                  anchor='w').place(x=0, y=50)

        self.btn7_1 = tk.Button(self.frame_context, text="確定", font=f3, bg=color_3, command=self.click_btn7_1)
        self.btn7_1.place(x=850, y=75, anchor='center')

        self.lab7_1 = tk.Label(self.frame_context, text="準時", font=f2, bg=color_2).place(x=179, y=125)
        self.lab7_2 = tk.Label(self.frame_context, text="遲到", font=f2, bg=color_2).place(x=279, y=125)
        self.lab7_3 = tk.Label(self.frame_context, text="未出席", font=f2, bg=color_2).place(x=372, y=125)

        self.lab7_4 = tk.Label(self.frame_context, text="是", font=f2, bg=color_2).place(x=700, y=125)
        self.lab7_5 = tk.Label(self.frame_context, text="否", font=f2, bg=color_2).place(x=770, y=125)
        self.lab7_6 = tk.Label(self.frame_context, text="無任務", font=f2, bg=color_2).place(x=823, y=125)

        global sheet, df_sheet, member_list, absence_value, mission_value
        sheet = wb_record.worksheet_by_title("出缺勤")
        df_sheet = sheet.get_as_df(has_header=False, index_column=None, include_tailing_empty=False)
        member_list = str(sheet_time.cell((18, 1)).value).split(',')
        absence_value = []
        mission_value = []

        try:
            absence = df_sheet.iloc[1:, 1].tolist()
            mission = df_sheet.iloc[1:, 2].tolist()
        except IndexError:
            absence = []
            mission = []

        for i in range(len(absence)):
            if absence[i] == '準時':
                absence[i] = 1
            if absence[i] == '遲到':
                absence[i] = 2
            if absence[i] == '未出席':
                absence[i] = 3

        for i in range(len(mission)):
            if mission[i] == '完成任務':
                mission[i] = 1
            if mission[i] == '未完成任務':
                mission[i] = 2
            if mission[i] == '無任務':
                mission[i] = 3

        canvas_height = 160  # 計算最後頁面會有多長
        for i in range(len(member_list)):
            tk.Label(self.frame_context, text=member_list[i], font=f2, bg=color_2).place(x=115, y=168 + 35 * i,
                                                                                         anchor='center')
            tk.Label(self.frame_context, text='是否完成指派任務？', font=f2, bg=color_2).place(x=515, y=155 + 35 * i)

            var_absence = tk.IntVar()
            absence_value.append(var_absence)
            if len(absence) > i:
                var_absence.set(absence[i])

            tk.Radiobutton(self.frame_context, variable=var_absence, value=1, bg=color_2).place(x=185, y=155 + 35 * i)
            tk.Radiobutton(self.frame_context, variable=var_absence, value=2, bg=color_2).place(x=285, y=155 + 35 * i)
            tk.Radiobutton(self.frame_context, variable=var_absence, value=3, bg=color_2).place(x=385, y=155 + 35 * i)

            var_mission = tk.IntVar()
            mission_value.append(var_mission)
            if len(mission) > i:
                var_mission.set(mission[i])

            tk.Radiobutton(self.frame_context, variable=var_mission, value=1, bg=color_2).place(x=700, y=155 + 35 * i)
            tk.Radiobutton(self.frame_context, variable=var_mission, value=2, bg=color_2).place(x=770, y=155 + 35 * i)
            tk.Radiobutton(self.frame_context, variable=var_mission, value=3, bg=color_2).place(x=840, y=155 + 35 * i)
            canvas_height += 35

        if canvas_height > 700:
            self.canvas.configure(scrollregion=(0, 0, 1000, canvas_height + 30))
        else:
            self.canvas.configure(scrollregion=(0, 0, 1000, 700))

    def click_btn7_1(self):
        for i in range(len(member_list)):
            if absence_value[i].get() == 1:
                df_sheet.iloc[i + 1, 1] = "準時"
            if absence_value[i].get() == 2:
                df_sheet.iloc[i + 1, 1] = "遲到"
            if absence_value[i].get() == 3:
                df_sheet.iloc[i + 1, 1] = "未出席"
            if mission_value[i].get() == 1:
                df_sheet.iloc[i + 1, 2] = "完成任務"
            if mission_value[i].get() == 2:
                df_sheet.iloc[i + 1, 2] = "未完成任務"
            if mission_value[i].get() == 3:
                df_sheet.iloc[i + 1, 2] = "無任務"

        sheet.set_dataframe(df_sheet, start='A1', copy_index=False, copy_head=False)

        self.page7.destroy()
        Page6()


class Page8:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘

        self.root = master
        self.page8 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page8.master.title("會議記錄")
        self.page8.grid()

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=15, family="源泉圓體 M")

        self.lbl8_1 = tk.Label(self.page8, text=" 會議記錄", height=1, width=15, font=f1, fg='white', bg=color_1,
                               anchor='w')
        self.record = tk.Text(self.page8, height=18, width=74, font=f2)
        self.btn8 = tk.Button(self.page8, text='確認', font=f2, bg=color_3, command=self.click_btn8)
        self.lbl8_2 = tk.Label(self.page8, text="會議時間：", font=f2, bg=color_2)

        global meeting_record
        meeting_record = wb_record.worksheet_by_title('Meeting record')

        if str(meeting_record.cell((1, 1)).value) != '':
            self.record.insert("1.0", str(meeting_record.cell((1, 1)).value))

        global var_times
        var_times = tk.StringVar()
        if str(meeting_record.cell((1, 2)).value) != '':
            var_times.set(meeting_record.cell((1, 2)).value)
        self.entry8 = tk.Entry(self.page8, textvariable=var_times, width=18, font=f2)

        self.lbl8_1.place(x=0, y=50)
        self.record.place(relx=0.5, rely=0.5, anchor='center')
        self.btn8.place(relx=0.5, y=620, anchor='center')
        self.entry8.place(x=705, y=105)
        self.lbl8_2.place(x=600, y=105)

    def click_btn8(self):
        meeting_record.update_value((1, 1), self.record.get("1.0", "end"))
        meeting_record.update_value((1, 2), var_times.get())

        self.page8.destroy()
        Page6()


class Page9:
    def _from_rgb(self, rgb):
        return "#%02x%02x%02x" % rgb

    def __init__(self, master=None):
        color_1 = self._from_rgb((68, 84, 106))  # 藍黑色
        color_2 = self._from_rgb((208, 224, 227))  # 湖水藍
        color_3 = self._from_rgb((255, 217, 102))  # 淡橘

        self.root = master
        self.page9 = tk.Frame(self.root, width=1000, height=700, bg=color_2)
        self.page9.master.title("會議已結束")
        self.page9.grid()

        self.canvas9 = tk.Canvas(self.page9, width=1000, height=700, bg=color_2)
        self.canvas9.place(x=0, y=0)
        self.slb9 = tk.Scrollbar(self.page9, orient='vertical')
        self.slb9.place(x=980, width=20, height=700)
        self.canvas9.configure(yscrollcommand=self.slb9.set, scrollregion=(0, 0, 1000, 1050))
        self.slb9.configure(command=self.canvas9.yview)
        self.frame_context9 = tk.Frame(self.canvas9, width=2000, height=10000, bg=color_2)
        self.canvas9.create_window((-2, -2), window=self.frame_context9, anchor='nw')

        f1 = tkfont.Font(size=30, family="源泉圓體 B")
        f2 = tkfont.Font(size=12, family="源泉圓體 M")
        f3 = tkfont.Font(size=15, family="源泉圓體 M")

        ws_1 = wb_record.worksheet_by_title('Meeting record')
        ws_2 = wb_record.worksheet_by_title('出缺勤')
        times = wb_record.worksheet_by_title('時間統計')

        df_ws_2 = ws_2.get_as_df(has_header=True, index_column=False, include_tailing_empty=False)
        member_list = str(times.cell((18, 1)).value).split(',')
        ontime_num = 0
        late_num = 0
        absence_num = 0
        ontime_member = []
        late_member = []
        absence_member = []

        done_num = 0
        undone_num = 0
        none_num = 0
        done_member = []
        undone_member = []
        none_member = []

        try:
            for i in range(len(member_list)):
                if df_ws_2.iat[i, 1] == "準時":
                    ontime_num += 1
                    ontime_member.append(str(df_ws_2.iat[i, 0]))
                if df_ws_2.iat[i, 1] == "遲到":
                    late_num += 1
                    late_member.append(str(df_ws_2.iat[i, 0]))
                if df_ws_2.iat[i, 1] == "未出席":
                    absence_num += 1
                    absence_member.append(str(df_ws_2.iat[i, 0]))
                    if df_ws_2.iat[i, 2] == "完成任務":
                        done_num += 1
                        done_member.append(str(df_ws_2.iat[i, 0]))
                    if df_ws_2.iat[i, 2] == "未完成任務":
                        undone_num += 1
                        undone_member.append(str(df_ws_2.iat[i, 0]))
                    if df_ws_2.iat[i, 2] == "無任務":
                        none_num += 1
                        none_member.append(str(df_ws_2.iat[i, 0]))
        except IndexError:
            pass

        self.lbl_title9 = tk.Label(self.frame_context9, text=' 會議已結束', height=1, width=15, font=f1, fg='white',
                                   bg=color_1, anchor='w').place(x=0, y=50)
        self.lbl9_1 = tk.Label(self.frame_context9, text="會議名稱：", font=f2, bg=color_2, fg=color_1)
        self.lbl9_2 = tk.Label(self.frame_context9, text=name, font=f2, bg=color_2)
        self.lbl9_3 = tk.Label(self.frame_context9, text="會議時間：", font=f2, bg=color_2, fg=color_1)
        self.lbl9_4 = tk.Label(self.frame_context9, text=str(ws_1.cell((1, 2)).value), font=f2, bg=color_2)
        self.lbl9_5 = tk.Label(self.frame_context9, text="會議記錄：", font=f2, bg=color_2, fg=color_1)
        self.lbl9_6 = tk.Label(self.frame_context9, text="準時", font=f2, bg=color_2, fg=color_1)
        self.lbl9_7 = tk.Label(self.frame_context9, text="遲到", font=f2, bg=color_2, fg=color_1)
        self.lbl9_8 = tk.Label(self.frame_context9, text="未出席", font=f2, bg=color_2, fg=color_1)
        self.lbl9_9 = tk.Label(self.frame_context9, text="完成任務", font=f2, bg=color_2, fg=color_1)
        self.lbl9_10 = tk.Label(self.frame_context9, text="未完成任務", font=f2, bg=color_2, fg=color_1)
        self.lbl9_11 = tk.Label(self.frame_context9, text="無任務", font=f2, bg=color_2, fg=color_1)
        self.btn9 = tk.Button(self.frame_context9, text="確定", font=f3, bg=color_3, command=self.click_btn9_1)

        self.lbl9_1.place(x=100, y=130)
        self.lbl9_2.place(x=180, y=130)
        self.lbl9_3.place(x=100, y=180)
        self.lbl9_4.place(x=180, y=180)
        self.lbl9_5.place(x=100, y=230)
        self.lbl9_6.place(x=475, y=485)
        self.lbl9_7.place(x=620, y=485)
        self.lbl9_8.place(x=755, y=485)
        self.lbl9_9.place(x=461, y=735)
        self.lbl9_10.place(x=595, y=735)
        self.lbl9_11.place(x=755, y=735)
        self.btn9.place(x=850, y=75, anchor='center')

        #  將會議記錄變成listbox配合scrollbar查閱
        self.scroll_meeting_record = tk.Scrollbar(self.frame_context9)
        self.scroll_meeting_record.place(x=824, y=261, relheight=0.0193)
        var_meeting_record = tk.StringVar()
        self.lst_meeting_record = tk.Listbox(self.frame_context9, listvariable=var_meeting_record, font=f2, width=80,
                                             height=10, yscrollcommand=self.scroll_meeting_record.set)
        self.lst_meeting_record.place(x=100, y=260)
        for record9 in str(ws_1.cell((1, 1)).value).split('\n'):
            self.lst_meeting_record.insert("end", record9)
        self.lst_meeting_record.config(yscrollcommand=self.scroll_meeting_record.set)
        self.scroll_meeting_record.config(command=self.lst_meeting_record.yview)

        # 出缺勤名單 準時
        self.scroll_ontime = tk.Scrollbar(self.frame_context9)
        self.scroll_ontime.place(x=533, y=511, relheight=0.0193)
        var_ontime = tk.StringVar()
        self.lst_ontime = tk.Listbox(self.frame_context9, listvariable=var_ontime, font=f2, width=10,
                                     height=10, yscrollcommand=self.scroll_ontime.set)
        self.lst_ontime.place(x=440, y=510)
        for ontime_member_s in ontime_member:
            self.lst_ontime.insert("end", ontime_member_s)
        self.lst_ontime.config(yscrollcommand=self.scroll_ontime.set)
        self.scroll_ontime.config(command=self.lst_ontime.yview)

        # 出缺勤名單 遲到
        self.scroll_late = tk.Scrollbar(self.frame_context9)
        self.scroll_late.place(x=678, y=511, relheight=0.0193)
        var_late = tk.StringVar()
        self.lst_late = tk.Listbox(self.frame_context9, listvariable=var_late, font=f2, width=10,
                                   height=10, yscrollcommand=self.scroll_late.set)
        self.lst_late.place(x=585, y=510)
        for late_member_s in late_member:
            self.lst_late.insert("end", late_member_s)
        self.lst_late.config(yscrollcommand=self.scroll_late.set)
        self.scroll_late.config(command=self.lst_late.yview)

        # 出缺勤名單 未出席
        self.scroll_absence = tk.Scrollbar(self.frame_context9)
        self.scroll_absence.place(x=823, y=511, relheight=0.0193)
        var_absence = tk.StringVar()
        self.lst_absence = tk.Listbox(self.frame_context9, listvariable=var_absence, font=f2, width=10,
                                      height=10, yscrollcommand=self.scroll_absence.set)
        self.lst_absence.place(x=730, y=510)
        for absence_member_s in absence_member:
            self.lst_absence.insert("end", absence_member_s)
        self.lst_absence.config(yscrollcommand=self.scroll_absence.set)
        self.scroll_absence.config(command=self.lst_absence.yview)

        # 是否完成任務名單 完成任務
        self.scroll_done = tk.Scrollbar(self.frame_context9)
        self.scroll_done.place(x=533, y=761, relheight=0.0193)
        var_done = tk.StringVar()
        self.lst_done = tk.Listbox(self.frame_context9, listvariable=var_done, font=f2, width=10,
                                   height=10, yscrollcommand=self.scroll_done.set)
        self.lst_done.place(x=440, y=760)
        for done_member_s in done_member:
            self.lst_done.insert("end", done_member_s)
        self.lst_done.config(yscrollcommand=self.scroll_done.set)
        self.scroll_done.config(command=self.lst_done.yview)

        # 是否完成任務名單 未完成任務
        self.scroll_undone = tk.Scrollbar(self.frame_context9)
        self.scroll_undone.place(x=678, y=761, relheight=0.0193)
        var_undone = tk.StringVar()
        self.lst_undone = tk.Listbox(self.frame_context9, listvariable=var_undone, font=f2, width=10,
                                     height=10, yscrollcommand=self.scroll_undone.set)
        self.lst_undone.place(x=585, y=760)
        for undone_member_s in undone_member:
            self.lst_undone.insert("end", undone_member_s)
        self.lst_undone.config(yscrollcommand=self.scroll_undone.set)
        self.scroll_undone.config(command=self.lst_undone.yview)

        # 是否完成任務名單 無任務
        self.scroll_none = tk.Scrollbar(self.frame_context9)
        self.scroll_none.place(x=823, y=761, relheight=0.0193)
        var_none = tk.StringVar()
        self.lst_none = tk.Listbox(self.frame_context9, listvariable=var_none, font=f2, width=10,
                                   height=10, yscrollcommand=self.scroll_none.set)
        self.lst_none.place(x=730, y=760)
        for none_member_s in none_member:
            self.lst_none.insert("end", none_member_s)
        self.lst_none.config(yscrollcommand=self.scroll_none.set)
        self.scroll_none.config(command=self.lst_none.yview)

        # 出缺勤圓餅圖
        labels1 = ["準時", "遲到", "未出席"]
        values1 = [ontime_num, late_num, absence_num]

        if ontime_num == 0:
            values1.remove(ontime_num)
            labels1.remove("準時")
        if late_num == 0:
            values1.remove(late_num)
            labels1.remove("遲到")
        if absence_num == 0:
            values1.remove(absence_num)
            labels1.remove("未出席")

        figure1 = plt.figure(figsize=(3, 2.5), dpi=105)
        figure1.set_facecolor(color_2)
        pie1 = plt.pie(values1, labels=labels1, pctdistance=0.6, autopct="%1.1f%%")
        plt.title("出缺勤", {"fontsize": 14})

        self.canvas1 = FigureCanvasTkAgg(figure1, self.frame_context9)
        self.canvas1.get_tk_widget().place(x=80, y=485)

        # 是否完成任務圓餅圖
        labels2 = ["完成任務", "未完成任務", "無任務"]
        values2 = [done_num, undone_num, none_num]

        if done_num == 0:
            values2.remove(done_num)
            labels2.remove("完成任務")
        if undone_num == 0:
            values2.remove(undone_num)
            labels2.remove("未完成任務")
        if none_num == 0:
            values2.remove(none_num)
            labels2.remove("無任務")

        figure2 = plt.figure(figsize=(3, 2.5), dpi=105)
        figure2.set_facecolor(color_2)
        pie2 = plt.pie(values2, labels=labels2, pctdistance=0.6, autopct="%1.1f%%")
        plt.title("是否完成指派任務", {"fontsize": 14})

        self.canvas = FigureCanvasTkAgg(figure2, self.frame_context9)
        self.canvas.get_tk_widget().place(x=80, y=740)

    def click_btn9_1(self):
        self.page9.destroy()
        try:
            wb = gc.open(str(name) + " in " + str(meeting_names[folder_location]))
            Page2()
        except (FileNotFoundError, NameError):
            Page1()


root.geometry("1000x700")
root.resizable(0, 0)
Page0(root)
root.mainloop()
