# -*- coding: utf-8 -*-
"""
Created on Tue Oct 22 15:52:13 2024

Version：v1.2.3

"""

import os
import csv

import matplotlib
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter
matplotlib.use('Qt5Agg') # 指定互動式後端

from datetime import datetime
from tkcalendar import DateEntry

import openpyxl
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.drawing.text import ParagraphProperties, CharacterProperties
from openpyxl.drawing.line import LineProperties

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 主程式模組---------------------------------------------------------------------------------------

# 11種線型
linetypes = ("solid","sysDash", "sysDashDot", "sysDashDotDot", "sysDot",
             "dash", "dashDot", "dot", "lgDash", "lgDashDot", "lgDashDotDot",)

# 54種循環顏色
colors = ("4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "2C4D75", "772C2A", "5F7530", "4D3B62", 
          "276A7C", "B65708", "729ACA", "CD7371", "AFC97A", "9983B5", "6FBDD1", "F9AB6B", "3A679C", "9F3B38", 
          "7E9D40", "664F83", "358EA6", "F3740B", "95B3D7", "D99694", "C3D69B", "B3A2C7", "93CDDD", "FAC090", 
          "254061", "632523", "4F6228", "403152", "215968", "984807", "84A7D1", "D38482", "B9CF8B", "A692BE", 
          "81C5D7", "F9B67E", "335A88", "8B3431", "6F8938", "594573", "2E7C91", "D56509", "A7C0DE", "DFA8A6", 
          "CDDDAC", "BFB2D0", "A5D6E2", "FBCBA3", )

# 資料類型
data_units = {"FREQ":"Frequency [Hz]", "IFB":"Current [mA]", "VFB":"Power [%]"}
data_scale = {"FREQ":500, "IFB":100, "VFB":50}

# Description
StartTrack   = "Mode/Status changed to: UD7_Stutas_StartTrack"
ModeStatus52 = "Status updated: ModeStatus=52. Errorcode=0"
StopCommand  = "User operation: User Send Stop UD7 Command."
UD7Alarm     = "UD7 Alarm"
Mode_changed = "Mode/Status changed to: UD7_Stutas_Ready"

# 讀取檔案 (路徑, 檔案類型)
def get_files_in_dir(f_path):
    # 取得目前目錄中的所有檔案名
    files = os.listdir(f_path)
    
    # 篩選出.csv結尾的檔名，並將它們儲存到清單中
    files_csv = [filename for filename in files if filename.endswith(".csv") or filename.endswith(".CSV")]

    # 合併.csv清單
    files = files_csv
    
    if len(files) == 0:
        messagebox.showerror("錯誤", "⚠️資料夾路徑下，沒有.CSV檔存在！")

    return files

# 讀取並合併資料 (路徑)
def CSV_Merge(f_path):
    try:
        # 取得文件名列表
        file_list = get_files_in_dir(f_path)
        
        # 讀取並合併資料
        Ori_data = []
        for f in file_list:
            with open(os.path.join(f_path, f), 'r', newline = '', encoding = 'utf-8') as file:
                reader = csv.reader(file)
                rows = [row for row in reader if row]
            
            if "FREQ" in rows[0] or "IFB" in rows[0] or "VFB" in rows[0]:
                # 消除開頭
                if file_list.index(f) > 0:
                    del rows[0]
                    
                # 合併資料
                Ori_data += rows
        #print(Ori_data)
    except Exception as e:
        print(e)
        
    return Ori_data

# UD7_HMI資料處理 (路徑, 起始時間, 終止時間)
def UD7_HMI(f_path, Start_time, End_time):
    # 儲存全部資料用
    all_data = []
    ws_titles = []
    
    # 儲存錯誤位置
    UD7_Error = []
    
    try:
        # 讀取並合併資料
        Ori_data = CSV_Merge(f_path)

        # 計算起始追頻點
        StartTrack_point = []
        for i1 in range(1, len(Ori_data)):
            
            # 取得資料時間
            Timestamp = datetime.strptime(Ori_data[i1][1], "%Y-%m-%d %H:%M:%S.%f")
            
            # 判斷追頻特徵資料
            if Ori_data[i1][5] == StartTrack and Timestamp >= Start_time and Timestamp <= End_time:
                StartTrack_point.append(i1)
                all_data.append([ ["Timestamp", "FREQ", "IFB", "VFB"] ])
                
                # 時間戳記做分頁名稱
                Timestamp = str(datetime.strptime(Ori_data[i1+1][1], "%Y-%m-%d %H:%M:%S.%f"))
                Timestamp = Timestamp[0:10] + str("_") + Timestamp[11:13] + str(".") + Timestamp[14:16] + str(".") + Timestamp[17:19]
                ws_titles.append(f"Track_{Timestamp}")
        
        # 分割數據
        for i2 in StartTrack_point:
            n1 = i2 + 1
            
            while n1 < len(Ori_data):
                if Ori_data[n1][5] == ModeStatus52:
                    Timestamp = datetime.strptime(Ori_data[n1][1], "%Y-%m-%d %H:%M:%S.%f")
                    VFB  = int(Ori_data[n1][6])
                    IFB  = int(Ori_data[n1][7])
                    FREQ = int(Ori_data[n1][8])
                    all_data[StartTrack_point.index(i2)].append([Timestamp, FREQ, IFB, VFB])
                    
                elif Ori_data[n1][5] == StopCommand:
                    break
                
                elif Ori_data[n1][5][0:9] == UD7Alarm:
                    # 錯誤訊息
                    Error1 = str("驅動器追頻發生錯誤：") + str(Ori_data[n1][1])
                    print(Error1)
                    UD7_Error.append(Error1)
                    break
                
                elif Ori_data[n1][5] == StartTrack:
                    # 尋找資料最後一項
                    last_One = all_data[StartTrack_point.index(i2)][-1][0]

                    # 尋找資料最後一項的下一項
                    n2 = 0
                    for j2 in range(1, len(Ori_data)):
                        Timestamp = datetime.strptime(Ori_data[j2][1], "%Y-%m-%d %H:%M:%S.%f")
                        if Timestamp == last_One:
                            n2 = j2 +1
                            break
                    
                    # 錯誤訊息
                    Error2 = str("驅動器 or HMI程式未正常關閉：") + str(Ori_data[n2][1])
                    print(Error2)
                    UD7_Error.append(Error2)
                    break
                
                elif Ori_data[n1][5] == Mode_changed:
                    Error3 = "操作模式切換，導致追頻終止：" + str(Ori_data[n1][1])
                    print(Error3)
                    UD7_Error.append(Error3)
                    break
                
                n1 += 1
                    
    except Exception as e:
        print("合併資料時，發生錯誤：", e)
    
    return all_data, ws_titles, UD7_Error

# 繪圖模組-----------------------------------------------------------------------------------------

# 設定圖表標題格式
def set_chart_title_size(chart, size = 1400):
    paraprops = ParagraphProperties()
    paraprops.defRPr = CharacterProperties(sz=size)

    for para in chart.title.tx.rich.paragraphs:
        para.pPr = paraprops
        
# 建立Excel圖表儲存位置
def Drawing_adress(n):
    adress = []
    n += 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        adress.append(chr(remainder + ord('A')))
    return ''.join(reversed(adress))

# Excel 圖表繪製 (資料/線顏色/線型/工作分頁)
def Drawing(DATA, colors, linetype, ws):
    # XY散佈圖
    chart = LineChart()
    chart.title = "Track-Test"
    set_chart_title_size(chart, size = 1400)
    chart.style = 13
    
    if len(DATA[0]) == 2:
        R = [[2, 3]]
    
    elif len(DATA[0]) >= 3:
        # 右Y軸
        chart2 = LineChart()
        chart2.y_axis.axId = 200
        chart2.y_axis.crosses = 'max'
        chart2.y_axis.majorGridlines = None  # 取消格線
        chart2.y_axis.majorTickMark  = 'out' # 刻度在外
        
        if len(DATA[0]) == 3:
            R = [[2, 3], [3, 4]]
            chart2.y_axis.title = data_units[DATA[0][2]]
            
        elif len(DATA[0]) == 4:
            R = [[2, 3], [3, 5]]
            chart2.y_axis.title = str(data_units[DATA[0][2]] + "\n" + data_units[DATA[0][3]])
    
    # 建立數據對應顏色標籤
    data_colors = []
    for i1 in range(1, len(DATA[0])):
        if DATA[0][i1] == "FREQ":
            data_colors.append(colors[0])

        elif DATA[0][i1] == "IFB":
            data_colors.append(colors[1])

        elif DATA[0][i1] == "VFB":
            data_colors.append(colors[2])
       
    # 調整Y軸上下限
    for i in range(1, len(DATA[0])):
        # 取得數據
        values = [int(row[i]) for row in DATA[1:]]
        if len(DATA) > 5:
            del values[2], values[1], values[0] # 忽略前3點數據
        
        # 取得數據基準
        scale = data_scale[DATA[0][i]]
        y_Base = lambda y: ((y // scale) + 1) * scale if (y % scale) > 0 else y
        
        # 取得數據Y軸上下限
        if DATA[0][i] == "FREQ":
            avg = y_Base(int(sum(values) / len(values)))
            min_scale = avg - scale * 2
            max_scale = avg + scale
            chart.y_axis.majorUnit = scale / 2 # 設定格線

        elif DATA[0][i] == "IFB":
            min_scale = y_Base(min(values)) - scale
            max_scale = y_Base(max(values)) + scale
        
        elif DATA[0][i] == "VFB":
            min_scale = 0
            max_scale = 120
        
        # 設定圖表上限下
        if i == 1:
            chart.y_axis.scaling.min = min_scale
            chart.y_axis.scaling.max = max_scale
            
        elif i == 2:
            chart2.y_axis.scaling.min = min_scale
            chart2.y_axis.scaling.max = max_scale
            
        elif i == 3 and DATA[0][i] == "VFB":
            chart2.y_axis.scaling.min = 0

    # X軸
    chart.x_axis.title = "Time"
    chart.x_axis.number_format = "h:mm:ss.000"
    x_values = Reference(ws, min_col = 1, max_col = 1, min_row = 2, max_row = len(DATA))

    # 左Y軸
    chart.y_axis.title = data_units[DATA[0][1]]
    chart.y_axis.majorGridlines = openpyxl.chart.axis.ChartLines() # 打開格線
    
    # 左右Y軸資料合併
    for i2 in range(len(R)):
        for y in range(R[i2][0], R[i2][1]):
            y_values = Reference(ws, min_col = y, min_row = 1, max_row = len(DATA))
            series = Series(y_values, title_from_data = True)
            line_properties = LineProperties(w = 12700, solidFill = data_colors[y - 2], prstDash = linetype[0])
            series.graphicalProperties.line = line_properties
            
            if i2 == 0:
                chart.append(series)
                
            elif i2 == 1:
                chart2.append(series)
        
    # 設定X軸標籤
    chart.set_categories(x_values)
                
    if len(DATA[0]) >= 3:
        chart += chart2
    
    # 圖表儲存位置
    adress = Drawing_adress(len(DATA[0])) + str("1") 
    
    chart.height = 15 # 設置高度
    chart.width  = 17 # 設置寬度
    
    return chart, adress

# Excel模組----------------------------------------------------------------------------------------

# Excel工作簿建立 (資料/分頁名稱/線顏色/線型)
def Excel_file(DATA, ws_titles, colors, linetypes):
    
    # 創建一個新的 Excel 工作簿
    wb = Workbook()  

    for i1 in range(len(DATA)):
        # 新增Excel分頁
        if i1 == 0:
            ws = wb.active
            ws.title = ws_titles[i1]
            
            for row in DATA[i1]:
                ws.append(row)
                
        else:
            ws = wb.create_sheet(title = ws_titles[i1])
            
            for row in DATA[i1]:
                ws.append(row)
                
        # 執行圖表繪製
        chart, adress = Drawing(DATA[i1], colors, linetypes, ws)
        ws.add_chart(chart, adress)
        
    return wb

# 檢查Excel是否成功生成 (工作簿/路徑)
def Save_Excel_file(wb, f_path):
    try:
        # 儲存檔案
        save_path = f'{f_path}/UD7_HMI_Output.xlsx'
        wb.save(save_path)
        
        # 成功訊息
        messagebox.showinfo("成功", "成功合併檔案：UD7_HMI_Output.xlsx")
         
    except Exception as e:
        print("⚠️檢查 1：讀確認UD7_HMI_Output.xlsx，檔案是否有開啟！")
        messagebox.showerror("錯誤", f"儲存檔案時，發生錯誤：\n{str(e)}\n\n>>>讀確認UD7_HMI_Output.xlsx，檔案是否有開啟！<<<")
        
# 起始/中止時間搜尋模組------------------------------------------------------------------------------

def Search_time_gap(s_date, s_hr, s_min, s_sec, e_date, e_hr, e_min, e_sec):
    # 時間
    Start_time = s_date + str(" ") + s_hr + str(":") + s_min + str(":") + s_sec
    Start_time = datetime.strptime(Start_time, "%Y-%m-%d %H:%M:%S")
    End_time   = e_date + str(" ") + e_hr + str(":") + e_min + str(":") + e_sec
    End_time   = datetime.strptime(End_time, "%Y-%m-%d %H:%M:%S")
    
    # 起始終止時間互換
    if End_time < Start_time:
        Temp_time = End_time
        End_time = Start_time
        Start_time = Temp_time
    
    return Start_time, End_time

# 預覽模組------------------------------------------------------------------------------------------

# 提取資料
def extract_data(all_data):
    timestamps = [row[0] for row in all_data[1:]]
    freq = [row[1] for row in all_data[1:]]
    ifb  = [row[2] for row in all_data[1:]]
    vfb  = [row[3] for row in all_data[1:]]
    return timestamps, freq, ifb, vfb

# GUI介面------------------------------------------------------------------------------------------

class UD7_HMI_App:
    def __init__(self, root):
        self.root = root
        self.root.title("UD7 HMI convert")
        self.window_width = 310
        self.window_height = 380
        self.root.geometry(f"{self.window_width}x{self.window_height}")  # 設置窗口大小
        self.root.resizable(False, False) # 限制視窗大小

        # 瀏覽資料夾框架
        self.f_path_frame = ttk.LabelFrame(root, text = "檔案路徑：", relief = "groove", borderwidth = 2)
        self.f_path_frame.place(x = 15, y = 10, width = 280, height = 60)
        self.f_path_frame.config(style = "Dashed.TFrame")
        
        # 創建Listbox
        self.listbox = tk.Listbox(self.f_path_frame, width = 30, height = 1)
        self.listbox.grid(row = 0, column = 0, padx = 5, pady = 5)

        # 創建瀏覽資料夾按鈕
        self.browse_button = ttk.Button(self.f_path_frame, text = "瀏覽...", command = self.browse_folder, width = 6.5)
        self.browse_button.grid(row = 0, column = 1, padx = 0, pady = 5)

        # 儲存最新選擇的資料夾路徑
        self.latest_folder = os.getcwd()
        self.update_listbox()

        # 第一行的選項
        self.frame1 = ttk.LabelFrame(root, text = "HMI驅動器數據資料：", relief = "groove", borderwidth = 2)
        self.frame1.place(x = 15, y = 80, width = 280, height = 60)
        self.frame1.config(style = "Dashed.TFrame")

        self.var_f = tk.BooleanVar(value = True) # f = 頻率 (Hz)
        self.var_c = tk.BooleanVar(value = True) # c = 電流 (mA)
        self.var_p = tk.BooleanVar()             # p = 功率 (%)
        
        self.check_f = ttk.Checkbutton(self.frame1, text = "頻率 (FREQ)", variable = self.var_f)
        self.check_c = ttk.Checkbutton(self.frame1, text = "電流 (IFB)",  variable = self.var_c)
        self.check_p = ttk.Checkbutton(self.frame1, text = "功率 (VFB)",  variable = self.var_p)

        self.check_f.grid(row = 0, column = 0, padx = 5, pady = 5)
        self.check_c.grid(row = 0, column = 1, padx = 5, pady = 5)
        self.check_p.grid(row = 0, column = 2, padx = 5, pady = 5, sticky = "w")
        
        
        # 時間搜尋框架
        self.frame_time = ttk.LabelFrame(root, text = "時間日期搜尋範圍：", relief = "groove", borderwidth = 2)
        self.frame_time.place(x = 15, y = 150, width = 280, height = 180)
        self.frame_time.config(style = "Dashed.TFrame")
        
        # 起始日期選擇器
        tk.Label(self.frame_time, text = "起始-搜尋日期：").grid(row = 0, column = 0, padx = 5, pady = 5)
        self.s_cal = DateEntry(self.frame_time,
                               date_pattern = 'yyyy-mm-dd',
                               width = 15)

        self.s_cal.grid(row = 0, column = 1, padx = 5, pady = 5, columnspan = 4, sticky = "w")
        
        # 起始時間選擇器
        tk.Label(self.frame_time, text = "起始-搜尋時間：").grid(row = 1, column = 0, padx = 5, pady = 5)

        self.s_hour_var   = tk.StringVar()
        self.s_minute_var = tk.StringVar()
        self.s_second_var = tk.StringVar()

        self.s_hour_spinbox = tk.Spinbox(self.frame_time, 
                                         from_ = 0, 
                                         to    = 23, 
                                         wrap  = True, 
                                         textvariable = self.s_hour_var, 
                                         width  = 3, 
                                         format = "%02.0f")

        self.s_minute_spinbox = tk.Spinbox(self.frame_time, 
                                           from_ = 0, 
                                           to    = 59, 
                                           wrap  = True, 
                                           textvariable = self.s_minute_var, 
                                           width  = 3, 
                                           format = "%02.0f")

        self.s_second_spinbox = tk.Spinbox(self.frame_time, 
                                           from_ = 0, 
                                           to    = 59, 
                                           wrap  = True, 
                                           textvariable = self.s_second_var, 
                                           width  = 3, 
                                           format = "%02.0f")

        self.s_hour_spinbox.grid  (row = 1, column = 1, padx = 5, pady = 5, sticky = "w")
        self.s_minute_spinbox.grid(row = 1, column = 2, padx = 5, pady = 5, sticky = "w")
        self.s_second_spinbox.grid(row = 1, column = 3, padx = 5, pady = 5, sticky = "w")

        # 分隔線
        split_line = "----------------------------------------------------"
        tk.Label(self.frame_time, text = split_line).grid(row = 2, column = 0, padx = 5, pady = 2, columnspan = 10)

        # 結束日期選擇器
        tk.Label(self.frame_time, text = "結束-搜尋日期：").grid(row = 3, column = 0, padx = 5, pady = 5)
        self.e_cal = DateEntry(self.frame_time,
                             date_pattern = 'yyyy-mm-dd',
                             width = 15)

        self.e_cal.grid(row = 3, column = 1, padx = 5, pady = 5, columnspan = 4, sticky = "w")
        
        # 結束時間選擇器
        tk.Label(self.frame_time, text = "結束-搜尋時間：").grid(row = 4, column = 0, padx = 5, pady = 5)

        self.e_hour_var   = tk.StringVar(value = 23)
        self.e_minute_var = tk.StringVar(value = 59)
        self.e_second_var = tk.StringVar(value = 59)

        self.e_hour_spinbox = tk.Spinbox(self.frame_time, 
                                         from_ = 0, 
                                         to    = 23, 
                                         wrap  = True, 
                                         textvariable = self.e_hour_var, 
                                         width  = 3, 
                                         format = "%02.0f")

        self.e_minute_spinbox = tk.Spinbox(self.frame_time, 
                                           from_ = 0, 
                                           to    = 59, 
                                           wrap  = True, 
                                           textvariable = self.e_minute_var, 
                                           width  = 3, 
                                           format = "%02.0f")

        self.e_second_spinbox = tk.Spinbox(self.frame_time, 
                                           from_ = 0, 
                                           to    = 59, 
                                           wrap  = True, 
                                           textvariable = self.e_second_var, 
                                           width  = 3, 
                                           format = "%02.0f")

        self.e_hour_spinbox.grid  (row = 4, column = 1, padx = 5, pady = 5, sticky = "w")
        self.e_minute_spinbox.grid(row = 4, column = 2, padx = 5, pady = 5, sticky = "w")
        self.e_second_spinbox.grid(row = 4, column = 3, padx = 5, pady = 5, sticky = "w")

        # 按鈕框架
        self.button_frame = ttk.Frame(root)
        self.button_frame.place(x = 20, y = 340)

        # 創建按鈕
        self.run_button     = ttk.Button(self.button_frame, text = "執行", command = self.run_action,         width = 10)
        self.drawing_button = ttk.Button(self.button_frame, text = "預覽", command = self.Matplotlib_Drawing, width = 10)
        self.close_button   = ttk.Button(self.button_frame, text = "離開", command = self.close_action,       width = 10)

        self.run_button.grid    (row = 0, column = 0, padx = 5, pady = 5)
        self.drawing_button.grid(row = 0, column = 1, padx = 5, pady = 5)
        self.close_button.grid  (row = 0, column = 2, padx = 5, pady = 5)

    # 瀏覽資料夾
    def browse_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.latest_folder = self.folder_path
            self.update_listbox()
            
            # 取代時間資料
            try:
                Ori_data = CSV_Merge(self.latest_folder)
                self.s_cal.set_date  (Ori_data[1][1][0:10])
                self.s_hour_var.set  (int(Ori_data[1][1][11:13]))
                self.s_minute_var.set(int(Ori_data[1][1][14:16]))
                self.s_second_var.set(int(Ori_data[1][1][17:19]))
                self.e_cal.set_date  (Ori_data[-1][1][0:10])
                self.e_hour_var.set  (int(Ori_data[-1][1][11:13]))
                self.e_minute_var.set(int(Ori_data[-1][1][14:16]))
                self.e_second_var.set(int(Ori_data[-1][1][17:19]) + 1)
                
            except Exception as e:
                if len(Ori_data) > 0:
                    print("錯誤發生：", e)
                    messagebox.showerror("錯誤", "數據解析發生錯誤！\n\n請檢查檔案是否正確！")

    # 清除Listbox的內容，並插入最新選擇的資料夾路徑    
    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        self.listbox.insert(tk.END, self.latest_folder)             
    
    # 執行
    def run_action(self):
        # 搜尋時間範圍
        Start_time, End_time = Search_time_gap(self.s_cal.get(), self.s_hour_var.get(), self.s_minute_var.get(), self.s_second_var.get(), 
                                               self.e_cal.get(), self.e_hour_var.get(), self.e_minute_var.get(), self.e_second_var.get())

        # 取得合併的資料
        all_data, ws_titles, UD7_Error = UD7_HMI(self.latest_folder, Start_time, End_time)
        
        try:
            # 判斷是否有追頻紀錄
            if all_data != []:
                
                # 判斷是否有選擇資料
                if self.var_p.get() == 1 or self.var_c.get() == 1 or self.var_f.get() == 1:
                    # 建立選擇資料
                    sec_data = []
                    
                    # 儲存時間資料
                    for i1 in range(len(all_data)):
                        sec_data.append([])
                        for j1 in range(len(all_data[i1])):
                            sec_data[i1].append([])
                            sec_data[i1][j1].append(all_data[i1][j1][0])
                    
                    # 選擇頻率
                    if self.var_f.get(): 
                        for i2 in range(len(all_data)):
                            for j2 in range(len(all_data[i2])):
                                sec_data[i2][j2].append(all_data[i2][j2][1])

                    # 選擇電流
                    if self.var_c.get(): 
                        for i3 in range(len(all_data)):
                            for j3 in range(len(all_data[i3])):
                                sec_data[i3][j3].append(all_data[i3][j3][2])

                    # 選擇功率
                    if self.var_p.get(): 
                        for i4 in range(len(all_data)):
                            for j4 in range(len(all_data[i4])):
                                sec_data[i4][j4].append(all_data[i4][j4][3])
                    
                    # 資料確認點
                    #print(sec_data)
                    
                    # 顯示錯誤
                    if len(UD7_Error) > 0:
                        Error_message = "錯誤項目：\n\n"
                        for i5 in range(len(UD7_Error)):
                            Error_message += str(i5 + 1) + str(". ") + UD7_Error[i5]
                            if (i5 + 1) < len(UD7_Error):
                                Error_message += str("\n\n")
                        messagebox.showerror("追頻中發生錯誤", Error_message)
                    
                    # 儲存資料
                    wb = Excel_file(sec_data, ws_titles, colors, linetypes)
                    Save_Excel_file(wb, self.latest_folder)
                    
                else:
                    print("⚠️注意：未選擇任何合併資料！")
                    messagebox.showerror("錯誤", "⚠️未選擇任何合併資料！")

            else:
                print("⚠️注意：未包含任何追頻資料！")
                messagebox.showerror("錯誤", "⚠️數據未包含任何追頻資料！")
                
        except Exception as e:
            print("Error code:", e)
    
    # 預覽繪圖
    def Matplotlib_Drawing(self):
        if self.var_p.get() or self.var_c.get() or self.var_f.get():
            # 搜尋時間範圍
            Start_time, End_time = Search_time_gap(self.s_cal.get(), self.s_hour_var.get(), self.s_minute_var.get(), self.s_second_var.get(), 
                                                   self.e_cal.get(), self.e_hour_var.get(), self.e_minute_var.get(), self.e_second_var.get())
                        
            # 取得合併的資料
            all_data, ws_titles, UD7_Error = UD7_HMI(self.latest_folder, Start_time, End_time)
            
            if len(all_data) == 0:
                print("⚠️注意：時間範圍內，未有任何資料！")
                messagebox.showerror("無法預覽", "⚠️注意：時間範圍內，未有任何資料！")
        
            # 繪製每組資料
            for i, dataset in enumerate(all_data):
                timestamps, freq, ifb, vfb = extract_data(dataset)
                
                fig, ax1 = plt.subplots(figsize = (12, 6))
                
                SHOW = [[freq, dataset[0][1], " [Hz]", str("#" + colors[0])], 
                        [ifb,  dataset[0][2], " [mA]", str("#" + colors[1])], 
                        [vfb,  dataset[0][3], " [%]",  str("#" + colors[2])]]
                
                if self.var_p.get() == 0:
                    del SHOW[2]
                    
                if self.var_c.get() == 0:
                    del SHOW[1]
                
                if self.var_f.get() == 0:
                    del SHOW[0]
                
                # 左側 Y 軸
                ax1.plot(timestamps, SHOW[0][0], label = SHOW[0][1], color = SHOW[0][3], linestyle = "-", linewidth = 1)
                ax1.set_xlabel(dataset[0][0])
                ax1.set_ylabel(str(SHOW[0][1]+SHOW[0][2]), color = SHOW[0][3])
                ax1.tick_params(axis = 'y', labelcolor = SHOW[0][3])
                ax1.grid()
                
                if len(SHOW) >= 2:
                    # 右側 Y 軸
                    ax2 = ax1.twinx()
                    ax2.plot(timestamps, SHOW[1][0], label = SHOW[1][1], color = SHOW[1][3], linewidth = 1)
                    ax2.tick_params(axis = 'y', labelcolor = SHOW[1][3])
                    ax2.set_ylabel(str(SHOW[1][1] + SHOW[1][2]), color = SHOW[1][3])
                    
                    if len(SHOW) == 3:
                        ax2.plot(timestamps, SHOW[2][0], label = SHOW[2][1], color = SHOW[2][3], linewidth = 1)
                        ax2.set_ylabel(str(SHOW[1][1] + SHOW[1][2] + "\n" + SHOW[2][1] + SHOW[2][2]), color = SHOW[1][3])
                        
                    # 圖例
                    ax2.legend(loc = "upper right")
    
                # 標題與圖例
                fig.suptitle(ws_titles[i])
                ax1.legend(loc = "upper left")

                # 時間格式化
                ax1.xaxis.set_major_formatter(DateFormatter('%Y-%m-%d %H:%M:%S'))
                plt.tight_layout()
                plt.show()
                
        else:
            print("⚠️注意：未選擇任何顯示資料！")
            messagebox.showerror("注意", "⚠️注意：未選擇任何顯示資料！")

    # 離開
    def close_action(self):
        root.destroy()

# 運行主循環
if __name__ == "__main__":
    root = tk.Tk()
    app = UD7_HMI_App(root)
    root.mainloop()

