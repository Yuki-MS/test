import bioread
import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
from datetime import datetime

class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("App_AcqFileConvert_ver-1.0")

        #　ベース画面　サイズ設定
        self.BASE_WINDOW_WIDTH = 1200    # << アプリサイズ（横幅）
        self.BASE_WINDOW_HEIGHT = 580    # << アプリサイズ（縦幅）
        self.BASE_WINDOW_POS_X = 25   # << アプリの左上位置（液晶画面の左端からの距離）
        self.BASE_WINDOW_POS_Y = 25    # << アプリの左上位置（液晶画面の上側からの距離）
        self.geometry(f"{self.BASE_WINDOW_WIDTH}x{self.BASE_WINDOW_HEIGHT}+{self.BASE_WINDOW_POS_X}+{self.BASE_WINDOW_POS_Y}")
        
        #　メインフレーム設置（ベース画面と同一サイズ）
        self.base_frame = tk.Frame(self, width=self.BASE_WINDOW_WIDTH, height=self.BASE_WINDOW_HEIGHT, bd=5, relief="ridge")
        self.base_frame.propagate(False)
        self.base_frame.pack()

        self.initial_dir = os.getcwd()

        #　１　タイトル設置【Title】
        self.clsApp01_Title_label01 = tk.Label(self.base_frame,
                                               text = "acqファイル EXCEL変換アプリ",
                                               font =  ("BIZ UDPゴシック", 20),
                                               height = 2)    # << タイトル高さ
        self.clsApp01_Title_label01.place(relx = 0.5,    # << タイトル設置の横位置（中間地点）
                                          y = 20,    # << タイトル設置の縦位置
                                          anchor = tk.N)
        
        #　２　説明文設置【Exp】
        self.clsApp02_Exp_POS_X = 50    # << 説明文タイトルの横位置
        self.clsApp02_Exp_POS_Y = 80    # << 説明文タイトルの縦位置
        self.clsApp02_Exp_SIDE_SPACE = 25    # << 説明文本文の左詰めスペース
        self.clsApp02_Exp_LINE_SPACE = 25    # << 説明文の行間
        self.clsApp02_Exp_FONT = ("BIZ UDPゴシック", 12)

        self.clsApp02_exp_list =["○ ソフト概要",
                                 "TI法で作成される「acqファイル」を、一括で「EXCELファイル」に変換します。その際、「acqファイル」自体の中身は変更されません。",
                                 "読込フォルダを指定し、変換実行ボタンをクリックすると、指定した読込フォルダ内にフォルダを作成して保存します。",
                                 "既に同名のフォルダがある場合は、上書きせずに新たにフォルダを作成して保存します。",
                                 "○ 注意点",
                                 "変換するファイルの測定条件（チャンネル数、サンプリングレート（測定ピッチ））は全て同じにしてください。",
                                 "同じ測定条件であれば、指定したフォルダ内のファイルを一括で変換可能です。"]

        for num, ind_exp in enumerate(self.clsApp02_exp_list):
            if num == 0 or num==4:
                ind_Exp_side_space = 0
            else:                
                ind_Exp_side_space = self.clsApp02_Exp_SIDE_SPACE
            ind_Exp_line_space = self.clsApp02_Exp_LINE_SPACE*num

            self.ind_Exp_label = tk.Label(self.base_frame, 
                                          text = ind_exp,
                                          font = self.clsApp02_Exp_FONT,
                                          justify = tk.LEFT)
            self.ind_Exp_label.place(x = self.clsApp02_Exp_POS_X + ind_Exp_side_space,
                                     y = self.clsApp02_Exp_POS_Y + ind_Exp_line_space)
        
        #　図形挿入（四角）
        canvas = tk.Canvas(self.base_frame, width=1100, height=200)
        canvas.create_rectangle(5, 5, 1095, 130)
        canvas.place(x=50, y=275)


        #　３　読込フォルダ設定【FldRef】
        self.clsApp03_FldRef_POS_X = 100    # << 説明の横位置
        self.clsApp03_FldRef_POS_Y = 295   # << 説明の縦位置
        self.clsApp03_FldRef_SIDE_SPACE = 100    # << 参照フォルダ表示設の、説明の横位置からの左詰めスペース
        self.clsApp03_FldRef_LINE_SPACE = 65   # << 説明文と参照ボタンの行間

        self.clsApp03_FldRef_entry_str = tk.StringVar(self, os.getcwd())

        #　３－１　ラベル１（説明文）   
        self.clsApp03_FldRef_label01 = tk.Label(self.base_frame,
                                                text="acqファイルを読み込むフォルダを指定してください。",
                                                font = ("BIZ UDPゴシック", 14,"bold", "underline"))
        self.clsApp03_FldRef_label01.place(x = self.clsApp03_FldRef_POS_X,
                                           y = self.clsApp03_FldRef_POS_Y)
        
        #　３－２　参照ボタン
        self.clsApp03_FldRef_button = tk.Button(self.base_frame,
                                                text = "参 照",
                                                font = ("BIZ UDPゴシック", 12),
                                                relief = "raised",
                                                width = 5,
                                                bd = 5,
                                                bg = "#E0E0E0",
                                                command = lambda:self.button_click_FldRef())
        self.clsApp03_FldRef_button.place(x = self.clsApp03_FldRef_POS_X,
                                          y = self.clsApp03_FldRef_POS_Y + self.clsApp03_FldRef_LINE_SPACE,
                                          anchor=tk.W)
        
        #　３－３　ラベル２（読込フォルダ表示）        
        self.clsApp03_FldRef_label02 = tk.Label(self.base_frame,
                                                textvariable = self.clsApp03_FldRef_entry_str,
                                                font = ("BIZ UDPゴシック", 11),
                                                relief = "sunken",
                                                width = 80,
                                                bd = 5,
                                                pady = 3,
                                                bg= "#D4E7F3",
                                                fg = "#333333",
                                                anchor=tk.W)
        self.clsApp03_FldRef_label02.place(x = self.clsApp03_FldRef_POS_X + self.clsApp03_FldRef_SIDE_SPACE,
                                           y = self.clsApp03_FldRef_POS_Y + self.clsApp03_FldRef_LINE_SPACE,
                                           anchor=tk.W)
        #　４　変換実行【FldRef】
        self.clsApp04_RunConv_POS_X = 50    # << 説明の横位置
        self.clsApp04_RunConv_POS_Y = 435   # << 説明の縦位置
        self.clsApp04_RunConv_SIDE_SPACE = 50    # << 変換実行ボタンの、説明の横位置からの左詰めスペース
        self.clsApp04_RunConv_LINE_SPACE = 70   # << 行間

        #　４－１　ラベル（説明文）   
        self.clsApp04_RunConv_label01 = tk.Label(self.base_frame,
                                                text="実行ボタンを押してしてください（エクセルへの変換が実行されます）",
                                                font = ("BIZ UDPゴシック", 14,"bold", "underline"))
        self.clsApp04_RunConv_label01.place(x = self.clsApp04_RunConv_POS_X,
                                            y = self.clsApp04_RunConv_POS_Y)
        
        #　４－２　変換実行ボタン
        self.clsApp04_RunConv_button = tk.Button(self.base_frame,
                                                 text = "変 換 実 行",
                                                 font = ("BIZ UDPゴシック", 16),
                                                 relief = "raised",
                                                 width = 28,
                                                 height = 2,
                                                 bd = 5,
                                                 bg = "#E0E0E0",
                                                 command = lambda:self.button_click_RunConv())
        self.clsApp04_RunConv_button.place(x = self.clsApp04_RunConv_POS_X + self.clsApp04_RunConv_SIDE_SPACE,
                                           y = self.clsApp04_RunConv_POS_Y + self.clsApp04_RunConv_LINE_SPACE,
                                           anchor=tk.W)
        
        # ５　終了ボタン【Exit】
        self.clsApp05_Exit_POS_X = self.BASE_WINDOW_WIDTH - 150  # 右下に配置（横位置）
        self.clsApp05_Exit_POS_Y = self.BASE_WINDOW_HEIGHT - 50  # 右下に配置（縦位置）

        self.clsApp05_Exit_button = tk.Button(self.base_frame,
                                            text="終 了",
                                            font=("BIZ UDPゴシック", 12),
                                            relief="raised",
                                            width=10,
                                            height=2,
                                            bd=5,
                                            bg="#E0E0E0",
                                            command=self.quit_app)
        self.clsApp05_Exit_button.place(x=self.clsApp05_Exit_POS_X,
                                        y=self.clsApp05_Exit_POS_Y,
                                        anchor=tk.CENTER)

        #　アプリの基本色
        self.tk_setPalette(background="#EBF4FA")

        self.lift()
        self.mainloop()
    
    def quit_app(self):
        """アプリを終了する処理"""
        self.destroy()

    #　ボタンクリック（参照ボタン）－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    def button_click_FldRef(self):
        dfBC_log = filedialog.askdirectory(initialdir = self.initial_dir)
        # ※ ファイル選択をキャンセルした時にパスが非表示になるのを防止
        if dfBC_log:    
            pass
        else:
            dfBC_log = self.clsApp03_FldRef_entry_str.get()
        self.clsApp03_FldRef_entry_str.set(dfBC_log)
    #　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

    #　ボタンクリック（変換実行ボタン）－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    def button_click_RunConv(self):
        #　読込フォルダに表示してあるディレクトリを取得して移動　－－－－－－－－－－－－
        try:
            self.working_dir = self.clsApp03_FldRef_entry_str.get()
        except:
            self.working_dir = self.initial_dir
        
        os.chdir(self.working_dir)
        #　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

        # acqファイルが存在するか確認
        self.file_list = [f for f in os.listdir() if ".lnk" not in f if ".acq" in f]
        if not self.file_list:
            self.show_no_file_window()
            return  # 処理を中断
            
        check_win = ChannelCheckWindow(self)
        self.wait_window(check_win)

        if check_win.result == "ok":
            self.run_conversion()
        else:
            pass
    #　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

    def show_no_file_window(self):
        """acqファイルがないときの警告ウィンドウ"""
        no_file_win = tk.Toplevel(self)
        no_file_win.title("ファイルなし")
        no_file_win.geometry("560x135")
        no_file_win.resizable(False, False)
        no_file_win.grab_set()

        label = tk.Label(no_file_win,
                        text="このフォルダにはacqファイルが存在していません。\n読込フォルダを正しく選択してください。",
                        font=("BIZ UDPゴシック", 14),
                        justify="center")
        label.pack(pady=(25, 10))

        ok_button = tk.Button(no_file_win,
                            text="OK",
                            font=("BIZ UDPゴシック", 14),
                            width=10,
                            command=no_file_win.destroy)
        ok_button.pack()
        
    #　変換実行　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    def run_conversion(self):
        progress_win = ProgressWindow(self)
        
        self.file_list = [f for f in os.listdir() if ".lnk" not in f if ".acq" in f]    # acqファイル一覧（ショートカットファイルを除く）
        self.file_name_list = [f.split(".")[0] for f in self.file_list]    # acqファイルのファイル名（拡張子なし）一覧
        self.file_count = len(self.file_name_list)    # acqファイル数

        self.data01 = bioread.read_file(self.file_list[0])    # acqファイル１番目のデータセット
        self.channel_name_list = [f.name for f in self.data01.channels]    # チャンネル名一覧（acqファイル１番目）
        self.channel_count = len(self.data01.channels)    # チャンネル数（acqファイル１番目）　※フォルダ内のファイルは、チャンネル数が一緒でないとエラー

        #　エクセルファイル保存フォルダを作成　－－－－－－－－－－－－－－－－－－－－－
        #　※ 保存フォルダがある場合は連番（(1)、(2)・・・）を作成
        self.save_folder_name = "エクセル変換ファイル（アプリ変換）"    #　保存フォルダ名
        
        if not os.path.isdir(self.save_folder_name):
            os.makedirs(self.save_folder_name)
            self.save_folder_fullpath = os.path.join(self.working_dir, self.save_folder_name)
        else:
            self.std_save_folder_name = self.save_folder_name
            for num in range(len(os.listdir())):       
                try:
                    self.save_folder_name = self.std_save_folder_name+"("+str(num+1)+")"
                    os.makedirs(self.save_folder_name)
                    self.save_folder_fullpath = os.path.join(self.working_dir, self.save_folder_name)
                    break
                except:
                    pass
        #　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

        total_3d_data = []
        each_file_data_count_max = 0

        file_data_list = []

        #　個別ファイル用フォルダを作成して個別ファイルデータを保存－－－－－－－－－－－
        # ループの先頭でリセット
        progress_win.reset_lower()
        progress_count_01 = 0

        folder_name_of_each_file = "ファイルごとのデータ"
        folder_fullpath_of_each_file = os.path.join(self.save_folder_fullpath, folder_name_of_each_file)    # 個別ファイル用の保存フォルダ（フルパス）
        os.makedirs(folder_fullpath_of_each_file)

        for each_file in self.file_list:            
            # 保存ファイル名を作成
            save_file_name = each_file.split(".")[0] + ".xlsx"    # 個別ファイル用の保存ファイル名（エクセル拡張子）
            save_file_name_fullpath = self.path_check(os.path.join(folder_fullpath_of_each_file, save_file_name))    # 個別ファイル用の保存ファイル名フルパス（エクセル拡張子）

            each_file_data = bioread.read_file(each_file)

            each_file_data_count = len(each_file_data.channels[0].data)    # 個別ファイル内のデータ数

            if each_file_data_count > each_file_data_count_max:
                each_file_data_count_max = each_file_data_count

            # 時間データを取得
            # ※タイムインデックスでデータを取得すると小数点がズレるため、サンプリングレートで割る
            each_file_time_data_np = np.arange(each_file_data_count)/each_file_data.channels[0].samples_per_second
            each_file_time_data = each_file_time_data_np.tolist()
            time_header = ["時間（sec）"]
            each_file_time_data_line = time_header + each_file_time_data

            each_file_total_data = []
            each_file_total_data.append(each_file_time_data_line)

            each_file_channel_name_list = []

            for each_file_channel_data in each_file_data.channels:
                each_file_channel_name = [each_file_channel_data.name]
                each_file_channel_data_line = each_file_channel_name + each_file_channel_data.data.tolist()
                each_file_total_data.append(each_file_channel_data_line)
                each_file_channel_name_list.append(each_file_channel_name)

            total_3d_data.append(each_file_total_data)

            each_file_total_data_pd = pd.DataFrame(each_file_total_data).T
            each_file_total_data_pd.to_excel(save_file_name_fullpath, index=False, header=False)
            
            each_file_channel_data01 = each_file_data.channels[0]
            each_file_data_list = ["",
                                   each_file.split(".")[0],
                                   len(each_file_channel_data01.data),
                                   f"{len(each_file_channel_data01.data)/each_file_channel_data01.samples_per_second}sec",
                                   f"{each_file_channel_data01.samples_per_second}Hz",
                                   f"{1/each_file_channel_data01.samples_per_second}sec"]
            each_file_data_list = each_file_data_list + each_file_channel_name_list + [""]
            file_data_list.append(each_file_data_list)

            progress_count_01 += 1
            progress = int((progress_count_01)/self.file_count*100)
            progress_win.update_lower(progress)
        
        progress_win.update_upper(33)
        #　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    
        #　チャンネル別ファイル用フォルダ作成をしてチャンネル別ファイルデータを保存－－－
        # ループの先頭でリセット
        progress_win.reset_lower()
        progress_count_02 = 0

        folder_name_of_each_channel = "チャンネルごとのデータ"
        folder_fullpath_of_each_channel = os.path.join(self.save_folder_fullpath, folder_name_of_each_channel)    # 個別ファイル用の保存フォルダ（フルパス）
        os.makedirs(folder_fullpath_of_each_channel)        
        
        total_3d_data_pd = pd.DataFrame(total_3d_data).T

        time_data_max_np = np.arange(each_file_data_count_max)/each_file_data.channels[0].samples_per_second
        time_data_max = time_data_max_np.tolist()
        time_data_max_line = time_header + time_data_max
        time_data_max_line_pd = pd.DataFrame(time_data_max_line)

        for i in range(self.channel_count):
            each_channel_data = total_3d_data_pd.loc[i+1]
            each_channel_data_pd = pd.DataFrame(list(each_channel_data)).T
            each_channel_data_pd.loc[0] = self.file_name_list    # ヘッダーをファイル名に変更
            each_channel_total_data_pd = pd.concat([time_data_max_line_pd, each_channel_data_pd], axis=1)
            each_channel_save_name_fullpath = self.path_check(os.path.join(folder_fullpath_of_each_channel, "CH"+str(i+1)+"データ.xlsx"))    # 個別ファイル用の保存ファイル名フルパス（エクセル拡張子）
            each_channel_total_data_pd.to_excel(each_channel_save_name_fullpath, index=False, header=False)
            
            progress_count_02 += 1
            progress = int((progress_count_02)/self.channel_count*100)
            progress_win.update_lower(progress)
        
        progress_win.update_upper(66)
        #　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－

        #　保存情報をまとめたファイルデータを保存－－－－－－－－－－－－－－－－－－－－
        # 先頭でリセット
        progress_win.reset_lower()
        progress_count_03 = 0

        file_data_save_name = "保存データ情報.xlsx" 
        file_data_save_fullpath = self.path_check(os.path.join(self.save_folder_name, file_data_save_name))        

        sfd_space =[[""]]

        sfd_01_01 = [["基本情報"]]
        sfd_01_02 = [["","変換したファイル数",f"{self.file_count}個"]]
        ti_test_date_list = set()
        for file in self.file_list:
            ind_creation_date = datetime.fromtimestamp(Path(file).stat().st_mtime).date()
            ti_test_date_list.add(ind_creation_date)
        ti_test_date_list = sorted(list(ti_test_date_list))
        ti_test_date_list = [ti_test_date_list[num].strftime("%Y/%m/%d") for num in range(len(ti_test_date_list))]
        sfd_01_03 = [["","試験日（ファイルが作成された日）"] + ti_test_date_list]
        
        sfd_02_01 = [["各データ一覧"]]
        sfd_02_02_add = [f"CH{i+1}名" for i in range(self.channel_count)]
        sfd_02_02 = [["","ファイル名","データ数","測定時間", "サンプリングレート", "測定ピッチ"] + sfd_02_02_add]

        save_file_data_01 = sfd_01_01 + sfd_01_02 + sfd_01_03 + sfd_space
        save_file_data_02 = sfd_02_01 + sfd_02_02 + file_data_list

        save_file_data = save_file_data_01 + save_file_data_02

        save_file_data_pd = pd.DataFrame(save_file_data)
        save_file_data_pd.to_excel(file_data_save_fullpath, index=False, header=False)
        
        progress_count_03 += 1
        progress = int((progress_count_03)/1*100)
        progress_win.update_lower(progress)

        progress_win.update_upper(100)
        #　－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－－
    
    def path_check(self, text_data):
        check01 = "\\"
        check02 = r"\\"
        check03 = "//"
        if check01 in text_data:
            text_data = text_data.replace(check01,"/")
        if check02 in text_data:
            text_data = text_data.replace(check02,"/")
        if check03 in text_data:
            text_data = text_data.replace(check03,"/")
        return text_data

class ProgressWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master  # ← アプリ本体を保持しておく
        self.title("進捗状況")
        self.geometry("600x150")
        self.resizable(False, False)

        # 上段バー（全体進捗）
        self.upper_label = tk.Label(self, text="全体進捗", font=("BIZ UDPゴシック", 12))
        self.upper_label.pack(pady=(10, 0))
        self.upper_bar = ttk.Progressbar(self, orient="horizontal", length=500, mode="determinate", maximum=100)
        self.upper_bar.pack(pady=(0, 10))

        # 下段バー（現在の処理進捗）
        self.lower_label = tk.Label(self, text="現在の処理進捗", font=("BIZ UDPゴシック", 12))
        self.lower_label.pack()
        self.lower_bar = ttk.Progressbar(self, orient="horizontal", length=500, mode="determinate", maximum=100)
        self.lower_bar.pack(pady=(0, 10))

    def update_upper(self, value):
        """上段バーの進捗を更新"""
        self.upper_bar["value"] = value
        self.upper_bar.update()

        # 100％になったら完了メッセージを表示して終了
        if value >= 100:
            self.after(300, self.show_completion_message)

    def show_completion_message(self):
        # 完了メッセージ用のサブウィンドウを作成
        self.completion_window = tk.Toplevel(self)
        self.completion_window.title("完了")
        self.completion_window.geometry("400x130")
        self.completion_window.resizable(False, False)
        self.completion_window.grab_set()  # 操作をこのウィンドウに限定

        label = tk.Label(self.completion_window,
                        text="ファイル変換が完了しました。\nOKボタンを押してください。\n操作がない場合、５秒後に自動で終了します。",
                        font=("BIZ UDPゴシック", 12),
                        justify="center")
        label.pack(pady=(20, 10))

        ok_button = tk.Button(self.completion_window,
                              text="OK",
                              font=("BIZ UDPゴシック", 12),
                              width=10,
                              command=self.manual_close)
        ok_button.pack()

        # ５秒後に自動終了
        self.completion_window.after(5000, self.auto_close)
    
    def manual_close(self):
        """OKボタンが押されたときの処理"""
        self.completion_window.destroy()
        self.close_window()
        self.master.destroy()

    def auto_close(self):
        """5秒後に自動終了する処理"""
        if self.completion_window.winfo_exists():
            self.completion_window.destroy()
            self.close_window()
            self.master.destroy()

    def update_lower(self, value):
        """下段バーの進捗を更新"""
        self.lower_bar["value"] = value
        self.lower_bar.update()

    def reset_lower(self):
        """下段バーをリセット"""
        self.lower_bar["value"] = 0
        self.lower_bar.update()

    def close_window(self):
        """サブウインドウを閉じる"""
        self.destroy()

class ChannelCheckWindow(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.title("チャンネル数チェック中")
        self.geometry("400x110")
        self.resizable(False, False)
        self.grab_set()

        label = tk.Label(self, text="チャンネル数を確認しています...", font=("BIZ UDPゴシック", 12))
        label.pack(pady=(20, 10))

        self.after(100, self.check_channels)

    def check_channels(self):
        file_list = [f for f in os.listdir() if ".lnk" not in f if ".acq" in f]

        channel_counts = []
        for f in file_list:
            try:
                data = bioread.read_file(f)
                channel_counts.append(len(data.channels))
            except:
                continue

        if len(set(channel_counts)) > 1:
            self.result = "error"
            self.show_error()
        else:
            self.result = "ok"
            self.grab_release()
            self.destroy()

    def show_error(self):
        for widget in self.winfo_children():
            widget.destroy()

        label = tk.Label(self, text="チャンネル数が一致しません。", font=("BIZ UDPゴシック", 14))
        label.pack(pady=(20, 10))

        ok_button = tk.Button(self, text="OK", font=("BIZ UDPゴシック", 14), width=10, command=self.close_window)
        ok_button.pack()

    def close_window(self):
        self.result = "error"
        self.grab_release()
        self.destroy()

if __name__ == "__main__":
    Application()