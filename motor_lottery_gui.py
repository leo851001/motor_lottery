
import tkinter as tk
import xlwings as xw
import time
from motor_lottery_logic import MotorLottery

class LotteryApp:
    def __init__(self, root, excel_path):
        self.root = root
        self.excel_path = excel_path
        self.participants = []
        self.parking_spaces = []
        self.excel_app = xw.App(visible=True)
        self.workbook = self.excel_app.books.open(self.excel_path)
        self.lottery = MotorLottery(self.participants, self.parking_spaces, workbook=self.workbook)
        self.automated = False

        # 視窗設定
        self.root.title("機車位抽籤")
        self.root.geometry("600x400")
        self.root.configure(bg="#f0f0f0")

        # 標題
        self.title_label = tk.Label(root, text="文華天際社區", font=("Arial", 30, "bold"), fg="#333333", bg="#f0f0f0")
        self.title_label.pack(pady=(20, 20))

        # 按鈕區塊
        button_frame = tk.Frame(root, bg="#f0f0f0")
        button_frame.pack(pady=10)

        # A棟抽籤＆B棟抽籤
        self.a_lottery_button = tk.Button(button_frame, text="A棟抽籤", command=self.start_a_lottery,
                                          bg="#4CAF50", fg="white", font=("Arial", 18, "bold"),
                                          width=14, relief="raised", padx=10, pady=8)
        self.a_lottery_button.grid(row=0, column=0, padx=5)

        self.b_lottery_button = tk.Button(button_frame, text="B棟抽籤", command=self.start_b_lottery,
                                          bg="#2196F3", fg="white", font=("Arial", 18, "bold"),
                                          width=14, relief="raised", padx=10, pady=8)
        self.b_lottery_button.grid(row=0, column=1, padx=5)

        # 結果顯示區塊
        self.result_frame = tk.Frame(root, bg="black", width=450, height=200, padx=15, pady=15, relief="sunken", borderwidth=4)
        self.result_frame.pack_propagate(False)
        self.result_frame.pack(pady=20)

        self.result_label = tk.Label(self.result_frame, text="抽籤準備中", font=("Arial", 50, "bold"), fg="white", bg="black", justify="left")
        self.result_label.pack(expand=True)

    # 讀取Ａ棟戶別＆車位
    def start_a_lottery(self):
        self.participants = MotorLottery.load_data_from_excel(self.excel_path, "A棟戶別", "A")
        self.parking_spaces = MotorLottery.load_data_from_excel(self.excel_path, "A棟車位", "A")
        if self.participants and self.parking_spaces:
            self.lottery.participants = self.participants
            self.lottery.parking_spaces = self.parking_spaces
            self.start_automatic_lottery()
        else:
            self.result_label.config(text="無法讀取A棟資料")

    # 讀取Ｂ棟戶別＆車位
    def start_b_lottery(self):
        self.participants = MotorLottery.load_data_from_excel(self.excel_path, "B棟戶別", "A")
        self.parking_spaces = MotorLottery.load_data_from_excel(self.excel_path, "B棟車位", "A")
        if self.participants and self.parking_spaces:
            self.lottery.participants = self.participants
            self.lottery.parking_spaces = self.parking_spaces
            self.start_automatic_lottery()
        else:
            self.result_label.config(text="無法讀取B棟資料")

    # 準備開始抽籤
    def start_automatic_lottery(self):
        self.a_lottery_button.config(state=tk.DISABLED)
        self.b_lottery_button.config(state=tk.DISABLED)
        self.automated = True
        self.result_label.config(text="抽籤進行中")
        self.root.after(1000, self.automatic_draw)

    # 開始抽籤
    def automatic_draw(self):
        if self.automated:
            self.result_label.config(text=f"戶別：\n車位：")
            self.root.update()
            time.sleep(0.5)
            participant = self.lottery.draw_participant()
            if participant:
                self.result_label.config(text=f"戶別： {participant}\n車位：")
                self.root.update()
                time.sleep(0.5)
                parking_spot = self.lottery.draw_parking_spot()
                if parking_spot:
                    self.result_label.config(text=f"戶別： {participant}\n車位： {parking_spot}")
                    self.root.update()
                    self.lottery.log_assignment()
                    self.lottery.record_to_excel(sheet_name="車位")
                    self.root.after(1000, self.automatic_draw)
                else:
                    self.result_label.config(text="所有車位皆已完成抽籤")
                    self.automated = False
                    self.a_lottery_button.config(state=tk.NORMAL)
                    self.b_lottery_button.config(state=tk.NORMAL)
            else:
                self.result_label.config(text="所有戶別皆已完成抽籤")
                self.automated = False
                self.a_lottery_button.config(state=tk.NORMAL)
                self.b_lottery_button.config(state=tk.NORMAL)

excel_path = "/Users/leohuang/Desktop/motor_lottery/文華天際-機車位抽籤.xlsx"

root = tk.Tk()
app = LotteryApp(root, excel_path=excel_path)
root.mainloop()
