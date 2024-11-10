import random
import openpyxl

class MotorLottery:
    # 初始化設定
    def __init__(self, participants, parking_spaces, workbook=None):
        self.participants = participants
        self.parking_spaces = parking_spaces
        self.assignments = []
        self.current_participant = None
        self.current_parking = None
        self.workbook = workbook

    # 抽取戶別
    def draw_participant(self):
        if not self.participants:
            return None
        self.current_participant = random.choice(self.participants)
        self.participants.remove(self.current_participant)
        return self.current_participant

    # 抽取車位
    def draw_parking_spot(self):
        if not self.parking_spaces:
            return None
        self.current_parking = random.choice(self.parking_spaces)
        self.parking_spaces.remove(self.current_parking)
        self.assignments.append((self.current_participant, self.current_parking))
        return self.current_parking

    # 紀錄抽籤結果
    def log_assignment(self, filename="lottery_log.txt"):
        if self.current_participant and self.current_parking:
            with open(filename, "a") as log_file:
                log_file.write(f"{self.current_participant} -> {self.current_parking}\n")
            print(f"Logged: {self.current_participant} -> {self.current_parking}")
        return f"Logged: {self.current_participant} -> {self.current_parking}"

    # 從excel讀取戶別＆車位
    def load_data_from_excel(file_path, sheet_name, column="A"):
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook[sheet_name]
        data = []
        for cell in sheet[column]:
            if cell.value is not None:
                if isinstance(cell.value, float):
                    data.append(str(int(cell.value)))
                else:
                    data.append(str(cell.value))
        workbook.close()
        return data

    # 從excel讀取對照表欄位
    def lookup_cell_location(self, excel_path, parking_space, lookup_sheet="對照表"):
        workbook = openpyxl.load_workbook(excel_path, data_only=True)
        sheet = workbook[lookup_sheet]
        cell_location = None

        try:
            cell_value = sheet[f"A{parking_space}"].value
            if cell_value:
                cell_location = str(cell_value)
        except KeyError:
            print(f"No entry found for parking space {parking_space} in {lookup_sheet}.")

        workbook.close()
        return cell_location

    # 即時寫入戶別至對應的車位
    def record_to_excel(self, sheet_name="車位"):
        if self.current_participant and self.current_parking:
            cell_location = self.lookup_cell_location(self.workbook.fullname, self.current_parking)
            if not cell_location:
                print(f"Could not find cell location for parking space {self.current_parking}")
                return

            sheet = self.workbook.sheets[sheet_name]
            sheet[cell_location].value = self.current_participant
