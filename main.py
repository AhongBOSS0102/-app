import traceback
import random
import openpyxl
from kivy.lang import Builder
from kivymd.app import MDApp

# 從 Excel 讀取餐廳名稱和位置
def read_restaurants_from_excel(file_path):
    restaurants = []
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active  # 獲取當前工作表

        # 跳過第一行（標題行），從第二行開始讀取
        for row in sheet.iter_rows(min_row=2, values_only=True):
            name = row[0]  # 餐廳名稱
            location = row[1]  # 餐廳位置
            if name and location:
                restaurants.append(f"{name} ({location})")  # 將餐廳名稱和位置放在一起
    except Exception as e:
        print("Error reading Excel file")
        traceback.print_exc()

    return restaurants

# 使用資料夾內的台中科技大學附近餐廳.xlsx
restaurants = read_restaurants_from_excel("assets/台中科技大學附近餐廳.xlsx")

class LunchApp(MDApp):
    def build(self):
        try:
            return Builder.load_file("yizhong.kv")
        except Exception as e:
            print("Error loading .kv file")
            traceback.print_exc()
            return None

    def pick_restaurant(self):
        try:
            result = random.choice(restaurants)
            self.root.ids.result_label.text = f"今天吃：{result}"
        except Exception as e:
            print("Error in pick_restaurant method")
            traceback.print_exc()

if __name__ == "__main__":
    try:
        LunchApp().run()
    except Exception as e:
        print("Error running the app")
        traceback.print_exc()
