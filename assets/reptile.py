import requests
import openpyxl

# 台中科技大學的經緯度
latitude = 24.1502200889548
longitude = 120.68708626931021

# Overpass API 查詢餐廳
def get_restaurants_nearby_osm(latitude, longitude):
    overpass_url = "http://overpass-api.de/api/interpreter"
    query = f"""
    [out:json];
    node["amenity"="restaurant"](around:500,{latitude},{longitude});
    out body;
    """

    # 向 Overpass API 發送請求
    response = requests.get(overpass_url, params={'data': query})
    if response.status_code != 200:
        print("API請求失敗，請檢查網路或 Overpass API 狀態")
        return []

    data = response.json()
    restaurant_data = []
    seen_restaurants = set()  # 用於檢查是否已經存在餐廳

    # 解析 API 返回的數據，提取餐廳名稱和位置
    for element in data['elements']:
        name = element.get('tags', {}).get('name', 'N/A')  # 默認為 'N/A'，如果沒有名稱
        lat = element.get('lat')
        lon = element.get('lon')
        location = f"{lat}, {lon}"

        # 如果餐廳名稱沒有重複，則添加
        if name not in seen_restaurants:
            restaurant_data.append([name, location])
            seen_restaurants.add(name)  # 記錄已處理過的餐廳名稱
            print(f"餐廳名稱: {name}, 位置: {location}")

    return restaurant_data

# 保存結果到 Excel 文件
def save_data_to_excel(data, filename='台中科技大學附近餐廳.xlsx'):
    # 創建一個新的 Excel 工作簿
    wb = openpyxl.Workbook()
    ws = wb.active  # 獲取當前工作表
    ws.title = "餐廳資料"  # 設定工作表名稱

    # 添加標題行
    ws.append(['餐廳名稱', '位置'])

    # 寫入餐廳數據
    for row in data:
        ws.append(row)

    # 保存工作簿為 Excel 文件
    wb.save(filename=filename)
    print(f"數據已保存至 {filename}")

if __name__ == '__main__':
    # 獲取台中科技大學附近的餐廳數據
    data = get_restaurants_nearby_osm(latitude, longitude)

    # 如果有數據，將其保存到 Excel 文件
    if data:
        save_data_to_excel(data)
