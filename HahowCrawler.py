#使用request發送請求去抓取資料
import requests
#將資料轉成excel
import pandas as pd

#使用陣列儲存要爬的所有資料
course_list = []

for index in range(44):
    #要爬的目標-Hahow的課程url
    url = "https://api.hahow.in/api/products/search?category=COURSE&filter=PUBLISHED&limit=24&page="
    url = url + str(index) + "&sort=TRENDING"

    #使用Hearder去模仿使用者取得資料
    headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }

    #使用requests取得Json檔
    response = requests.get(url, headers=headers)
    #若狀態碼為200代表成功
    if response.status_code == 200:
        #將資料轉成json
        data = response.json()
        products = data['data']['courseData']['products']
        #將資料用陣列儲存起來
        for product in products:
            #資料包含 課程名稱、評分、價格、和購課人數
            course_data = [
                product['title'],
                product['averageRating'],
                product['price'],
                product['numSoldTickets']
            ]
            #將資料加入陣列
            course_list.append(course_data)
        print(str(index + 1) + " page saved successfully!")
    else:
        print("無法印出網頁")
print("All page saved successfully!")

df = pd.DataFrame(course_list, columns = ["課程名稱","評分","價格","購課人數"])
#使用excel套件OpenPyXl
df.to_excel('result.xlsx', index=False, engine="openpyxl")