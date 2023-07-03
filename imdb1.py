from bs4 import BeautifulSoup
import requests
import time
import openpyxl


headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"}

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top 250 IMDB Movies'
sheet.append(['Movie Rank','Movie Name','Year of Release','Run-time','IMDB-Rating'])

try:
    source = requests.get("https://www.imdb.com/chart/top/",headers=headers)
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text,'lxml')
    
    movies = soup.find("div", class_="sc-debe89e-2 fFnNlf ipc-page-grid__item ipc-page-grid__item--span-2").find_all("li")
    
    for movie in movies:
        main = movie.find("div",class_="sc-14dd939d-0 fBusXE cli-children")
        title_text = main.a.h3.text
        
        split_title = title_text.split(".")
        rank = split_title[0]
        title = split_title[1]
        
        data = main.find("div",class_="sc-14dd939d-5 cPiUKY cli-title-metadata")
        
        year = data.find_all("span")[0].text
        run_time = data.find_all("span")[1].text
        rating = main.find("div",class_="sc-951b09b2-0 hDQwjv sc-14dd939d-2 fKPTOp cli-ratings-container").find("span").text
        
        print(rank,title,year,run_time,rating)
        sheet.append([rank,title,year,run_time,rating])
        time.sleep(0.1)
        
        
except Exception as e:
    print("There's an error while loading...")
    
excel.save('Top 250 IMDB Movies.xlsx')