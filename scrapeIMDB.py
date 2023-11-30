from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()

print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)

sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])

URL='https://www.imdb.com/chart/top/'

HEADERS=({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36', 'Accept-Langauge': 'en-US, en;q=0.5'})

try:
    source=requests.get(URL,headers=HEADERS)
    source.raise_for_status

    soup = BeautifulSoup(source.text , 'html.parser')

    movies = soup.find('ul', class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-9d2f6de0-0 iMNUXk compact-list-view ipc-metadata-list--base").find_all('li')
    
    for movie in movies:
        
        name = movie.find('h3', class_='ipc-title__text').get_text(strip=True).split('.')[1]
        
        rank = movie.find('h3', class_='ipc-title__text').get_text(strip=True).split('.')[0]
        
        year = movie.find('span', class_='sc-479faa3c-8 bNrEFi cli-title-metadata-item').text

        rating=movie.find('span', class_='ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating').text
        
        print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])
    
except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')