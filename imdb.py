from bs4 import BeautifulSoup
import openpyxl as op
import requests

excel = op.Workbook()
sheet = excel.active
sheet.title = 'Top rated movies'
sheet.append(['Rank', 'Name', 'Year', 'IMDB'])

try:
    source = requests.get("https://www.imdb.com/chart/top/")
    source.raise_for_status() #error in case the website cant be accessed
    soup = BeautifulSoup(source.text,'html.parser')
    movies = soup.find('tbody',class_="lister-list").find_all('tr')

    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split(".")
        year = movie.find('span', class_="secondaryInfo").text.strip("()")
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        row = [rank, name, year, rating]
        print(row)
        sheet.append(row)
except Exception as e:
    print(e)
excel.save('IMDB Top 250 movies.xlsx')


"""
Intersting links:
openpyxl - adjust column width size -- https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size




"""


"""print(name2[0])
        print(rank[0])
        print(name)
        print(year)
        name2 = rank[1].split("(") #option that I came on by myown playing around with the strings
        year2 = name2[1].split(")") / or .strip(")")"""

