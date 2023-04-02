#Import all necessary modules for the web scraping application

from bs4 import BeautifulSoup
import requests, openpyxl

#Create excel workbook to store the links and data gathered by the web scraping application
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top rated movies"
print(excel.sheetnames)
#We want to extract the movie rank, name of the movie as well as the year and rating
sheet.append(["Movie Rank","Movie Name","Year of release","IMDB Rating"])


try:
    #Connect to the website through the GET method
   source = requests.get("https://www.imdb.com/chart/top")
   source.raise_for_status()
    #We want to extract the source code of the website to extract the links from.

   soup = BeautifulSoup(source.text,'html.parser')
    #In the source code we want to find where the information about the movies is stored, so we find the tag that contains the list and give it the variable movies.

   movies = soup.find("tbody",class_="lister-list").find_all('tr')
   for movie in movies:
       #This loop will iterate over the list of moves, collecting its title, rating, year of release as well as rating.
       name = movie.find("td",class_ = "titleColumn").a.text
       rank = movie.find("td",class_="titleColumn").get_text(strip = True).split(".")[0]
       year = movie.find("td",class_ = "titleColumn").span.text.strip("()")
       rating = movie.find("td",class_ = "ratingColumn imdbRating").strong.text


       print(rank,name,year,rating)

       sheet.append([rank,name,year,rating])

except Exception as e:
    print(e)

excel.save("IMDB Movie Rating.xlsx")















