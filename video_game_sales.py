import pandas as pd
import os

pwd = os.getcwd()

data = pd.read_csv(pwd + "\\vgsales.csv") # Read CSV file into a dataframe


data['Global_Sales'] = data['Global_Sales'] * 1000000
data['NA_Sales'] = data['NA_Sales'] * 1000000
data['EU_Sales'] = data['EU_Sales'] * 1000000
data['JP_Sales'] = data['JP_Sales'] * 1000000
data['Other_Sales'] = data['Other_Sales'] * 1000000
#Global Sales of the years
gs = data.groupby(['Year'])['Global_Sales'].sum()

#Sales by Genre Global
genre = data.groupby(['Genre'])['Global_Sales'].sum()

#Sales by Genre Rest of the World
genre_OS = data.groupby(['Genre'])['Other_Sales'].sum()

#Sales by Genre USA
US_genre = data.groupby(['Genre'])['NA_Sales'].sum()

#Sales by Genre EU
EU_genre = data.groupby(['Genre'])['EU_Sales'].sum()

#Sales by Genre Japan
JP_genre = data.groupby(['Genre'])['JP_Sales'].sum()

#Sales by Genre each Year
year_genre = data.groupby(['Year', 'Genre'])['Global_Sales'].sum()





#Export dataset to excel file 
with pd.ExcelWriter(pwd + "\\GS_total_years.xlsx") as writer:
    gs.to_excel(writer, sheet_name ='Global Sales by Year' )
    genre.to_excel(writer, sheet_name ='Total Sales by Genre')
    genre_OS.to_excel(writer, sheet_name = 'Rest of The World Total Sales by Genre')
    US_genre.to_excel(writer, sheet_name = 'USA Total Sales by Genre' )
    EU_genre.to_excel(writer, sheet_name = 'EU Total Sales by Genre' )
    JP_genre.to_excel(writer, sheet_name = 'Japan Total Sales by Genre' )
    year_genre.to_excel(writer, sheet_name = 'Genre Global Total Sales Years' )
    