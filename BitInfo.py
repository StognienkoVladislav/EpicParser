#!/usr/bin/env python3

#Импортируем нужные нам библиотеки для проекта
import csv
import requests
from bs4 import BeautifulSoup


def get_html(url):
    #Вытягиваем html
    r = requests.get(url)
    return r.text

#def get_total_pages(html):
#    soup = BeautifulSoup(html, 'lxml')
#
#    pages = soup.find("div", class_="pagination-pages").find_all('a', class_ = "pagination-page")[-1].get("href")
#    total_pages = pages.split("=")[1].split("&")[0]

#    return int(total_pages)


def write_csv(data):
    with open('cripto.csv', 'a') as f:
        writer = csv.writer(f)
        #Записываем в csv относительно наших ключей
        writer.writerow((data['title'],
                         data['price'],
                         data['for_12_hours'],
                         data['for_7_days'],
                         data['cap'],
                         data['count_of_cap'],
                         data['change'],
                         ))



def get_page_data(html):
    #Библиотека для парсинга
    soup = BeautifulSoup(html, "lxml")

    #Парсим таблицу по строкам
    ads = soup.find("table", class_ = 'table-bordered').find_all('tr', class_ = 'ptr')
    #Парсим по кускам, потом ,в конце for, будем собирать
    for ad in ads:

        #Title of crypto-currency
        try:
            title = ad.find_all("td")[0].find("span").find("a").text.strip()

        except:
            title = ""


        #Price
        try:
            price = ad.find_all("td", attrs = {'data-val'} )[1].find("a", class_ = "conv_cur").text.strip()

        except:
            price = ""

        #in 12 hours
        try:
            for_12_hours = ad.find_all("td")[1].find("span", class_ = "text-error").find("b").text.strip()

        except:
            for_12_hours = ""


        #in 7 days
        try:
            for_7_days = ad.find_all("td")[1].find("span", class_ = "text-success").find("b").text.strip()

        except:
            for_7_days = ""



        #Capitalization
        try:
            cap = ad.find_all("td")[3].find("a").find("span").text.strip()

        except:
            cap = ""


        try:
            count_of_cap = ad.find_all("td")[3].find("span").find("b").text.strip()

        except:
            count_of_cap = ""


        #Count of change
        try:
            change = ad.find_all("td")[4].find_all("span")[2].find("span").text.strip()

        except:
            change = ""




        data = {'title'       : title,
                'price'       : price,
                'for_12_hours': for_12_hours,
                'for_7_days'  : for_7_days,
                'cap'         : cap,
                'count_of_cap': count_of_cap,
                'change'      : change}

        write_csv(data)


def main():
    #Урл из которого мы парсим всю нужную нам информацию
    url = "https://bitinfocharts.com/ru/markets/"

    #total_pages = get_total_pages(get_html(url))

    #Парс всех страниц total_pages(если бы было больше страниц)
    html = get_html(url)
    get_page_data(html)


if __name__ == '__main__':
    #Просто вызов main()
    main()