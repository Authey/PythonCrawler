# Author: Authey
# Date: 10/05/2020


import os
from bs4 import BeautifulSoup
import re
from urllib import request
from urllib import error
from urllib import parse
import xlwt
import sqlite3


def main():
    base_url = 'https://movie.douban.com/top250?start='
    movie_list = get_data(base_url)
    saved_path = './top250.xls'
    db_path = './movie.db'
    save_data(movie_list, saved_path)
    if not os.path.exists(db_path):
        init_db(db_path)
    save_db(movie_list, db_path)


# Link
find_link = re.compile(r'<a href="(.*?)">')
# Image
find_image = re.compile(r'<img alt="(.*?)" class="" src="(.*?)" width="100"/>', re.S)
# Title
find_title = re.compile(r'<span class="title">(.*?)</span>')
# Rating
find_rating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
# Number of people rated
find_num = re.compile(r'<span>(\d*)人评价</span>')
# Introduction
find_intro = re.compile(r'<span class="inq">(.*)</span>')
# Information
find_info = re.compile(r'<p class="">(.*)</p>\n<div class="star">', re.S)


def ask_url(asked_url):
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.138 Safari/537.36'
    }
    req = request.Request(url=asked_url, headers=header)
    page = ''
    try:
        response = request.urlopen(req)
        page = response.read().decode('utf-8')
    except error.URLError as e:
        if hasattr(e, 'code'):
            print(e.code)
        if hasattr(e, 'reason'):
            print(e.reason)
    return page


def get_data(init_url):
    movie_data = []
    for i in range(0, 10):
        url = init_url + str(i * 25)
        each_page = ask_url(url)
        soup = BeautifulSoup(each_page, 'html.parser')
        for item in soup.find_all('div', class_="item"):
            each_movie = {}
            item = str(item)
            each_link = re.findall(find_link, item)[0]
            each_image = re.findall(find_image, item)[0][1]
            each_title = re.findall(find_title, item)[0]
            each_rating = re.findall(find_rating, item)[0]
            each_num = re.findall(find_num, item)[0]
            each_intro = re.findall(find_intro, item)[0] if re.findall(find_intro, item) else ''
            each_info = re.findall(find_info, item)[0]
            each_info = re.sub('\n\s*', '', each_info)
            each_info = re.sub('(\xa0)+', ' ', each_info)
            each_info = re.sub('<br/>', ' ', each_info)
            each_movie['link'] = each_link
            each_movie['image'] = each_image
            each_movie['title'] = each_title
            each_movie['rating'] = each_rating
            each_movie['num'] = each_num
            each_movie['intro'] = each_intro
            each_movie['info'] = each_info
            movie_data.append(each_movie)
    return movie_data


def init_db(path):
    db_sheet = '''
        create table top250(
        id integer primary key autoincrement,
        link text,
        image text,
        title varchar,
        rating numeric,
        reviews numeric,
        intro text,
        info text
        )
    '''
    db = sqlite3.connect(path)
    db_cursor = db.cursor()
    db_cursor.execute(db_sheet)
    db.commit()
    db.close()


def save_data(movie_data, path):
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('movie data', cell_overwrite_ok=True)
    worksheet.write(0, 0, 'Link')
    worksheet.write(0, 1, 'Image')
    worksheet.write(0, 2, 'Title')
    worksheet.write(0, 3, 'Rating')
    worksheet.write(0, 4, 'Reviews')
    worksheet.write(0, 5, 'Introduction')
    worksheet.write(0, 6, 'Information')
    for i in range(0, len(movie_data)):
        worksheet.write(i + 1, 0, movie_data[i]['link'])
        worksheet.write(i + 1, 1, movie_data[i]['image'])
        worksheet.write(i + 1, 2, movie_data[i]['title'])
        worksheet.write(i + 1, 3, movie_data[i]['rating'])
        worksheet.write(i + 1, 4, movie_data[i]['num'])
        worksheet.write(i + 1, 5, movie_data[i]['intro'])
        worksheet.write(i + 1, 6, movie_data[i]['info'])
    workbook.save(path)


def save_db(movie_data, path):
    db_clear = '''
                DELETE FROM top250
            '''
    db_delete = '''
                DELETE FROM sqlite_sequence WHERE name = 'top250'
            '''
    db = sqlite3.connect(path)
    cur = db.cursor()
    cur.execute(db_clear)
    cur.execute(db_delete)
    for each_movie in movie_data:
        for index in each_movie:
            if index == 'rating' or index == 'num':
                pass
            else:
                each_movie[index] = '"' + each_movie[index] + '"'

        db_sheet = '''
            insert into top250(
            link, image, title, rating, reviews, intro, info)
            values({0})
        '''.format(','.join(each_movie.values()))
        cur.execute(db_sheet)
        db.commit()
    db.close()


if __name__ == "__main__":
    main()
