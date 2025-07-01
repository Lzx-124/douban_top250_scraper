# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random

def get_html(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    }
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        response.encoding = response.apparent_encoding
        return response.text
    except requests.RequestException as e:
        print(f"请求URL时出错: {url}, 错误: {e}")
        return None

def parse_page(html):
    movie_list = []
    soup = BeautifulSoup(html, 'html.parser')
    container = soup.find('div', class_='article')
    if not container:
        print("页面结构不符，未找到电影主容器。")
        return []

    ol_tag = container.find('ol', class_='grid_view')
    if not ol_tag:
        print("页面中未找到电影列表（ol.grid_view）。")
        return []

    movie_items = ol_tag.find_all('li')
    for item in movie_items:
        movie_info = {}
        try:
            movie_info['No'] = item.find('em').text.strip()

            title_tag = item.find('div', class_='hd')
            titles = [t.text.strip() for t in title_tag.find_all('span', class_='title')]
            other = title_tag.find('span', class_='other')
            if other:
                titles.append(other.text.strip())
            movie_info['title'] = ' / '.join(titles)

            p_tag = item.find('div', class_='bd').find('p')
            p_text = p_tag.get_text().strip()
            lines = [line.strip() for line in p_text.split('\n') if line.strip()]
            line1 = lines[0]
            line2 = lines[1] if len(lines) > 1 else ''

            director = line1.split('主演:')[0].replace('导演:', '').strip()
            movie_info['director'] = director

            parts = line2.split('\xa0/\xa0')
            movie_info['year'] = parts[0].strip() if len(parts) > 0 else 'N/A'
            movie_info['country'] = parts[1].strip() if len(parts) > 1 else 'N/A'
            movie_info['genre'] = parts[2].strip() if len(parts) > 2 else 'N/A'

            rating = item.find('span', class_='rating_num')
            movie_info['rating'] = rating.text.strip() + '分' if rating else 'N/A'

            quote = item.find('span', class_='inq')
            movie_info['quote'] = quote.text.strip() if quote else 'N/A'

            movie_list.append(movie_info)

        except Exception as e:
            print(f"解析出错，跳过该项：{e}")
            continue

    return movie_list

def save_to_excel(data_list, filename):
    if not data_list:
        print("没有收集到任何电影信息，无法生成Excel文件。")
        return
    df = pd.DataFrame(data_list)
    columns_order = ['No', 'title', 'director', 'year', 'country', 'genre', 'rating', 'quote']
    df = df.reindex(columns=[col for col in columns_order if col in df.columns])
    try:
        df.to_excel(filename, index=False)
        print(f"数据成功保存到: {filename}")
    except Exception as e:
        print(f"保存失败: {e}")

if __name__ == '__main__':
    print("开始爬取豆瓣Top250电影...")
    all_movies = []
    base_url = "https://movie.douban.com/top250?start="
    max_pages = 10
    print(f"目标页数：{max_pages}")

    for page in range(max_pages):
        url = base_url + str(page * 25)
        print(f"正在爬取第 {page+1} 页: {url}")
        html = get_html(url)
        if html:
            movie_on_page = parse_page(html)
            if movie_on_page:
                all_movies.extend(movie_on_page)
                print(f"  -> 成功解析到 {len(movie_on_page)} 部电影")
            else:
                print(f"  -> 第 {page+1} 页没有提取到电影数据")
        else:
            print(f"  -> 第 {page+1} 页请求失败")
        time.sleep(random.uniform(3, 5))

    save_to_excel(all_movies, '豆瓣Top250电影.xlsx')
    print(f"\n爬取完成，共收集 {len(all_movies)} 部电影。")
