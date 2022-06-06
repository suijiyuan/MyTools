#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas
import requests
import random
import bs4
import time
import numpy
import sys

# -- config --

user = "Michael"

count = 1

file_path_pattern = r"C:\Users\%s\Desktop\tmp\word_%s.xlsx"

url_pattern = "https://www.iciba.com/word?w=%s"

user_agent_pc = [
    'Mozilla/5.0.html (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.html.2171.71 Safari/537.36',
    'Mozilla/5.0.html (X11; Linux x86_64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.html.1271.64 Safari/537.11',
    'Mozilla/5.0.html (Windows; U; Windows NT 6.1; en-US) AppleWebKit/534.16 (KHTML, like Gecko) Chrome/10.0.html.648.133 Safari/534.16',
    'Mozilla/5.0.html (Windows NT 6.1; WOW64; rv:34.0.html) Gecko/20100101 Firefox/34.0.html',
    'Mozilla/5.0.html (X11; U; Linux x86_64; zh-CN; rv:1.9.2.10) Gecko/20100922 Ubuntu/10.10 (maverick) Firefox/3.6.10',
    'Mozilla/5.0.html (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.html.2171.95 Safari/537.36 OPR/26.0.html.1656.60',
    'Mozilla/5.0.html (compatible; MSIE 9.0.html; Windows NT 6.1; WOW64; Trident/5.0.html; SLCC2; .NET CLR 2.0.html.50727; .NET CLR 3.5.30729; .NET CLR 3.0.html.30729; Media Center PC 6.0.html; .NET4.0C; .NET4.0E; QQBrowser/7.0.html.3698.400)',
    'Mozilla/5.0.html (Windows NT 5.1) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.html.963.84 Safari/535.11 SE 2.X MetaSr 1.0.html',
    'Mozilla/5.0.html (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/30.0.html.1599.101 Safari/537.36',
    'Mozilla/5.0.html (Windows NT 6.1; WOW64; Trident/7.0.html; rv:11.0.html) like Gecko',
    'Mozilla/5.0.html (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/38.0.html.2125.122 UBrowser/4.0.html.3214.0.html Safari/537.36',
]

result_pattern = "%s %s; "

# -- function --

def get_user_agent_pc():
    return random.choice(user_agent_pc)

def translate(word):
    url = url_pattern % word

    session = requests.session()
    session.keep_alive = False
    session.headers = {"User-Agent": get_user_agent_pc()}
    get_result = session.get(url)

    if get_result is None:
        return -1, None

    if get_result.status_code != 200:
        return get_result.status_code, None

    soup = bs4.BeautifulSoup(get_result.content, "lxml")

    if soup is None:
        return 200, None

    translate_list = soup.find(attrs={"class": "Mean_part__UI9M6"})

    if translate_list is None:
        return 200, None

    translate_li_list = translate_list.find_all("li")

    if translate_li_list is None:
        return 200, None

    result = ""
    for translate_li_item in translate_li_list:
        if translate_li_item is None:
            return 200, None
        word_type_item = translate_li_item.find("i")
        if word_type_item is None:
            return 200, None
        word_type = translate_li_item.find("i").text
        word_translate_list = translate_li_item.find_all("span")
        word_translate_string = ""
        for word_translate_item in word_translate_list:
            if word_translate_item is None:
                return 200, None
            word_translate_string += word_translate_item.text
        result += result_pattern % (word_type, word_translate_string)

    return 200, result


def swap_rows(data, left, right):
    left_item, right_item = data.iloc[left,
                                      :].copy(), data.iloc[right, :].copy()
    data.iloc[left, :], data.iloc[right, :] = right_item, left_item
    return data


# -- main --

if __name__ == "__main__":
    translate_success_count = 0
    translate_fail_count = 0
    translate_4xx_count = 0

    while True:
        print("start-%s" % count)
        excel_file_path = file_path_pattern % (user, count)
        data = pandas.read_excel(excel_file_path)
        data_length = len(data)
        for index, row in data.iterrows():
            if not pandas.isnull(row["t"]):
                continue

            word = row["w"]
            if pandas.isnull(word):
                continue

            print("translating... %s/%s, success count: %s, fail count: %s, 4xx count: %s" %
                  (index + 1, data_length, translate_success_count, translate_fail_count,  translate_4xx_count))
            time.sleep(random.randint(1, 3))
            translate_code, word_translate = translate(word)
            if translate_code != 200:
                translate_4xx_count += 1
            elif word_translate is None:
                translate_fail_count += 1
            else:
                translate_success_count += 1
                data.loc[index, "t"] = word_translate

        for index in range(data_length):
            swap_rows(data, index, random.randint(0, data_length - 1))

        for index in range(data_length):
            data.loc[index, "c1"] = index + 1
            data.loc[index, "c2"] = index + 1
            data.loc[index, "a"] = numpy.nan

        data = data[["c1", "w", "a", "c2", "t"]]
        print(data)

        target_excel_file_path = file_path_pattern % (user, count + 1)
        data.to_excel(target_excel_file_path, index=False)

        if translate_4xx_count != 0:
            print("translate 4xx count is %s, try again" % translate_4xx_count)
            count += 1
            translate_fail_count = 0
            translate_4xx_count = 0
        elif translate_fail_count != 0:
            print("translate fail count is %s, please check file" %
                  translate_fail_count)
            sys.exit(1)
        else:
            print("finish")
            sys.exit(0)
