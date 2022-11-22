import json
import re
from random import randint
from time import sleep

import requests
import urllib3
import xlrd
from bs4 import BeautifulSoup
from scholarly import scholarly
import os
from tqdm import tqdm

# 论文数量
publication_num = 20

# 系统变量
os.environ["http_proxy"] = "http://127.0.0.1:7890"
os.environ["https_proxy"] = "http://127.0.0.1:7890"
requests.DEFAULT_RETRIES = 5
http = urllib3.PoolManager(cert_reqs='CERT_NONE')
scholarly.set_retries(5)
scholarly.set_timeout(20)
# 定义数据存储的词典
data = {}


# 清空字符串前后空白
def clear_blank(name):
    return name.strip()


# 比较两个英文名大概是否为同一个人
def compare_name(name1, name2):
    name1 = list(map(normalize, re.split("[ -·.'“\"]", clear_blank(name1))))
    name2 = list(map(normalize, re.split("[ -·.'“\"]", clear_blank(name2))))

    count = compare_list(name1, name2)
    if count >= 2:
        return True
    else:
        return False


# 比对两个list相同的元素数量
def compare_list(list1, list2):
    count = 0
    for i in list1:
        if i in list2:
            count += 1
        else:
            # 比较首字母是否相等
            for j in list2:
                if len(i) == 1 or len(j) == 1:
                    if i[0] == j[0]:
                        count += 1
                        break
    return count


# 英文名规范化
def normalize(name):
    i = 1
    length = len(name)
    if length <= 1:
        # 转大写
        return name.upper()
    ans = [0 for x in range(length)]
    for i in range(1, len(name)):
        if ord(name[0]) >= 97:
            ans[0] = chr(ord(name[0]) - 32)
        else:
            ans[0] = name[0]
        if ord(name[i]) < 97:
            ans[i] = chr(ord(name[i]) + 32)
        else:
            ans[i] = name[i]
    strans = ''
    for a in range(len(ans)):
        strans = strans + ans[a]
    return strans


# 初始化递归搜索
def search_author(names):
    names = clear_blank(names)
    # Retrieve the author's data, fill-in, and print
    # Get an iterator for the author results
    global data
    data = {}
    data['姓名'] = names
    search_query = scholarly.search_author(names)
    try:
        author_result = search_author_oneByOne(names, search_query)
        author_result_json = json.dumps(author_result, default=lambda o: o.__dict__, sort_keys=False, indent=4)
        print(author_result['name'] + "匹配成功，开始获取作者论文信息")
        data["显示姓名"] = author_result['name']
        search_author_publication(author_result, names)
    except StopIteration:
        print("没有发现该作者：" + names)


# 挨个搜索作者
def search_author_oneByOne(name, search_query):
    # Retrieve the first result from the iterator
    author_result = next(search_query)
    # scholarly.pprint(first_author_result)
    if compare_name(name, author_result['name']):
        # print(first_author_result_json)
        return author_result
    else:
        return search_author_oneByOne(name, search_query)


# 写入数据到词典
def write_data(Key1, Key2, infoKey, datas):
    global data
    try:
        if Key2 is None:
            if infoKey == "谷歌学术链接":
                data[
                    infoKey] = "https://scholar.google.com/citations?view_op=view_citation&hl=zh-cn&citation_for_view=" + \
                               datas[Key1]
            else:
                if infoKey == "文章引文":
                    data[infoKey] = get_citations("https://scholar.google.com/" + datas[Key1])
                    data["文章ID"] = datas[Key1]
                else:
                    data[infoKey] = datas[Key1]
        else:
            data[infoKey] = datas[Key1][Key2]
    except KeyError:
        return

    # 获取谷歌印文


def get_citations(url, retry=0):
    headers = {
        'authority': 'scholar.google.com',
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'cache-control': 'no-cache',
        'connection': 'close',
        'dnt': '1',
        'pragma': 'no-cache',
        'sec-ch-ua': '"Chromium";v="110", "Not A(Brand";v="24", "Google Chrome";v="110"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"',
        'sec-fetch-dest': 'document',
        'sec-fetch-mode': 'navigate',
        'sec-fetch-site': 'none',
        'sec-fetch-user': '?1',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36',
    }
    try:
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        item = soup.find_all('div', class_="gs_citr")
        return item[0].text
    except (requests.exceptions.ProxyError, requests.exceptions.SSLError, urllib3.exceptions.MaxRetryError, IndexError):
        print("连接错误")
        if retry <= 5:
            sleep(randint(1, 5))
            return get_citations(url, retry + 1)
        else:
            print("连接错误且重试失败")
            return None


def search_author_publication(author_result, name):
    # Retrieve all the details for the author
    author = scholarly.fill(author_result, publication_limit=publication_num, sections=['basics', 'publications'])
    # scholarly.pprint(author)
    # author转JSONObj
    author_json = json.dumps(author, default=lambda o: o.__dict__, sort_keys=False, indent=4)
    # print(author_json)
    count = len(author['publications'])
    print("获取出版物数量：" + str(count))
    sleep(1)
    global data
    init_data = data.copy()
    for i in tqdm(range(count), desc="获取" + name + "的出版物信息", colour="green", unit="篇"):
        data = init_data.copy()
        title = author['publications'][i]['bib']['title']
        # print(author['publications'][i])
        write_data('bib', 'title', "出版物标题", author['publications'][i])
        write_data('bib', 'pub_year', "发表年份", author['publications'][i])
        write_data('bib', 'citation', "出版商", author['publications'][i])
        write_data('num_citations', None, "被引数量", author['publications'][i])
        write_data('author_pub_id', None, "谷歌学术链接", author['publications'][i])
        search_query = scholarly.search_pubs(title)
        print("搜索文章：" + author['publications'][i]['bib']['title'])
        pubInfo = next(search_query)
        pubJson = json.dumps(pubInfo, default=lambda o: o.__dict__, sort_keys=False, indent=4)
        # print(pubJson)
        write_data('pub_url', None, "文章链接", pubInfo)
        write_data('bib', 'abstract', "文章摘要", pubInfo)
        write_data('bib', 'venue', "文章会场", pubInfo)
        write_data('url_scholarbib', None, "文章引文", pubInfo)
        if data is not {}:
            data_json = json.dumps(data, default=lambda o: o.__dict__, sort_keys=False, ensure_ascii=False)
            # print(data_json)
            append_text_to_file(data_json, "datas.json")
    # 词典转jsonObj
    sleep(1)
    print("获取作者论文完成：" + name)


# 追加文本到文件
def append_text_to_file(text, file_path):
    # 如果文件不存在就创建
    if not os.path.exists(file_path):
        open(file_path, 'w').close()
    with open(file_path, 'a', encoding='utf-8') as f:
        f.write(text)
        f.write("\n")
    # print("写入文件完成：" + file_path)


if __name__ == '__main__':
    excelData = xlrd.open_workbook(r'人员信息采集.xlsx')
    dataSheets = excelData.sheets()[0]
    row = dataSheets.nrows
    print("获取人员数量：" + str(row - 1))
    for i in tqdm(range(1, row), desc="获取作者信息", colour="yellow", unit="个"):
        name = dataSheets.cell(i, 1).value
        print("开始获取作者：" + name)
        search_author(name)
        print("获取作者完成：" + name)
        sleep(randint(1, 10))
