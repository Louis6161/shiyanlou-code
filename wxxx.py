# 导入交互界面包
import PySimpleGUI as sg
# 导入requests包
import requests
# 导入bs4包 # 可以用pip install bs4安装
from bs4 import BeautifulSoup
# 导入pandas包，做数据存储
import pandas as pd
# 导入xlwt包，创建excel # 用于将数据写入xls
import xlwt
# 解析数据的包用lxml
from lxml import etree
# 导入生成随机数的包
import random
# 导入时间包
import time

# 创建表格
wb = xlwt.Workbook()
# 创建表单sheet
sh = wb.add_sheet("数据")
# 在0行0列写入标题
sh.write(0, 0, '标题')
# 在0行1列写入作者
sh.write(0, 1, '作者')
# 在0行2列写入摘要
sh.write(0, 2, '摘要')
# 在0行3列写入DOI
sh.write(0, 3, 'DOI')

# 构造请求头
# 作用:
# 伪装爬虫代码
# cookie:自己的一些信息
# User-Agent:伪装成一个浏览器发起的请求
# 都可以从浏览器直接复制黏贴
headers = {
    "Cookie": 'pm-csrf=Qi3P9ivuKbvkq1DE72nwoDjUt24jKPh6pgyCFAn6JMJ9LCBbR5WwAvr59EZTSedf; pm-sessionid=9ozg2l670zuthgb5ldepbhnijz12cpjj; ncbi_sid=9208389436DB1AE3_23301SID; _gid=GA1.2.1265363938.1668147239; pm-sb=pubdate; pm-ps=20; pm-sid=Vvv_k13q1TR0FLitHAgxNw:eafdbdbb5ee90dff7eefcae351162f09; pm-adjnav-sid=c7qDaHKKosduBY9q2apGjQ:58520709e0075042969bf88deab8a0a6; _gat_ncbiSg=1; pm-iosp=; _ga_DP2X732JSX=GS1.1.1668147238.1.1.1668147310.0.0.0; _ga=GA1.2.1747509766.1668147239; _gat_dap=1; ncbi_pinger=N4IgDgTgpgbg+mAFgSwCYgFwgAwDEDsAQtiQIyEBs+AnABynYCsJJAzK6QCwBMAwt9UIBBapwB0pMQFs4tEABoQAVwB2AGwD2AQ1QqoADwAumUN0zglAIylR0i1ubBWbdkJ3MBnKFogBjRNAeSmrGiozmCiCkpOZmitzY5nhELORUdAzMLOxcfALCohLSspHcMVhO1rYYlS4YXj7+gcGGGAByAPJtAKKlZhXOtmIqvpbIw2pSw8iIYgDmGjCl1OakjKwUkayJWLQU7vblUeub9v0gtNwHIOzmAGZaal5b7liGEEpQW3JYWytYpFoNDibh2IG4jFIFAcik4Diw2DErHwYkSsNeynU2l0BlCbnCCMi63MnC4RIx+G4IMYm0JYXwSUiFCOjHo4UUFFpIAafkQIAAvvygA==',
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36 SLBrowser/8.0.0.3161 SLBChan/33",
}
# GUI交互代码
list_sizes = ['10', '20', '50', '100', '200']
list_sorts = ['Best match', 'Most recent',
              'Publication date', 'First author', 'Journal']

# 定义布局，确定行数
layout = [
    [sg.Text('关键词', size='60', font=('黑体', 15)), sg.InputText(
        'lung cancer', key='str', size=(30, None), font=('黑体', 15))],
    [sg.Text('页数', size='60', font=('黑体', 15)),
     sg.InputText('5', key='page', size=(30, None), font=('黑体', 15))],
    [sg.Text('每页数量',  size='60', font=('黑体', 15)), sg.Drop(
        list_sizes, key='num', size=(30, None), readonly=True, font=('黑体', 15))],
    [sg.Text('排序方式',  size='60', font=('黑体', 15)), sg.Drop(
        list_sorts, key='sort', readonly=True, size=(30, None), font=('黑体', 15))],
    [sg.Button('确定', font=('黑体', 10), expand_x=True),
     sg.Button('取消', font=('黑体', 10), expand_x=True)],
]
# 创建窗口
window = sg.Window('设置检索基本信息', layout)

# 事件循环
while True:
    event, values = window.read()
    if event == None:
        break
    if event == '确定':
        print(values)
        strings = values['str']
        pages = values['page']
        nums = values['num']
        sorts = values['sort']
        break
    if event == '取消':
        break
# window.close()
print(type(nums))
print(nums)
print(int(pages) > int(nums))


def get_data():
    # 接受输入的关键词
    # 作用:构建url
    # strings = input('请输入关键词:')
    # pages = input('请输入页数:')
    # nums = input('请选择每页论文数:')
    # sorts = input('请选择论文排序方式:')
    # 定义变量jj，用于控制写入表格的行数
    jj = 0
    # 1861
    # 循环控制页数，共5页
    # 作用：用于构造url，用于爬虫代码的翻页
    for oo in range(1, int(pages)+1):
        # 构造url
        url = f'https://pubmed.ncbi.nlm.nih.gov/?term={strings}&format=abstract&sort={sorts}&size={nums}&page={oo}'
        # requests发送请求
        resp = requests.get(url, headers=headers)
        # 生成随机数
        random3 = random.uniform(1, 2)    # 随机数
        # 让程序休息1~2秒
        time.sleep(random3)
        # 用bs4解析数据
        soup = BeautifulSoup(resp.text, 'html.parser')
        # css选择器提取数据
        biaoti = soup.select('article.article-overview')
        # 循环提取数据
        # 共20条数据
        for u in range(len(biaoti)):
            # 给相关变量赋初始值
            zuozher = ''
            DOI = ''
            is_m = ''
            biaoti_mak = ''
            # 每循环一次，jj+1,对应这xls的行数加一
            jj += 1
            # 捕获异常语句
            try:
                # 提取数据
                zuozher = soup.select('article.article-overview')[u].select('header.heading > div.full-view > div.inline-authors')[0].text.replace(
                    '\n', '').replace('1', '').replace('2', '').replace('3', '').replace('4', '').replace('5', '').replace('6', '')
            except:

                pass
            try:
                # 提取DOI
                DOI1 = soup.select(
                    'article.article-overview')[u].select('header.heading > div.full-view > ul.identifiers > li')  # 循环判断
                for g in DOI1:
                    # 循环判断哪个是DOI
                    if 'DOI:' in g.text:
                        # 给DOI赋值
                        DOI = g.text.replace('\n', '').replace('DOI:', '')
            except:
                pass
            try:
                # 提取摘要
                is_m = soup.select(
                    'article.article-overview')[u].select('div.abstract > div.abstract-content.selected')[0].text.replace('\n', '')
            except:
                pass
            try:
                # 提取标题
                biaoti_mak = soup.select(
                    'article.article-overview')[u].select('header.heading > div.full-view > h1.heading-title')[0].text.replace('\n', '')
            except:
                pass
            # 将数据写入表格
            sh.write(jj, 0, biaoti_mak)
            sh.write(jj, 1, zuozher)
            sh.write(jj, 2, is_m)
            sh.write(jj, 3, DOI)
            print(DOI)
    # 保存数据到对应的文件
    wb.save(r'数据.xls')


if __name__ == '__main__':
    get_data()

window.close()
