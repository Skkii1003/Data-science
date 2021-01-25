import urllib.request
import urllib.error
import urllib.parse
import xlwt
import time
from lxml import etree


def saveData(new_list, time_list, comment_list, repost_list, like_list, comment_link_list, num):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet("微博评论")
    col = ("新闻", "时间", "评论数", "转发数", "点赞数", "评论链接")

    for i in range(5, 6):
        sheet.write(0, i, col[i])   # 给excel表添加列名
    for i in range(0, len(time_list)):
        sheet.write(i+1, 0, new_list[i])
        sheet.write(i+1, 1, time_list[i])
        sheet.write(i+1, 2, comment_list[i])
        sheet.write(i+1, 3, repost_list[i])
        sheet.write(i+1, 4, like_list[i])
        sheet.write(i+1, 5, comment_link_list[i])
        # sheet.write(i+1, 4, link_list)

    book.save('./头条新闻/头条新闻微博'+str(num)+'.xls')


def askURL(url, headers):
    request = urllib.request.Request(url, headers=headers)  # 发送请求
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read() # 获取网页内容
        # print(html)
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


def get_data(html):
    parsed_html = etree.HTML(html)
    like_list = []
    repost_list = []
    comment_list = []
    news_list = []
    time_list = []
    comment_link_list = []
    try:

        div_len = len(parsed_html.xpath('//*/div[@id]'))
        path = ''
        for i in range(1, div_len):
            path = '//*/div[@id][' + str(i) + ']/div'
            news_list.append(parsed_html.xpath(path + '[1]/span[@class="ctt"]/a[1]/text()'))
            if len(parsed_html.xpath(path)) == 2:
                index_len = len(parsed_html.xpath(path + '[2]/a/text()'))
                comment_list.append(parsed_html.xpath(path + '[2]/a['+str(index_len)+']/text()'))
                repost_list.append(parsed_html.xpath(path + '[2]/a['+str(index_len-1)+']/text()'))
                like_list.append(parsed_html.xpath(path + '[2]/a['+str(index_len-2)+']/text()'))
                time_list.append(parsed_html.xpath(path + '[2]/span[@class="ct"]/text()'))
                comment_link_list.append(parsed_html.xpath(path + '[2]/a['+str(index_len)+']/@href'))
            if len(parsed_html.xpath(path)) == 1:
                index_len = len(parsed_html.xpath(path + '[1]/a/text()'))
                comment_list.append(parsed_html.xpath(path + '[1]/a['+str(index_len-1)+']/text()'))
                repost_list.append(parsed_html.xpath(path + '[1]/a['+str(index_len-2)+']/text()'))
                like_list.append(parsed_html.xpath(path + '[1]/a[' + str(index_len-3) + ']/text()'))
                time_list.append(parsed_html.xpath(path + '[1]/span[@class="ct"]/text()'))
                comment_link_list.append(parsed_html.xpath(path + '[1]/a[' + str(index_len-1) + ']/@href'))
    except AttributeError as ae:
        if hasattr(ae, "code"):
            print(ae.code)
        if hasattr(ae, "reason"):
            print(ae.reason)
    # saveData(news_list, time_list, comment_list, repost_list, like_list, comment_link_list)
    return news_list, time_list, comment_list, repost_list, like_list, comment_link_list


if __name__ == '__main__':

    headers = {
        "cookie": "_T_WM=26998148397; H5_wentry=H5; _T_WL=1; WEIBOCN_WM=3349; backURL=https%3A%2F%2Fweibo.cn; ALF=1613975296; SCF=AvXNCxHHmXBcTrDlD7aCpH4O3WQIxM7T0HJZE8moEmsc1lr7TL1qIzI8b6AdUQ7OE0MHPyH2VLYn-YiAgOuMGVo.; SUB=_2A25ND68PDeRhGeFL7lEU9C3Nwj2IHXVu8zFHrDV6PUJbktANLVbWkW1NfeGMuWI9cvRAxLp4gJy_XHuDp6htPVY7; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WFVn8iEbGIZegiDqWEhxf9F5NHD95QNSK-0SKB0eK.pWs4DqcjMi--NiK.Xi-2Ri--ciKnRi-zNeKeES0M0e05ceBtt; SSOLoginState=1611390815",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                             "Chrome/87.0.4280.141 Safari/537.36 Edg/87.0.664.75"
    }

    print("开始爬了，搁这等着")
    start_time = time.time()
    base_url = 'https://weibo.cn/breakingnews?page='
    temp1 = []
    temp2 = []
    temp3 = []
    temp4 = []
    temp5 = []
    temp6 = []
    for i in range(0, 5):
        time.sleep(10)
        for j in range(1400+i*200, 1600+i*200):
            url = base_url + str(j)
            html = askURL(url, headers)
            a, b, c, d, e, f = get_data(html)
            temp1 += a
            temp2 += b
            temp3 += c
            temp4 += d
            temp5 += e
            temp6 += f
        saveData(temp1, temp2, temp3, temp4, temp5, temp6, i)
    print('spend '+str((time.time()-start_time))+' s per 740 pages.')