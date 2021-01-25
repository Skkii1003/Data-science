import urllib.request
import urllib.error
import urllib.parse
import xlwt
import time
import xlrd
from lxml import etree
import eventlet


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
    id_list = []
    comment_list = []
    like_list = []
    # div_len = 0
    try:
        # div_len = len(parsed_html.xpath('//*/div[@id]'))
        # path = ''
        for i in range(3, 6):#TODO 记得改参数(已改）
            id_list.append(parsed_html.xpath('//*/div[@id]['+str(i)+']/@id'))
            like_list.append(parsed_html.xpath('//*/div[@id]['+str(i)+']/span[@class="cc"]/a[1]/text()'))
            comment_list.append(parsed_html.xpath('//*/div[@id]['+str(i)+']/span[@class="ctt"]/text()'))
    except AttributeError as ae:
        print("ae error")
    finally:
        return id_list, like_list, comment_list


def save_data(mark_list, time_list, id_list, like_list, comment_list, xls_name):
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet(xls_name)
    col = ('标记', 'id', '时间', '点赞', '评论')
    for i in range(0, 5):
        sheet.write(0, i, col[i])
    for j in range(1, len(id_list)+1):
        sheet.write(j, 0, mark_list[j-1])
        sheet.write(j, 1, id_list[j-1])
        sheet.write(j, 2, time_list[j-1])
        sheet.write(j, 3, like_list[j-1])
        sheet.write(j, 4, comment_list[j-1])
    book.save('./评论爬取/'+xls_name+'.xls')


if __name__ == '__main__':
    headers = {
        "cookie": "_T_WM=26998148397; H5_wentry=H5; _T_WL=1; WEIBOCN_WM=3349; backURL=https%3A%2F%2Fweibo.cn; ALF=1613975296; SCF=AvXNCxHHmXBcTrDlD7aCpH4O3WQIxM7T0HJZE8moEmsc1lr7TL1qIzI8b6AdUQ7OE0MHPyH2VLYn-YiAgOuMGVo.; SUB=_2A25ND68PDeRhGeFL7lEU9C3Nwj2IHXVu8zFHrDV6PUJbktANLVbWkW1NfeGMuWI9cvRAxLp4gJy_XHuDp6htPVY7; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WFVn8iEbGIZegiDqWEhxf9F5NHD95QNSK-0SKB0eK.pWs4DqcjMi--NiK.Xi-2Ri--ciKnRi-zNeKeES0M0e05ceBtt; SSOLoginState=1611390815",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/87.0.4280.141 Safari/537.36 Edg/87.0.664.75"
    }
    eventlet.monkey_patch()
    base_url = 'https://weibo.cn/comment/Iplj1awIy?uid=2656274875&rl=0&page='
    xls_name = "评论爬取pd2"
    book = xlrd.open_workbook('./pd2.xls')    #TODO
    data_temp_id = []
    data_temp_comment = []
    data_temp_like = []

    mark_list = []
    title_list = []
    id_list = []
    comment_list = []
    like_list = []

    sheet_link = book.sheet_by_index(0)
    row = sheet_link.nrows
    for i in range(1, row):  # TODO 记得改参数(已改）
        time.sleep(2)
        try:
            with eventlet.Timeout(10, False):
                base_url = str(sheet_link.cell_value(i, 4)).split('#')[0] + "&page=1"
                html = askURL(base_url, headers)
                html_parse = etree.HTML(html)
                # comment_pages = html_parse.xpath('//*/div[@id="pagelist"]/form/div[1]/input[1]/@value')
                # for j in range(1, 2):  # TODO 记得改参数 int(comment_pages[0])（已改）
                # html = askURL(base_url, headers)
                for k in range(0, 3):
                    mark_list.append(str(sheet_link.cell_value(i, 0)))
                    title_list.append(str(sheet_link.cell_value(i, 2)))
                data_temp_id, data_temp_like, data_temp_comment = get_data(html)
                id_list += data_temp_id
                like_list += data_temp_like
                comment_list += data_temp_comment
                print("未跳过第{}个".format(i))
        finally:
            print("爬了第{}个新闻的热评".format(i))
            continue
    print("爬完一次")
    save_data(mark_list, title_list, id_list, like_list, comment_list, xls_name)
