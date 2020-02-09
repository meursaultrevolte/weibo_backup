# -*-coding:utf8-*-
# 需要的模块
import os
import urllib
from urllib import request, parse
import urllib.request
import time
import json
import xlwt

proxy_addr = "122.241.72.191:808"

print('先输入你的微博数字id，可尝试在网页上登录weibo.cn，点击自己的微博数量，weibo.cn/xxx/profile,xxx就是你的id,'
      '随后会让你选择从第几页开始抓取，有一定几率抓取失败，可再次执行任务从结束的页面重新抓取')
id = input("你的微博数字ID：")


def use_proxy(url, proxy_addr):
    req = urllib.request.Request(url)
    req.add_header('User-Agent',
                   'Mozilla/6.0 (iPhone; CPU iPhone OS 8_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/8.0 Mobile/10A5376e Safari/8536.25')
    proxy = urllib.request.ProxyHandler({'http': proxy_addr})
    opener = urllib.request.build_opener(proxy, urllib.request.HTTPHandler)
    urllib.request.install_opener(opener)
    data = urllib.request.urlopen(req).read().decode('utf-8', 'ignore')
    return data


# 获取微博主页的containerid
def get_containerid(url):
    data = use_proxy(url, proxy_addr)
    content = json.loads(data).get('data')
    for data in content.get('tabsInfo').get('tabs'):
        if (data.get('tab_type') == 'weibo'):
            containerid = data.get('containerid')
    return containerid


# 获取微博账号的用户基本信息
def get_userInfo(id):
    url = 'https://m.weibo.cn/api/container/getIndex?type=uid&value=' + id
    data = use_proxy(url, proxy_addr)
    content = json.loads(data).get('data')
    profile_image_url = content.get('userInfo').get('profile_image_url')
    description = content.get('userInfo').get('description')
    profile_url = content.get('userInfo').get('profile_url')
    verified = content.get('userInfo').get('verified')
    guanzhu = content.get('userInfo').get('follow_count')
    name = content.get('userInfo').get('screen_name')
    fensi = content.get('userInfo').get('followers_count')
    gender = content.get('userInfo').get('gender')
    urank = content.get('userInfo').get('urank')
    print("微博昵称：" + name + "\n" + "微博主页地址：" + profile_url + "\n" + "微博头像地址：" + profile_image_url + "\n" + "是否认证：" + str(
        verified) + "\n" + "微博说明：" + description + "\n" + "关注人数：" + str(guanzhu) + "\n" + "粉丝数：" + str(
        fensi) + "\n" + "性别：" + gender + "\n" + "微博等级：" + str(urank) + "\n")
    return name


# 保存图片
def savepic(pic_urls, created_at, page, num):
    pic_num = len(pic_urls)
    srcpath = '尝试抓取的微博图片/'
    if not os.path.exists(srcpath):
        os.makedirs(srcpath)
    picpath = str(created_at) + 'page' + str(page) + 'num' + str(num) + 'pic'
    for i in range(len(pic_urls)):
        picpathi = picpath + str(i)
        path = srcpath + picpathi + ".jpg"
        urllib.request.urlretrieve(pic_urls[i], path)


# 获取微博内容信息,并保存到文本中
def get_weibo(id, file):
    i = int(input("从第几页开始："))
    page = int(input("到第几页(包含输入的页面)（输入0则为抓取全部微博）：")) + 1
    while True:
        url = 'https://m.weibo.cn/api/container/getIndex?type=uid&value=' + id
        weibo_url = 'https://m.weibo.cn/api/container/getIndex?type=uid&value=' + id + '&containerid=' + get_containerid(
            url) + '&page=' + str(i)
        print(weibo_url)
        try:
            data = use_proxy(weibo_url, proxy_addr)
            content = json.loads(data).get('data')
            cards = content.get('cards')
            if (len(cards) > 0):
                for j in range(len(cards)):
                    print("-----正在爬取第" + str(i) + "页，第" + str(j) + "条微博------")
                    card_type = cards[j].get('card_type')
                    if (card_type == 9):
                        mblog = cards[j].get('mblog')
                        attitudes_count = mblog.get('attitudes_count')  # 点赞数
                        comments_count = mblog.get('comments_count')  # 评论数
                        created_at = mblog.get('created_at')  # 发布时间
                        reposts_count = mblog.get('reposts_count')  # 转发数
                        scheme = cards[j].get('scheme')  # 微博地址
                        text = mblog.get('text')  # 微博内容
                        pictures = mblog.get('pics')  # 正文配图，返回list
                        pic_urls = []  # 存储图片url地址
                        if pictures:
                            for picture in pictures:
                                pic_url = picture.get('large').get('url')
                                pic_urls.append(pic_url)
                        # print(pic_urls)

                        # 保存文本
                        with open(file, 'a', encoding='utf-8') as fh:
                            if len(str(created_at)) < 6:
                                created_at = str(created_at)
                            # 页数、条数、微博地址、发布时间、微博内容、点赞数、评论数、转发数、图片链接
                            fh.write(str(i) + '\t' + str(j) + '\t' + str(scheme) + '\t' + str(
                                created_at) + '\t' + text + '\t' + str(attitudes_count) + '\t' + str(
                                comments_count) + '\t' + str(reposts_count) + '\t' + str(pic_urls) + '\n')

                        # 保存图片
                        savepic(pic_urls, created_at, i, j)
                i += 1
                if page == 0:
                    break
                if i == page:
                    break

                '''休眠1s以免给服务器造成严重负担'''
                time.sleep(1)
            else:
                break
        except Exception as e:
            print(e)
            pass


# txt转换为xls
def txt_xls(filename, xlsname):
    """
    :文本转换成xls的函数
    :param filename txt文本文件名称、
    :param xlsname 表示转换后的excel文件名
    """
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            xls = xlwt.Workbook()
            # 生成excel的方法，声明excel
            sheet = xls.add_sheet('sheet1', cell_overwrite_ok=True)
            # 页数、条数、微博地址、发布时间、微博内容、点赞数、评论数、转发数
            sheet.write(0, 0, '爬取页数')
            sheet.write(0, 1, '爬取当前页数的条数')
            sheet.write(0, 2, '微博地址')
            sheet.write(0, 3, '发布时间')
            sheet.write(0, 4, '微博内容')
            sheet.write(0, 5, '点赞数')
            sheet.write(0, 6, '评论数')
            sheet.write(0, 7, '转发数')
            sheet.write(0, 8, '图片链接')
            x = 1
            while True:
                # 按行循环，读取文本文件
                line = f.readline()
                if not line:
                    break  # 如果没有内容，则退出循环
                for i in range(0, len(line.split('\t'))):
                    item = line.split('\t')[i]
                    sheet.write(x, i, item)  # x单元格行，i 单元格列
                x += 1  # excel另起一行
        xls.save(xlsname)  # 保存xls文件
    except:
        raise


if __name__ == "__main__":
    name = get_userInfo(id)
    file = str(name) + id + ".txt"
    get_weibo(id, file)

    txtname = file
    xlsname = str(name) + id + "的微博内容.xls"
    txt_xls(txtname, xlsname)

print('本次抓取完毕，若未抓取完毕，可从最后的页数开始抓取')
print('')
input("按回车键关闭程序")
