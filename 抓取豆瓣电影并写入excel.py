#导入库
import parsel
import requests
import MySQLdb
import xlsxwriter
#模拟浏览器请求
heads = {'user-agent': 'Mozilla/5.0(Windows NT 10.0; Win64; x64)\
         AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari537.36'}
#获取电影信息
url = 'https://movie.douban.com/top250'
r = requests.get(url, headers=heads)
selector = parsel.Selector(r.text)  #把r.text文本数据转换成selecor对象
lins = selector.css('.grid_view li').getall() #抓取li中信息进行下一步解析
title = selector.css('.hd a span:nth-child(1)::text').getall()  #影片名
urls = selector.css('.pic a::attr(href)').getall()  #详情链接
showtxt = selector.css('.inq::text').getall()  #介绍文字
mark = selector.css('.rating_num::text').getall()  #豆瓣评分
markNum = selector.css('.star span:nth-child(4)::text').getall()  #评分人数
lis=selector.css('.pic a img::attr(src)').getall()  #电影海报
for n in range(len(lis)):  #将海报图片保存在.py同路径文件夹下
    img =requests.get(lis[n]).content  
    with open(f'./images/{title[n]}.webp' , 'wb') as f:
        f.write(img)

#写入excel
wb=xlsxwriter.Workbook('豆瓣电影.xlsx')
ws=wb.add_worksheet('豆瓣电影Top250')
ws.set_column('A:A',7)
ws.set_column('B:B',30)
ws.set_column(2,3,50)
ws.set_column('G:G',30)
# 标题行
headings=['海报','电影名','详情','介绍','评分','评分人数']
# 设置excel风格
head_format=wb.add_format({'bold':1,'fg_color':'cyan','align':'center','font_name':u'微软雅黑','valign':'vcenter'})
cell_format=wb.add_format({'bold':0,'align':'center','font_name':u'微软雅黑','valign':'vcenter'})
ws.write_row('A1',headings,head_format)
# 将获取到的各种标签信息写入excel
for k in range(len(lins)):
    ws.set_row(k+1, 60)
    ws.insert_image('A'+str(k+2),f'./images/{title[k]}.jpg',{'x_scale':0.2,'y_scale':0.2})
    ws.write(k+1,1, title[k], cell_format)
    ws.write(k+1,2, urls[k], cell_format)
    ws.write(k+1,3, showtxt[k], cell_format)
    ws.write(k+1,4, mark[k], cell_format)
    ws.write(k+1,5, markNum[k],cell_format)
wb.close()

