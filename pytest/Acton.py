from selenium import webdriver
import xlwt
import time
import pymysql
import xlrd
import os

path="H:\\桌面\\pai.xls "
url='https://www.bilibili.com/ranking/bangumi/13/1/3/?spm_id_from=333.334.b_72616e6b696e675f74696d696e675f62616e67756d69.11'
print('导入excel中，请稍等')
#elements=wd.find_elements_by_tag_name('span')

def save_xls():
    wbk = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = wbk.add_sheet('sheet 1', cell_overwrite_ok=True)
    first_col = sheet.col(0)  # xlwt中是行和列都是从0开始计算的
    sec_col = sheet.col(1)
    first_col.width = 256 * 70  # 列宽

    # 表头
    table_top_list = ['动漫', '播放量', '弹幕', '追番人数', '综合得分']
    for c, top in enumerate(table_top_list):
    # row_list.append(top.text)
        sheet.write(0, c, top)

    # 按列写入数据
    names = wd.find_elements_by_class_name('title')
    for r, name in enumerate(names, 1):
        sheet.write(r, 0, name.text)

    for i in range(51):
        viewpath = '//*[@id="app"]/div[2]/div/div[1]/div[2]/div[3]/ul/li[' + str(i) + ']/div[2]/div[2]/div[2]/span[1]'  # xppath只有一处不同，用循环抓取所有数据
        views = wd.find_elements_by_xpath(viewpath)
        for r, view in enumerate(views, 1):
            sheet.write(i, 1, view.text)

    for i in range(51):
        danmupath = '//*[@id="app"]/div[2]/div/div[1]/div[2]/div[3]/ul/li[' + str(i) + ']/div[2]/div[2]/div[2]/span[2]'
        views = wd.find_elements_by_xpath(danmupath)
        for r, view in enumerate(views, 1):
            sheet.write(i, 2, view.text)

    for i in range(51):
        fanpath = '//*[@id="app"]/div[2]/div/div[1]/div[2]/div[3]/ul/li[' + str(i) + ']/div[2]/div[2]/div[2]/span[3]'
        views = wd.find_elements_by_xpath(fanpath)
        for r, view in enumerate(views, 1):
            sheet.write(i, 3, view.text)

    for i in range(51):
        fanpath = '//*[@id="app"]/div[2]/div/div[1]/div[2]/div[3]/ul/li[' + str(i) + ']/div[2]/div[2]/div[3]/div'
        views = wd.find_elements_by_xpath(fanpath)
        for r, view in enumerate(views, 1):
            sheet.write(i, 4, view.text)

    wbk.save(path)
    print('导入成功,请在桌面查看')


def open_excel():
    try:
        book = xlrd.open_workbook(path)  # 文件名，把文件与py文件放在同一目录下
    except:
        print("open excel file failed!")
    try:
        sheet = book.sheet_by_name("sheet 1")  # execl里面的worksheet1
        return sheet
    except:
        print("locate worksheet in excel failed!")


# 连接数据库
try:
    db = pymysql.connect(host="127.0.0.1",
                         user="root",
                         passwd="123456",
                         db="acton",
                         charset='utf8')
except:
    print("could not connect to mysql server")


def insert_data():
    print('导入数据库中，请稍后\n')
    sheet = open_excel()
    cursor = db.cursor()
    for i in range(1, sheet.nrows):  # 第一行是标题名，对应表中的字段名所以应该从第二行开始，计算机以0开始计数，所以值是1

        name = sheet.cell(i, 0).value  # 取第i行第0列
        view = sheet.cell(i, 1).value  # 取第i行第1列，下面依次类推
        danmu = sheet.cell(i,2).value
        zhui = sheet.cell(i,3).value
        grade = sheet.cell(i,4).value
        value = (name, view, danmu, zhui, grade, i)
        print(value)
        sql = "replace into ac(动漫,播放量,弹幕,追番人数,综合得分,排名) values (%s,%s,%s,%s,%s,%s)"
        cursor.execute(sql, value)  # 执行sql语句
        db.commit()

    cursor.close()  # 关闭连接
    db.close()  # 关闭数据
    print('导入成功')

if __name__ == '__main__':

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('log-level=3')
    # chrome_options = Options()
    chrome_options.add_argument('--headless')  # 隐藏窗口chrome_options=chrome_options
    chrome_options.add_argument('--disable-gpu')

    if (os.path.exists(path)):
        print('excel文件已存在')
        os.remove(path)
        print('更新数据中，请稍后')
        wd = webdriver.Chrome('H:\桌面\chromedriver.exe',options=chrome_options)
        wd.get(url)
        time.sleep(1)
        save_xls()
        insert_data()
    else:
        wd = webdriver.Chrome('H:\桌面\chromedriver.exe',options=chrome_options)
        wd.get(url)
        time.sleep(1)
        save_xls()
        insert_data()
        
raw=input("please enter any key to exit")
wd.quit()
