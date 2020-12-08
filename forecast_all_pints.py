from selenium import webdriver
import time
import xlrd
import xlsxwriter


driver = webdriver.Chrome()
driver.set_window_size(2400, 1400)

def points_read():
    work_book = xlrd.open_workbook('./points.xlsx')
    sheet = work_book.sheet_by_name('点位表')
    name = sheet.col_values(0)[3:]
    coordinate = sheet.col_values(1)[3:]
    altitude = sheet.col_values(2)[3:]
    lat = [round(int(i[1:3]) + int(i[4:6])/60 + int(i[7:9])/3600, 4) for i in coordinate]
    lon = [round(int(i[12:14]) + int(i[15:17])/60 + int(i[18:20])/3600, 4) for i in coordinate]
    name_dic = {}
    for i in range(len(name)):
        name_dic[name[i]] = [round(lat[i], 3), round(lon[i], 3), int(altitude[i])]
    return name_dic


def get_forecast(name, lat, lon, alt):
    date = time.strftime("%Y-%m-%d", time.localtime())
    url = 'https://www.windy.com/' + str(lat) + '/' + str(lon)
    print('生成URL成功 ...')
    print('获取信息 ...')
    driver.get(url) 
    print('信息已获取 ... 正在处理 ...')
    time.sleep(6)
    driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/div[2]').click()  # 未来十天
    time.sleep(2)
    table = driver.find_element_by_xpath('''/html/body/div[4]/div[1]/div[4]/div[2]/div[1]/table''')  # 数据表格
    print('保存图片 ...')
    table.screenshot('./' + name + '.png')
    print(name + '已保存。')
    format_normal = {
    'align':'center',#水平位置设置：居中
    'valign':'vcenter',#垂直位置设置，居中
    'font_size':10,#'字体大小设置'
    'font_name':'仿宋_GB2312',#字体设置
    'border':1,#边框设置样式1
    #'border_color':'green',#边框颜色
    #'bg_color':'#c7ffec',#背景颜色设置
    }
    
    format_title = {
    'align':'center',#水平位置设置：居中
    'valign':'vcenter',#垂直位置设置，居中
    'font_size':16,#'字体大小设置'
    'font_name':'方正小标宋简体',#字体设置
    'border':0,#边框设置样式1
    #'border_color':'green',#边框颜色
    #'bg_color':'#c7ffec',#背景颜色设置
    }
    
    format_table = {
    'align':'left',#水平位置设置：居中
    'valign':'vcenter',#垂直位置设置，居中
    'font_size':10,#'字体大小设置'
    'font_name':'楷体_GB2312',#字体设置
    'border':0,#边框设置样式1
    #'border_color':'green',#边框颜色
    #'bg_color':'#c7ffec',#背景颜色设置
    }
    
    format_date = {
    'align':'right',#水平位置设置：居中
    'valign':'vcenter',#垂直位置设置，居中
    'font_size':10,#'字体大小设置'
    'font_name':'楷体_GB2312',#字体设置
    'border':0,#边框设置样式1
    }
    
    format_first = {
    'align':'center',#水平位置设置：居中
    'valign':'vcenter',#垂直位置设置，居中
    'font_size':10,#'字体大小设置'
    'font_name':'黑体',#字体设置
    'border':1,#边框设置样式1
    }
    
    book = xlsxwriter.Workbook('./' + date + '.xlsx')
    style = book.add_format(format_normal)
    style_title = book.add_format(format_title)
    style_table = book.add_format(format_table)
    style_date = book.add_format(format_date)
    style_first = book.add_format(format_first)
    sheet = book.add_worksheet('forecast')
    sheet.merge_range('A1:D1', '各点位未来10天天气预报', style_title)
    sheet.merge_range('A2:C2', '制表：气象导航办公室', style_table)
    sheet.write('D2', '生成时间：' + time.strftime("%Y年%m月%d日 %H时%M分", time.localtime()), style_date)
    sheet.set_column('A:A', 13)
    sheet.set_column('B:B', 13)
    sheet.set_column('C:C', 7)
    sheet.set_column('D:D', 200) 
    for i in range(100):
        sheet.set_row(i+3, 137)
    
    for i in POINTS:
        sheet.write('A' + str(POINTS[i][3]), i, style)
        sheet.write('B' + str(POINTS[i][3]), 'N' + str(POINTS[i][0]) + '\n' + ' E' + str(POINTS[i][1]), style)
        sheet.write('C' + str(POINTS[i][3]), str(POINTS[i][2]), style)
        sheet.write('D' + str(POINTS[i][3]), '', style)
    
    sheet.write('A3', '点位', style_first)
    sheet.write('B3', '坐标', style_first)
    sheet.write('C3', '海拔', style_first)
    sheet.write('D3', '天气预报详情', style_first)
    
    for i in POINTS:
        sheet.insert_image('D' + str(POINTS[i][3]), './'+ i + '.png')
        print(i + '-天气信息已插入。')
    book.close()


POINTS = points_read()
I = 4
for i in POINTS:
    POINTS[i].append(I)
    I += 1


while 1:
    try:
        for i in POINTS:
            get_forecast(i, POINTS[i][0], POINTS[i][1], POINTS[i][2])
    except:
        pass


