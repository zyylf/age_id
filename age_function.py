import openpyxl, arrow
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, colors


wb = openpyxl.load_workbook('d:\\age\\age20200622.xlsx')#读取源表
s = wb.active


def age_function(i):

    sfz = s['p'+str(i)].value  # 读取身份证号码
    #截取出生年月日
    birthday = str(sfz)[6:10] + '-' + str(sfz)[10:12] + '-' + str(sfz)[12:14] #截取出生日期

    #计算出生年月日
    b1 = arrow.get(birthday, 'YYYY-MM-DD')
    b = b1.format('YYYY-MM-DD')

    #计算年龄
    a = arrow.now() - b1
    age = int(a.days/365)#取整

    #计算退休时间
    age_old1 = b1.shift(years=60)
    age_old = age_old1.format('YYYY-MM-DD')

    #计算距离退休月数
    year = age_old1.year - arrow.now().year  #年减年
    month = age_old1.month - arrow.now().month #月减月


    age_old_months = year*12 + month  # 年数*12 +月数＝总月数


    #计算总工龄
    gzsj = s['L'+str(i)].value  # 读取参加工作时间
    gzsj = arrow.get(gzsj)#日期格式转换
    gl_year = arrow.now().year - gzsj.year
    gl_month = arrow.now().month - gzsj.month
    gl_months = gl_year*12 + gl_month
    #

    #计算公司内部工龄
    gsgzsj = s['N'+str(i)].value  # 读取进入公司时间
    gsgzsj = arrow.get(gsgzsj)#日期格式转换
    gsgl_year = arrow.now().year - gsgzsj.year
    gsgl_month = arrow.now().month - gsgzsj.month
    gsgl_months = gsgl_year*12 + gsgl_month

    s['R'+str(i)].value = b
    s['S'+str(i)].value = age
    s['T'+str(i)].value = gl_months
    s['U'+str(i)].value = age_old_months
    s['V'+str(i)].value = gsgl_months
    s['W'+str(i)].value = age_old


s['R4'].value = '出生年月日'
s['S4'].value = '截至目前年龄'
s['T4'].value = '总工龄（月数）'
s['U4'].value = '距离退休月数'
s['V4'].value = '公司内部工龄（月数）'
s['W4'].value = '退休日期'


for i in range(5, 31):#根据人数设置
    age_function(i)




#设置单元格格式
font1 = Font(name='黑体',size=24)
font2 = Font(size=12)
border2 = Border(left=Side('thin'),
                right=Side('thin'),
                top=Side('thin'),
                bottom=Side('thin'))
alignment = Alignment(horizontal='center', vertical='center')



for cells in s['R4:w30']:#设置边框
    for cell in cells:
        cell.border = border2
        cell.alignment = alignment



#按当前日期保存
wb.save('d:\\age\\age'+str(arrow.now().format('YYYY-MM-DD'))+ '.xlsx')
