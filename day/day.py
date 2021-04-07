from xlutils.copy import copy
import xlwt
import xlrd
import datetime

tem_excel = xlrd.open_workbook('e:/day/test.xls', formatting_info=True)
tem_sheet = tem_excel.sheet_by_index(0)

new_excel = copy(tem_excel)
new_sheet = new_excel.get_sheet(0)

# ~~~样式区域 start~~~

# 内容样式
style = xlwt.XFStyle()

# 第二行样式
style2 = xlwt.XFStyle()

# 字体样式
font = xlwt.Font()
font.height = 200
style.font = font
style2.font = font

# 居中设置
alianment = xlwt.Alignment()
alianment.horz = xlwt.Alignment.HORZ_CENTER
alianment.vert = xlwt.Alignment.VERT_CENTER
style.alignment = alianment
style2.alignment = alianment

# 表格设置
borders = xlwt.Borders()
borders.left = 2
borders.right = 2
borders.top = 2
borders.bottom = 2

borders2 = xlwt.Borders()
borders2.left = 2
borders2.right = 2
borders2.top = 6
borders2.bottom = 2

style.borders = borders
style2.borders = borders2
# ~~~样式区域 end~~~

# ~~~日报内容 start~~~
# 姓名
new_sheet.write(1, 1, '梁泽', style2)
# 日期
nowTime = datetime.datetime.now().strftime('%Y-%m-%d') # 现在
nowWeekDay = str(datetime.datetime.now().isoweekday()) # 星期
# 星期处理
if nowWeekDay == '1':
    nowWeekDay = '星期一'
elif nowWeekDay == '2':
    nowWeekDay = '星期二'
elif nowWeekDay == '3':
    nowWeekDay = '星期三'
elif nowWeekDay == '4':
    nowWeekDay = '星期四'
else:
    nowWeekDay = ''

# 日期
new_sheet.write(1, 3, nowTime, style2)
# 星期
new_sheet.write(1, 5, nowWeekDay, style2)
# 内容
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~编辑区域~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
new_sheet.write_merge(2, 2, 0, 5, 'h5接口： 对应问题的解决', style)
new_sheet.write_merge(3, 3, 0, 5, 'web后台：处理商品问题', style)
new_sheet.write_merge(4, 4, 0, 5, 'web后台：处理提现查询的结算详情问题', style)
new_sheet.write_merge(5, 5, 0, 5, 'web后台：优化提现记录-脱敏-显示对应提现方式', style)
new_sheet.write_merge(6, 6, 0, 5, 'web后台：优化分销商列表界面', style)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~编辑区域~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# ~~~日报内容 end~~~
# 保存
nowDay = datetime.datetime.now().strftime('%m-%d')#现在
new_excel.save('e:/day/梁泽每日工作记录 '+nowDay+'.xls')

