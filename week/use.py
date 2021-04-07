from xlutils.copy import copy
import xlwt
import xlrd
import datetime

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~def区域~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
"""
  获取指定的某天是某个月中的第几周
  周一作为一周的开始
"""
def get_week_of_month():
    nowYear = datetime.datetime.now().year
    nowMonth = datetime.datetime.now().month
    nowDay = datetime.datetime.now().day

    end = int(datetime.datetime(nowYear, nowMonth, nowDay).strftime("%W"))
    begin = int(datetime.datetime(nowYear, nowMonth, 1).strftime("%W"))
    return end - begin + 1
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~def区域~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# read excel
week_excel = xlrd.open_workbook('E:/day/week_test.xls')
week_sheet = week_excel.sheet_by_index(0)

# copy new excel
new_excel = copy(week_excel)
new_sheet = new_excel.get_sheet(0)


"""
    需要用到的参数
"""
# 标题动态时间
titleTime = (datetime.datetime.now() - datetime.timedelta(days = 4)).strftime("%Y-%m-%d") + '——' + datetime.datetime.now().strftime("%Y-%m-%d")
# 周一到周六的时间
week_one = (datetime.datetime.now() - datetime.timedelta(days = 4)).strftime("%Y-%m-%d")


"""
    标题
"""
title = "部门：技术部  （小邻通组）                 姓名：梁泽           岗位：                 直接主管： 吴振飞                                                                                             日期：" + titleTime

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~编辑区域~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
new_sheet.write_merge(2, 2, 1, 15, title)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~编辑区域~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# 保存日期
saveTime = str(datetime.datetime.now().year) + '.' + str(datetime.datetime.now().month)
# 保存
new_excel.save('e:/day/周工作任务与问效表-技术部-小邻通项目组-梁泽' + saveTime + '第' + str(get_week_of_month()) + '周.xls')