# 总模板
import xlwt
from datetime import datetime, date

# 第一行样式
def set_style(name, height, bold=False, format_str='', align='center'):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name  # 'Times New Roman'
    font.bold = bold
    font.height = height

    borders = xlwt.Borders()  # 为样式创建边框
    borders.left = 2
    borders.right = 2
    borders.top = 2
    borders.bottom = 2

    alignment = xlwt.Alignment()  # 设置排列
    if align == 'center':
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
    else:
        alignment.horz = xlwt.Alignment.HORZ_LEFT
        alignment.vert = xlwt.Alignment.VERT_BOTTOM

    style.font = font
    style.borders = borders
    style.num_format_str = format_str
    style.alignment = alignment

    return style

workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Liangze',cell_overwrite_ok=True)

# 添加第一行
worksheet.write_merge(
    0,
    0,
    0,
    5,
    '每日工作记录',
    set_style(
        'Times New Roman',
        320,
        bold=True,
        format_str=''
    )
)

# ~~内容~~
worksheet.write_merge(
    2,
    2,
    0,
    5,
    '',
    set_style(
        'Times New Roman',
        320,
        bold=False,
        format_str=''
    )
)

worksheet.write_merge(
    3,
    3,
    0,
    5,
    '',
    set_style(
        'Times New Roman',
        320,
        bold=False,
        format_str=''
    )
)

worksheet.write_merge(
    4,
    4,
    0,
    5,
    '',
    set_style(
        'Times New Roman',
        320,
        bold=False,
        format_str=''
    )
)

worksheet.write_merge(
    5,
    5,
    0,
    5,
    '',
    set_style(
        'Times New Roman',
        320,
        bold=False,
        format_str=''
    )
)

worksheet.write_merge(
    6,
    6,
    0,
    5,
    '',
    set_style(
        'Times New Roman',
        320,
        bold=False,
        format_str=''
    )
)
# ~~内容~~

# 第二行 的 样式1
styleOK = xlwt.easyxf()
pattern = xlwt.Pattern()  # 一个实例化的样式类
pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # 固定的样式

borders = xlwt.Borders()  # 为样式创建边框
borders.left = 2
borders.right = 2
borders.top = 6
borders.bottom = 2

font = xlwt.Font()  # 为样式创建字体
font.name = 'Times New Roman'
font.height = 220

alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
alignment.vert = xlwt.Alignment.VERT_CENTER

styleOK.pattern = pattern
styleOK.borders = borders
styleOK.font = font
styleOK.alignment = alignment

# 第二行 的 样式2
styleOK2 = xlwt.easyxf()
pattern2 = xlwt.Pattern()  # 一个实例化的样式类
pattern2.pattern = xlwt.Pattern.SOLID_PATTERN  # 固定的样式

borders2 = xlwt.Borders()  # 为样式创建边框
borders2.left = 2
borders2.right = 2
borders2.top = 6
borders2.bottom = 2

font2 = xlwt.Font()  # 为样式创建字体
font2.name = 'Times New Roman'
font2.height = 220
font2.bold = True

alignment2 = xlwt.Alignment()
alignment2.horz = xlwt.Alignment.HORZ_CENTER
alignment2.vert = xlwt.Alignment.VERT_CENTER

styleOK2.pattern = pattern2
styleOK2.borders = borders2
styleOK2.font = font2
styleOK2.alignment = alignment2

# 第二行
rows = ['姓名','','日期','','星期','']
for index, val in enumerate(rows):
    worksheet.col(index).width = 150 * 30
    if(index % 2 != 1):
        worksheet.write(1, index, val, styleOK2)
    else:
        worksheet.write(1, index, val, styleOK)

workbook.save('d:/Workspace/python-pro/day/test.xls')