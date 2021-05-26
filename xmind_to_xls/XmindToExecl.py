# -*- coding: utf-8 -*-
#@Time    : 4/7/21 4:37 PM
#@Author  : SHAUN-coyote
#@Email   : coyotezxy@163.com
#@File    : XmindToExecl.py

import xlwt
from xmindparser import xmind_to_dict

class XlwtSeting(object):
    @staticmethod  # 静态方法装饰器，使用此装饰器装饰后，可以直接使用类名.方法名调用（XlwtSeting.styles()），并且不需要self参数
    def template_one(worksheet):
        dicts = {"horz": "CENTER", "vert": "CENTER"}
        sizes = [15, 15, 30, 60, 45, 45, 15, 15]
        se = XlwtSeting()
        style = se.styles()
        style.alignment = se.alignments(**dicts)
        style.font = se.fonts(bold=True)
        style.borders = se.borders()
        style.pattern = se.patterns(7)
        se.heights(worksheet, 0)
        for i in range(len(sizes)):
            se.widths(worksheet, i, size=sizes[i])
        return style

    @staticmethod
    def template_two():
        dicts2 = {"vert": "CENTER"}
        se = XlwtSeting()
        style = se.styles()
        style.borders = se.borders()
        style.alignment = se.alignments(**dicts2)
        return style

    @staticmethod
    def template_three():
        dicts3 = {"horz": "CENTER","vert": "CENTER"}
        se = XlwtSeting()
        style = se.styles()
        style.borders = se.borders()
        style.alignment = se.alignments(**dicts3)
        return style

    @staticmethod
    def styles():
        """设置单元格的样式的基础方法"""
        style = xlwt.XFStyle()
        return style

    @staticmethod
    def borders(status=1):
        """设置单元格的边框，
        细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:
        11，粗双点划线:12，斜点划线:13"""
        border = xlwt.Borders()
        border.left = status
        border.right = status
        border.top = status
        border.bottom = status
        return border

    @staticmethod
    def heights(worksheet, line, size=4):
        """设置单元格的高度"""
        worksheet.row(line).height_mismatch = True
        worksheet.row(line).height = size * 256

    @staticmethod
    def widths(worksheet, line, size=11):
        """设置单元格的宽度"""
        worksheet.col(line).width = size * 256

    @staticmethod
    def alignments(wrap=1, **kwargs):
        """设置单元格的对齐方式，
        ：接收一个对齐参数的字典{"horz": "CENTER", "vert": "CENTER"}horz（水平），vert（垂直）
        ：horz中的direction常用的有：CENTER（居中）,DISTRIBUTED（两端）,GENERAL,CENTER_ACROSS_SEL（分散）,RIGHT（右边）,LEFT（左边）
        ：vert中的direction常用的有：CENTER（居中）,DISTRIBUTED（两端）,BOTTOM(下方),TOP（上方）"""
        alignment = xlwt.Alignment()

        if "horz" in kwargs.keys():
            alignment.horz = eval(f"xlwt.Alignment.HORZ_{kwargs['horz'].upper()}")
        if "vert" in kwargs.keys():
            alignment.vert = eval(f"xlwt.Alignment.VERT_{kwargs['vert'].upper()}")
        alignment.wrap = wrap  # 设置自动换行
        return alignment

    @staticmethod
    def fonts(name='宋体', bold=False, underline=False, italic=False, colour='black', height=11):
        """设置单元格中字体的样式，
        默认字体为宋体，不加粗，没有下划线，不是斜体，黑色字体"""

        font = xlwt.Font()
        # 字体
        font.name = name
        # 加粗
        font.bold = bold
        # 下划线
        font.underline = underline
        # 斜体
        font.italic = italic
        # 颜色
        font.colour_index = xlwt.Style.colour_map[colour]
        # 大小
        font.height = 20 * height
        return font

    @staticmethod
    def patterns(colors=1):
        """设置单元格的背景颜色，该数字表示的颜色在xlwt库的其他方法中也适用，默认颜色为白色
        0 = Black, 1 = White,2 = Red, 3 = Green, 4 = Blue,5 = Yellow, 6 = Magenta, 7 = Cyan,
        16 = Maroon, 17 = Dark Green,18 = Dark Blue, 19 = Dark Yellow ,almost brown), 20 = Dark Magenta,
        21 = Teal, 22 = Light Gray,23 = Dark Gray, the list goes on..."""

        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = colors
        return pattern

class XmindToXsl(XlwtSeting):
    def __init__(self,name,people):
        """调用类时，读取xmind文件，并生成excel表格"""
        self.per = people
        try:
            self.out = xmind_to_dict(name)
            #print(self.out)
            self.excelname = name.split('/')[-1].split('.')[0] + '.xls'
            self.xmind_cat(self.out[0]['topic']['topics'], self.excelname)
        except Exception as e:
            print(f"打开xmind文件失败:{e}")

    def resolvePath(dict, lists, title):
        #处理xmind内容，生成一个列表lists，每个内容为一条用例
        # title去除首尾空格
        title = title.strip()
        if len(title) == 0:
            concatTitle = dict['title'].strip()
        else:
            concatTitle = title + '\t' + dict['title'].strip()
            if dict.__contains__('makers') == True:
                concatTitle = concatTitle + '\t' +  dict['makers'][0]
        if dict.__contains__('topics') == False:
            lists.append(concatTitle)
            #print(lists)
        else:
            for d in dict['topics']:
                XmindToXsl.resolvePath(d,lists,concatTitle)

    def xmind_cat(self,list, excelname):
        #生成Excel
        f = xlwt.Workbook()
        #生成excel文件，单sheet，sheet名为：xmind下面的画布名
        #目前只支持一个画布
        sheet_name = self.out[0]['title']
        sheet = f.add_sheet(sheet_name, cell_overwrite_ok=True)#第二参数用于确认同一个cell单元是否可以重设值。
        row0 = ['序号', '模块', '用例','步骤','预期结果','优先级','执行结果','bug对应链接','执行人员']
        # 生成第一行中固定表头内容
        style2 = XlwtSeting.template_one(sheet)
        for i in range(0,len(row0)):
            sheet.write(0,i,row0[i],style2)

        style = XlwtSeting.template_two()
        style3 = XlwtSeting.template_three()

        # 增量索引
        index = 0
        #定义mode：Excel的模块列
        mode = str()

        for h in range(0, len(list)):
            lists = []
            XmindToXsl.resolvePath(list[h], lists, '')
            #print(lists)

            for j in range(0, len(lists)):
                lists[j] = lists[j].split('\t')
                #print(lists[j])

                for n in range(0, len(lists[j])):
                    #print(lists[j][n])
                    if 'priority' not in lists[j][n+1]:
                        mode = mode + lists[j][n]
                        sheet.write(j + index + 1,1,mode,style)
                    else:
                        #用例
                        test_name = lists[j][n]
                        #优先级：只取1，2，3
                        priority = lists[j][n+1][-1]
                        sheet.write(j + index + 1,2,test_name,style)
                        sheet.write(j + index + 1,5,priority,style)
                        if len(lists[j]) >= n+3:
                            #步骤
                            step = lists[j][n+2]
                            #结果
                            result = lists[j][n+3]
                            sheet.write(j + index + 1,3,step,style)
                            sheet.write(j + index + 1,4,result,style)
                        else:
                            sheet.write(j + index + 1, 3, '', style)
                            sheet.write(j + index + 1, 4, '', style)
                        mode = ''
                        break
                    #序号
                    sheet.write(j + index + 1, 0, j + index + 1,style3)
                    #执行人员
                    sheet.write(j + index + 1, 8, self.per, style3)
                    #执行结果
                    sheet.write(j + index + 1, 6, '', style3)
                    #bug对应链接
                    sheet.write(j + index + 1, 7, '', style)
        f.save(excelname)

#     def maintest(filename):
#         out = xmind_to_dict(filename)
#         excelname = filename.split('/')[-1].split('.')[0] + '.xls'
#         XmindToXsl.xmind_cat(out[0]['topic']['topics'], excelname)
#
#
# if __name__ == '__main__':
#     filename = '/Users/coyote/Desktop/test/xmindtoexcel/测试用例test.xmind'
#     XmindToXsl.maintest(filename)



