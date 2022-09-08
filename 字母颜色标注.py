from unicodedata import digit
import xlrd,xlwt
from itertools import groupby
import re
x1 = xlrd.open_workbook("J:\浏览器下载文件\single beads 转换.xls")
x2 = open("J:\分割ABC.txt")
sort_list = x2.readlines()
# print(x2.readlines())
print(x1.sheet_names())
sheet1 = x1.sheet_by_name("A")
book1 = xlwt.Workbook()
sheet = book1.add_sheet('')

def main():
    for i in range(1):
        row_value = sheet1.row_values(i)
        if row_value[0] == "":
            continue
        while "" in row_value: row_value.remove("")
        one_list = get_whole_ep(row_value,sort_list)
        print(one_list)
        sheet.write(i,0,one_list[0][0])
        for j in range(len(one_list)-1):  
            x = one_list[j]  
            color_list = []
            for s in x[1:]:
                print(s)
                digit,num = search_num(s)               
                bgx = get_bgcolor(color_site,sheet,3,digit+3)
                color_s = []
                color_s.append(s)
                color_s.append(bgx)
                color_list.append(color_s)
            write_sheet(sheet,i,j,color_list)
def get_whole_ep(row_value,sort_list):
    one_list = []
    for x in row_value[1:]:
        for y in sort_list:
            y = y.strip('\n')
            list_y = y.split('\t')
            while '' in list_y: list_y.remove("")
            if x in list_y:
                list_ep = list_y
                list_ep.insert(0,row_value[0])
                one_list.append(list_ep)
    return one_list
                

# print(all_list)
# 获取颜色
list_ep
color_site = xlrd.open_workbook("J:\多态位点分布.xls",formatting_info=True)
sheet = color_site.sheets()[0]
s = '80T'

                    

def search_num(s):
    digit = re.search(r'\d+',s)
    letter = re.search(r"\D",s)        
    return digit.group(),letter.group()

sheet0 = color_site.sheets()[0]
def get_bgcolor(Book, sheet, row, col,all_list):
    """获取单元格背景颜色"""
    xfx = sheet.cell_xf_index(3, col)
    xf = Book..xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    # 加个判断：
    print(bgx)
    return bgx 
# 根据获取的单元格颜色，写入字体颜色

def wite_sheet(sheet,row,col,color_list):
    font = xlwt.Font()
    row_list = []
    font.colour_index = color_list[0][1]
    row_list.append([color_list[0][0],font])
    if len(color_list) > 1:
        for x in color_list[1:]:          
            font.colour_index = x[1]
            link_letter = re.search(r'\D',x[0]).group()
            row_list.append([link_letter,font])
    print(row_list)
    sheet.write(row,col,row_list)



                
                
                
        
        
        