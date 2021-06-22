# -*- coding: utf-8 -*-
# @Time    : 2021/6/21 16:46
# @Author  : CuiShuangqi
# @Email   : 2807481686@qq.com
# @File    : Test.py
# 创建一个workbook 设置编码
# import xlwt
import xml.etree.ElementTree as ET
#
# workbook = xlwt.Workbook(encoding='utf-8')
# # 创建一个worksheet
# worksheet1 = workbook.add_sheet("a")
# worksheet2 = workbook.add_sheet("b")
# # 写入excel
# # 参数对应 行, 列, 值
# worksheet1.write(0, 0, label='this is test')
# worksheet2.write(0, 0, label='this is test')
# # 保存
# workbook.save(f"a.xls")

# tree = ET.parse(r"E:\Python-Projects\DK-TranslateAnalysis-byCSQ\DK_Setting\res_ext\values-zh-rCN\mtk_strings.xml")
tree = ET.parse(r"E:\Python-Projects\DK-TranslateAnalysis-byCSQ\DK_Setting\res\values\arrays.xml")
root = tree.getroot()
for child in root:
    try:
        print(f"{child}节点的【name】属性为{child.attrib['name']}")
        print(f"{child}节点的【text】属性为{child.text}")
        print(type(child.text))
        if child.text is None:
            print("============================")
        # print(len(child.text))
        elif len(child.text.strip()) == 0:
            for next_cild in child:
                if next_cild.text is None:
                    pass
                else:
                    print(f"二儿子的值：{next_cild.text}")
                    print(f"二儿子的值类型：{type(next_cild.text)}")
    except:
        print(f"{child}节点的name属性为【空】")
    # 遍历第二个子节点
