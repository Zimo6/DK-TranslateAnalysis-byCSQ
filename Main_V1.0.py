# -*- coding: utf-8 -*-
# @Time    : 2021/6/21 17:04
# @Author  : CuiShuangqi
# @Email   : 2807481686@qq.com
# @File    : Main_V1.0.py
# 脚本当前目录
import os
import xml.etree.ElementTree as ET
import xlwt

# 定义全局变量
CUR_PATH = os.path.dirname(os.path.realpath(__file__))
RES_PATH = os.path.join(CUR_PATH, 'DK_Setting\\res')
RES_EXT_PATH = os.path.join(CUR_PATH, f'DK_Setting\\res_ext')
# 准备工作，新建解析结果的文件夹
analysis_path = os.path.join(CUR_PATH, f'DK_Setting_AnalysisResult')
if not os.path.exists(analysis_path):
    os.mkdir(analysis_path)
analysis_res_path = f"{analysis_path}\\res"
if not os.path.exists(analysis_res_path):
    os.mkdir(analysis_res_path)
analysis_res_ext_path = f"{analysis_path}\\res_ext"
if not os.path.exists(analysis_res_ext_path):
    os.mkdir(analysis_res_ext_path)


# 遍历DK_Setting\res的语言目录
for root, dirs, files in os.walk(RES_PATH):
    for res_dir in dirs:
        # 创建一个workbook 设置编码
        workbook = xlwt.Workbook(encoding='utf-8')
        # 设置Excel样式
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = '微软雅黑'
        style.font = font
        print(f"正在解析目录：【{res_dir}】...")
        # 遍历目录下的xml文件
        for r, d, f in os.walk(f"{RES_PATH}\\{res_dir}"):
            for res_file in f:
                # xml文件名
                # print(res_file)
                sheet_name = res_file.strip(".xml")
                sheet_tmp = workbook.add_sheet(sheet_name)
                # 设置列宽
                sheet_tmp.col(0).width = 256*50
                sheet_tmp.col(1).width = 256*50
                # 写入excel
                res_xml_path = f"{RES_PATH}\\{res_dir}\\{res_file}"
                tree = ET.parse(res_xml_path)
                root = tree.getroot()
                # 列表存放值
                name_list = []
                value_list = []
                for child in root:
                    try:
                        name_list.append(child.attrib['name'])
                        # 这里说明有节点嵌套
                        if child.text is None:
                            value_list.append(" ")
                        # 这里说明有子节点
                        elif len(child.text.strip()) == 0:
                            next_value_list = []
                            for next_child in child:
                                if next_child.text is not None:
                                    next_value_list.append(str(next_child.text + ' | '))
                                # 如果子节点再有节点嵌套
                                # if next_child.text is None:
                                #     break
                            value_list.append(next_value_list)
                        else:
                            value_list.append(child.text)
                    except:
                        name_list.append(None)
                        value_list.append(None)
                for i in range(len(value_list)):
                    # 设置行高
                    sheet_tmp.row(i).height = 50 * 80
                    sheet_tmp.write(i, 0, name_list[i], style)
                    sheet_tmp.write(i, 1, value_list[i], style)
        workbook.save(f"{analysis_res_path}\\{res_dir}.xls")

print("===================================================================")

# 遍历DK_Setting\res_ext的语言目录
for root, dirs, files in os.walk(RES_EXT_PATH):
    for res_ext_dir in dirs:
        # 创建一个workbook 设置编码
        workbook = xlwt.Workbook(encoding='utf-8')
        # 设置Excel样式
        style = xlwt.XFStyle()
        font = xlwt.Font()
        font.name = '微软雅黑'
        style.font = font
        print(f"正在解析目录：【{res_ext_dir}】...")
        # 遍历目录下的xml文件
        for r, d, f in os.walk(f"{RES_EXT_PATH}\\{res_ext_dir}"):
            for res_ext_file in f:
                # xml文件名
                # print(res_ext_file)
                sheet_name = res_ext_file.strip(".xml")
                sheet_tmp = workbook.add_sheet(sheet_name)
                # 设置列宽
                sheet_tmp.col(0).width = 256*50
                sheet_tmp.col(1).width = 256*50
                # 写入excel
                res_ext_xml_path = f"{RES_EXT_PATH}\\{res_ext_dir}\\{res_ext_file}"
                tree = ET.parse(res_ext_xml_path)
                root = tree.getroot()
                # 列表存放值
                name_list = []
                value_list = []
                for child in root:
                    try:
                        name_list.append(child.attrib['name'])
                        # 这里说明有节点嵌套
                        if child.text is None:
                            value_list.append(" ")
                        # 这里说明有子节点
                        elif len(child.text.strip()) == 0:
                            next_value_list = []
                            for next_child in child:
                                if next_child.text is not None:
                                    next_value_list.append(str(next_child.text + ' | '))
                                # 如果子节点再有节点嵌套
                                # if next_child is None:
                                #     break
                            value_list.append(next_value_list)
                        else:
                            value_list.append(child.text)
                    except:
                        name_list.append(None)
                        value_list.append(None)
                for i in range(len(value_list)):
                    sheet_tmp.write(i, 0, name_list[i], style)
                    sheet_tmp.write(i, 1, value_list[i], style)
        workbook.save(f"{analysis_res_ext_path}\\{res_ext_dir}.xls")
