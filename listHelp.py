# -*- coding: utf-8 -*-
"""
Created on Tue Feb 13 08:27:51 2018
    Python系统可以通过dir()、type()函数和__doc__属性获得指定模块的各项名称、属性的类
型，以及函数的内置说明。本程序通过利用上述3个函数提起指定模块的信息，并保存在excel文
件内，以便于以后的查询阅读。
    2018-2-14:
        经过2天多的努力完成了我认可的第一个能用的版本。
        发现问题：1.type()函数的参数应是对象类型，不能直接用模块名字符串代替。此问题
    目前没有办法解决。2.module对象转换为str类型后，字符串中存在不能作为Excel.Sheet的
    名称的特殊字符，通过RE提取module名解决。3.Python序列默认从0开始，而Excel单元格序
    列默认从1开始，在处理时要做转换。
    模块内置函数：
    1.ge_help(module):对传入的模块提取相关的帮助内容，并返回含有帮助的元组。
    2.write_help(module,in_help,filename):将传入的帮助内容元组（in_help）写入Excel
    文件（filename）中的表（module）中。
    3.help_to_excel(module,filename):调用get_help、write_help，将module模块的帮助
    写入filename文件（Excel格式）已module命名的表中。
    
@author: ccds_stx
"""
def format_typestr(type_string):
    """
    2018-3-1,增加本函数，用于将类型名称字符串进行格式化，去除多余内容，仅留下类型名称。
    注意：实参必须时字符串，本函数内不进行类型校验。
    """
    import re
    type_re = re.compile(r'\'.*\'')
    type_match = type_re.search(type_string)
    #type_name = type_match.group().replace("'",'')
    type_name = type_match.group()[1:-2] #使用字符串切片方式 将两端的“'”删除
    return type_name

def get_help(module):
    """
    本函数通过给定的模块名，获得包括模块内部项目名称、类型、内置帮助信息等内容，并
    保存在列表（或元组）中。
    输入：module，要提取帮助信息的模块。注：object类型。
    输出：返回包含帮助信息的列表（元组）名。
    """
    name_list = dir(module)
    header=('名称','类型','内部帮助')  #定义表头信息
    content = [] #列表，保存获得的所有内容
    content.append(header)
    for i in range(0,len(name_list)):
        x = name_list[i] # 变量x保存名称
        y = str(type(getattr(module,x))) #变量y保存类型
        #2018-3-1 增加格式化类型名称处理
        y = format_typestr(y)
        z = getattr(module,x).__doc__ #变量z保存帮助内容
        content.append((x,y,z))
    return tuple(content)

def write_help(module_name,in_help,filename):
    """
    本函数将获得模块内容帮助内容写入指定Excel文件中。
    输入：
    1.module_name：模块名称，str类型。
    2.in_help：获得的帮助内容，tuple类型。
    3.filename：保存文件名称，即Excel文件名称，str类型。
    输出：True或False
    """
    import openpyxl
    module_name = str(module_name)
    import re
    temp_name = re.search(r'\'.*?\'',module_name)
    if temp_name == None:
        return 'Error:模块名无效，退出'  
    else:
        tmp = temp_name.group()
        tmp = tmp.replace('\'','')
        module_name = tmp
    from os.path import exists as fileexists
    if not fileexists(filename):
        wb = openpyxl.Workbook()
    else:
        wb = openpyxl.load_workbook(filename)
    if  module_name in wb.sheetnames:
        return 'Error:已经有相关模块的帮助了，没必要重新写入。'
    else:
        ws = wb.create_sheet(title=module_name)
        for row in range(1,len(in_help)+1):
            for col in range(1,len(in_help[row-1])+1):
                ws.cell(row,col).value = in_help[row-1][col-1]
        wb.save(filename)
        return 'Ok'
    
def help_to_excel(module,filename):
    """
    完成指定模块的帮助信息的收集以及写入Excel文件的过程。
    输入：(1)module，指定的模块对象。(2)filename，写入帮助的Excel文件名。
    输出：str，“Ok”正常完成处理结束；其他，错误。
    """
    lines = get_help(module)
    str_write = write_help(module,lines,filename)
    return str_write
    
        
