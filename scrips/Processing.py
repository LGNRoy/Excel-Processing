import xlwt
import numpy as np
import pandas as pd

def read_file():
    print("  READ_FILE")
    mainsource = pd.read_csv('../files/mainsource.csv')
    switch_stuff = pd.read_csv('../files/switch_stuff.csv')
    target_stuff = pd.read_csv('../files/target_stuff.csv')

    print("    mainsource  : {}".format(mainsource.shape))
    # print(mainsource[:2])
    print("    switch_stuff: {}".format(switch_stuff.shape))
    # print(switch_stuff[:2])
    print("    target_stuff: {}".format(target_stuff.shape))
    # print(target_stuff[:2])

    target_list = target_stuff[['序号','单位名称','身份证号码','姓名']]
    print("    target_list:  {}".format(target_list.shape))
    # print(target_list[:2])
    
    print("  _____________________________")
    return mainsource, switch_stuff, target_list

def join_file(mainsource, target_list):
    print("  JOIN_FILE")

    i_targetlist = 1
    i_mainsource = 1
    # last_index = i_mainsource # 为了避免在list里面有，在main里面没有

    result1 = []
    result2 = []

    for i_targetlist in range(1, len(target_list)+1):
        id = target_list[i_targetlist-1:i_targetlist]['身份证号码'].values[0]
        # print(id)
        while (True):
            id_current = mainsource[i_mainsource-1:i_mainsource]['身份证号码'].values[0]
            # 身份证号对应上了
            if (id_current == id):
                result1.append(mainsource[i_mainsource-1:i_mainsource]['2017年月收入'].values[0])
                result2.append(mainsource[i_mainsource-1:i_mainsource]['2018年养老月基数'].values[0])
                break
            # 没对上
            else: 
                i_mainsource += 1
                

    result = target_list
    result.insert(result.shape[1],'2017年月收入',result1)
    result.insert(result.shape[1],'2018年养老月基数',result2)
    print("    result:       {}".format(result.shape))
    print(result[:2])
    print("  _____________________________")
    return result

def join_file_swith(switch_stuff, target_list):
    print("  JOIN_FILE_SWITH")

    switch_list = switch_stuff[['姓名','调动生效时间']]
    print("    switch_list:  {}".format(switch_list.shape))
    # print(target_list[:2])

    i_targetlist = 1
    i_switchlist = 1

    result1 = []

    for i_targetlist in range(1, len(target_list)+1):
        name = target_list[i_targetlist-1:i_targetlist]['姓名'].values[0]
        time_str = ''
        for i_switchlist in range(1, len(switch_list)+1):
            name_current = switch_list[i_switchlist-1:i_switchlist]['姓名'].values[0]
            # 身份证号对应上了
            if (name_current == name):
                time_str = switch_list[i_switchlist-1:i_switchlist]['调动生效时间'].values[0]
                break
        time = 12 
        
        if (time_str != ''):
            time = int(time_str.split('年')[1].split('月')[0])
        result1.append(time)

    result = target_list
    result.insert(result.shape[1],'调动生效时间',result1)
    print("    result:       {}".format(result.shape))
    print(result[:2])
    print("  _____________________________")
    return result

def write_file(dataset ,address):
    print("  WRITE_FILE")

    wb = xlwt.Workbook()
    ws = wb.add_sheet('Sheet1')

    style1 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',num_format_str='#,##0.00')
    
    i, j = 0, 0
    print("    写入表头")
    for col in dataset.columns:
        ws.write(i, j, col)
        j += 1
    i += 1
    j = 0

    print("    写入数据")
    for row in dataset.values:
        for cell in row:
            ws.write(i, j, cell)
            j += 1
        i += 1
        j = 0
            
    wb.save('../files/result.xls')
    print("    success")
    return True

mainsource, switch_stuff, target_list = read_file()
result = join_file(mainsource, target_list)
result_switch = join_file_swith(switch_stuff, result)
write_file(result ,'../files/result.xlsx')

# for col in result.columns:
#     print(col)
# # print(result.columns)
# print(result[:3].values)