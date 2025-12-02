import numpy as np
import pandas as pd
from pandas import ExcelWriter

from feature_values import settings_V2, settings_V3
from excel_Optimization import export_excel


def device_info(file1, file2, file4):
    df1_1 = pd.read_excel(file1, dtype=str, sheet_name="设备档案")
    df1_2 = pd.read_excel(file1, dtype=str, sheet_name="输出模板")
    df2 = pd.read_excel(file2, dtype=str)

    for index, series in df1_2.iterrows():
        if str(series['数据项（特征）编码'])[-3:] == '000':
            df1_2.at[index, '测点（通道）类型'] = '温度'
            df1_2.at[index, '数据项（特征）名称'] = '轴承温度'
            df1_2.at[index, '数据项（特征）类型'] = '温度'
            df1_2.at[index, '数据类型'] = '模拟量'
            df1_2.at[index, '单位'] = '℃'

    rows_to_be_deleted = []
    # for index, series in df1_2.iterrows():
    #     if (str(series['测点（点位）编码'])[-2:-1] in ['X', 'Y', 'Z'] and
    #             str(series['数据项（特征）编码'])[-3:] not in ['001', '003', '004', '005', '008'] and
    #             series['测点（通道）类型'] == '加速度'):
    #         rows_to_be_deleted.append(index)
    #
    # df1_2 = df1_2.drop(rows_to_be_deleted)  # 删除无线传感器的多余特征项

    # 建立device_info
    df4 = df1_2
    equipname_list = []
    equipcod_list = []
    pointname_list = []
    pointcod_list = []
    dataname_list = []
    datatype_list = []
    routecod_list = []
    unit_list = []
    pointtype_list = []
    sensor_list = []
    for index, series in df4.iterrows():
        if series['测点（通道）类型'] in ['加速度', "应力波", "位移", '速度', '温度', '电流谱', '电压谱', '声音',
                                        '径向位移', '轴向位移', '转速']:
            # 筛选出需要的行
            equipname_list.append(series['设备名称'])
            equipcod_list.append(series['设备编码'])
            pointname_list.append(series['测点（点位）名称'])
            pointcod_list.append(series['测点（点位）编码'])
            dataname_list.append(series['数据项（特征）名称'])
            datatype_list.append(series['数据项（特征）编码'])
            routecod_list.append(series['通道编码'])
            unit_list.append(series['单位'])
            pointtype_list.append(series['测点（通道）类型'])

            sensor_list.append(
                'ACCELEROMETER' if series['测点（通道）类型'] in ['加速度', '应力波', '电流谱', '电压谱', '声音',
                                                                '径向位移', '轴向位移']
                else 'VELOCITY' if series['测点（通道）类型'] == '速度'
                else 'DISPLACEMENT' if series['测点（通道）类型'] == '位移'
                else 'TEMPERATURE' if series['测点（通道）类型'] == '温度'
                else 'SPEED' if series['测点（通道）类型'] == '转速'
                else 'Unknown')

    df_deviceinfo = pd.DataFrame(
        columns=['区域', '设备名称', '设备编码', '测点名称', '测点编号', '数据项名称', '数据项编码', 'MAC地址',
                 '通道编号', '通道类型', '通道值', '单位', '测点类型']
    )
    df_deviceinfo['设备名称'] = equipname_list
    df_deviceinfo['设备编码'] = equipcod_list
    df_deviceinfo['测点名称'] = pointname_list
    df_deviceinfo['测点编号'] = pointcod_list
    df_deviceinfo['数据项名称'] = dataname_list
    df_deviceinfo['数据项编码'] = datatype_list
    df_deviceinfo['通道编号'] = routecod_list  # 创建dataframe对象，填入表头和四列数据
    df_deviceinfo['通道类型'] = sensor_list
    df_deviceinfo['单位'] = unit_list
    df_deviceinfo['测点类型'] = pointtype_list

    typecod_keylist = df2['数据项代号'].to_list()
    routekey_valuelist = df2["数据项（特征）类型"].to_list()
    dict1 = dict(zip(typecod_keylist, routekey_valuelist))  # 创建数据项编码后三位编号和通道值的对应字典

    equipcod_keylist = df1_1['*设备编码'].to_list()
    area_valuelist = df1_1["* 所属区域"].to_list()
    dict2 = dict(zip(equipcod_keylist, area_valuelist))
    # 创建一个标记删除的列表
    rows_to_drop = []
    for index, series in df_deviceinfo.iterrows():
        equip_cod = series['设备编码']
        if equip_cod in dict2:
            df_deviceinfo.at[index, '区域'] = dict2[equip_cod]
        else:
            df_deviceinfo.at[index, '区域'] = 'Unknown'  # 可以设置一个默认值，以便调试

        if series['测点编号'][-2:-1] in ['X', 'Y', 'Z', 'S']:
            mac2 = str(series['通道编号'])[:-2]
            df_deviceinfo.at[index, 'MAC地址'] = mac2  # 填写无线传感器的’MAC地址‘，实则为网关的sn号
        else:
            mac = str(series['通道编号'])[:-3]
            df_deviceinfo.at[index, 'MAC地址'] = mac

        data_type_cod3 = series['数据项编码'][-3:]
        # print("data_type_cod3:", data_type_cod3)
        if series['测点编号'][-2] in ['Z'] and series['测点类型'] == '加速度':
            print("series['测点编号'][-2]:", series['测点编号'][-2])
            if data_type_cod3 == '001':
                df_deviceinfo.at[index, '通道值'] = 'integratRMS'
            elif data_type_cod3 == '003':
                df_deviceinfo.at[index, '通道值'] = 'rmsValues'
            elif data_type_cod3 == '004':
                df_deviceinfo.at[index, '通道值'] = 'diagnosisPk'
            elif data_type_cod3 == '005':
                df_deviceinfo.at[index, '通道值'] = 'envelopEnergy'
            elif data_type_cod3 == '008':
                df_deviceinfo.at[index, '通道值'] = 'integratPk'
            elif data_type_cod3 == '000':
                df_deviceinfo.at[index, '通道值'] = 'TemperatureBot'
        elif series['测点编号'][-2] in ['X', 'Y'] and series['测点类型'] == '加速度':
            if data_type_cod3 == '001':
                df_deviceinfo.at[index, '通道值'] = 'integratRMS'
            elif data_type_cod3 == '003':
                df_deviceinfo.at[index, '通道值'] = 'rmsValues'
            elif data_type_cod3 == '004':
                df_deviceinfo.at[index, '通道值'] = 'diagnosisPk'
            elif data_type_cod3 == '005':
                rows_to_drop.append(index)
            elif data_type_cod3 == '008':
                df_deviceinfo.at[index, '通道值'] = 'integratPk'
            elif data_type_cod3 == '000':
                df_deviceinfo.at[index, '通道值'] = 'TemperatureBot'
        else:
            df_deviceinfo.at[index, '通道值'] = dict1.get(data_type_cod3, 'Unknown')  # 使用dict1查找通道值，默认Unknown
    # 删除标记的行
    df_deviceinfo_cleaned = df_deviceinfo.drop(rows_to_drop)
    # df_deviceinfo_cleaned.to_excel(file4, sheet_name='Sheet1', index=False)  # 保存为device-info_new.xlsx
    export_excel(df_deviceinfo_cleaned, file4, "deviceinfo")
    # return df_deviceinfo_cleaned


def tupuSetting_V2(file1, file4):
    df1_2 = pd.read_excel(file1, dtype=str, sheet_name="输出模板")

    df_tupu = pd.DataFrame(columns=['设备名称', '设备编码', '测点（点位）名称', '测点（点位）编码', '测点（通道）类型',
                                    '数据项（特征）名称', '数据项（特征）编码', '数据项（特征）类型', '数据类型', '单位',
                                    '抽样频率（Hz）', '采样时长(s)', '高通滤波（Hz）', '分析截止频率（Hz）',
                                    '采样点数（需求）'])
    pointcod_list_only = []
    for index, series in df1_2.iterrows():
        if series['测点（点位）编码'] not in pointcod_list_only and series['测点（通道）类型'] in list(settings_V2.keys()):
            if series['测点（通道）类型'] == "加速度" and series['测点（点位）编码'][-2:-1] in ['X', 'Y', 'Z']:
                series['测点（通道）类型'] = "无线传感器"
            number = len(settings_V2[series['测点（通道）类型']][0])
            new_data = pd.DataFrame({
                '设备名称': [series['设备名称']] * number,
                '设备编码': [series['设备编码']] * number,
                '测点（点位）名称': [series['测点（点位）名称']] * number,
                '测点（点位）编码': [series['测点（点位）编码']] * number,
                '测点（通道）类型': [series['测点（通道）类型']] * number,
                '数据项（特征）名称': settings_V2[series['测点（通道）类型']][0],
                '数据项（特征）编码': settings_V2[series['测点（通道）类型']][1],
                '数据项（特征）类型': settings_V2[series['测点（通道）类型']][2],
                '数据类型': settings_V2[series['测点（通道）类型']][3],
                '单位': settings_V2[series['测点（通道）类型']][4],
                '抽样频率（Hz）': settings_V2[series['测点（通道）类型']][5],
                '采样时长(s)': settings_V2[series['测点（通道）类型']][6],
                '高通滤波（Hz）': settings_V2[series['测点（通道）类型']][7],
                '分析截止频率（Hz）': settings_V2[series['测点（通道）类型']][8],
                '采样点数（需求）': settings_V2[series['测点（通道）类型']][9],
            })
            # 合并新的数据框到 df_tupu
            df_tupu = pd.concat([df_tupu, new_data], axis=0)
            pointcod_list_only.append(series['测点（点位）编码'])
        else:
            continue
        df_tupu['测点（通道）类型'] = df_tupu['测点（通道）类型'].replace('无线传感器', '加速度')
        # df_tupu.to_excel(file4, sheet_name='Sheet1', index=False)
        export_excel(df_tupu, file4, "tupusetting")


def tupuSetting_V3(file1, file4):
    df1_1 = pd.read_excel(file1, dtype=str, sheet_name="设备档案")
    df1_2 = pd.read_excel(file1, dtype=str, sheet_name="输出模板")

    df_tupu = pd.DataFrame(columns=['设备区域', '设备名称', '设备编码', '测点（点位）名称', '测点（点位）编码', '测点（通道）类型',
                                    '波形数据名称', '波形数据编码', '波形数据类型', '数据类型', '单位',
                                    '抽样频率（Hz）', '采样时长(s)', '高通滤波（Hz）', '分析截止频率（Hz）',
                                    '采样点数（需求）'])
    equipcod_keylist = df1_1['*设备编码'].to_list()
    area_valuelist = df1_1["* 所属区域"].to_list()
    dict2 = dict(zip(equipcod_keylist, area_valuelist))

    pointcod_list_only = []
    for index, series in df1_2.iterrows():
        equip_cod = series['设备编码']
        # if equip_cod in dict2:
        #     df_deviceinfo.at[index, '区域'] = dict2[equip_cod]
        # else:
        #     df_deviceinfo.at[index, '区域'] = 'Unknown'  # 可以设置一个默认值，以便调试
        if series['测点（点位）编码'] not in pointcod_list_only and series['测点（通道）类型'] in list(settings_V3.keys()):
            print(series['测点（通道）类型'], pointcod_list_only, series['测点（通道）类型'])
            if series['测点（通道）类型'] == "加速度" and series['测点（点位）编码'][-2:-1] in ['X', 'Y', 'Z']:
                series['测点（通道）类型'] = "无线传感器"
            number = len(settings_V3[series['测点（通道）类型']][0])
            new_data = pd.DataFrame({
                '设备区域': [dict2[equip_cod]] * number,
                '设备名称': [series['设备名称']] * number,
                '设备编码': [series['设备编码']] * number,
                '测点（点位）名称': [series['测点（点位）名称']] * number,
                '测点（点位）编码': [series['测点（点位）编码']] * number,
                '测点（通道）类型': [series['测点（通道）类型']] * number,
                '波形数据名称': settings_V3[series['测点（通道）类型']][0],
                '波形数据编码': settings_V3[series['测点（通道）类型']][1],
                '波形数据类型': settings_V3[series['测点（通道）类型']][2],
                '数据类型': settings_V3[series['测点（通道）类型']][3],
                '单位': settings_V3[series['测点（通道）类型']][4],
                '抽样频率（Hz）': settings_V3[series['测点（通道）类型']][5],
                '采样时长(s)': settings_V3[series['测点（通道）类型']][6],
                '高通滤波（Hz）': settings_V3[series['测点（通道）类型']][7],
                '分析截止频率（Hz）': settings_V3[series['测点（通道）类型']][8],
                '采样点数（需求）': settings_V3[series['测点（通道）类型']][9],
            })
            # 合并新的数据框到 df_tupu
            df_tupu = pd.concat([df_tupu, new_data], axis=0)
            pointcod_list_only.append(series['测点（点位）编码'])
        else:
            # print(series['测点（通道）类型'], pointcod_list_only, 1111111)
            continue
    # print(df_tupu.head())
    df_tupu['测点（通道）类型'] = df_tupu['测点（通道）类型'].replace('无线传感器', '加速度')
    # df_tupu.to_excel(file4, sheet_name='Sheet1', index=False)
    export_excel(df_tupu, file4, "tupusetting")


if __name__ == "__main__":
    device_info(r"H:\chaos项目资料\特征解析工具汇编\测试文件\data_all - 平台导入表(电流电压).xlsx",
                "后台文件/my_def_对应注释.xlsx", "device.xlsx")
    # tupuSetting_V3(r"H:\chaos项目资料\特征解析工具汇编\测试文件\data_all - 平台导入表(电流电压).xlsx", "tupusetting.xlsx")

