import os
import numpy as np
import pandas as pd
from pandas import DataFrame, ExcelWriter


# 确定输出模板
def output_template(parm_data, bearing_data):
    """
    :param parm_data: 一行设备参数
    :return: 可计算的特征值列表('v','null')
    """

    def ismy_null(value):
        return not (pd.isna(value) or value == "" or value is None or pd.isnull(value) or value == "/" or value == "\\")

    # 确定可计算的特征值列表
    (vel_pass_rms, vel_low_rms, acc_rms, acc_p, vibration_impulse, acc_kurtosis, acc_skew, vel_p,
     RF_1X, RF_2X, RF_3X, RF_4X, RF_5X, RF_1_2X, RF_1_3X, RF_1_4X, RF_1_5X, DPF_1X, DPF_2X, DPF_3X, DPF_4X, DPF_5X,
     GDE_ratio_1X, GDE_ratio_2X, GDE_ratio_3X, GDE_ratio_4X, GDE_ratio_5X, RFE_ratio_1X, RFE_ratio_2X, RFE_ratio_3X,
     RFE_ratio_4X, RFE_ratio_5X, RLE_ratio_1X, RLE_ratio_2X, RLE_ratio_3X, RLE_ratio_4X, RLE_ratio_5X, BPFI_1X, BPFI_2X,
     BPFI_3X, BPFI_4X, BPFI_5X, BPFO_1X, BPFO_2X, BPFO_3X, BPFO_4X, BPFO_5X, FTF_1X, FTF_2X, FTF_3X, FTF_4X, FTF_5X,
     BSF_1X, BSF_2X, BSF_3X, BSF_4X, BSF_5X, GMF_1X, GMF_2X, GMF_3X, GMF_4X, GMF_5X, GLE_sum_1X, GLE_sum_2X, GLE_sum_3X,
     GLE_sum_4X, GLE_sum_5X, GUE_sum_1X, GUE_sum_2X, GUE_sum_3X, GUE_sum_4X, GUE_sum_5X, Whirl_energy_sum, BPF_1X,
     BPF_2X, BPF_3X, BPF_4X, BPF_5X, DBPF_1X, DBPF_2X, DBPF_3X, DBPF_4X, DBPF_5X, ISE_sum, GSE_sum, EDF1_1X, EDF1_2X,
     EDF1_3X, EDF1_4X, EDF1_5X, EDF2_1X, EDF2_2X, EDF2_3X, EDF2_4X, EDF2_5X, EDF1_ratio_1X, EDF1_ratio_2X,
     EDF1_ratio_3X, EDF1_ratio_4X, EDF1_ratio_5X, EDF2_ratio_1X, EDF2_ratio_2X, EDF2_ratio_3X, EDF2_ratio_4X,
     EDF2_ratio_5X, EDF1_sum, EDF2_sum, DCValues_acc, vel_pass_rms_sudu, vel_low_rms_sudu, vel_p_sudu, DCValues_temp,
     ylb_SWE, ylb_SWPE, ylb_SWPA, ylb_vel_rms, ylb_kur, ylb_acc_rms, ylb_impulse, DCValues_ylb, Current_RMS, Current_PK,
     Current_CF, Current_Power_RMS, Current_THDF, Current_Odd_THD, Current_Even_THD, Current_Pos_THD, Current_Neg_THD,
     Current_Zero_THD, Current_Total_THD, Voltage_RMS, Voltage_PK, Voltage_CF, Voltage_Power_RMS, Voltage_THDF,
     Voltage_Odd_THD, Voltage_Even_THD, Voltage_Pos_THD, Voltage_Neg_THD, Voltage_Zero_THD, Voltage_Total_THD,
     Noise_RMS, Noise_PK, Noise_Kurt, Noise_Imp, dis_voltgap, dis_pp, dis_peak, dis_rms, dis_amp1x, dis_phase2x,
     dis_phase2x, dis_amp2x, dis_phase_1_2x, dis_amp_1_2x, dis_ampRt, dis_kurt, dis_skew, dis_mean, dis_amp3x,
     dis_amp4x, dis_amp5x, dis_amp_1_2x, dis_amp_1_4x, dis_amp_1_5x, DCValues_mc, acc_rms_mc, acc_peak_mc, acc_kurt_mc,
     acc_skew_mc, acc_crest_mc, acc_shape_mc, acc_pulse_mc, acc_margin_mc, hf_impulse_mc, lf_impulse_mc,
     vel_pass_rms_mc, speed) = ['null'] * 179

    eq_name, eq_code, point_name, point_code, channel_id, sensor_type, DW_type, L, N, nc, n, f0, m, Bearing_designation, \
        Manufacturer, Z, vane, G_vane, EDF1, EDF2, fc1, fb1, fc2, fb2, F_min1, F_max1, F_min2, F_max2 = parm_data
    # 中间计算值
    PPF = "/"

    if sensor_type == '应力波':
        ylb_SWE, ylb_SWPE, ylb_SWPA, ylb_vel_rms, ylb_kur, ylb_acc_rms, ylb_impulse, DCValues_ylb = ['v'] * 8
    if sensor_type == '加速度':
        # 判断加速度传感器是无线还是有线
        if str(point_code)[-2:-1] not in ['X', 'Y', 'Z']:
            vel_pass_rms, vel_low_rms, acc_rms, acc_p, vibration_impulse, acc_kurtosis, acc_skew, vel_p, DCValues_acc = [
                                                                                                                            'v'] * 9
        elif point_code[-2:-1] in ['X', 'Y']:
            vel_pass_rms, acc_rms, acc_p, vel_p = ['v'] * 4
        elif point_code[-2:-1] in ['Z']:
            vel_pass_rms, acc_rms, acc_p, vibration_impulse, vel_p = ['v'] * 5
        # 非应力波、N为数字或外部键相、N不为空时，得到倍频特征值
        if sensor_type != '应力波' and ismy_null(N) and ismy_null(N):
            RF_1X, RF_2X, RF_3X, RF_4X, RF_5X = ['v'] * 5
            RF_1_2X, RF_1_3X, RF_1_4X, RF_1_5X = ['v'] * 4
        # f0不为空时，得到电机故障特征
        if ismy_null(f0):
            DPF_1X, DPF_2X, DPF_3X, DPF_4X, DPF_5X = ['v'] * 5
        if ismy_null(n) and ismy_null(nc) and ismy_null(f0):
            PPF = 'v'
            GDE_ratio_1X, GDE_ratio_2X, GDE_ratio_3X, GDE_ratio_4X, GDE_ratio_5X = ['v'] * 5
        if ismy_null(PPF) and ismy_null(N):
            RFE_ratio_1X, RFE_ratio_2X, RFE_ratio_3X, RFE_ratio_4X, RFE_ratio_5X = ['v'] * 5
        if ismy_null(PPF) and ismy_null(m):
            RLE_ratio_1X, RLE_ratio_2X, RLE_ratio_3X, RLE_ratio_4X, RLE_ratio_5X = ['v'] * 5
        # N不为空时且轴承型号、轴承厂家都不为空时，得到轴承故障特征
        if ismy_null(N) and ismy_null(str(Bearing_designation)) and ismy_null(str(Manufacturer)):
            bearing_one = bearing_data.loc[
                (bearing_data['轴承型号'] == str(Bearing_designation).split(".")[0]) & (
                        bearing_data['轴承厂家'] == Manufacturer)]
            if not bearing_one.empty:
                BPFI_1X, BPFI_2X, BPFI_3X, BPFI_4X, BPFI_5X = ['v'] * 5
                BPFO_1X, BPFO_2X, BPFO_3X, BPFO_4X, BPFO_5X = ['v'] * 5
                FTF_1X, FTF_2X, FTF_3X, FTF_4X, FTF_5X = ['v'] * 5
                BSF_1X, BSF_2X, BSF_3X, BSF_4X, BSF_5X = ['v'] * 5
        # N不为空时且Z不为空时，得到齿轮故障特征
        if ismy_null(N) and ismy_null(Z):
            GMF_1X, GMF_2X, GMF_3X, GMF_4X, GMF_5X = ['v'] * 5
            GLE_sum_1X, GLE_sum_2X, GLE_sum_3X, GLE_sum_4X, GLE_sum_5X = ['v'] * 5
            GUE_sum_1X, GUE_sum_2X, GUE_sum_3X, GUE_sum_4X, GUE_sum_5X = ['v'] * 5
        # N不为空时且vane不为空时，得到叶片故障特征（其他故障特征）
        if ismy_null(N) and ismy_null(vane):
            BPF_1X, BPF_2X, BPF_3X, BPF_4X, BPF_5X = ['v'] * 5
            ISE_sum = "v"
        if ismy_null(N) and ismy_null(G_vane):
            DBPF_1X, DBPF_2X, DBPF_3X, DBPF_4X, DBPF_5X = ['v'] * 5
            GSE_sum = "v"
        if ismy_null(N) and Bearing_designation == "滑动轴承":
            Whirl_energy_sum = "v"
        # EDF1不为空时，得到自定义故障特征
        if ismy_null(EDF1):
            EDF1_1X, EDF1_2X, EDF1_3X, EDF1_4X, EDF1_5X = ['v'] * 5
        if ismy_null(EDF2):
            EDF2_1X, EDF2_2X, EDF2_3X, EDF2_4X, EDF2_5X = ['v'] * 5
        if ismy_null(fc1) and ismy_null(fb1):
            EDF1_ratio_1X, EDF1_ratio_2X, EDF1_ratio_3X, EDF1_ratio_4X, EDF1_ratio_5X = ['v'] * 5
        if ismy_null(fc2) and ismy_null(fb2):
            EDF2_ratio_1X, EDF2_ratio_2X, EDF2_ratio_3X, EDF2_ratio_4X, EDF2_ratio_5X = ['v'] * 5
        if ismy_null(F_min1) and ismy_null(F_max1):
            EDF1_sum = "v"
        if ismy_null(F_min2) and ismy_null(F_max2):
            EDF2_sum = "v"
    elif sensor_type == '速度':
        vel_pass_rms_sudu, vel_low_rms_sudu, vel_p_sudu = ['v'] * 3
    elif sensor_type == '电流谱':
        (Current_RMS, Current_PK, Current_CF, Current_Power_RMS, Current_THDF, Current_Odd_THD, Current_Even_THD,
         Current_Pos_THD, Current_Neg_THD, Current_Zero_THD, Current_Total_THD) = ['v'] * 11
    elif sensor_type == '电压谱':
        (Voltage_RMS, Voltage_PK, Voltage_CF, Voltage_Power_RMS, Voltage_THDF, Voltage_Odd_THD, Voltage_Even_THD,
         Voltage_Pos_THD, Voltage_Neg_THD, Voltage_Zero_THD, Voltage_Total_THD) = ['v'] * 11
    elif sensor_type == '电压谱':
        (Voltage_RMS, Voltage_PK, Voltage_CF, Voltage_Power_RMS, Voltage_THDF, Voltage_Odd_THD, Voltage_Even_THD,
         Voltage_Pos_THD, Voltage_Neg_THD, Voltage_Zero_THD, Voltage_Total_THD) = ['v'] * 11
    elif sensor_type == '声音':
        Noise_RMS, Noise_PK, Noise_Kurt, Noise_Imp = ['v'] * 4
    elif sensor_type == '径向位移':
        (dis_voltgap, dis_pp, dis_peak, dis_rms, dis_amp1x, dis_phase2x, dis_phase2x, dis_amp2x, dis_phase_1_2x,
         dis_amp_1_2x, dis_ampRt, dis_kurt, dis_skew, dis_amp3x, dis_amp4x, dis_amp5x, dis_amp_1_2x, dis_amp_1_4x,
         dis_amp_1_5x) = ['v'] * 19
    elif sensor_type == '轴向位移':
        dis_mean = ['v'] * 1
    elif sensor_type == '冲击脉冲':
        (acc_rms_mc, acc_peak_mc, acc_kurt_mc, acc_skew_mc, acc_crest_mc, acc_shape_mc, acc_pulse_mc, acc_margin_mc,
         hf_impulse_mc, lf_impulse_mc,
         vel_pass_rms_mc, DCValues_mc) = ['v'] * 12
    elif sensor_type == '温度':
        DCValues_temp = ['v']* 1
    elif sensor_type == '转速':
        speed = ['v'] * 1

    res_type = [vel_pass_rms, vel_low_rms, acc_rms, acc_p, vibration_impulse, acc_kurtosis, acc_skew, vel_p,
                RF_1X, RF_2X, RF_3X, RF_4X, RF_5X, RF_1_2X, RF_1_3X, RF_1_4X, RF_1_5X, DPF_1X, DPF_2X, DPF_3X, DPF_4X,
                DPF_5X, GDE_ratio_1X, GDE_ratio_2X, GDE_ratio_3X, GDE_ratio_4X, GDE_ratio_5X, RFE_ratio_1X, RFE_ratio_2X,
                RFE_ratio_3X, RFE_ratio_4X, RFE_ratio_5X, RLE_ratio_1X, RLE_ratio_2X, RLE_ratio_3X, RLE_ratio_4X, RLE_ratio_5X,
                BPFI_1X, BPFI_2X, BPFI_3X, BPFI_4X, BPFI_5X, BPFO_1X, BPFO_2X, BPFO_3X, BPFO_4X, BPFO_5X, FTF_1X, FTF_2X, FTF_3X, FTF_4X,
                FTF_5X, BSF_1X, BSF_2X, BSF_3X, BSF_4X, BSF_5X, GMF_1X, GMF_2X, GMF_3X, GMF_4X, GMF_5X, GLE_sum_1X, GLE_sum_2X,
                GLE_sum_3X, GLE_sum_4X, GLE_sum_5X, GUE_sum_1X, GUE_sum_2X, GUE_sum_3X, GUE_sum_4X, GUE_sum_5X, Whirl_energy_sum,
                BPF_1X, BPF_2X, BPF_3X, BPF_4X, BPF_5X, DBPF_1X, DBPF_2X, DBPF_3X, DBPF_4X, DBPF_5X, ISE_sum, GSE_sum, EDF1_1X,
                EDF1_2X, EDF1_3X, EDF1_4X, EDF1_5X, EDF2_1X, EDF2_2X, EDF2_3X, EDF2_4X, EDF2_5X, EDF1_ratio_1X, EDF1_ratio_2X,
                EDF1_ratio_3X, EDF1_ratio_4X, EDF1_ratio_5X, EDF2_ratio_1X, EDF2_ratio_2X, EDF2_ratio_3X, EDF2_ratio_4X,
                EDF2_ratio_5X, EDF1_sum, EDF2_sum, DCValues_acc, vel_pass_rms_sudu, vel_low_rms_sudu, vel_p_sudu,
                DCValues_temp, ylb_SWE, ylb_SWPE, ylb_SWPA, ylb_vel_rms, ylb_kur, ylb_acc_rms, ylb_impulse, DCValues_ylb, Current_RMS, Current_PK,
                Current_CF, Current_Power_RMS, Current_THDF, Current_Odd_THD, Current_Even_THD, Current_Pos_THD, Current_Neg_THD,
                Current_Zero_THD, Current_Total_THD, Voltage_RMS, Voltage_PK, Voltage_CF, Voltage_Power_RMS,
                Voltage_THDF, Voltage_Odd_THD, Voltage_Even_THD, Voltage_Pos_THD, Voltage_Neg_THD, Voltage_Zero_THD,
                Voltage_Total_THD, Noise_RMS, Noise_PK, Noise_Kurt, Noise_Imp, dis_voltgap, dis_pp, dis_peak, dis_rms, dis_amp1x,
                dis_phase2x, dis_phase2x, dis_amp2x, dis_phase_1_2x, dis_amp_1_2x, dis_ampRt, dis_kurt, dis_skew, dis_mean,
                dis_amp3x, dis_amp4x, dis_amp5x, dis_amp_1_2x, dis_amp_1_4x, dis_amp_1_5x, DCValues_mc, acc_rms_mc, acc_peak_mc,
                acc_kurt_mc, acc_skew_mc, acc_crest_mc, acc_shape_mc, acc_pulse_mc, acc_margin_mc, hf_impulse_mc, lf_impulse_mc, vel_pass_rms_mc, speed]
    return res_type


def output_template_all(excel_path, my_deftable, output_path, need_channel_id=True):
    """
    :param my_def: 特征对应注释
    :param excel_path: 设备参数表格位置
    :param my_null: 空值格式
    :param output_path: 输出的表格位置
    :return: 在指定位置输出模板表格
    """
    template_dataframe = pd.read_excel(my_deftable)
    input_data = pd.read_excel(excel_path, sheet_name='输入参数')
    desired_columns = ['设备名称', '设备编码', '测点名称', '测点编码', '通道编码', '传感器类型', '网关型号',
                       '传感器量程', '工作转速', '电机额定转速', '电机同步转速', '电源频率', '电机转子条数', '轴承型号',
                       '轴承生产厂家', '齿轮齿数Z', '叶轮叶片数目', '导叶叶片数目', '自定义频率1', '自定义频率2',
                       '自定义能量比1-中心频率', '自定义能量比1-边带频率', '自定义能量比2-中心频率',
                       '自定义能量比2-边带频率', '自定义频带能量和1-频率下限', '自定义频带能量和1-频率上限',
                       '自定义频带能量和2-频率下限', '自定义频带能量和2-频率上限']
    input_data = input_data[desired_columns]
    input_device_profile = pd.read_excel(excel_path, sheet_name='设备档案')
    bearing_data = pd.read_excel("后台文件/Bearing.xlsx", sheet_name="轴承库数据库配置")
    """找出设备档案和输入参数中设备不对应的地方"""
    # 去重，转换为集合, 找出缺少的项
    missing_in_profile = set(input_data["设备名称"].dropna()) - set(input_device_profile["*设备名称"].dropna())
    missing_in_data = set(input_device_profile["*设备名称"].dropna()) - set(input_data["设备名称"].dropna())
    output_dir = os.path.dirname(output_path)
    output_file = os.path.join(output_dir, '设备缺失项.txt')
    output_file_True = False
    if missing_in_profile or missing_in_data:
        output_file_True = True
        with open(output_file, 'w') as file:
            if missing_in_profile:
                file.write(f"设备参数中缺少: {missing_in_profile}\n")
            if missing_in_data:
                file.write(f"输入参数中缺少: {missing_in_data}\n")

    columns_name = ['设备名称', '设备编码', '测点（点位）名称', '测点（点位）编码', '测点（通道）类型', '数据项（特征）名称',
                    '数据项（特征）编码', '数据项（特征）类型', '数据类型', '单位', '通道编码']
    value_name = list(template_dataframe['数据项（特征）名称'])
    value_code = list(template_dataframe['数据项代号'])
    value_fea_type = list(template_dataframe['数据项（特征）类型'])
    # print(value_fea_type)
    value_type = list(template_dataframe['数据类型'])
    value_unit = list(template_dataframe['单位'])
    tmp_data = []
    input_data = input_data[input_data['设备名称'].notna()]
    for df_index, df_row in input_data.iterrows():
        eq_name, eq_code, point_name, point_code, channel_id, sensor_type = df_row.iloc[:6]
        if eq_name is None or eq_name == "" or eq_name == np.nan or pd.isna(eq_name) or pd.isnull(eq_name):
            continue
        result_type = output_template(df_row, bearing_data)
        # print(result_type)
        # if sensor_type == '应力波':
        #     # tmp_data.append(
        #     #     [eq_name, eq_code, point_name, point_code, sensor_type, '偏置电压',
        #     #      str(point_code) + '999', '应力波特征', '模拟量', 'V', channel_id])
        #     pass
        # elif sensor_type == "加速度" and point_code[-2:-1] not in ['X', 'Y', 'Z']:
        #     # tmp_data.append(
        #     #     [eq_name, eq_code, point_name, point_code, sensor_type, '偏置电压',
        #     #      str(point_code) + '999', '时域特征', '模拟量', 'V', channel_id])
        #     pass
        # elif sensor_type == "速度":
        #     # tmp_data.append(
        #     #     [eq_name, eq_code, point_name, point_code, sensor_type, '偏置电压',
        #     #      str(point_code) + '999', '时域特征', '模拟量', 'V', channel_id])
        #     pass
        # elif sensor_type == "冲击脉冲":
        #     # tmp_data.append(
        #     #     [eq_name, eq_code, point_name, point_code, sensor_type, '偏置电压',
        #     #      str(point_code) + '999', '时域特征', '模拟量', 'V', channel_id])
        #     pass
        # elif sensor_type == "温度":
        #     # tmp_data.append(
        #     #     [eq_name, eq_code, point_name, point_code, sensor_type, '轴承温度',
        #     #      str(point_code) + '000', '温度', '模拟量', '℃', channel_id])
        #     pass
        # elif sensor_type == "转速":
        #     # tmp_data.append(
        #     #     [eq_name, eq_code, point_name, point_code, sensor_type, '转速',
        #     #      str(point_code) + '777', '工艺', '模拟量', 'rpm', channel_id])
        #     pass
        tmp_data += [
            [eq_name, eq_code, point_name, point_code, sensor_type, value_name[res_index],
             str(point_code) + str(value_code[res_index]).zfill(3), value_type[res_index],
             "模拟量", value_unit[res_index], channel_id
             ] for res_index, res_type_one in enumerate(result_type) if res_type_one != 'null'
        ]
    output_data = DataFrame(tmp_data, columns=columns_name)
    if not need_channel_id:
        output_data = output_data.drop(columns=['通道编码'])

    # export_to_excel(output_data, output_path, "输出模板")
    # 自适应列宽实现
    def excel_widths(excel_dataframe):
        # 计算表头的字符宽度
        column_widths = (excel_dataframe.columns.to_series()
                         .apply(lambda x: len(x.encode('utf-8'))).values
                         ) * 0.8
        #  计算每列的最大字符宽度
        max_widths = (excel_dataframe.astype(str)
                      .applymap(lambda x: len(x.encode('utf-8')))
                      .agg(max).values
                      ) * 0.8
        # 计算整体最大宽度
        widths = np.max([column_widths, max_widths], axis=0)
        return widths

    with ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        input_device_profile.to_excel(writer, sheet_name='设备档案', index=False)
        worksheet_profile = writer.sheets['设备档案']
        # 设置边框格式
        border_format = workbook.add_format({'border': 1, 'border_color': 'black'})
        worksheet_profile.conditional_format(0, 0, input_device_profile.shape[0], input_device_profile.shape[1] - 1,
                                             {'type': 'no_errors', 'format': border_format})

        title_format = workbook.add_format(
            {'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#BFBFBF'})
        for col_num, value in enumerate(input_device_profile.columns.values):
            worksheet_profile.write(0, col_num, value, title_format)
        wrap_format = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
        for i in range(input_device_profile.shape[1]):
            worksheet_profile.set_column(i, i, 10, wrap_format)

        output_data.to_excel(writer, sheet_name="输出模板", index=False)
        worksheet_data = writer.sheets["输出模板"]
        worksheet_data.conditional_format(0, 0, output_data.shape[0], output_data.shape[1] - 1,
                                          {'type': 'no_errors', 'format': border_format})
        # 设置表头格式
        title_format = workbook.add_format(
            {'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#BFBFBF'})
        for col_num, value in enumerate(output_data.columns.values):
            worksheet_data.write(0, col_num, value, title_format)
        for i, width in enumerate(excel_widths(output_data)):
            worksheet_data.set_column(i, i, width)
        worksheet_data.set_column(4, 4, 20)
        worksheet_data.set_column(10, 10, 20)

    return output_file_True


if __name__ == "__main__":
    output_template_all(r"H:\chaos项目资料\特征解析工具汇编\测试文件\data_all(模板).xlsx",
                        r"后台文件\my_def_对应注释.xlsx",
                        r"H:\chaos项目资料\特征解析工具汇编\测试文件\data_all - 平台导入表(电流电压).xlsx", False)
