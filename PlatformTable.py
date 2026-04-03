import os
import numpy as np
import pandas as pd
from pandas import DataFrame, ExcelWriter

TEMPLATE_REQUIRED_COLUMNS = [
    '数据项（特征）名称',
    '数据项代号',
    '数据项（特征）类型',
    '数据类型',
    '单位',
    '通道类型',
]


def normalize_template_dataframe(template_dataframe: pd.DataFrame) -> pd.DataFrame:
    """
    模板标准化：
    1. 检查必要列
    2. 清洗空白
    3. 将“加速度”中 数据类型 != “时域特征” 的行，自动复制给
       “XY无线加速度” 和 “Z无线加速度”
    4. 构造唯一码：(通道类型, 数据项（特征）类型, 数据类型)
    5. 校验唯一码不能重复
    """
    missing_cols = [c for c in TEMPLATE_REQUIRED_COLUMNS if c not in template_dataframe.columns]
    if missing_cols:
        raise ValueError(f"模板表缺少必要列: {missing_cols}")

    df = template_dataframe.copy()

    for col in TEMPLATE_REQUIRED_COLUMNS:
        df[col] = df[col].astype(str).str.strip()

    # 去掉空白/伪空值
    df = df.replace("nan", "").replace("None", "")
    df = df[
        (df['数据项（特征）类型'] != "") &
        (df['数据类型'] != "") &
        (df['通道类型'] != "")
    ].copy()

    # =========================================================
    # 扩充模板：
    # 将 通道类型=加速度 且 数据类型!=时域特征 的行，
    # 复制给 XY无线加速度 和 Z无线加速度
    # =========================================================
    base_acc_rows = df[
        (df["通道类型"] == "加速度") &
        (df["数据类型"] != "时域特征")
    ].copy()

    if not base_acc_rows.empty:
        # 先构造当前已存在的键，避免复制后撞重
        existing_keys = set(zip(
            df["通道类型"],
            df["数据项（特征）类型"],
            df["数据类型"]
        ))

        extra_rows = []

        for target_channel_type in ["XY无线加速度", "Z无线加速度"]:
            cloned = base_acc_rows.copy()
            cloned["通道类型"] = target_channel_type

            cloned["__temp_key__"] = list(zip(
                cloned["通道类型"],
                cloned["数据项（特征）类型"],
                cloned["数据类型"]
            ))

            # 只保留当前模板中还不存在的键
            cloned = cloned[~cloned["__temp_key__"].isin(existing_keys)].copy()
            cloned = cloned.drop(columns="__temp_key__", errors="ignore")

            # 更新 existing_keys，防止两轮复制内部也重复
            existing_keys.update(zip(
                cloned["通道类型"],
                cloned["数据项（特征）类型"],
                cloned["数据类型"]
            ))

            if not cloned.empty:
                extra_rows.append(cloned)

        if extra_rows:
            df = pd.concat([df] + extra_rows, ignore_index=True)

    # =========================================================
    # 构造内部唯一码
    # =========================================================
    df["__template_key__"] = list(zip(
        df["通道类型"],
        df["数据项（特征）类型"],
        df["数据类型"]
    ))

    duplicated = df[df["__template_key__"].duplicated(keep=False)]
    if not duplicated.empty:
        dup_keys = duplicated["__template_key__"].tolist()
        raise ValueError(
            "模板表中存在重复的唯一码 (通道类型, 数据项（特征）类型, 数据类型)，请清理重复项：\n"
            f"{dup_keys}"
        )

    return df


def format_feature_code(point_code, feature_code) -> str:
    """
    生成输出表中的“数据项（特征）编码”
    """
    point_code = "" if pd.isna(point_code) else str(point_code).strip()

    if pd.isna(feature_code):
        raise ValueError("模板中的“数据项代号”不能为空")

    try:
        feature_code = str(int(float(feature_code))).zfill(3)
    except Exception:
        feature_code = str(feature_code).strip()

    return f"{point_code}{feature_code}"


def resolve_channel_type(sensor_type, point_code):
    """
    根据传感器类型 + 测点编码，解析模板筛选时使用的“通道类型”
    """
    sensor_type = "" if pd.isna(sensor_type) else str(sensor_type).strip()
    point_code_str = "" if pd.isna(point_code) else str(point_code).strip()

    if sensor_type == '加速度':
        axis_flag = point_code_str[-2:-1] if len(point_code_str) >= 2 else ""

        if axis_flag in ['X', 'Y']:
            return 'XY无线加速度'
        elif axis_flag in ['Z']:
            if point_code_str[-1] == "T":
                return 'S无线加速度'
            return 'Z无线加速度'
        else:
            return '加速度'

    return sensor_type


def output_template(parm_data, bearing_data):
    """
    返回当前测点可启用的“数据项（特征）类型”集合
    注意：这里只返回 特征类型，不返回 数据类型
    数据类型由 output_template_all 在模板中过滤时决定
    """

    def ismy_null(value):
        return not (
                pd.isna(value) or value == "" or value is None or
                pd.isnull(value) or value == "/" or value == "\\"
        )

    enabled_features = set()

    def enable(*feature_types):
        for ft in feature_types:
            enabled_features.add(ft)

    (eq_name, eq_code, point_name, point_code, channel_id, sensor_type, DW_type, L, N, nc, n, f0, m,
     Bearing_designation, Manufacturer, Z, vane, G_vane, EDF1, EDF2, fc1, fb1, fc2, fb2,
     F_min1, F_max1, F_min2, F_max2
     ) = parm_data

    PPF = "/"

    if sensor_type == '应力波':
        enable(
            'ylb_SWE', 'ylb_SWPE', 'ylb_SWPA', 'ylb_vel_rms',
            'ylb_kur', 'ylb_acc_rms', 'ylb_impulse', 'DCValues'
        )

    elif sensor_type == '加速度':
        # 判断加速度传感器是无线还是有线
        point_code_str = str(point_code)
        channel_type = resolve_channel_type(sensor_type, point_code_str)
        if channel_type == "加速度":
            enable('vel_pass_rms', 'vel_low_rms', 'acc_rms', 'acc_p','vibration_impulse', 'acc_kurtosis',
                   'acc_skew', 'vel_p', 'DCValues')
        elif channel_type == "XY无线加速度":
            enable('rmsValues', 'diagnosisPk', 'integratRMS', 'kurtosis')
        elif channel_type == "Z无线加速度":
            enable('rmsValues', 'diagnosisPk', 'integratRMS', 'envelopEnergy', 'kurtosis')
        elif channel_type == "S无线加速度":
            enable('TemperatureBot')

        # 倍频特征
        if ismy_null(N):
            enable(
                'RF_1X', 'RF_2X', 'RF_3X', 'RF_4X', 'RF_5X',
                'RF_1_2X', 'RF_1_3X', 'RF_1_4X', 'RF_1_5X'
            )

        # 电机故障特征
        if ismy_null(f0):
            enable('DPF_1X', 'DPF_2X', 'DPF_3X', 'DPF_4X', 'DPF_5X')

        if ismy_null(n) and ismy_null(nc) and ismy_null(f0):
            PPF = 'v'
            enable('GDE_ratio_1X', 'GDE_ratio_2X', 'GDE_ratio_3X', 'GDE_ratio_4X', 'GDE_ratio_5X')

        if ismy_null(PPF) and ismy_null(N):
            enable('RFE_ratio_1X', 'RFE_ratio_2X', 'RFE_ratio_3X', 'RFE_ratio_4X', 'RFE_ratio_5X')

        if ismy_null(PPF) and ismy_null(m):
            enable('RLE_ratio_1X', 'RLE_ratio_2X', 'RLE_ratio_3X', 'RLE_ratio_4X', 'RLE_ratio_5X')

        # 轴承故障特征
        if ismy_null(N) and ismy_null(str(Bearing_designation)) and ismy_null(str(Manufacturer)):
            bearing_one = bearing_data.loc[
                (bearing_data['轴承型号'] == str(Bearing_designation).split(".")[0]) &
                (bearing_data['轴承厂家'] == Manufacturer)
                ]
            if not bearing_one.empty:
                enable('BPFI_1X', 'BPFI_2X', 'BPFI_3X', 'BPFI_4X', 'BPFI_5X')
                enable('BPFO_1X', 'BPFO_2X', 'BPFO_3X', 'BPFO_4X', 'BPFO_5X')
                enable('FTF_1X', 'FTF_2X', 'FTF_3X', 'FTF_4X', 'FTF_5X')
                enable('BSF_1X', 'BSF_2X', 'BSF_3X', 'BSF_4X', 'BSF_5X')

        # 齿轮故障特征
        if ismy_null(N) and ismy_null(Z):
            enable('GMF_1X', 'GMF_2X', 'GMF_3X', 'GMF_4X', 'GMF_5X')
            enable('GLE_sum_1X', 'GLE_sum_2X', 'GLE_sum_3X', 'GLE_sum_4X', 'GLE_sum_5X')
            enable('GUE_sum_1X', 'GUE_sum_2X', 'GUE_sum_3X', 'GUE_sum_4X', 'GUE_sum_5X')

        # 叶片故障特征
        if ismy_null(N) and ismy_null(vane):
            enable('BPF_1X', 'BPF_2X', 'BPF_3X', 'BPF_4X', 'BPF_5X', 'ISE_sum')

        if ismy_null(N) and ismy_null(G_vane):
            enable('DBPF_1X', 'DBPF_2X', 'DBPF_3X', 'DBPF_4X', 'DBPF_5X', 'GSE_sum')

        if ismy_null(N) and Bearing_designation == "滑动轴承":
            enable('Whirl_energy_sum')

        # 自定义故障特征
        if ismy_null(EDF1):
            enable('EDF1_1X', 'EDF1_2X', 'EDF1_3X', 'EDF1_4X', 'EDF1_5X')

        if ismy_null(EDF2):
            enable('EDF2_1X', 'EDF2_2X', 'EDF2_3X', 'EDF2_4X', 'EDF2_5X')

        if ismy_null(fc1) and ismy_null(fb1):
            enable('EDF1_ratio_1X', 'EDF1_ratio_2X', 'EDF1_ratio_3X', 'EDF1_ratio_4X', 'EDF1_ratio_5X')

        if ismy_null(fc2) and ismy_null(fb2):
            enable('EDF2_ratio_1X', 'EDF2_ratio_2X', 'EDF2_ratio_3X', 'EDF2_ratio_4X', 'EDF2_ratio_5X')

        if ismy_null(F_min1) and ismy_null(F_max1):
            enable('EDF1_sum')

        if ismy_null(F_min2) and ismy_null(F_max2):
            enable('EDF2_sum')

    elif sensor_type == '速度':
        enable('vel_pass_rms_sudu', 'vel_low_rms_sudu', 'vel_p_sudu')

    elif sensor_type == '电流谱':
        enable(
            'Current_RMS', 'Current_PK', 'Current_CF', 'Current_Power_RMS', 'Current_THDF',
            'Current_Odd_THD', 'Current_Even_THD', 'Current_Pos_THD', 'Current_Neg_THD',
            'Current_Zero_THD', 'Current_Total_THD'
        )

    elif sensor_type == '电压谱':
        enable(
            'Voltage_RMS', 'Voltage_PK', 'Voltage_CF', 'Voltage_Power_RMS', 'Voltage_THDF',
            'Voltage_Odd_THD', 'Voltage_Even_THD', 'Voltage_Pos_THD', 'Voltage_Neg_THD',
            'Voltage_Zero_THD', 'Voltage_Total_THD'
        )

    elif sensor_type == '声音':
        enable('Noise_RMS', 'Noise_PK', 'Noise_Kurt', 'Noise_Imp')

    elif sensor_type == '径向位移':
        enable(
            'dis_voltgap', 'dis_pp', 'dis_peak', 'dis_rms', 'dis_amp1x', 'dis_phase2x',
            'dis_amp2x', 'dis_phase_1_2x', 'dis_amp_1_2x', 'dis_ampRt',
            'dis_kurt', 'dis_skew', 'dis_amp3x', 'dis_amp4x', 'dis_amp5x',
            'dis_amp_1_4x', 'dis_amp_1_5x'
        )

    elif sensor_type == '轴向位移':
        enable('dis_mean')

    elif sensor_type == '冲击脉冲':
        enable(
            'acc_rms_mc', 'acc_peak_mc', 'acc_kurt_mc', 'acc_skew_mc', 'acc_crest_mc',
            'acc_shape_mc', 'acc_pulse_mc', 'acc_margin_mc',
            'hf_impulse_mc', 'lf_impulse_mc', 'vel_pass_rms_mc', 'DCValues'
        )

    elif sensor_type == '温度':
        enable('DCValues')

    elif sensor_type == '转速':
        enable('speed')

    return enabled_features


def select_template_rows(template_dataframe: pd.DataFrame, channel_type: str,
                         enabled_feature_types: set) -> pd.DataFrame:
    """
    先按通道类型，再按特征类型过滤模板
    """
    channel_type = "" if pd.isna(channel_type) else str(channel_type).strip()

    candidate_rows = template_dataframe[
        template_dataframe["通道类型"].astype(str).str.strip() == channel_type
        ].copy()

    if candidate_rows.empty:
        return candidate_rows

    selected_rows = candidate_rows[
        candidate_rows["数据项（特征）类型"].isin(enabled_feature_types)
    ].copy()

    return selected_rows


def output_template_all(excel_path, my_deftable, output_path, need_channel_id=True):
    """
    :param my_deftable: 特征对应注释模板
    :param excel_path: 设备参数表格位置
    :param output_path: 输出位置
    :param need_channel_id: 是否保留通道编码
    :return: 是否生成了设备缺失项.txt
    """
    template_dataframe = pd.read_excel(my_deftable)
    template_dataframe = normalize_template_dataframe(template_dataframe)

    input_data = pd.read_excel(excel_path, sheet_name='输入参数')
    desired_columns = [
        '设备名称', '设备编码', '测点名称', '测点编码', '通道编码', '传感器类型', '网关型号',
        '传感器量程', '工作转速', '电机额定转速', '电机同步转速', '电源频率', '电机转子条数',
        '轴承型号', '轴承生产厂家', '齿轮齿数Z', '叶轮叶片数目', '导叶叶片数目',
        '自定义频率1', '自定义频率2',
        '自定义能量比1-中心频率', '自定义能量比1-边带频率',
        '自定义能量比2-中心频率', '自定义能量比2-边带频率',
        '自定义频带能量和1-频率下限', '自定义频带能量和1-频率上限',
        '自定义频带能量和2-频率下限', '自定义频带能量和2-频率上限'
    ]
    input_data = input_data[desired_columns]
    input_device_profile = pd.read_excel(excel_path, sheet_name='设备档案')
    bearing_data = pd.read_excel("后台文件/Bearing.xlsx", sheet_name="轴承库数据库配置")

    # 找出设备档案和输入参数中设备不对应的地方
    missing_in_profile = set(input_data["设备名称"].dropna()) - set(input_device_profile["*设备名称"].dropna())
    missing_in_data = set(input_device_profile["*设备名称"].dropna()) - set(input_data["设备名称"].dropna())

    output_dir = os.path.dirname(output_path)
    output_file = os.path.join(output_dir, '设备缺失项.txt')
    output_file_True = False
    if missing_in_profile or missing_in_data:
        output_file_True = True
        with open(output_file, 'w', encoding='utf-8') as file:
            if missing_in_profile:
                file.write(f"设备参数中缺少: {missing_in_profile}\n")
            if missing_in_data:
                file.write(f"输入参数中缺少: {missing_in_data}\n")

    columns_name = [
        '设备名称', '设备编码', '测点（点位）名称', '测点（点位）编码', '测点（通道）类型',
        '数据项（特征）名称', '数据项（特征）编码', '数据项（特征）类型', '数据类型', '单位', '通道编码'
    ]

    tmp_data = []
    input_data = input_data[input_data['设备名称'].notna()]

    for _, df_row in input_data.iterrows():
        eq_name, eq_code, point_name, point_code, channel_id, sensor_type = df_row.iloc[:6]

        if eq_name is None or eq_name == "" or pd.isna(eq_name) or pd.isnull(eq_name):
            continue

        channel_type = resolve_channel_type(sensor_type, point_code)
        enabled_feature_types = output_template(df_row, bearing_data)
        if df_row[3][-2:-1] in ['X', 'Y']:
            print(df_row[3], ":", enabled_feature_types)

        # 关键：先按通道类型过滤，再按特征类型过滤
        selected_template_rows = select_template_rows(
            template_dataframe=template_dataframe,
            channel_type=channel_type,
            enabled_feature_types=enabled_feature_types
        )

        for _, meta_row in selected_template_rows.iterrows():
            tmp_data.append([
                eq_name,
                eq_code,
                point_name,
                point_code,
                sensor_type,
                meta_row['数据项（特征）名称'],
                format_feature_code(point_code, meta_row['数据项代号']),
                meta_row['数据项（特征）类型'],
                meta_row['数据类型'],
                meta_row['单位'],
                channel_id
            ])

    output_data = DataFrame(tmp_data, columns=columns_name)

    if not need_channel_id:
        output_data = output_data.drop(columns=['通道编码'])

    # 自适应列宽
    def excel_widths(excel_dataframe):
        if excel_dataframe.empty:
            return np.array([10] * len(excel_dataframe.columns))

        column_widths = (excel_dataframe.columns.to_series()
                         .apply(lambda x: len(str(x).encode('utf-8'))).values
                         ) * 0.8

        max_widths = (excel_dataframe.astype(str)
                      .applymap(lambda x: len(x.encode('utf-8')))
                      .agg(max).values
                      ) * 0.8

        return np.max([column_widths, max_widths], axis=0)

    with ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book

        # 设备档案
        input_device_profile.to_excel(writer, sheet_name='设备档案', index=False)
        worksheet_profile = writer.sheets['设备档案']

        border_format = workbook.add_format({'border': 1, 'border_color': 'black'})
        worksheet_profile.conditional_format(
            0, 0, input_device_profile.shape[0], input_device_profile.shape[1] - 1,
            {'type': 'no_errors', 'format': border_format}
        )

        title_format = workbook.add_format(
            {'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#BFBFBF'}
        )
        for col_num, value in enumerate(input_device_profile.columns.values):
            worksheet_profile.write(0, col_num, value, title_format)

        wrap_format = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
        for i in range(input_device_profile.shape[1]):
            worksheet_profile.set_column(i, i, 10, wrap_format)

        # 输出模板
        output_data.to_excel(writer, sheet_name="输出模板", index=False)
        worksheet_data = writer.sheets["输出模板"]

        worksheet_data.conditional_format(
            0, 0, output_data.shape[0], output_data.shape[1] - 1,
            {'type': 'no_errors', 'format': border_format}
        )

        title_format = workbook.add_format(
            {'bold': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#BFBFBF'}
        )
        for col_num, value in enumerate(output_data.columns.values):
            worksheet_data.write(0, col_num, value, title_format)

        for i, width in enumerate(excel_widths(output_data)):
            worksheet_data.set_column(i, i, width)

        if output_data.shape[1] > 4:
            worksheet_data.set_column(4, 4, 20)
        if need_channel_id and output_data.shape[1] > 10:
            worksheet_data.set_column(10, 10, 20)

    return output_file_True


if __name__ == "__main__":
    output_template_all(
        r"后台文件/data_all(模板).xlsx",
        r"后台文件/my_def_对应注释.xlsx",
        r"后台文件/平台导入表.xlsx",
        False
    )
