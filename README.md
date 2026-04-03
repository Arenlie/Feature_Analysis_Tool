# Feature_Analysis_Tool（特征解析工具）

> 一个基于 PyQt6 的桌面工具，用于把 `data_all` 数据模板转换为多类下游交付物：平台导入表、DW 导入表、JSON 配置、deviceInfo、tupuSetting（V2/V3）等。

---

## 1. 项目整体目标

本项目面向设备特征解析与配置落地场景，核心输入是 `data_all.xlsx`（至少包含“输入参数”“设备档案”两个工作表），核心输出是：

1. 平台导入表（带“设备档案 + 输出模板”）；
2. DW 导入表（按主机 MAC 分 sheet，补齐板卡/通道）；
3. JSON 三件套（`ChannelSettings.json` / `FeatureCalc.json` / `Features.json`）；
4. `deviceInfo.xlsx`；
5. `tupuSetting`（V2/V3 版本）。

整个系统由 **GUI 编排层 + 数据转换层 + 特征规则层 + Excel 美化层 + 静态资源层** 组成。

---

## 2. 一图看懂总体流程

```text
用户在 GUI 上传 data_all.xlsx
        |
        v
 main.py（按钮触发 Worker 线程）
        |
        +--> PlatformTable.output_template_all
        |         |
        |         +--> 平台导入表.xlsx
        |         +-->（可选）设备缺失项.txt
        |
        +--> dataToDWTable.dataToDWTable
        |         |
        |         +--> DW-导入表.xlsx
        |
        +--> fea_json.feature_json_all
        |         |
        |         +--> 每个 MAC 文件夹
        |                 |- ChannelSettings.json
        |                 |- FeatureCalc.json（有高速通道时）
        |                 |- Features.json（有特征时）
        |
        +--> deviceInfo_And_tupusetting.device_info
        |         |
        |         +--> deviceInfo.xlsx
        |
        +--> deviceInfo_And_tupusetting.tupuSetting_V2 / V3
                  |
                  +--> tupuSettingV2.xlsx / tupuSettingV3.xlsx
```

---

## 3. 运行方式

### 3.1 环境依赖

安装依赖：

```bash
pip install -r requirements.txt
```

### 3.2 启动 GUI

```bash
python main.py
```

启动后可在界面中：
- 上传 `data_all`；
- 下载模板；
- 选择不同按钮输出目标文件。

---

## 4. 关键业务流程（用户操作）

### 4.1 上传 data_all
- 入口：`main.py -> MyMainWindow.load_file`
- 动作：用户选择 `.xlsx`，路径保存到 `self.data_all_edit`。

### 4.2 输出“平台导入表”
- 入口：`main.py -> Worker1`
- 主逻辑：`PlatformTable.output_template_all(...)`
- 结果：生成带“设备档案”“输出模板”的 Excel；若输入参数与设备档案设备名不匹配，则额外生成 `设备缺失项.txt`。

### 4.3 输出“DW-导入表”
- 入口：`main.py -> Worker5`
- 主逻辑：`dataToDWTable.dataToDWTable(...)`
- 结果：按主机 MAC 分 sheet 输出，自动补齐板卡与通道、设置启用状态与键相类型，并进行合并单元格美化。

### 4.4 输出 JSON
- 入口：`main.py -> Worker4`
- 主逻辑：先 `dataToDWTable(...)` 生成中转 DW 表，再 `fea_json.feature_json_all(...)`
- 结果：按每个 MAC 单独创建文件夹并输出 JSON；若校验失败，生成 `error_list.json`。

### 4.5 输出 deviceInfo
- 入口：`main.py -> Worker2`
- 主逻辑：先生成中转“平台导入表”，再 `deviceInfo_And_tupusetting.device_info(...)`
- 结果：输出 `deviceInfo.xlsx`。

### 4.6 输出 tupuSetting（V2 / V3）
- 入口：`main.py -> Worker3 / Worker6`
- 主逻辑：先生成中转“平台导入表”，再调用 `tupuSetting_V2(...)` 或 `tupuSetting_V3(...)`
- 结果：输出对应版本 tupuSetting 文件。

---

## 5. 每个文件的流程说明

> 下面按“入口文件 / 核心逻辑 / 输入输出 / 在全链路中的位置”说明。

### 5.1 根目录 Python 文件

#### 1) `main.py`
- **角色**：GUI 主入口 + 多任务线程调度。
- **流程**：
  1. 初始化 PyQt 主窗口、样式、按钮事件；
  2. 用户上传 `data_all`；
  3. 按按钮启动对应 `Worker`（QThread）避免阻塞 UI；
  4. Worker 内调用业务模块并回传结果到 UI 标签。
- **输出职责**：不直接做复杂业务计算，主要负责流程编排与状态反馈。

#### 2) `PlatformTable.py`
- **角色**：生成“平台导入表”的核心引擎。
- **流程**：
  1. 读取并标准化 `my_def_对应注释.xlsx` 模板，校验必需列与唯一码唯一性；
  2. 读取 `data_all` 的“输入参数/设备档案”；
  3. 对每个测点根据传感器类型与参数，计算启用特征集合（`output_template`）；
  4. 先按通道类型，再按特征类型筛模板行；
  5. 生成输出模板行并拼接“数据项（特征）编码”；
  6. 写入 Excel（设备档案 + 输出模板）并格式化。
- **附加能力**：设备名差异检测，输出 `设备缺失项.txt`。

#### 3) `dataToDWTable.py`
- **角色**：把 `data_all` 转为 DW 设备接入导入表。
- **流程**：
  1. 从“输入参数”读取原始行；
  2. 识别主机 MAC、板卡编号、通道编号、传感器类型；
  3. 记录已存在板卡/通道结构；
  4. 根据网关型号（DW2700 / DW2300）补齐缺失板卡与通道；
  5. 输出按 MAC 分 sheet 的 Excel；
  6. 合并重复单元格并居中对齐。
- **约束**：目前仅支持特定传感器类型（如加速度、温度、转速等）。

#### 4) `fea_json.py`
- **角色**：从 DW 导入表生成 JSON 配置。
- **流程**：
  1. 读取 DW 导入表全部 sheet；
  2. 每个 sheet（每个 MAC）单独生成输出目录；
  3. 首轮校验字段完整性与类型合法性，异常写 `error_list.json`；
  4. 通过模板 JSON 与常量特征 ID 组装：
     - `ChannelSettings.json`
     - `FeatureCalc.json`
     - `Features.json`
  5. 保存到对应 MAC 目录。
- **依赖**：`feature_values.py` 常量、`后台文件/tmp_*.json` 模板、轴承库 CSV。

#### 5) `deviceInfo_And_tupusetting.py`
- **角色**：派生生成 `deviceInfo` 与 `tupuSetting`。
- **流程（device_info）**：
  1. 读取平台导入表和注释映射；
  2. 构建设备/测点/数据项字段；
  3. 做通道值映射、区域补齐、无线点位特殊处理；
  4. 输出格式化 Excel。
- **流程（tupuSetting_V2 / V3）**：
  1. 以“输出模板”为输入去重测点；
  2. 根据通道类型从 `settings_V2/V3` 拉取波形参数模板；
  3. 组装成目标结构并导出。

#### 6) `feature_values.py`
- **角色**：集中维护特征 ID 常量与 tupuSetting 模板参数。
- **内容**：
  - 大量 `kXXXX` 特征编号常量；
  - `settings_V2` 与 `settings_V3` 字典（波形项、编码、采样率、时长等）。
- **位置**：属于规则/配置层，被 `fea_json.py` 和 `deviceInfo_And_tupusetting.py` 复用。

#### 7) `excel_Optimization.py`
- **角色**：Excel 导出样式与自适应列宽工具。
- **流程**：
  1. 根据中英文字符宽度估算列宽；
  2. 应用标题样式（字体、背景色、边框、居中）；
  3. 输出统一风格 Excel。
- **调用方**：`deviceInfo_And_tupusetting.py`。

#### 8) `dataTo2700table.py`
- **角色**：历史版本转换脚本（早期/兼容用途）。
- **说明**：逻辑与 `dataToDWTable.py` 类似，但规则较旧（如“动态电压”等）；主流程已由 `dataToDWTable.py` 承担。

#### 9) `resr.py`
- **角色**：独立绘图实验脚本。
- **流程**：生成模拟冲击显度热力图并保存 `output.png`。
- **说明**：与主业务链路无直接耦合。

#### 10) `app.spec`
- **角色**：PyInstaller 打包配置。
- **用途**：定义主脚本、图标、输出名称等，用于生成桌面可执行程序。

---

### 5.2 资源与数据文件目录

#### 11) `images/`
- **角色**：GUI 资源与界面文件。
- **关键文件流程**：
  - `UImain.ui`：Qt Designer 源文件；
  - `UImain.py`：由 `.ui` 转换得到的 Python UI 类，`main.py` 直接加载；
  - `custom.css`：`qt_material` 叠加样式；
  - `*.png/*.ico`：按钮、背景、Logo 等视觉资源。

#### 12) `后台文件/`
- **角色**：后台模板与静态数据仓。
- **关键文件流程**：
  - `data_all(模板).xlsx`：用户下载的输入模板；
  - `my_def_对应注释.xlsx`：特征元数据模板（平台导入表映射核心）；
  - `Bearing.xlsx` / `t_bearing_head.csv`：轴承库，用于轴承相关特征判断；
  - `tmp_ChannelSettings.json` / `tmp_features.json` / `tmp_ChannelFeatureCalc.json`：JSON 生成模板。

#### 13) `中转文件/`
- **角色**：中间产物缓存目录。
- **文件流程**：
  - `平台导入表.xlsx`：为 deviceInfo / tupuSetting 生成准备的中转文件；
  - `DW-导入表.xlsx`：JSON 生成前的中转 DW 表；
  - `设备缺失项.txt`：设备档案与输入参数差异提示。

#### 14) `requirements.txt`
- **角色**：Python 依赖清单。
- **用途**：环境安装入口。

---

## 6. 数据与文件关系（输入 / 中间 / 输出）

### 6.1 输入
- 用户输入：`data_all.xlsx`
- 后台静态模板：`后台文件/*`

### 6.2 中间文件
- `中转文件/平台导入表.xlsx`
- `中转文件/DW-导入表.xlsx`

### 6.3 最终输出
- 平台导入表（用户指定路径）
- DW 导入表（用户指定路径）
- JSON 文件夹（用户指定目录下按 MAC 分目录）
- deviceInfo（用户指定路径）
- tupuSetting V2/V3（用户指定路径）

---

## 7. 常见异常与排查建议

1. **提示“访问权限限制，请关闭相关文件”**
   - 原因：目标 Excel 正在被占用；
   - 处理：关闭已打开的同名文件后重试。

2. **提示“缺少列 / 文件格式错误”**
   - 原因：输入表头不符合约定；
   - 处理：优先使用“下载模板”按钮拿到标准 `data_all(模板).xlsx`。

3. **JSON 生成失败并出现 error_list.json**
   - 原因：DW 表字段校验失败（板卡类型、通道类型、键相参数等不一致）；
   - 处理：先检查 DW 导入表对应行，再回溯 data_all 源数据。

4. **出现“设备缺失项.txt”**
   - 原因：“输入参数”和“设备档案”中的设备名集合不一致；
   - 处理：补齐两张表中的设备名称。

---

## 8. 二次开发建议

- 新增传感器/特征：优先修改 `PlatformTable.py` 与 `feature_values.py`；
- 新增输出文件：在 `main.py` 新增 Worker + 按钮绑定，业务逻辑放独立模块；
- 统一导出风格：复用 `excel_Optimization.export_excel`；
- 调整打包：修改 `app.spec` 并重新执行 PyInstaller。

---

## 9. 项目维护速查

1. 先跑通 `python main.py`；
2. 使用模板生成一次全量输出；
3. 按输出类型逆向看代码：
   - 平台导入表 -> `PlatformTable.py`
   - DW 导入表 -> `dataToDWTable.py`
   - JSON -> `fea_json.py`
   - deviceInfo / tupuSetting -> `deviceInfo_And_tupusetting.py`
4. 若改规则，最后检查 `feature_values.py` 是否同步。

