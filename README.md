# Express7 CytoSummary

**Express7 数据汇总 · Flow Cytometry Data Summary**

当前版本：**V3.4.4**

作者：Guibin Huang

基于 Python、SQLite 和 Tkinter/CustomTkinter 的桌面工具，用于流式细胞术数据汇总、验证统计和 Excel 一致性核查。

## 主要功能

- 按项目和 Panel 管理数据，维护受试者信息、样本结果及操作记录；工作台和受试者库支持搜索、分页。
- 导出 Summary、计算明细、变更记录和复管 CV，支持配置小数位或有效数字。
- 手动设置验证批次及 Baseline，计算批内 CV 和与 Baseline 均值的 CV。
- 比较两个 Excel 的指定工作表和列，生成逐样本核查、差异明细等报告，保留原文件。

## 安装与启动

在 Windows 上安装带有 Tkinter 的 Python。将本仓库下载并解压，在项目文件夹中打开终端，运行：

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe express7_v2.py
```

`express7_v2.py` 与 `consistency_check.py` 必须放在同一目录。主程序文件名沿用旧版命名，程序版本为 V3.4.4。

启动后按提示新建或选择已有数据库。已有用户可选择原数据库，无需重新导入已有数据。不要同时用两个版本修改同一个数据库。

## V3.4.4 要点

- 工作台和受试者库支持搜索及每页 100/200/500 条查看；工作台可按 Run 筛选，列排序作用于全部筛选结果。
- 点击“数据异常”数字，只查看资料冲突或 Beads 待确认的记录；点击“显示全部”恢复。F 标记本身不计入该异常数字。
- “导入异常”单独显示问题文件；表格搜索、异常筛选和分页不改变 Summary 的导出范围。
- 结果预览、复管 CV 及验证计算/导出在后台执行；预览支持分页。批量导入显示当前文件和文件进度。
- Summary 状态同时检查数据、配置和“排除 F”条件，并显示上次导出的时间、条件及记录数。
- 移除独立历史页面；后台的版本链、来源、操作记录及导出变更记录保留。

## V3.4.3 要点

- Summary 的 Average、CV% 统计行遵循“rpt复管CV”配置。
- Summary 时间点采用 Cycle、Visit、采血阶段、Pre / Post、时长五列。
- 验证统计按手动设置的批次与 Baseline 计算，增加“验证Summary”工作表。
- 一致性核查增加“逐样本核查”第一页，按比较精度标红差异。

详细计算口径、缺失值处理、操作步骤和原版检查记录见 [使用说明](使用说明.txt)。

## 文件说明

| 文件 | 用途 |
| --- | --- |
| `express7_v2.py` | 主程序、数据库管理、汇总及验证导出 |
| `consistency_check.py` | Excel 一致性核查模块 |
| `requirements.txt` | 固定版本的第三方依赖 |
| `使用说明.txt` | 原版使用说明和版本记录 |

## 数据与版本说明

本仓库以 `Express7_V3_4_3_PairedCheck_no_bat.zip` 为初始版本，V3.4.4 改善查看和后台处理体验。仓库不附带样本 Excel、业务数据库或数据库备份；这些文件及本地环境文件已加入 `.gitignore`。

本次更新使用模拟数据检查分页、异常筛选、导出条件状态、后台预览及完整范围导出，未重新验证全部业务计算。Excel 中的原始指标通常保存为数值，ABS 和部分 CV 保留公式，需要 Excel 计算并保存后才能供依赖公式缓存的核查工具读取。
