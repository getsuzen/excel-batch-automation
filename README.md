# Python Excel 自动合并 Demo

这是一个适合初学者阅读和演示的 Excel 自动化项目。脚本会读取 `input` 目录中的全部 `.xlsx` 文件，合并第一张工作表的数据，去除整行重复记录，并把结果保存到 `output\summary.xlsx`。

## 功能

- 按文件名顺序读取 `input` 目录中的全部 `.xlsx` 文件
- 每个文件使用第一张工作表，并只保留一次表头
- 按整行内容去重，保留第一次出现的记录
- 忽略全空行和 Excel 临时文件（文件名以 `~$` 开头）
- 在终端显示有效文件数、原始行数、去重后行数和删除重复数
- 自动创建 `output` 目录并生成带基础格式的 `summary.xlsx`
- 遇到空工作簿、损坏文件或表头不一致时给出提示，并继续处理其他文件

## 目录结构

```text
demo_excel_tool/
├─ input/                  # 放入待合并的 Excel 文件
│  ├─ sales_01.xlsx
│  ├─ sales_02.xlsx
│  └─ sales_03.xlsx
├─ output/
│  └─ summary.xlsx         # 运行后生成
├─ create_sample_files.py  # 重新生成演示数据
├─ merge_excel.py          # 合并与去重脚本
├─ requirements.txt
└─ README.md
```

## 安装

在 Windows PowerShell 中进入项目目录：

```powershell
cd D:\Freelance\demo_excel_tool
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

如果 `.venv` 已经存在，只需激活环境并安装依赖。

## 运行

直接合并 `input` 中现有的 Excel 文件：

```powershell
python merge_excel.py
```

如需重新生成演示文件，再运行合并：

```powershell
python create_sample_files.py
python merge_excel.py
```

`create_sample_files.py` 会覆盖同名的 3 个演示文件。演示数据共有 7 行，其中 1 行完全重复，所以正常结果应为：

```text
原始行数: 7
去重后行数: 6
删除重复数: 1
完成！输出文件: output\summary.xlsx
```

## 输入文件要求

- 文件扩展名为 `.xlsx`
- 要合并的数据位于第一张工作表
- 第一行是表头，各文件表头应完全一致
- 从第二行开始是数据

损坏或空的 Excel 文件会被跳过。表头不一致的文件也会被跳过，以免不同列的数据被错误合并。如果没有可用文件，脚本不会生成新的结果文件。

每次正常运行都会覆盖已有的 `output\summary.xlsx`，不会追加旧结果。
