# Excel 数据比对工具使用说明

## 简介

这个 Python 脚本 `find_common_data.py` 可以帮助你比较 Excel 文件中两个工作表的数据，找出指定列中相同的内容，并将结果保存到一个新的 Excel 文件中。

## 使用前准备

1. 确保你的电脑上安装了 Python（建议使用 Python 3.6 或更高版本）。
2. 安装必要的 Python 库。打开命令提示符（Windows）或终端（Mac/Linux），运行以下命令：

   ```bash
   pip install pandas openpyxl
   ```

3. 将 `find_common_data.py` 文件下载到你的电脑上。

## 使用方法

1. 打开命令提示符（Windows）或终端（Mac/Linux）。

2. 切换到 `find_common_data.py` 文件所在的目录。例如：

   ```bash
   cd C:\Users\YourName\Documents\excel_tools
   ```

3. 运行脚本，使用以下格式：

   ```bash
   python find_common_data.py <Excel文件路径> <工作表1名称> <工作表2名称> <列名>
   ```

   注意：<Excel文件路径> 可以是任何位置的文件，不必与脚本在同一目录。

   例如，要比对 "D:\code\01.xlsx" 文件中的两个工作表，比较 "3、学号" 列的相同数据：

   ```bash
   python find_common_data.py D:\code\01.xlsx Sheet1 Sheet2 3、学号
   ```

4. 脚本会在当前目录（即 `find_common_data.py` 所在的目录）下的 `data` 文件夹中生成一个新的 Excel 文件，文件名格式为 `common_data_年月日_时分秒.xlsx`。如果 `data` 文件夹不存在，脚本会自动创建它。

## 参数说明

- `<Excel文件路径>`: 你要比对的 Excel 文件的完整路径。
- `<工作表1名称>`: 第一个要比对的工作表的名称。
- `<工作表2名称>`: 第二个要比对的工作表的名称。
- `<列名>`: 你想要比对的列的名称。这个列名必须在两个工作表中都存在。

## 注意事项

1. 确保你有权限读取指定的 Excel 文件和创建新文件。
2. 指定的列名必须在两个工作表中都存在。
3. 如果遇到错误，脚本会显示相应的错误信息，帮助你解决问题。
4. 新生成的文件会包含两个工作表中共同的数据，所以可能后续还需要人工操作一下

## 输出结果

脚本会创建一个新的 Excel 文件，包含以下内容：

1. 一个名为 "Common Data" 的工作表，其中包含两个工作表中指定列相同的所有数据。
2. 保留了原始 Excel 文件中的字体样式。

脚本运行完成后，会显示找到的相同数据数量和新文件的保存位置。

如果你在使用过程中遇到任何问题，请检查错误信息并确保你正确地输入了所有参数。