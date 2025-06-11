# WordImageInserter
---------------------------------------------------------------------------
本程序用于将文件夹中的图片按照分层目录结构排序后，以表格方式批量插入 Word 报告模板中。
---------------------------------------------------------------------------
# 所需文件结构 
主目录（任意命名）
├── WordImageInserter.exe
├── data
├── 文件名_级别.xlsx
├── 图片命名映射表.xlsx
├── template.docx

---------------------------------------------------------------------------
# 所需文件说明 （这些文件必须处于同一目录）
1. 图片命名映射表.xlsx
- 第 1 列：原始图片文件名（不含文件后缀）
- 第 2 列：映射名称（将显示于 Word 中的图片标题）
- 必须从第二行开始写，第一行是表头不会被读取

2. Word 模板 Template
- 可预设页眉页脚、字体等格式要求，主体部分空白即可

3. data文件夹
- 将图片文件夹放入该文件夹

4. 文件名_级别.xlsx
- 用于存储各级文件夹中每一个文件的读取优先级

---------------------------------------------------------------------------
# 功能说明 
- 自动读取文件夹层数
- 按照模板每页插入 6 张图片（3行2列表格）
- 自动分页、生成页码与总页数
- 图片下方自动生成名称
- 图片命名可映射，如：5.2 → file2 （来自图片命名映射表.xlsx）简化图片命名

---------------------------------------------------------------------------
# 使用说明 
第一步：准备图像目录
- 图片必须为 .png, .jpg, .jpeg, .bmp 中一种
- 所有目录名必须与 Excel 表格中一致（区分大小写，不能多空格）
- 将该文件夹放入data文件夹中

第二步：填写优先级和命名映射 Excel
- 每两列一组（如：第一级文件夹名 + 优先级）

第三步：运行程序
-  点击  WordImageInserter.exe 打开程序
-  程序开启后会自动读取一遍 文件名_级别.xlsx 表格，读取到的文件优先级信息会显示在程序上半部分
-  如果在程序运行中修改了 文件名_级别.xlsx 表格内容，需要点击”重新读取优先级“按钮重新读取
- 点击“写入文档”
- 等待几秒后程序同目录下会有一个”output.docx“，这个就是写入后包含照片的文档

---------------------------------------------------------------------------
# 常见问题
图片顺序不对？→ 检查 Excel 是否为所有文件夹定义了优先级，文件命名是否区分大小写、是否有拼写错误。
图片标题错误？→ 图片名是否正确填写到“命名映射表”中（不含文件后缀）
图片未插入？ → 确保图片格式正确（必须是.png, .jpg, .jpeg, .bmp）、Excel 正确保存

对于没有在‘文件名_级别.xlsx’中声明的所有文件（包括有拼写错误的文件），程序会默认将该文件优先级赋值为999，同时再点击写入文档按钮之后程序底部的运行日志框中会出现如下内容：
*******************************************
（注意！）C:\Users\79192\OneDrive\桌面\WordImageInserter\data\folder\1.1\2.2\3.1\4.1\5.1PNG.PNG -> 排序 key: (1, 2, 1, 1, 999)
该路径中可能存在文件名拼写错误或未在‘文件名_级别.xlsx’中声明该文件，请检查。
*******************************************
可以看见这里的图片文件夹有五级，它们各自的优先级为(1, 2, 1, 1, 999)，最后一级文件也就是”1024QAMlightPNG.PNG“优先度被赋值999，代表该文件没有被声明或存在拼写或大小写错误。
对于这种情况，请更改文件名或者在”文件名_级别.xlsx“中添加这个文件的声明。

---------------------------------------------------------------------------
# 注意事项
- 所有文件夹名、图片名区分大小写
- 如果图片较多建议分批处理

-----------------------------------------------------------------------------

# WordImageInserter
---------------------------------------------------------------------------
This program is used to batch insert images from nested folders into a Word report template, sorted according to a multi-level directory structure.
---------------------------------------------------------------------------
# Required File Structure 
Main Directory (any name)
├── WordImageInserter.exe
├── data
├── filename_priority.xlsx
├── image_name_mapping.xlsx
├── template.docx

---------------------------------------------------------------------------
# Required File Descriptions (all files must be in the same directory)
1. image_name_mapping.xlsx
- Column 1: Original image filename (without extension)
- Column 2: Mapped name (to be shown as the image title in Word)
- Start from row 2; row 1 is the header and will not be read

2. Word Template (template.docx)
- Can predefine headers, footers, fonts, etc.; the body can remain empty

3. data folder
- Place your image folders inside this folder

4. filename_priority.xlsx
- Stores the reading priority of each folder/file at every level

---------------------------------------------------------------------------
# Features
- Automatically detects folder levels
- Inserts 6 images per page (3 rows, 2 columns per table)
- Automatic pagination and total page numbering
- Automatically adds image captions
- Supports name mapping, e.g., 5.2 → file2 (from image_name_mapping.xlsx)

---------------------------------------------------------------------------
# Instructions
Step 1: Prepare image directories
- Images must be one of: .png, .jpg, .jpeg, .bmp
- Folder names must match Excel entries exactly (case-sensitive, no extra spaces)
- Place the folder into the "data" directory

Step 2: Fill in priority and name mapping Excel files
- Every two columns define a level (e.g., level 1 folder + priority)

Step 3: Run the program
- Launch WordImageInserter.exe
- Upon launch, the program will read filename_priority.xlsx and display the folder priorities
- If you modify filename_priority.xlsx during runtime, click "Reload Priority" to refresh
- Click “Write to Document”
- After a few seconds, a file named “output.docx” will be generated in the same directory

---------------------------------------------------------------------------
# Common Issues
Wrong image order? → Check if all folders are defined in Excel, and ensure correct case and spelling.
Wrong image title? → Ensure the image name is correctly listed in "image_name_mapping.xlsx" (without extension)
Image not inserted? → Make sure the image format is correct and Excel is properly saved

If a file is not declared in ‘filename_priority.xlsx’ (e.g. due to a typo), the program will assign it priority 999 and display the following message in the log box after clicking “Write to Document”:
*******************************************
(Warning!) C:\Users\Example\WordImageInserter\data\folder\1.1\2.2\3.1\4.1\5.1PNG.PNG.PNG -> sort key: (1, 2, 1, 1, 999)
This may indicate a filename typo or missing declaration in 'filename_priority.xlsx'. Please check.
*******************************************

---------------------------------------------------------------------------
# Notes
- Folder and file names are case-sensitive
- For large numbers of images, consider processing in smaller batches
