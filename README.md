# 烟订单打印
## 这是做什么的？
新商盟订烟单excel处理文件。删掉一些不必要的行和列，并插入日期到最后一行，方便打印出来核对。

# 环境搭建
方法二：使用pip
```shell
pip install openpyxl,pandas,

```
方法二：使用 Anaconda 
```shell
conda create --name <env> --file requirements.txt
```

# 用法：
把resources下的
source.xls或者source.csv替换成你要修改的excel文件，运行main.py或者maincsv.py即可生成一个处理后的表格

# 快捷方式
写一个bat文件，把要打印的excel拖拽到该bat文件上面，就会在当前目录下生成一个printer的文件夹，并将目标文件存放到里面
```bat
G:\software\anaconda\envs\smoke_excel\python.exe "H:\OneDrive/Project/pyCharmProject/excel_utils/main.py" %1
pause
```
# 参数说明
- `G:\software\anaconda\envs\smoke_excel\python.exe`: python可执行文件位置
- `"H:\OneDrive/Project/pyCharmProject/excel_utils/main.py"`: 此项目main.py文件位置
- `%1`: 在bat脚本中, %0表示bat本身, %1被视为传入第一个参数，如果拖拽文件到bat，则%1就是拖拽文件的绝对路径