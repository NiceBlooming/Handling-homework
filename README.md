# Handling-homework
该文件用来处理每次上交的作业。

## 代码使用场景
1. 每次的作业都会保存在某一个文件夹【第i章作业】下面；
2. 存在花名册【花名册.xls】；
3. 只需要判断是否提交，无关对错。

## 代码用法
`python homework_process.py 作业文件夹 i(第i次作业)`  
例如：  
`python homework_process.py 第三章作业 3`

## 修改部分
* execl文件的位置  
`15. execl_file = xlrd.open_workbook(r'作业提交情况点名册.xls')`  
* 学生名字的位置  
`30. name_list.append(sheet.cell(i, 2).value)`
`31. name_str += sheet.cell(i, 2).value`  
* 需要添加的位置  
`57. table.write(loc + 8, homework_number + 5, 10)`

## 注意
目前只之前中文名称，即学生的名字为中文  
使用截图：  
![Image text](https://github.com/NiceBlooming/Handling-homework/blob/df0166a9a5476d743bf73aeac15a96e7b84f9d0c/%E8%BF%90%E8%A1%8C%E6%88%AA%E5%9B%BE.jpg)
