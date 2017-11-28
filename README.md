# EISC_KPIHelper
- 用于EISC线下绩效考核

## 环境
- python3.6 + openpyxl

## 使用

- 根据需要填写config文件，填写方法见config文件
- 分类放入所有资料，包括：
    - data/(主席团名) 所有主席的3个表
    - data/(部门名) 出勤表
    - data/(部门名)/leader/ltol 部长对其他部长评价表
    - data/(部门名)/leader/ltom 部长对干事评价表
    - data/(部门名)/member 干事自评表
    - model 3份反馈表模板
 - 用python3执行main.py，（可添加check参数先对资料进行检验）

## 说明

- 该程序只对线下的绩效表格进行整理计算然后写入到反馈表，不能上传到数据库或任何线上行为
- 该程序设计的所有表格事先准备好（data和model文件中有范例），若需要更改表格则需要更改相应代码
- config文件和表格中涉及的人名的顺序不影响，但总数要和config文件保持一致
- 每年干事数量不同不影响，主席数量不影响，部门数量有改变则可能有影响反馈表的美观
- 根据每年干事数量的不同需要修改对应反馈表的行数，不用修改代码
- borden.py是对反馈表单元格边框做调整，可以根据需要更改
- 涉及计算相关的改变需要直接更改代码 

## 历史

 - 2017.11.15前：
   + 完成基本功能
   + 完成check模块
   + 完成进度条模块
   
 - 2017.11.28：
   + 追加计算部门加减分部分
