# 对象



### 创建对象

```python
wb = Workbook()
```

### 活动的工作表

```python
ws = wb.active
```

### 

# 文件



### 保存

```python
wb.save("text.xlsx")
```

### 读取

```python
llshw = load_workbook("llshw.xlsx")
```



# 工作表



### 打印所有工作表

```python
print(llshw.sheetnames)
```

### 创建工作表

```python
answer.create_sheet("Sheet2",1)
```

### 取工作表

```python
mysheet = answer["Sheet2"]
```

### 移动工作表

```python
answer.move_sheet(mysheet,-1)
```

### 删除工作表

```python
del answer["Sheet2"]
```

### 

# 单元格



### 单元格赋值

```python
#单元格赋值1
acsheet["L25"] = "2316"
#单元格赋值2
cell2 = acsheet.cell(13,1,"LL")
#单元格赋值3
cell2.value = "ILLL" 
```

### 打印单元格坐标

```python
print(cell2.coordinate)
print(cell2.row)
print(cell2.column)
print(cell2.column_letter)
```

### 遍历单元格

```python
#按列遍历从A列到C列
acsheet["A:C"]
#按行遍历从1行到5行
acsheet[1:5]
#行优先遍历指定区域A1到C4
acsheet["A1:C4"]
```

### 合并单元格

```python
acsheet.merge_cells("A5:B6")
```

### 插入单元格

```python
#在第1行位置插入2行
acsheet.insert_rows(1,2)
#在第2行位置插入3列
acsheet.insert_cols(2,3)
```

### 删除单元格

```python
#在第1行位置删除2行
acsheet.delete_rows(1,2)
```

### 移动单元格

```python
#将C5:D6移动1行,-1列
acsheet.move_range("C5:D6",1,-1)
```

### 最下方插入一行数据

```python
acsheet.append([2,3,1,6])
```



# 公式



### 使用公式

```python
from openpyxl.utils import FORMULAE
acsheet["F1"] = "=SUM(A1:E1)"
```

### 翻译公式

```python
from openpyxl.formula.translate import Translator
acsheet["F2"] = Translator(formula="=SUM(A1:E1)",origin="F1").translate_formula("F2")
for cell in acsheet["F3:F10"]:
	cell[0].value = Translator(formula="=SUM(A1:E1)",origin="F1").translate_formula(cell[0].coordinate)
```


