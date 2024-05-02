from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border,Side

#llshw = load_workbook("llshw.xlsx",data_only=True)

#zhsheet = myhw["sheet1"]

answer = Workbook()
sheet = answer["Sheet"]
sheet["A1"].value = "所属二级所"
sheet["B1"].value = "所属实验室"
sheet["C1"].value = "团队名称"
sheet["D1"].value = "经费类别"
sheet["E1"].value = "国家基金"
sheet["F1"].value = "科技部"
sheet["G1"].value = "国家其他"
sheet["H1"].value = "中科院"
sheet["I1"].value = "浙江省"
sheet["J1"].value = "宁波市"
sheet["K1"].value = "国际合作"
sheet["L1"].value = "高技术"
sheet["M1"].value = "其他"
sheet["N1"].value = "总计"
sheet["O1"].value = "转出经费"
sheet["P1"].value = "净收入"

sheet["A5"].value = "材料所"
sheet["A143"].value = "先进制造所"
sheet["A173"].value = "新能源所"
sheet["A224"].value = "医工所总计"
sheet["A257"].value = "其他"
sheet["A272"].value = "工研院总计"

sheet["B5"].value = "纳米"
sheet["B26"].value = "海洋"
sheet["B68"].value = "高分子"
sheet["B107"].value = "先进能源材料工程实验室"
sheet["B119"].value = "磁材实验室"
sheet["B134"].value = "磁性材料应用技术研究中心"
sheet["B137"].value = "稀土永磁材料联合创新中心"
sheet["B143"].value = "材料所总计"
sheet["B146"].value = "机器人与智能制造装备技术实验室"
sheet["B161"].value = "激光极端制造研究中心"
sheet["B167"].value = "特种飞行器系统工程研究中心"
sheet["B173"].value = "先进制造所总计"
sheet["B173"].value = "新能源"
sheet["B215"].value = "动力锂电"
sheet["B221"].value = "新能源所总计"
sheet["B224"].value = "医工所"
sheet["B254"].value = "医工所总计"
sheet["B257"].value = "平台"
sheet["B260"].value = "科技发展部"
sheet["B263"].value = "转移办"
sheet["B266"].value = "人力"
sheet["B269"].value = "党政办"

for i in range(5,275):
	if(i%3 == 2):
		sheet.cell(i,4,"2024年合同")
	elif(i%3 == 0):
		sheet.cell(i,4,"2024年到位")
	else:
		sheet.cell(i,4,"2023年到位")

Ctext = ["许高杰","曹彦伟","曹鸿涛","刘兆平","诸葛飞","先进纳米实验室","纳米小计",\
"江南","王立平","汪爱英","宋振纶","曾志翔","张涛","茅东升","蒲吉斌","黄政仁","陈涛","刘小青","常可可","海洋关键材料重点实验室","海洋小计",\
"朱锦","郑文革","姚强","方省众","陈鹏","刘富","张永刚","王震","张浩","张健","余海斌","高分子与复合材料实验室","高分子小计",\
"黄庆","张蕾","先进能源材料工程实验室","先进能源小计",\
"王军强","李国伟","李润伟","中科院磁性材料与器件重点实验室","磁材小计",\
"满其奎","闫阿儒","二级所","","张驰","祝颖丹","蒋俊","肖江剑","赵夙","张文武","韦超","陈新民","二级所","",\
"宋伟杰","叶继春","向超宇","葛子义","陈亮","况永波","尹宏峰","张亚杰","姚霞银","官万兵","杨明辉","夏永高","陆之毅","谢银君","二级所","",\
"赵一天","张建涛","吴爱国","李华","左国坤","王荣","郑建萍","郭明全","李辉","二级所"]

#B列合并数组
Bmerge = [21,42,39,12,15,3,3,3,3,15,6,3,3,3,12,30,3,3,30]

for i in range(0,83):
	sheet.cell(3*i+5,3,Ctext[i])

for i in range(1,17):
	sheet.merge_cells(start_row=1,start_column=i,end_row=4,end_column=i)
	cell = sheet.cell(1,i)
	cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("A5:A145")
cell = sheet.cell(5,1)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("A146:A175")
cell = sheet.cell(146,1)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("A176:A223")
cell = sheet.cell(176,1)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("A224:A256")
cell = sheet.cell(224,1)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("A257:A271")
cell = sheet.cell(257,1)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("A272:C274")
cell = sheet.cell(272,1)
cell.alignment = Alignment(horizontal='center', vertical='center')

i = 5
for count in range(83):
	sheet.merge_cells(start_row=i,start_column=3,end_row=i+2,end_column=3)
	cell = sheet.cell(i,3)
	cell.alignment = Alignment(horizontal='center', vertical='center')
	i += 3

'''
sheet.merge_cells("B5:B25")
cell = sheet.cell(5,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B26:B67")
cell = sheet.cell(26,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B68:B103")
cell = sheet.cell(68,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B104:B115")
cell = sheet.cell(104,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B116:B130")
cell = sheet.cell(116,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B131:B133")
cell = sheet.cell(131,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B134:B136")
cell = sheet.cell(134,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B137:B139")
cell = sheet.cell(137,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B140:C142")
cell = sheet.cell(140,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B143:B160")
cell = sheet.cell(143,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B161:B166")
cell = sheet.cell(161,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B167:B169")
cell = sheet.cell(167,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B170:C172")
cell = sheet.cell(170,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B173:B214")
cell = sheet.cell(173,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B215:B217")
cell = sheet.cell(215,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B218:B220")
cell = sheet.cell(218,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B221:C223")
cell = sheet.cell(221,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B224:B253")
cell = sheet.cell(224,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B254:C256")
cell = sheet.cell(254,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B5:B25")
cell = sheet.cell(5,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B257:C259")
cell = sheet.cell(257,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B260:C262")
cell = sheet.cell(260,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B263:C265")
cell = sheet.cell(263,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B266:C268")
cell = sheet.cell(266,2)
cell.alignment = Alignment(horizontal='center', vertical='center')

sheet.merge_cells("B269:C271")
cell = sheet.cell(269,2)
cell.alignment = Alignment(horizontal='center', vertical='center')
'''


#设置边框
border = Border(left=Side(border_style='thin'), 
                right=Side(border_style='thin'), 
                top=Side(border_style='thin'), 
                bottom=Side(border_style='thin'))
cell_range = "A1:P274"
for row in sheet[cell_range]:
	for cell in row:
		cell.border = border

answer.save("answer.xlsx")