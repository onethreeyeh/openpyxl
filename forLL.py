from openpyxl import Workbook,load_workbook
from openpyxl.styles import Alignment

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
sheet["B104"].value = "先进能源材料工程实验室"
sheet["B116"].value = "磁材实验室"
sheet["B131"].value = "磁性材料应用技术研究中心"
sheet["B134"].value = "稀土永磁材料联合创新中心"
sheet["B140"].value = "材料所总计"
sheet["B161"].value = "激光极端制造研究中心"
sheet["B170"].value = "先进制造所总计"
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

Ctext = ["许高杰","曹彦伟","曹鸿涛","向超宇","诸葛飞","先进纳米实验室","纳米小计",\
"江南","王立平","汪爱英","宋振纶","曾志翔","张涛","茅东升","蒲吉斌","黄政仁","陈涛","刘小青","常可可","中科院海洋新材料与应用技术重点实验室","海洋小计",\
"朱锦","郑文革","姚强","方省众","陈鹏","刘富","张永刚","王震","张浩","余海斌","高分子与复合材料实验室","高分子小计",\
"黄庆","张蕾","先进能源材料工程实验室","先进能源小计",\
"王军强","李国伟","李润伟","中科院磁性材料与器件重点实验室","磁材小计",\
"满其奎","闫阿儒","二级所","","张驰","祝颖丹","陈新民","蒋俊","肖江剑","赵夙","张文武","韦超","二级所","",\
"宋伟杰","叶继春","陆之毅","葛子义","陈亮","况永波","尹宏峰","张亚杰","姚霞银","官万兵","杨明辉","张建","夏永高","谢银君","刘兆平","二级所","",\
"赵一天","张建涛","吴爱国","李华","左国坤","王荣","郑建萍","郭明全","李辉","二级所"]

for i in range(0,83):
	sheet.cell(3*i+5,3,Ctext[i])

for i in range(1,17):
	sheet.merge_cells(start_row=1,start_column=i,end_row=4,end_column=i)
	cell = sheet.cell(1,i)
	cell.alignment = Alignment(horizontal='center', vertical='center')

answer.save("answer.xlsx")