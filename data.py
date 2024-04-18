from openpyxl import Workbook,load_workbook

llshw = load_workbook("llshw.xlsx",data_only=True)

zhsheet = llshw["工作表"]
ygssheet = llshw["医工所项目表"]
hzwsheet = llshw["杭州湾项目表"]

#团队负责人
TeamLeader = ["许高杰","曹彦伟","曹鸿涛","刘兆平","诸葛飞","江南","王立平","汪爱英","宋振纶",
"曾志翔","张涛","茅东升","蒲吉斌","黄政仁","陈涛","刘小青","常可可","朱锦","郑文革","姚强",
"方省众","陈鹏","刘富","张永刚","王震","张浩","张建","余海斌","黄庆","张蕾","王军强","李国伟",
"李润伟","满其奎","闫阿儒","张驰","祝颖丹","蒋俊","肖江剑","赵夙","张文武","韦超","陈新民",
"宋伟杰","叶继春","向超宇","葛子义","陈亮","况永波","尹宏峰","张亚杰","姚霞银","官万兵","杨明辉",
"夏永高","陆之毅","谢银君","赵一天","张建涛","吴爱国","李华","左国坤","王荣","郑建萍","郭明全",
"李辉"]

#23年到位
data23 = [[6,99,0,0,0,345,0,127,0,47],[191,445,0,20,35,165,0,0,13,224.9],
[42,109,0,90,44,0,120,0,0,60.82],[0,414,0,0,35,3,0,0,0,153.48],
[89,0,0,0,0,6,0,0,0,0],[60,247,0,17,0,409,45,1437,30,283.83],
[1036,503,0,840,35,441,90,1023,130,795.64],[150,41,0,0,20,252,0,11,0,127.2],
[2,44,0,0,30,200,60,0,50,0],[50,0,0,0,0,45,0,43,115,0],
[126,0,0,0,35,78,0,0,6,0],[18,0,0,0,0,27,32,0,0,0],
[267,248,0,0,410,445,0,300,18,327.76],[188,1334,0,0,0,0,0,346,0,955],
[234,161,0,0,90,192,60,0,15,49.85],[233,152,0,0,10,150,0,260,0,38],
[39,365,0,54,0,5,0,0,0,322],[0,0,0,75,0,132,0,0,0,0],
[69,604,0,0,103,300,0,0,114,725.94],[24,0,0,0,256,365,0,33,0,6]
]


#创建空数组
data_rows = 200
data_cols = 10
data = []
for i in range(data_rows):
	row = []
	for j in range(data_cols):
		row.append(0)
	data.append(row)


for row in zhsheet:
	if(row[74].value == "科技发展部" or row[74].value == "科技发展部（重任处）"):
		#高技术
		if(row[73].value == "高技术"):
			#到位经费
			if (row[64].value == "YES"):
				TeamIndex = TeamLeader.index(row[67].value)
				if(row[62].value != None):
					data[3*TeamIndex + 1][7] += row[62].value
				if(row[63].value != None):
					data[3*TeamIndex + 1][9] += row[63].value
			#合同经费
			if(row[7].value == 24):
				TeamIndex = TeamLeader.index(row[67].value)
				data[3*TeamIndex][7] += row[8].value
		#国际合作
		if(row[73].value == "国际（地区）合作与交流项目" or \
			row[73].value == "国际工程科技高端论坛" or \
			row[73].value == "国际合作" or row[73].value == "国际会议" or \
			row[73].value == "国际交流" or row[73].value == "国际人才" or \
			row[73].value == "国际载体" or row[73].value == "海外及港澳学者合作研究基金" or \
			row[73].value == "外国青年学者项目" or row[73].value == "外国学者研究基金(资深项目)"):
			#到位经费
			if (row[64].value == "YES"):
				TeamIndex = TeamLeader.index(row[67].value)
				if(row[62].value != None):
					data[3*TeamIndex + 1][6] += row[62].value
				if(row[63].value != None):
					data[3*TeamIndex + 1][9] += row[63].value
			#合同经费
			if(row[7].value == 24):
				TeamIndex = TeamLeader.index(row[67].value)
				data[3*TeamIndex][6] += row[8].value

for i in range(200) :
	if(i%3 == 0):
		print(TeamLeader[int(i/3)])
	print(data[i])