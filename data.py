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

data_rows = 66
data_cols = 10

data = []
for i in range(data_rows):
	row = []
	for j in range(data_cols):
		row.append(0)
	data.append(row)
