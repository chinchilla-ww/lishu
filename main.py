import pandas as pd
import matplotlib.pyplot as plt

#平均数计算
def avg(mylist = []):
    sum = 0
    for i in range(0,len(mylist)):
        sum += int(mylist[i])
    avge = sum / len(mylist)
    return avge

#计算自动控制理论A相应所需数值
data_1 = pd.read_excel('附件3.1-自动控制原理A成绩及其达成度分析.xlsx')
data_1 = data_1[['平时成绩1', '平时成绩2', '期末成绩1', '期末成绩2']]
a_1 = avg(list(data_1['平时成绩1'].tolist()))
b_1 = avg(list(data_1['平时成绩2'].tolist()))
c_1 = avg(list(data_1['期末成绩1'].tolist()))
d_1 = avg(list(data_1['期末成绩2'].tolist()))
print(a_1, b_1, c_1, d_1)

A_1 = 10
B_1 = 10
C_1 = 30
D_1 = 50
obj_1_1 = (a_1+b_1+d_1)/(A_1+B_1+D_1)
obj_1_2 = (a_1+b_1+c_1+d_1)/(A_1+B_1+C_1+D_1)
obj_1_3 = (a_1+c_1+d_1)/(A_1+C_1+D_1)
obj_1_4 = (a_1+b_1+c_1)/(A_1+B_1+C_1)
obj_1 = 0.2*obj_1_1+0.2*obj_1_2+0.4*obj_1_3+0.2*obj_1_4
gr_1_12 = 0.4*obj_1_1+0.6*obj_1_2
gr_1_13 = 0.2*obj_1_2+0.5*obj_1_3+0.3*obj_1_4

# 定义数据
data1 = [obj_1_1, obj_1_2, obj_1_3, obj_1_4, obj_1]
print(data1)
# 绘制柱状图
plt.bar(range(len(data1)), data1)
# 设置横轴标签和标题
plt.xlabel('Data Index')
plt.ylabel('Value')
plt.title('Course 1 Data Visualization')
plt.xticks(range(5),(1,2,3,4,sum))
#保存图像
plt.savefig('自动控制理论A课程目标达成度数据可视化结果.jpg')
# 显示图形
plt.show()

# 将数据框输出为Excel文件
data = {'课程名称': ['自动控制理论A'], '课程目标1达成度': [obj_1_1], '课程目标2达成度': [obj_1_2], '课程目标3达成度': [obj_1_3], '课程目标4达成度': [obj_1_4], '课程目标总体达成度': [obj_1], '毕业要求12': [gr_1_12], '毕业要求13': [gr_1_13]}
df = pd.DataFrame(data)
df.to_excel('自动控制理论A输出结果.xlsx', index=False)


#计算散控制系统（DCS）相应所需数值
data_2 = pd.read_excel('附件3.2-分散控制系统（DCS）综合实践A成绩.xlsx')
data_2 = data_2[['平时成绩', '实验结果', '实验报告']]
a_2 = avg(list(data_2['平时成绩'].tolist()))
b_2 = avg(list(data_2['实验结果'].tolist()))
c_2 = avg(list(data_2['实验报告'].tolist()))
print(a_2, b_2, c_2)

A_2 = 20
B_2 = 40
C_2 = 40
obj_2_1 = (a_2+b_2)/(A_2+B_2)
obj_2_2 = (a_2+c_2)/(A_2+C_2)
obj_2_3 = (a_2+c_2+b_2)/(A_2+C_2+B_2)
obj_2_4 = (b_2+c_2)/(B_2+C_2)
obj_2 = 0.2*obj_2_1+0.2*obj_2_2+0.4*obj_2_3+0.2*obj_2_4
gr_2_51 = 0.5*obj_2_1+0.3*obj_2_3+0.2*obj_2_4
gr_2_52 = 0.3*obj_2_2+0.7*obj_2_4
gr_2_53 = 0.4*obj_2_2+0.6*obj_2_3

# 定义数据
data2 = [obj_2_1, obj_2_2, obj_2_3, obj_2_4, obj_2]
print(data2)

# 绘制柱状图
plt.bar(range(len(data2)), data2)
# 设置横轴标签和标题
plt.xlabel('Data Index')
plt.ylabel('Value')
plt.title('Course 2 Data Visualization')
plt.xticks(range(5),(1,2,3,4,sum))
#保存图像
plt.savefig('计算散控制系统（DCS）课程目标达成度数据可视化结果.jpg')
# 显示图形
plt.show()

# 将数据框输出为Excel文件
data = {'课程名称': ['计算散控制系统（DCS）'], '课程目标1达成度': [obj_2_1], '课程目标2达成度': [obj_2_2], '课程目标3达成度': [obj_2_3], '课程目标4达成度': [obj_2_4], '课程目标总体达成度': [obj_2], '毕业要求51': [gr_2_51], '毕业要求52': [gr_2_52], '毕业要求53': [gr_2_53]}
df = pd.DataFrame(data)
df.to_excel('计算散控制系统（DCS）输出结果.xlsx', index=False)