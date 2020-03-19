import xlrd
import pandas as pd
import numpy as np
from collections import Counter
'''
#打开test.xlsx
book = xlrd.open_workbook('test.xlsx')
print('sheet页内容：',book.sheet_names())
sheet = book.sheet_by_index(0)
rows = sheet.nrows
cols = sheet.ncols
print("%d行,%d列" %(rows,cols))
#第3行
print(sheet.row_values(2))
print(type(sheet.row_values(2)))
#第2行第2列
print(sheet.cell_value(1,1))
'''
#建立处理的类
class handle_excel(object):
    def __init__(self,file_name):
        #文件的名字
        self.name=file_name
        #self.num = tem
    #读入xlsx文件
    #返回3个dataframe格式的文件
    def read_file(self):
        #print(pd.read_excel(io=self.name))
        df = pd.read_excel(io=self.name)
        #B_ID==1的信息
        df1= df[df['B_ID']==1]
        #B_ID==2的信息
        df2 = df[df['B_ID'] == 2]
        #B_ID==3的信息
        df3 = df[df['B_ID'] == 3]
        #print(df1)
        #print(df.groupby(df['B_ID']))
        #ID_info = df['B_ID'].unique()
        '''
                for tem_ID in ID_info:
            tem_data = df[df['class'].isin([tem_ID])]
            exec("df%s = tem_data"%tem_ID)
        print("df1")
        print(df1)
        print("df1")
        print(df2)
        print("df1")
        print(df3)
        '''
        return df1,df2,df3
    #计算数字的位数，在检测标准1的第三个方面用
    def getLength(self,number):
        Length = 0
        while number != 0:
            Length += 1
            number = number // 10  # 关键，整数除法去掉最右边的一位
        return Length
    #检测标准1
    def handle_one(self,data_name):

        #计算xlsx的列数
        cols_num = data_name.shape[1]
        # 计算xlsx的行数
        rows_num = data_name.shape[0]
        #正值的数量，第一个方面用
        positive_num = [0] * cols_num
        #负值的数量，第一个方面用
        negative_num = [0] * cols_num
        #diff = [[0]*rows_num]*(cols_num-2)
        #每一列中每个位置的元素的位数
        diff=[[0 for _ in range(rows_num)] for _ in range(cols_num-2)]
        positive_value=[]
        negative_value=[]
        #print(diff[1][2])
        #print(diff)
        #计算每一列中正值和负值的数量
        i=0
        for index,col in data_name.iteritems():
            if index=="Equity " or index == "B_ID":
                continue
            for x in col.values:
                #print("##",index)
                if np.isnan(x) == False and x > 0:
                # print("***",row[i])
                    positive_num[i] = positive_num[i] + 1
                    positive_value.append(x)
                if np.isnan(x) == False and x < 0:
                # print("###",row[i])
                    negative_num[i] = negative_num[i] + 1
                    negative_value.append(x)
                if x==0:
                    print(data_name['B_ID'].values[1], index, x)
            if len(positive_value)==1:
                print(data_name['B_ID'].values[1], index, positive_value[0])
            if len(negative_value)==1:
                print(data_name['B_ID'].values[1], index, negative_value[0])
            positive_value = []
            negative_value = []
            i=i+1

            #print(row)
            '''

            for i in range(2,cols_num):
                if np.isnan(row[i]) == False and row[i] > 0:
                    # print("***",row[i])
                    positive_num[i] = positive_num[i] + 1
                if np.isnan(row[i]) == False and row[i] < 0:
                    # print("###",row[i])
                    negative_num[i] = negative_num[i] + 1
                if row[i] == 0:
                    print("0值")
                        '''
        #for i in range(2,cols_num):
            #print(positive_num[i],negative_num[i])
        data_name1=data_name[['2016_Q1_A','2016_Q2_A','2016_Q3_A','2016_Q4_A','2016_A_A']]
        #data_name = data_name.drop(['Equity'])
        #print(data_name1)
        j=0
        #print(diff)
        #计算每个位置的元素的位数，若是空值记为0
        for inedx,col in data_name1.iteritems():
            #print(diff)
            for i in range(rows_num):
                #print(col.values[i])
                if np.isnan(col.values[i])==False:
                    diff[j][i]=self.getLength(abs(col.values[i]))
                else:
                    diff[j][i] =0
                #print(diff[j][i])
            #print(diff)
            j=j+1
            #print("j==",j)
        #print(diff)
        #i表示第几列
        i=0
        #符号要求的次数的集合
        result3_value=[]
        #符号要求的列的集合
        result3_index=[]
        for list_tem in diff:
            #制作字典，key是位数，value是出现的次数
            b = dict(Counter(list_tem))
            #次数为rows_num-1的字典
            tem={key:value for key,value in b.items() if value==rows_num-1}
            #如果存在这样的字典，继续处理
            if tem:
                #print(i)
                result3_index.append(i)
                result3_value.append(list(tem.keys())[0])
                #print(list(tem.keys())[0])
            i=i+1
        #print(result3_index)
        #print(result3_value)
        #print("ZZZZZ",result3_index,result3_value)
        #print("diff==",diff)
        for index,value in zip(result3_index,result3_value):
            x=0
            for i in diff[index]:

                if i !=value and i !=0:
                    #print("i==",i)
                    #print("value==",value)
                    #print(i,x)
                    #print("UUUUUUUUUUUUUUUUUUU")
                    #print("**1",data_name['B_ID'].values[1])
                    #print("**2",data_name1.columns[index])
                    #print("**3",data_name1[data_name1.columns[index]].values[x])
                    print(data_name['B_ID'].values[1],data_name1.columns[index],data_name1[data_name1.columns[index]].values[x])
                x=x+1

        #print(data_name[data_name.columns[1]])

    # 检测标准2
    def handle_two(self, data_name):
        #data_name1 = data_name[['2016_Q1_A', '2016_Q2_A', '2016_Q3_A', '2016_Q4_A', '2016_A_A']]
        for index, col in data_name.iteritems():
            if index=="Equity " or index == "B_ID":
                continue
            #print("index==",index)
            #去掉空值
            col_list = [x for x in col.values if np.isnan(x)==False]
            #print(col_list)
            #每一列的平均数
            average_a = np.mean(col_list)
            # 每一列的中位数
            median_a = np.median(col_list)
            #print(average_a,median_a)
            for i in col_list:
                col_tem = list(col_list)
                #去掉一个数
                col_tem.remove(i)
                #print("***",col_tem)
                #去掉一个数之后的平均值
                average_t = np.mean(col_tem)
                # 去掉一个数之后的中位数
                median_t = np.median(col_tem)
                #print(average_t, median_t)
                #判断共识
                juage_average=(average_a-average_t)/average_t
                juage_median = (median_a - median_t) / median_t
                if juage_average>0.2:
                    print(data_name['B_ID'].values[1], index,i)
                if juage_median>0.2:
                    print(data_name['B_ID'].values[1], index,i)

    # 检测标准3
    def handle_three(self, data_name):
        cols_num = data_name.shape[1]
        rows_num = data_name.shape[0]
        sum_4=[0]*rows_num
        #(sum_4)
        for index, col in data_name.iteritems():
            if index=="Equity " or index == "B_ID":
                continue
            #不是A_A的情况，相加
            if index.find("A_A")==-1:
                #print(col.values)
                sum_4=np.sum([sum_4,col.values],axis=0)
            # 是A_A的情况
            if index.find("A_A")!=-1:
                sum_A = col.values
                #print(sum_4)
                sum_4_left = sum_4*0.8
                sum_4_right = sum_4*1.2
                #print(sum_4_left,sum_4_right)
                for i in range(len(sum_A)):
                    if sum_A[i]<sum_4_left[i] or sum_A[i]>sum_4_right[i]:
                        print(data_name['B_ID'].values[1], index,sum_A[i])
                sum_4 = [0] * rows_num
             #   print(index)
#主程序
if __name__ == '__main__':
    file_name = input("请输入文件名: ")
    test_file = handle_excel(file_name)
    test_data1,test_data2,test_data3 = test_file.read_file()
    #print(test_data1)
    #print(test_data2)
    #print(test_data3)
    print("检测标准1的结果:")
    test_file.handle_one(test_data1)
    test_file.handle_one(test_data2)
    test_file.handle_one(test_data3)
    print("检测标准2的结果:")
    test_file.handle_two(test_data1)
    test_file.handle_two(test_data2)
    test_file.handle_two(test_data3)
    print("检测标准3的结果:")
    test_file.handle_three(test_data1)
    test_file.handle_three(test_data2)
    #test_file.handle_three(test_data3)
    #test_file.handle_one(test_data2)
    #test_file.handle_one(test_data3)
