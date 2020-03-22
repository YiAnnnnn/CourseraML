import xlrd
import pandas as pd
import numpy as np
from collections import Counter
from glob import glob
import os

#建立处理的类
class handle_excel(object):
    df_result = pd.DataFrame(columns = ["B_ID", "name", "date", "result"])
    def __init__(self,file_name):
        #文件的名字
        self.name=file_name
        #self.num = tem
    #读入xlsx文件
    #返回3个dataframe格式的文件
    def read_file(self):
        #print(pd.read_excel(io=self.name))
        df = pd.read_excel(io=self.name)
        listType = df['B_ID'].unique()
        for i in range(len(listType)):
            df_list=df[df['B_ID'].isin([listType[i]])]
            self.handle_one(df_list)
            self.handle_two(df_list)
            self.handle_three(df_list)
        #print(self.df_result)
        df_sorted=self.df_result.sort_values(by=['B_ID','date'],ascending=(True,True)).drop_duplicates()
        rows_num = df_sorted.shape[0]
        for i in range(rows_num):
            print("B_ID=",df_sorted.iloc[i].values[0],df_sorted.iloc[i].values[1],"Column=",df_sorted.iloc[i].values[2],df_sorted.iloc[i].values[3],";")

    #计算数字的位数，在检测标准1的第三个方面用
    def getLength(self,number):
        Length = 0
        while number != 0:
            Length += 1
            number = number // 10  # 关键，整数除法去掉最右边的一位
        return Length

        # 检测标准1
    def handle_one(self, data_name):
        # 计算xlsx的列数
        cols_num = data_name.shape[1]
        # 计算xlsx的行数
        rows_num = data_name.shape[0]
        # 正值的数量，第一个方面用
        positive_num = [0] * cols_num
        # 负值的数量，第一个方面用
        negative_num = [0] * cols_num
        # diff = [[0]*rows_num]*(cols_num-2)
        # 每一列中每个位置的元素的位数
        diff = [[0 for _ in range(rows_num)] for _ in range(cols_num - 2)]
        positive_value = []
        negative_value = []
        positive_index = []
        negative_index = []
        # 计算每一列中正值和负值的数量
        i = 0
        for index, col in data_name.iteritems():
            if index == "Equity " or index == "B_ID" or index == "Name":
                continue
            j = 0
            for x in col.values:
                # print("##",index)
                if np.isnan(x) == False and x > 0:
                    # print("***",row[i])
                    positive_num[i] = positive_num[i] + 1
                    positive_value.append(x)
                    positive_index.append(j)
                if np.isnan(x) == False and x < 0:
                    # print("###",row[i])
                    negative_num[i] = negative_num[i] + 1
                    negative_value.append(x)
                    negative_index.append(j)
                if x == 0:
                    #print(data_name['B_ID'].values[j], data_name['Name'].values[j], index, x)
                    self.df_result.loc[len(self.df_result)] = list(
                        [data_name['B_ID'].values[j], data_name['Name'].values[j], index, x])
                j = j + 1
            if len(positive_value) == 1:
                #print(data_name['B_ID'].values[positive_index[0]], data_name['Name'].values[positive_index[0]],index, positive_value[0])
                self.df_result.loc[len(self.df_result)]=list([data_name['B_ID'].values[positive_index[0]], data_name['Name'].values[positive_index[0]],
                      index, positive_value[0]])
            if len(negative_value) == 1:
                #print(data_name['B_ID'].values[negative_index[0]], data_name['Name'].values[negative_index[0]],index, negative_value[0])
                #print("*****",type(list([data_name['B_ID'].values[negative_index[0]], data_name['Name'].values[negative_index[0]],index, negative_value[0]])))
                #self.df_result.append(list([data_name['B_ID'].values[negative_index[0]], data_name['Name'].values[negative_index[0]],index, negative_value[0]]),ignore_index=True)
                self.df_result.loc[len(self.df_result)]=list([data_name['B_ID'].values[negative_index[0]], data_name['Name'].values[negative_index[0]],index, negative_value[0]])
                #print(self.df_result)
            positive_value = []
            negative_value = []
            positive_index = []
            negative_index = []
            i = i + 1
        j = 0
        # 计算每个位置的元素的位数，若是空值记为0
        for index, col in data_name.iteritems():
            if index=="Equity " or index == "B_ID" or index=="Name":
                continue
            for i in range(rows_num):
                # print(col.values[i])
                if np.isnan(col.values[i]) == False:
                    diff[j][i] = self.getLength(abs(col.values[i]))
                else:
                    diff[j][i] = 0
            j = j + 1
        # i表示第几列
        i = 3
        # 符号要求的次数的集合
        result3_value = []
        # 符号要求的列的集合
        result3_index = []
        for list_tem in diff:
            # 制作字典，key是位数，value是出现的次数
            b = dict(Counter(list_tem))
            # 次数为rows_num-1的字典
            tem = {key: value for key, value in b.items() if value == rows_num - 1}
            # 如果存在这样的字典，继续处理
            if tem:
                # print(i)
                result3_index.append(i)
                result3_value.append(list(tem.keys())[0])
                # print(list(tem.keys())[0])
            i = i + 1
        for index, value in zip(result3_index, result3_value):
            x = 0
            for i in diff[index-3]:

                if i != value and i != 0:
                    #print(data_name['B_ID'].values[x], data_name['Name'].values[x],data_name.columns[index],data_name[data_name.columns[index]].values[x])
                    self.df_result.loc[len(self.df_result)] =list([data_name['B_ID'].values[x], data_name['Name'].values[x],data_name.columns[index],
                          data_name[data_name.columns[index]].values[x]])
                x = x + 1

    # 检测标准2
    def handle_two(self, data_name):
        #data_name1 = data_name[['2016_Q1_A', '2016_Q2_A', '2016_Q3_A', '2016_Q4_A', '2016_A_A']]
        for index, col in data_name.iteritems():
            if index=="Equity " or index == "B_ID" or index=="Name":
                continue
            #去掉空值
            col_list = [x for x in col.values if np.isnan(x)==False]
            #每一列的平均数
            average_a = np.mean(col_list)
            # 每一列的中位数
            median_a = np.median(col_list)
            for i in range(len(col)):
                if np.isnan(col.values[i])==True:
                    continue
                col_tem = list(col_list)
                #去掉一个数
                col_tem.remove(col.values[i])
                #去掉一个数之后的平均值
                average_t = np.mean(col_tem)
                # 去掉一个数之后的中位数
                median_t = np.median(col_tem)
                #判断共识
                juage_average=(average_a-average_t)/average_t
                juage_median = (median_a - median_t) / median_t
                if juage_average>0.2:
                    #print(data_name['B_ID'].values[i], data_name['Name'].values[i],index,col.values[i])
                    self.df_result.loc[len(self.df_result)]=list([data_name['B_ID'].values[i], data_name['Name'].values[i],index,col.values[i]])
                if juage_median>0.2:
                    #print(data_name['B_ID'].values[i], data_name['Name'].values[i],index,col.values[i])
                    self.df_result.loc[len(self.df_result)]=list([data_name['B_ID'].values[i], data_name['Name'].values[i],index,col.values[i]])

    # 检测标准3
    def handle_three(self, data_name):
        cols_num = data_name.shape[1]
        rows_num = data_name.shape[0]
        sum_4=[0]*rows_num
        #(sum_4)
        i=0
        for index, col in data_name.iteritems():
            if index=="Equity " or index == "B_ID"or index=="Name":
                continue
            #不是A_A的情况，相加
            if index.find("A_A")==-1:
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
                        #print(data_name['B_ID'].values[i], data_name['Name'].values[i],index,sum_A[i])
                        self.df_result.loc[len(self.df_result)] =list([data_name['B_ID'].values[i], data_name['Name'].values[i],index,sum_A[i]])
                sum_4 = [0] * rows_num
                i=i+1
#主程序
if __name__ == '__main__':
    file_name = input("请输入文件所在文件夹: ")
    filename_join=os.path.join(file_name,'*.xlsx')
    filename_list=glob(filename_join)
    #print(filename_list)
    for file in filename_list:
        print(file)
        test_file = handle_excel(file)
        test_file.read_file()

