from Utility_Class import *
import os
import pandas as pd
import numpy as np
from xlsxwriter.workbook import Workbook

class Data_Processing_Type_1:
    # 投资治理表,委托投资表问题1-18数据处理类
    '''
    一、统计内容：
    1.第0列统计字符'√'
    2.第一列'其他，请填写：战略投资部'，取出'战略投资部'
    3.第2列统计普通字符
    4.首页的公司信息
    二、数据类型：
    1.无内容处：显示为空白,NaN型
    2.'√'：字符型
    3.题号：整数型
    三、注：
    1.大坑：pandas读取excel空白部分默认为NaN数据类型，务必将其替换为字符等类型！不要使用np.isnan()等方法做类型判断，坑。
    '''
    def __init__(self, data_path, result_path, sheet_name):
        '''
        :param data_path: The path of the excel file to be processed
        :param result_path: The path of the excel file to be output
        :return:
        '''
        self.data_path = data_path
        self.result_path = result_path
        # self.sheet_name = '投资治理'
        # self.sheet_name = '委托投资'
        self.sheet_name = sheet_name


    def __get_sheet_result(self, excel_path):
        '''
        :param excel_path: The path of the excel file to be processed.
        :return: Processed data. Data type is a ndarray
        '''
        # Load the file
        data = pd.read_excel(excel_path, sheet_name = self.sheet_name)
        
        # use 'missing' to replase 'NaN' 
        data = data.fillna('missing')

        # Get the column data. The lenth of the column data is same.
        data_column_1 = data.iloc[:,0]  # column 1, '√'
        data_column_2 = data.iloc[:,1]  # column 2, '其他，请填写：'
        data_column_3 = data.iloc[:,2]  # column 3, answer of question
        
        # extraction data
        
        # Record the index of question
        data_location = []
        for i in range(len(data_column_1)):
            if isInt(data_column_1[i]):
                data_location.append(i)
                
        
        data_result = []
        data_result = np.array(data_result)
        # Dataframe to ndarray
        data_column_1 = data_column_1.values
        data_column_2 = data_column_2.values
        data_column_3 = data_column_3.values
        
        # 注：第三列数据处理方法保留
        for i in range(len(data_location)):
            
            # Processing the first column
            if i+1 == len(data_location):
                problem_data = list_split(data_column_1, data_location[i])
            else:
                problem_data = list_split(data_column_1, data_location[i], data_location[i+1])
            
            # # Processing the second column
            # if problem_data[-1] == '√':
            #     string = data_column_2[data_location[i]+len(problem_data)-1]
            #     if string[:2] == '其他':
            #         # use "战略投资部" to replase '√'
            #         problem_data[-1] = string_split(string, 7)
            
            # Processing the second column
            string = data_column_2[data_location[i]+len(problem_data)-1]
            if string[:2] == '其他':
                # use "战略投资部" to replase '√'
                problem_data[-1] = string_split(string, 7)


            # Processing the third column
            if len(problem_data) == 1:
                problem_data = np.append(problem_data, data_column_3[data_location[i]])
            
            # delet the index
            problem_data = np.delete(problem_data, 0)
            
            # store the result
            data_result = np.append(data_result, problem_data)


        # Use 'None' to replace 'NaN' data structure
        for i in range(len(data_result)):
            if data_result[i] == 'missing':
                data_result[i] = None       
        
        print(problem_data)
        # print(problem_data)
        # print(problem_data)


        return data_result
  
    
    def get_company_message(self, excel_path):
        # get company
        sheet_name = '首页'
        data = pd.read_excel(excel_path, sheet_name = sheet_name)
        data = data.fillna('missing')  # use 'missing' to replase 'NaN' 
        data_column_4 = data.iloc[:,3]  # The information is in column 4
        data_column_4 = data_column_4.values
        data_result = []
        data_result = np.array(data_result)
        for i in data_column_4:
            if i != 'missing':
                data_result = np.append(data_result, i)
        return data_result

    
    def get_result(self):
        # merge the company_message and sheet_result
        data_path = self.data_path
        excel_path = get_filenames(data_path)
        end_xls = self.result_path
        sheet_name = self.sheet_name
        lst = []
        for file_names in excel_path:
            user_message = self.get_company_message(file_names)
            data = self.__get_sheet_result(file_names)
            data = np.append(user_message, data)
            lst.append(data)
        list_to_excel(end_xls, sheet_name, lst)
        print('Sheet: '+self.sheet_name+' 处理成功')


class Data_Processing_Type_2:
    # 委托投资表问题数据处理类
    # 输入内容为'委托投资'
    '''
    一、统计内容：
    1.委托投资表中，问题19的内容
    2.首页的公司信息
    二、数据类型：
    1.无内容处：显示为空白,NaN型
    2.内容处：字符型，浮点/整数型
    '''
    def __init__(self, data_path, result_path, sheet_name):
        '''
        :param data_path: The path of the excel file to be processed
        :param result_path: The path of the excel file to be output
        :return:
        '''
        self.data_path = data_path
        self.result_path = result_path
        self.sheet_name = sheet_name
        # self.sheet_name = '委托投资'


    def get_company_message(self, excel_path):
        # get company
        sheet_name = '首页'
        data = pd.read_excel(excel_path, sheet_name = sheet_name)
        data = data.fillna('missing')  # use 'missing' to replase 'NaN' 
        data_column_4 = data.iloc[:,3]  # The information is in column 4
        data_column_4 = data_column_4.values
        data_result = []
        data_result = np.array(data_result)
        for i in data_column_4:
            if i != 'missing':
                data_result = np.append(data_result, i)
        return data_result


    def __get_sheet_result(self, excel_path):
        '''
        :param excel_path: The path of the excel file to be processed.
        :return: Processed data. Data type is a ndarray
        '''       
        data = pd.read_excel(excel_path, sheet_name = self.sheet_name)
        data = data.fillna('missing')  # 使用missing填充np.nan，防止后期问题
        data_column_5 = data.iloc[:,-1]  # 以最后一列数据数量为基准，读取有效信息的行数
        data_column_5 = data_column_5.values
        data_valid_row_num = 0  # 有效信息的行数
        for i in data_column_5:
            if i != 'missing':
                data_valid_row_num = data_valid_row_num+1
        data_valid = data.iloc[:,list(range(2, 15))]  # 提取问题19第3-15列，三列为管理人名称
        # 提取行，第102行是问题19第一个回答，（102+有效信息量-1）是为最后一个回答
        # excel内行号为1的，被删去，行号为2的为第0行。因此，行号102的，在pandas对应过来为100
        data_valid = data_valid.iloc[100:100+data_valid_row_num-1,]  
        data_valid = data_valid.values
        return data_valid  # 返回2维数组

    

    def get_result(self):
        # merge the company_message and sheet_result
        data_path = self.data_path
        excel_path = get_filenames(data_path)
        end_xls = self.result_path
        sheet_name = self.sheet_name
        lst = []
        for file_names in excel_path:
            user_message = self.get_company_message(file_names)
            data = self.__get_sheet_result(file_names)
            for data_valid in data:
                lst.append(np.append(user_message, data_valid))
        list_to_excel(end_xls, sheet_name, lst)
        print('Sheet: '+self.sheet_name+' 处理成功')


class Data_Processing_Type_3:
    # 委托投资表问题19专用数据处理类
    # 需要把问题19单独列出一个表
    '委托投资19'
    '''
    一、统计内容：
    1.C-O列所有数据
    2.首页的公司信息
    二、数据类型：
    1.无内容处：显示为空白,NaN型
    2.内容处：字符型，浮点/整数型
    '''
    def __init__(self, data_path, result_path, sheet_name):
        '''
        :param data_path: The path of the excel file to be processed
        :param result_path: The path of the excel file to be output
        :return:
        '''
        self.data_path = data_path
        self.result_path = result_path
        self.sheet_name = sheet_name
        # self.sheet_name = '委托投资19'
    
    
    def get_company_message(self, excel_path):
        # get company
        sheet_name = '首页'
        data = pd.read_excel(excel_path, sheet_name = sheet_name)
        data = data.fillna('missing')  # use 'missing' to replase 'NaN' 
        data_column_4 = data.iloc[:,3]  # The information is in column 4
        data_column_4 = data_column_4.values
        data_result = []
        data_result = np.array(data_result)
        for i in data_column_4:
            if i != 'missing':
                data_result = np.append(data_result, i)
        return data_result


    def __get_sheet_result(self, excel_path):
        '''
        :param excel_path: The path of the excel file to be processed.
        :return: Processed data. Data type is a ndarray
        '''
        data = pd.read_excel(excel_path, sheet_name = self.sheet_name)
        data = data.fillna('missing')  # 使用missing填充np.nan，防止后期问题
        data_column_5 = data.iloc[:,-1]  # 以最后一列数据数量为基准，读取有效信息的行数
        data_column_5 = data_column_5.values
        data_valid_row_num = 0  # 有效信息的行数
        for i in data_column_5:
            if i != 'missing':
                data_valid_row_num = data_valid_row_num+1
        data_valid = data.iloc[:,list(range(2, 15))]  # 提取第3-15列，三列为管理人名称
        data_valid = data_valid.iloc[0:3,]  # 提取0-2行，第0行即为数据
        data_valid = data_valid.values
        return data_valid  # 返回2维数组
        

    def get_result(self):
        # merge the company_message and sheet_result
        data_path = self.data_path
        excel_path = get_filenames(data_path)
        end_xls = self.result_path
        sheet_name = self.sheet_name
        lst = []
        for file_names in excel_path:
            user_message = self.get_company_message(file_names)
            data = self.__get_sheet_result(file_names)
            for data_valid in data:
                lst.append(np.append(user_message, data_valid))
        list_to_excel(end_xls, sheet_name, lst)
        print('Sheet: '+self.sheet_name+' 处理成功')
