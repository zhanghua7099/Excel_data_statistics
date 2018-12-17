from data_processing import *
import time


start = time.time()
a = Data_Processing_Type_1("./data", "./result/target.xlsx", 'sheet_name')
a.get_result()

