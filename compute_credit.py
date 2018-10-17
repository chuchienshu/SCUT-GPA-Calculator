# -*- coding: utf-8 -*-
'''
Usage: this script is used to calculate the GPA of all classmates.
Time: 2018/10/2
copyright@ 2018 chuchienshu.
'''
import sys, os, multiprocessing, csv
from urllib import request, error
from PIL import Image
from io import BytesIO
from tqdm import tqdm
from collections import OrderedDict
import random
import xlrd
import xlwt
from credit import credits
from termcolor import *

static=['4、中国特色社会主义理论与实践研究','5、综合英语','7、自然辩证法概论' ]
major_credit4jike = ['6、高级操作系统与分布式系统',
                    '8、高级数据库系统' ,
                    '9、人工智能' ,
                    '10、高级软件体系结构' ,
                    '11、高级计算机网络' ]

major_credit4wangan = ['6、高级操作系统与分布式系统',
                    '8、高级数据库系统' ,
                    '13、高级信息安全' ]

def change_score(original, mean):
    # print(type(mean))
    # print(mean)
    return 85.0 * original / mean

nonmajor_factor = 0.8

def parse_data_xls(data_file):
    dataset = []
    workbook = xlrd.open_workbook(data_file)
    table = workbook.sheets()[0]# 
    for row in range(table.nrows):
        dataset.append(table.row_values(row))
    course_name = [c.strip() for c in dataset[0]]

    table1 = workbook.sheets()[2]# 
    aver_credit = table1.row_values(1)

    print(aver_credit)

    del dataset[0] # remove the row title.

    wk = xlwt.Workbook()
    st = wk.add_sheet('final')
    
    for id, ds in enumerate(dataset):
        # print(ds)
        major_credits = 0.
        nonmajor_credits = 0.
        all_scores = 0.
        print('#################################################################')
        for i in range(3, len(course_name)):
            if ds[i] == '':
                continue
            
            if course_name[i] in static:
                major_credits += credits[course_name[i]]
                all_scores += (ds[i] * credits[course_name[i]])
                # print(colored('公共课: %s ，\t得分： %f \t学分： %d ' % ( course_name[i], ds[i], credits[course_name[i]]),"yellow"))
            
            elif ds[0] == 1 and course_name[i] in major_credit4jike:# just for jike.
                major_credits += credits[course_name[i]]
                standard_score = change_score(ds[i], aver_credit[i])
                all_scores += (standard_score * credits[course_name[i]])
                # print(colored('必修课: %s ，\t得分： %f \t学分： %d ' % ( course_name[i], ds[i], credits[course_name[i]]),"red"))
            elif ds[0] == 2 and course_name[i] in major_credit4wangan:# just for wangan.
                major_credits += credits[course_name[i]]
                standard_score = change_score(ds[i], aver_credit[i])
                all_scores += (standard_score * credits[course_name[i]])
                # print(colored('必修课: %s ，\t得分： %f \t学分： %d ' % ( course_name[i], ds[i], credits[course_name[i]]),"red"))
            
            else:
                nonmajor_credits += credits[course_name[i]]
                standard_score = change_score(ds[i], aver_credit[i])
                all_scores += (standard_score * credits[course_name[i]] * nonmajor_factor)
                # print('选修课: %s ，\t得分： %f \t学分： %d ' % ( course_name[i], ds[i], credits[course_name[i]]))

        all_credits = major_credits + nonmajor_credits * nonmajor_factor
        print(colored('%d %s 的必修课学分 %d, 选修课学分 %d,总学分 %d, 最终成绩为 %4f' % (ds[0],ds[2],major_credits, nonmajor_credits,major_credits + nonmajor_credits, all_scores/all_credits),"green"))
        st.write(id,0,ds[2])
        st.write(id,1,all_scores/all_credits)
        st.write(id,2,ds[0])

    wk.save('final.xls') 
    return dataset

def loader():
    # if len(sys.argv) != 2:
    #     print('Syntax: {} <data_file.csv> <output_dir/>'.format(sys.argv[0]))
    #     sys.exit(0)
    # data_file = sys.argv[1]
    data_file = 'wangan.csv'
    # xls_file = 'credit_76.xls'
    xls_file = '85.xls'

    data = parse_data_xls( xls_file)

if __name__ == '__main__':
    loader()
