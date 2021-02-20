# -*- coding: utf-8 -*-
"""
Created on Wed Jan 20 09:45:00 2021
@author: Sunshuo, Pengcheng Song
"""

from ppt_maker.maker import PrsMaker
from pptx.chart.data import ChartData
import time as time

prs = PrsMaker("./test/test_template.pptx")

text_data = {
        'month': '2021年1月',
        'update_user': 'Pengcheng Song',
        'update_date': '2021/02/20'
    }

chart_data = {
        'chart_1': {
            'title': '猪肉炖粉条配料比例',
            'categories': ['猪肉','白菜','粉条','水'],
            'series': [
                {
                    'name': '含量',
                    'values': (0.2,0.3,0.2,0.3)
                }
            ]
        },
        'chart_2': {
            'title': '',
            'categories': ['2019/1','2019/2','2019/3','2019/4','2019/5','2019/6','2019/7','2019/8','2019/9','2019/10','2019/11','2019/12'],
            'series': [
                {
                    'name': '猪肉',
                    'values': (111,143,279,245,148,131,144,146,158,111,122,149)
                },
                {
                    'name': '白菜',
                    'values': (67,64,82,85,76,84,73,71,74,76,74,84)
                },
                {
                    'name': '粉条',
                    'values': (7,6,5,19,35,13,15,16,15,15,34,42)
                },
                {
                    'name': '水',
                    'values': (4,5,5,5,4,4,5,5,4,4,5,5)
                },
                {
                    'name': '当月每份价格',
                    'values': (32.99,29.99,25.99,31.99,21.99,28.99,40.99,34.99,37.99,33.99,32.99,26.99)
                }
            ]
        }
    }

table_data = [['万元','2019', '2020','同比'],
    ['市场日均销量','1000','1500','50%'],
    ['本店日均销量','125','200','66%'],
    ['营业支出','3,000','4,000','33%'],
    ['原材料','2,500','3,000','20%'],
    ['猪肉','1,000','1,200','20%'],
    ['白菜','800','1000','25%'],
    ['粉条','400','500','25%'],
    ['水','300','300','0%'],
    ['人工','500','1000','100%']]

prs.text_inject(text_data)
prs.chart_inject(chart_data)
prs.table_inject(table_data)
                    
prs.save('./test/test.pptx')