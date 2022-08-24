# -*- coding: utf-8 -*-
"""
Created on Wed Aug 24 18:02:28 2022

@author: lovro
v 0.1.0
"""

import pandas as pd
from pandas import ExcelWriter

MC = pd.read_csv("MC.csv", usecols=["Email Address"])
MC.rename(columns={"Email Address": 'email'}, inplace=True)
MC['source'] = 'MC'

ML = pd.read_excel('ML.xlsx', usecols=[0], names=['email'])
ML.email = ML.email.str.lower()
ML.sort_values(['email'], inplace=True)
ML.drop_duplicates(subset=['email'], inplace=True)
ML['source'] = 'ML'
JOINED = pd.concat([ML, MC])
JOINED.sort_values(['email'], inplace=True)
JOINED.drop_duplicates(subset=['email'], inplace=True, keep=False)
drop_MC = JOINED['source'] == 'MC'
JOINED = JOINED[~drop_MC]
JOINED.drop(columns=['source'], axis=1, inplace=True)

# =============================================================================
# # To excel
# =============================================================================
excel = ExcelWriter("ML_not in MC.xlsx", engine='xlsxwriter')
JOINED.to_excel(excel, 'Mail List', index=False)
excel.save()
