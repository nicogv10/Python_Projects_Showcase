import numpy as np
import pandas as pd
import geopandas as gpd
import re
import itertools
from itertools import combinations
from collections import Counter

#here I am reading in the CBP 1987 csv file
CBP87 = pd.read_csv('/Users/nicog-v/Documents/Python Mapping GPDs/Ellison Glaeser Index/4dig_sic.csv')


#in these steps I am using itertools to get every unique combination of 2 SIC codes and putting them into a list
SIC = CBP87['sic']
SIC_uniq = set(SIC)
SIC_combos = list(combinations(SIC_uniq, 2))


#here I storing xm_squared & xm for each state in its own dataframe - will use at the end when I join my
#dataframes and do calculations
CBP87_data = CBP87.loc[:,['fipstate', 'st_share_agg_emp', 'xm_squared']]
#getting rid of duplicate lines (this would mess up our join later)
CBP87_data.drop_duplicates(subset=['fipstate'], keep= "first", inplace=True)

#here I am setting up the empty columns and list that will be used to store my output in a new
#dataframe (which will then be saved in Excel)
cols = ['combo', 'gamma']
lst = []

#SIC_combos is now a list of combos - I can access the industry code by indexing
for combo in SIC_combos:
    #here I am creating 2 separate dataframes - one has info for industry 1 and one has info for industry 2
    industry_1 = CBP87[CBP87['sic'] == combo[0]]
    industry_2 = CBP87[CBP87['sic'] == combo[1]]
    #here I am renaming the emp column and deleting the SIC column (and some other unused columns) for both
    #dataframes, so that I can then join all 3 dataframes with no issues
    industry_1.rename(columns= {'emp':f"emp_{combo[0]}"}, inplace=True)
    industry_1.rename(columns= {'st_ind_share':'state_ind1_share'}, inplace=True)
    del industry_1['sic']
    del industry_1['industry_emp']
    del industry_1['ind_emp_by_st']
    del industry_1['st_share_agg_emp']
    del industry_1['xm_squared']
    industry_2.rename(columns= {'emp':f"emp_{combo[1]}"}, inplace=True)
    industry_2.rename(columns= {'st_ind_share':'state_ind2_share'}, inplace=True)
    del industry_2['sic']
    del industry_2['industry_emp']
    del industry_2['ind_emp_by_st']
    del industry_2['st_share_agg_emp']
    del industry_2['xm_squared']

    #here I am merging/joining the 3 tables so we can do final calculations
    CBP87_merged = industry_1.merge(industry_2, on='fipstate', how='outer')
    #replace any NA values with 0
    CBP87_merged.fillna(0, inplace=True)

    CBP87_merged_full = CBP87_data.merge(CBP87_merged, on='fipstate', how='left')
    #here I calc sm1 minus xm and sm2 minus xm
    CBP87_merged_full['sm1_minus_xm'] = CBP87_merged_full['state_ind1_share'] - CBP87_merged_full['st_share_agg_emp']
    CBP87_merged_full['sm2_minus_xm'] = CBP87_merged_full['state_ind2_share'] - CBP87_merged_full['st_share_agg_emp']

    #here I am doing final calculations to calculate gamma (Pairwise Coagglomeration)
    CBP87_merged_full['diff1_x_diff2'] = CBP87_merged_full['sm1_minus_xm'] * CBP87_merged_full['sm2_minus_xm']
    sum_of_diffs = CBP87_merged_full['diff1_x_diff2'].sum()
    sum_xm_squared = CBP87_merged_full['xm_squared'].sum()
    gamma = sum_of_diffs / (1 - sum_xm_squared)
    lst.append([combo, gamma])

#here I am saving combo and gamma results in a results dataframe which I will then save in Excel
df_results = pd.DataFrame(lst, columns=cols)

#saving the results in an Excel file
with pd.ExcelWriter(r'/Users/nicog-v/Documents/Python Mapping GPDs/Ellison Glaeser Index/Gamma_4dig_Results.xlsx') as writer:
    df_results.to_excel(writer, sheet_name='Results', index=False)

