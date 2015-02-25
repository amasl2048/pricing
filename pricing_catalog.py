#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Combine Product Group, Material Category Name and Disc.

2015 Feb
'''
import sys
import yaml
import pandas as pd
#from pandas import read_csv
from numpy import loadtxt

print "Starting..."

Conf = yaml.load(open("./base/pricing_conf.yml"))

par_file = "./base/partners3.yml"
Par = yaml.load(open(par_file))
Partners = Par.keys()

category_conf = yaml.load(open("./base/category_conf.yml"))

pfile = category_conf["part_file"]
xl1 = pd.ExcelFile(pfile)
print xl1.sheet_names
part = xl1.parse(category_conf["part_sheet"])#, index_col = "partnum")

cat_file = category_conf["cat_file"]
xl2 = pd.ExcelFile(cat_file)
print xl2.sheet_names
category = xl2.parse(category_conf["cat_sheet"])#, index_col = "Part Number")
category = category.rename(columns={"Part Number": "partnum"})
category = category.set_index("partnum")
#print category.head()

# Changing partcatalog name
cat1 = loadtxt(Conf["cat1.csv"], dtype=str, usecols=(0,))
def change_catalog1(df):
    if (df["partnum"] in cat1): df["partcatalog"] = Conf["new1"]
    return df

part = part.apply(change_catalog1, axis = 1)
part = part.set_index("partnum")

#part["PG"] = category["Product Group"]
part["Material Category Name"] = category["Material Category Name"]

def catalog(Company, cat):
    '''
    Get discount value from partcatalog
    '''
    if (cat in Par[Company]['discount'].keys() ): return cat
    if (cat in Conf['catalog'].keys() ): return Conf['catalog'][cat]
    else:
        print "ERROR name!", cat, Company
        sys.exit(0)
    return
    
def disc_calc(Series):
    return Par[Company]["discount"][catalog(Company,Series["partcatalog"])]

def na(Series):
    if (Series != "NA"): return Series

for Company in Partners:
    Disc = part.apply(disc_calc, axis = 1)
    Disc = Disc[ Disc >= 0 ] # delete empty price items
    Disc = Disc.apply(na) # delete "NA"
    part[Company] = Disc * 100

#print part.head()

mcn = part["Material Category Name"].unique()
for each in mcn:
    print each

print mcn.size

pgroup = part["partdisc"].unique()
print pgroup.size

part.to_excel("./all_categories.xls", index=True)
