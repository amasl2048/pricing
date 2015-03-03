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
from numpy import loadtxt, unique
#import csv

print "Starting..."

Conf = yaml.load(open("./base/pricing_conf.yml"))

par_file = "./base/partners3.yml"
Par = yaml.load(open(par_file))
Partners = Par.keys()

category_conf = yaml.load(open("./base/category_conf.yml"))

group_file = category_conf["group_file"]
groups = pd.ExcelFile(group_file).parse(category_conf["group_sheet"])

c = groups[["Product Group", "Material Category Name", "Disc. group"]]
c = c.drop_duplicates()
#c.to_excel("./groups_unique3.xls", index=False)

def mba(Series):
    s = "MBA"
    if (Series in category_conf["new_cat"]): return s
    return Series

d = groups[["Product Group", "Disc. group"]]
d["Disc. group"] = groups["Disc. group"].apply(mba)
d = d.drop_duplicates()
d.to_excel("./groups_unique4.xls", index=False)


pfile = category_conf["part_file"]
xl1 = pd.ExcelFile(pfile)
#print xl1.sheet_names
part = xl1.parse(category_conf["part_sheet"])#, index_col = "partnum")

cat_file = category_conf["cat_file"]
xl2 = pd.ExcelFile(cat_file)
#print xl2.sheet_names
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
part = part.rename(columns={"partdisc": "Product Group"})
part["Material Category Name"] = category["Material Category Name"]
part["LMSRP"] = category["Price [EUR]"]

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

def part2group(Series):
    '''
    Get discont name from partcatalog
    '''
    if (Series in Par[Company]['discount'].keys() ): return Series
    if (Series in Conf['catalog'].keys() ): return Conf['catalog'][Series]
    else:
        print "ERROR name!", Series
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
part["Disc. group"] = part["partcatalog"].apply(part2group)

mcn = part["Material Category Name"].unique()
'''
for each in mcn:
    print each
print mcn.size
'''

pgroup = part["Product Group"].unique()
pgroup.sort()


'''
for row in pgroup:
    print row
print pgroup.size
'''

#part.to_excel("./all_categories.xls", index=True)

# change index
a = part.set_index("Product Group")
a.sort_index(inplace=True)
#a.to_excel("./groups5.xls", index=True)

b = a[["partcatalog","Material Category Name"]]
b = b.reset_index()
b = b.drop_duplicates()
#b.to_excel("./groups_unique2.xls", index=False)

print "Done."
#raw_input()