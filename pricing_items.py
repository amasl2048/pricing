#!/usr/bin/env python
# -*- coding: utf-8 -*-
import yaml
import pandas as pd
import sys
'''
Combine Product Group, Material Category Name and Disc.
Check non-discount items
Check k

2015 Dec
'''
print "Starting..."

# load config
category_conf = yaml.load(open("./base/items_conf.yml")) # new conf
cross = category_conf["cross"]

# load partners data
Par = yaml.load(open("./base/partners.yml"))
Partners = Par.keys()

# load prices
dist = pd.ExcelFile(category_conf["dist_file"]).parse(category_conf["export"], index_col = "Part Number")
si = pd.ExcelFile(category_conf["si_file"]).parse(category_conf["export"], index_col = "Part Number")

# Read prepared discounts (Disc. group) per product group
d = pd.ExcelFile(category_conf["prod_groups"]).parse("Sheet1") # groups_unique
d = d.set_index("Product Group")
d = d["Disc. group"]

# Read all part.numbers, product groups, partrefp
# 'part_file' need to be prepared from items.csv filtered from actual LMSRP partnum
part = pd.ExcelFile(category_conf["part_file"]).parse(category_conf["part_sheet"])
part = part.rename(columns={"partdisc": "Product Group"})

# Read category file (extract from LMSRP price)
category = pd.ExcelFile(category_conf["cat_file"]).parse(category_conf["export"])#, index_col = "Part Number")
category = category.rename(columns={"Part Number": "partnum"})
category = category.set_index("partnum")
#print category.head()

f = lambda x: round(x, 2)

def change_u(Series):
    return unicode(Series)

part["partnum"] = part["partnum"].apply(change_u)

def new_disc(Series):
    '''
    get new discount group from Product Group: "RC" -> "MBA"
    In : d.loc[part["Product Group"][0]]
    Out: u'MBA'
    '''
    if pd.isnull(part["Product Group"]).any(): # should be no empty items!
        print "Error: empty product group..."
        sys.exit(0)
    return d.loc[Series]

# adding new calculating columns
###part["New disc"] = part["Product Group"].apply(new_disc)
part = part.set_index("partnum")
part["old_dist_buy"] = dist["Price [USD]"]
part["old_si_buy"] = si["Price [EUR]"]
part["Material Category Name"] = category["Material Category Name"]
part["LMSRP"] = category["Price [EUR]"]

def min_price(df):
    if (df["old_dist_buy"]/cross < df["old_si_buy"] ): return df["old_dist_buy"]/cross
    #elif: (df["old_si_buy"] == ""): return 
    return df["old_si_buy"]

def k_ref(df):
    if (df["partrefp"] > 0): 
        k = df["min_buy_eur"]/df["partrefp"]
        #if (k < 1.099): return k
        return k
    elif (df["partrefp"] == 0):
        return 0
    return -1

# adding new calculating columns
f = lambda x: round(x, 2)
part["min_buy_eur"] = part.apply(min_price, axis = 1).apply(f)
part["k"] = part.apply(k_ref, axis = 1).apply(f)
part["di/si"] = (part["old_dist_buy"] / cross / part["old_si_buy"]).apply(f)

# adding new calculating columns "Company"
P_disc = pd.DataFrame() # partners discount per product group

mcn = part["Material Category Name"].unique()
for each in mcn:
    print each
print "\nMaterial Category Names: ", mcn.size

pgroup = part["Product Group"].unique()
pgroup.sort()
'''
for row in pgroup:
    print row
'''
print "Items: ", part.index.size
print "Product groups: ", pgroup.size

# Change index
part.reset_index(inplace = True)
# Whole sorted output from part DataFrame
#a = part.set_index("Product Group")

part = part.rename(columns={"old_dist_buy": category_conf["gold_dist"] + ", " + Par[category_conf["gold_dist"]]["cur"]})
part = part.rename(columns={"old_si_buy": category_conf["gold_si"] + ", " + Par[category_conf["gold_si"]]["cur"]})

a = part.set_index("partnum")
a.sort_index(inplace=True)
try: 
    a.to_excel("./buy_new.xls", index=True)
except:
    print "\nError: 'buy_new.xls' is busy..."

print "Done."
#raw_input()
