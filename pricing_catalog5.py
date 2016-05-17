#!/usr/bin/env python
# -*- coding: utf-8 -*-
import yaml
import pandas as pd
import sys
'''
Combine Product Group, Material Category Name and Disc.
Check koef. k and calc new k_new & buy_p

Create files:
	buy_new3.xls
	lmsrp_ru3.xls
	partners_discounts.xls
	partners_discounts_uniq.xls

2016 May
'''
print "Starting..."

# load config
category_conf = yaml.load(open("./base/category_conf.yml"))
cross = category_conf["cross"]

# load partners data
Par = yaml.load(open(category_conf["partner_disc"]))
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
part["New disc"] = part["Product Group"].apply(new_disc)
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

def k_new(df):
    k_min = category_conf["k_min"]
    if (df["k"] == -1): return -1
    #if (df["New disc"] == "Service"): return 1
    if (df["Material Category Name"] == category_conf["category_1"]):
        if ( df["k"] > category_conf["max_cat1"] ): return category_conf["max_cat1"]
        elif ( df["k"] == 0): return 0
        elif ( df["k"] < k_min): return k_min
        return df["k"]
    if ( category_conf["category_2"] in df["Material Category Name"]):
        if (df["New disc"] == category_conf["product_1"]):
            if ( df["k"] > k_min ): return k_min
            elif ( df["k"] == 0): return 0
            elif ( df["k"] < k_min): return k_min
            return df["k"]
        else:
            if ( df["k"] > category_conf["max_cat2"] ): return category_conf["max_cat2"]
            elif ( df["k"] == 0): return 0
            elif ( df["k"] < k_min): return k_min
            return df["k"]
    if ( df["k"] > 1 and df["k"] < k_min): return k_min
    return df["k"]

def buy_01(df):
    if ( (df["buy_new"] == 0) and (df["partmsrp"] == 0) ): return 0.01
    if ( (df["buy_new"] == 0) and  not (df["partmsrp"] == 0) ): return df["min_buy_eur"]
    if (df["k"] == -1): return df["min_buy_eur"]
    #if (df["partrefp"] == ""): return df["min_buy_eur"]
    return df["buy_new"]
        
def new_lmsrp(df):
    if (df["New disc"] == category_conf["special_catalog"]): 
        return df["buy_new"] / (1 - category_conf["max_disc_sip"])
    else: 
        return df["buy_new"] / (1 - category_conf["max_disc"])

# adding new calculating columns
part["min_buy_eur"] = part.apply(min_price, axis = 1)#.apply(f)
part["k"] = part.apply(k_ref, axis = 1).apply(f)
part["k_new"] = part.apply(k_new, axis = 1).apply(f)
part["buy_new"] = part["partrefp"] * part["k_new"]
part["buy_new"] = part.apply(buy_01, axis = 1)
part["new LMSRP"] = part.apply(new_lmsrp, axis = 1).apply(f)
part["diff LMSRP"] = part["new LMSRP"] - part["LMSRP"]

def disc_calc(df):
    new_grp = d.loc[df["Product Group"]]
    return Par[Company]["discount"][new_grp]

def na(Series):
    if (Series != "NA"): return Series

def minus(Series):
    '''
    buy price part = 1 - discount
    '''
    if (Series != "NA"): return 1 - Series

# adding new calculating columns "Company"
P_disc = pd.DataFrame() # partners discount per product group
for Company in Partners:
    Disc = part.apply(disc_calc, axis = 1)
    Disc = Disc[ Disc >= 0 ] # delete empty price items
    Disc = Disc.apply(na) # delete "NA"
    P_disc[Company] = Disc * 100 # 100%
    if ("USD" not in Company):
        part[Company] = part["new LMSRP"] * Disc.apply(minus)
    else:
        part[Company] = part["new LMSRP"] * Disc.apply(minus) * cross
    part[Company] = part[Company].apply(f)

# calculation difference new-old
part["diff_si_buy"] = part[category_conf["gold_si"]] - part["old_si_buy"]
part["diff_dist_buy"] = part[category_conf["gold_dist"]] - part["old_dist_buy"]
#part.to_excel("./all_categories.xls", index=True)

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

# Preparing product discount per product group per partner 
b = part[["Product Group", "Material Category Name", "New disc"]]
for Company in Partners:
    b[Company] = P_disc[Company]

# Change index
part.reset_index(inplace = True)

# Whole sorted output from part DataFrame
#a = part.set_index("Product Group")
a = part.set_index("partnum")
a.sort_index(inplace=True)
try: 
    a.to_excel("./buy_new3.xls", index=True)
except:
    print "\nError: 'buy_new3.xls' is busy..."

# New lmsrp_ru
lmsrp_ru = part[["partnum", "partlabel", "partmsrp", "partrefp", "partxferbasep", "Product Group"]]
lmsrp_ru = lmsrp_ru.rename(columns={"Product Group": "partdisc"})
lmsrp_ru["lmsrp_ru"] = part["new LMSRP"]
lmsrp_ru.reset_index(inplace = True)
lmsrp_ru.to_excel("./lmsrp_ru3.xls", index=False)

# Unique product groups
b.set_index("Product Group", inplace = True)
b.sort_index(inplace=True)
b.reset_index(inplace = True)
#print b.head()
b = b.drop_duplicates()
b.to_excel("./partners_discounts.xls", index=False)
del b["Material Category Name"] # only unique remain
b = b.drop_duplicates()
b.to_excel("./partners_discounts_uniq.xls", index=False)

print "Done."
#raw_input()
