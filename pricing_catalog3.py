#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Combine Product Group, Material Category Name and Disc.
Check koef. k and calc new k_new & buy_p

2015 March
'''
import yaml
import pandas as pd
#from numpy import unique

print "Starting..."

Conf = yaml.load(open("./base/pricing_conf2.yml"))

par_file = "./base/partners4.yml"
Par = yaml.load(open(par_file))
Partners = Par.keys()

category_conf = yaml.load(open("./base/category_conf.yml"))

dist = pd.ExcelFile(category_conf["dist_file"]).parse(category_conf["export"], index_col = "Part Number")

si = pd.ExcelFile(category_conf["si_file"]).parse(category_conf["export"], index_col = "Part Number")

groups = pd.ExcelFile(category_conf["group_file"]).parse(category_conf["group_sheet"])

c = groups[["Product Group", "Material Category Name", "Disc. group"]]
c = c.drop_duplicates()
#c.to_excel("./groups_unique3.xls", index=False)

def mba(Series):
    s = u"MBA"
    if (Series in category_conf["new_cat"]): return s
    return Series

# Prepare dependence discount group from product group
'''
d = groups[["Product Group", "Disc. group"]]
d["Disc. group"] = groups["Disc. group"].apply(mba)
d = d.drop_duplicates()
#d.to_excel("./groups_unique4.xls", index=False)
d = d.set_index("Product Group")
d = d["Disc. group"]
'''
# Read prepared product groups discounts
d = pd.ExcelFile("./base/groups_unique4.xls").parse("Sheet1")
d = d.set_index("Product Group")
d = d["Disc. group"]

part = pd.ExcelFile(category_conf["part_file"]).parse(category_conf["part_sheet"])
part = part.rename(columns={"partdisc": "Product Group"})

# Category file
cat_file = category_conf["cat_file"]
xl2 = pd.ExcelFile(cat_file)
category = xl2.parse(category_conf["export"])#, index_col = "Part Number")
category = category.rename(columns={"Part Number": "partnum"})
category = category.set_index("partnum")
#print category.head()

f = lambda x: round(x, 2)

def change_u(Series):
    return unicode(Series)

part["partnum"] = part["partnum"].apply(change_u)

def new_disc(Series):
    return d.loc[Series]

part["New disc"] = part["Product Group"].apply(new_disc)
part = part.set_index("partnum")
part["old_dist_buy"] = dist["Price [USD]"]
part["old_si_buy"] = si["Price [EUR]"]
part["Material Category Name"] = category["Material Category Name"]
part["LMSRP"] = category["Price [EUR]"]

def min_price(df):
    if (df["old_dist_buy"]/Conf["cross"] < df["old_si_buy"] ): return df["old_dist_buy"]/Conf["cross"]
    else: return df["old_si_buy"]

def k_ref(df):
    if (df["partrefp"] > 0): 
        k = df["min_buy_eur"]/df["partrefp"]
        #if (k < 1.099): return k
        return k
    elif (df["partrefp"] == 0):
        return 0

def new_lmsrp(df):
    dist_eur = df["old_dist_buy"]/Conf["cross"]
    if ( dist_eur < df["old_si_buy"] ): min = dist_eur
    else: min = df["old_si_buy"]
    if (df["New disc"] == category_conf["special_catalog"]): return min / (1 - category_conf["max_disc_sip"])
    else: return min / (1 - category_conf["max_disc"])

def k_new(df):
    if (df["Material Category Name"] == category_conf["category_1"]):
        if ( df["k"] > 2 ): return 2
        elif ( df["k"] == 0): return 0
        elif ( df["k"] < 1.1): return 1.1
        return df["k"]
    if ( category_conf["category_2"] in df["Material Category Name"]):
        if (df["New disc"] == category_conf["product_1"]):
            if ( df["k"] > 1.1 ): return 1.1
            elif ( df["k"] == 0): return 0
            elif ( df["k"] < 1.1): return 1.1
            return df["k"]
        else:
            if ( df["k"] > 1.24 ): return 1.24
            elif ( df["k"] == 0): return 0
            elif ( df["k"] < 1.1): return 1.1
            return df["k"]
    return df["k"]

part["min_buy_eur"] = part.apply(min_price, axis = 1)#.apply(f)
part["k"] = part.apply(k_ref, axis = 1).apply(f)
part["k_new"] = part.apply(k_new, axis = 1).apply(f)
part["buy_new"] = part["partrefp"] * part["k_new"]
part["new LMSRP"] = part.apply(new_lmsrp, axis = 1)#.apply(f)
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

for Company in Partners:
    Disc = part.apply(disc_calc, axis = 1)
    Disc = Disc[ Disc >= 0 ] # delete empty price items
    Disc = Disc.apply(na) # delete "NA"
    #if (Company == "Distr.(Bronze) EUR"): discounts = Disc * 100 # 100%
    if ("USD" not in Company):
        part[Company] = part["new LMSRP"] * Disc.apply(minus)
    else:
        part[Company] = part["new LMSRP"] * Disc.apply(minus) * Conf["cross"]
    part[Company] = part[Company].apply(f)


part["diff_si_buy"] = part[category_conf["gold_si"]] - part["old_si_buy"]
part["diff_dist_buy"] = part[category_conf["gold_dist"]] - part["old_dist_buy"]

#part.to_excel("./all_categories.xls", index=True)

mcn = part["Material Category Name"].unique()

for each in mcn:
    print each
print mcn.size


pgroup = part["Product Group"].unique()
pgroup.sort()
'''
for row in pgroup:
    print row
'''
print "Items: ", part.index.size
print "Product groups: ", pgroup.size

# Change index
b = part[["Product Group", "Material Category Name"]]
#b["discounts"] = discounts
print b.head()
part.reset_index(inplace = True)

#a = part.set_index("Product Group")
a = part.set_index("partnum")
a.sort_index(inplace=True)
a.to_excel("./groups7.xls", index=True)


# Unique product groups
b.set_index("Product Group", inplace = True)
b.sort_index(inplace=True)
b.reset_index(inplace = True)
print b.head()
b = b.drop_duplicates()
#b.to_excel("./bronze_dist.xls", index=False)

print "Done."
#raw_input()
