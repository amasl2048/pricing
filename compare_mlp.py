#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
Сравнение MLP
2017 July

by amaslennikov
'''
import sys
import yaml
import pandas as pd
import time
import pprint

print("Starting...")

# Load price sources 
config_file = "./base/pricing_comp.yml"

print("\nConfig file: ", config_file)
try:
    CONF = yaml.load(open(config_file))
except: 
    print("Error: no %s file" % config_file)
    sys.exit(0)
pprint.pprint(CONF)

def is_duplicate(Series):
    pn = Series.values.tolist()
    full_len = len(pn)
    set_len = len( set(pn) )
    if full_len != set_len:
        print("mlp full len: %s, uniq len: %s, diff: %s" % (full_len, set_len, full_len - set_len))
        # http://stackoverflow.com/questions/9835762/find-and-list-duplicates-in-python-list
        import collections
        print([item for item, count in collections.Counter(pn).items() if count > 1])
        sys.exit(0)

# Get part.num. list
mlp_new = pd.ExcelFile(CONF["mlp_new"]).parse(CONF["sheet_new"])
is_duplicate(mlp_new["Part Number"])  # pandas ValueError: cannot reindex from a duplicate axis
mlp_new = mlp_new.set_index("Part Number")
PN_NEW = set(mlp_new.index)  #;print(PN_NEW)

mlp_old = pd.ExcelFile(CONF["mlp_old"]).parse(CONF["sheet_old"])
is_duplicate(mlp_old["Part Number"])
mlp_old = mlp_old.set_index("Part Number")
PN_OLD = set(mlp_old.index)  #;print(PN_OLD)

# Read prepared product groups discounts
disc = pd.ExcelFile(CONF["prod_groups"]).parse(CONF["prod_sheet"])
disc = disc.set_index("MPG")
MPG = set(disc.index)  ;print(MPG)

# Part numbers common list
PNUMS = list(PN_OLD.intersection(PN_NEW))
PNUMS.sort()
print("part.nums: %s old, %s new, %s common" % (len(PN_OLD), len(PN_NEW), len(PNUMS)))

# Create new dataframe
cols = ["MPG", "DN", "old", "new", "diff", "percent"]
mlp_diff = pd.DataFrame(index=PNUMS, columns=cols)

mlp_diff["MPG"] = mlp_new["Material Pricing Group"]
mlp_diff["DN"] = mlp_new["Designation EN"]

mlp_diff["old"] = mlp_old["MLP"]
mlp_diff["new"] = mlp_new["List price [USD]"]

mlp_diff["diff"] = mlp_diff["new"] - mlp_diff["old"]
mlp_diff["percent"] = mlp_diff["diff"] / mlp_diff["old"] * 100

# Groups
mlp_sum = mlp_diff.groupby("MPG").sum()
mlp_sum["percent"] = mlp_sum["diff"] / mlp_sum["old"] * 100

# Write xls
writer = pd.ExcelWriter("./MLP_compare_" + time.strftime("%Y%m%d") + ".xls")
mlp_diff.reset_index(inplace = True) 
mlp_diff.to_excel(writer, "Prices", index=False)

mlp_sum.reset_index(inplace = True) 
mlp_sum.to_excel(writer, "Sum", index=False)

try:
    writer.save()
except:
    print("\n!!!Error: '.xls' is busy!")

print("Done.")
input()