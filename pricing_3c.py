#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Расчёт скидок и генерация файлов buy*.xls и jde*.xls
на основе Product groups
2015 March
'''
import sys
import yaml
from pandas import read_csv, ExcelFile

print "Starting..."

config_file = "base/pricing_conf2.yml"
try:
    Conf = yaml.load(open(config_file))
except: 
    print "Error: no %s file" % conf_file
    sys.exit(0)
cross = Conf["cross"]    # ex-rate eur/usd
kz = Conf["kz"]

#msrp = read_csv("msrp_dot.csv", ";", header=0, index_col=False)
msrp = ExcelFile(Conf["msrp_ru"]).parse(Conf["sheet"])

Par = yaml.load(open(Conf["partners"]))
Partners = Par.keys()

f = lambda x: round(x, 2)

# Read prepared product groups discounts
d = ExcelFile(Conf["prod_groups"]).parse(Conf["prod_sheet"])
d = d.set_index("Product Group")
d = d["Disc. group"]

msrp = msrp.rename(columns={"partnum": "Part Number",
                            "partlabel":"Designation EN", 
                            "partdisc": "Product Group",
                            "lmsrp_ru": "Price [EUR]"})

lmsrp_ru = msrp[["Part Number", "Designation EN", "Product Group"]]
price = msrp["Price [EUR]"] # LMSRP_RU
lmsrp_ru["Price [EUR]"] = price.map(f) # round to 2 digits
lmsrp_ru.to_excel("LMSRP_RU.xls", index=False)
#print lmsrp_ru

lmsrp_kz = msrp[["Part Number", "Designation EN", "Product Group"]]
price_kz = price * kz
lmsrp_kz["Price [EUR]"] = price_kz.map(f)
lmsrp_kz.to_excel("LMSRP_KZ.xls", index=False)
#print lmsrp_kz

def minus(Series):
    """
    buy price part = 1 - discount
    """
    if (Series != "NA"): return 1 - Series

def disc_calc(df):
    '''
    New disc calc from Product Groups
    '''
    try:
        new_grp = d.loc[df["Product Group"]]
    except:
        print "Error: no '%s' product group" % df["Product Group"]
        sys.exit(0)
    return Par[Company]["discount"][new_grp]

buy = lmsrp_ru[["Part Number", "Designation EN"]]
for Company in Partners:
    print Company
    Pcode = Par[Company]["p_code"]
    Disc = msrp.apply(disc_calc, axis = 1)
    Disc = Disc.apply(minus)
    if ("USD" not in Company):
        col4 = "Price [EUR]"
    else:
        Disc = Disc * cross
        col4 = "Price [USD]"
    a = lmsrp_ru[["Part Number", "Designation EN", "Product Group"]]
    jde = lmsrp_ru[["Part Number", "Designation EN"]]
    rus = price * Disc
    rus = rus.map(f) # round to 2 digits
    a.loc[:, col4] = rus
    jde.loc[:, col4] = rus
    jde.loc[:, "JDE"] = Pcode
    jde = jde.reindex(columns = ["JDE", "Part Number", "Designation EN", col4])
    
    # save buy price xls
    fname = "Buy_Price_" + str(Company) + "_" + str(Pcode) + ".xls"
    b_price = a[ a[col4] >= 0 ] # delete empty price items
    b_price.to_excel(fname, index=False)
    
    # save jde xls
    fname2 = "jde_" + fname
    jde2 = jde[ jde[col4] >= 0 ] # delete empty price items
    jde2.to_excel(fname2, index=False)
    
    buy.loc[:, Company] = a[col4]

buy.loc[:, "LMSRP_RU"] = lmsrp_ru["Price [EUR]"]
buy.loc[:, "MSRP"] = msrp["partmsrp"]
buy.loc[:, "Ref."] = msrp["partrefp"]
buy.loc[:, "Trans."] = msrp["partxferbasep"]

# summary table for approval
#print buy
buy.to_excel("./Buy_prices.xls", index=False)

print "Ex-rate EUR/USD:", cross
raw_input()