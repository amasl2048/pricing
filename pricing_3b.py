#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Расчёт скидок и генерация файлов buy*.xls и jde*.xls

1 аргумент: файл msrp_dot.csv
2 аргумент: файл со скидками partners2.yml

2014 Dec
'''
import sys
import yaml
from pandas import read_csv
from numpy import loadtxt

print "Starting..."

Conf = yaml.load(open("base/pricing_conf.yml"))
cross = Conf["cross"]    # кросс-курс eur/usd
kz = Conf["kz"]  # множитель KZ

file = sys.argv[1]
#file = "msrp_dot.csv"
msrp = read_csv(file, ";", header=0, index_col=False)

par_file = sys.argv[2]
#par_file = "partners2.yml"

cat1 = loadtxt(Conf["cat1.csv"], dtype=str, usecols=(0,))
def change_catalog1(df):
    if (df["partnum"] in cat1): df["partcatalog"] = Conf["new1"]
    return df
msrp = msrp.apply(change_catalog1, axis = 1)

#Product = sys.argv[2]
f = lambda x: round(x, 2)

msrp = msrp.rename(columns={"partnum": "Part Number",
                            "partlabel":"Designation EN", 
                            "partdisc": "Product Group",
                            "lmsrp_ru": "Price [EUR]"})

lmsrp_ru = msrp[["Part Number", "Designation EN", "Product Group"]]
price = msrp["Price [EUR]"] # LMSRP_RU
lmsrp_ru["Price [EUR]"] = price.map(f) # округление до 2 знаков после запятой
lmsrp_ru.to_excel("LMSRP_RU.xls", index=False)
#print lmsrp_ru

lmsrp_kz = msrp[["Part Number", "Designation EN", "Product Group"]]
price_kz = price * kz
lmsrp_kz["Price [EUR]"] = price_kz.map(f)
lmsrp_kz.to_excel("LMSRP_KZ.xls", index=False)
#print lmsrp_kz

Par = yaml.load(open(par_file))
Partners = Par.keys()

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

def minus(Series):
    '''
    buy price part = 1 - discount
    '''
    if (Series != "NA"): return 1 - Series
    
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
    rus = rus.map(f) # округление до 2 знаков после запятой
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