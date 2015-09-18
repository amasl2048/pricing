#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Расчёт скидок и генерация файлов buy*.xls и jde*.xls
на основе Product groups
2015 Sept
'''
import sys
import yaml
from pandas import read_csv, ExcelFile
import time

print "Starting..."
log_file = open("pricing.log", "w")
log = []

config_file = "base/pricing_conf2.yml"
try:
    Conf = yaml.load(open(config_file))
except: 
    print "Error: no %s file" % conf_file
    sys.exit(0)
cross = Conf["cross"]    # ex-rate eur/usd

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


# LMSRP
countries = Conf["countries"]
lmsrp = {}
prices = {}
for count in countries:
    lmsrp[count] = msrp[["Part Number", "Designation EN", "Product Group", ]]
    prices[count] = msrp["Price [EUR]"] * Conf["countries"][count]
    lmsrp[count]["Price [EUR]"] = prices[count].map(f) # round to 2 digits
    lmsrp[count].to_excel(count + "_EUR_LMSRP_" + time.strftime("%Y%m%d") + ".xls", index=False)
    #import ipdb;ipdb.set_trace()

def disc_calc(df):
    '''
    New disc calc from Product Groups
    '''
    try:
        new_grp = d.loc[df["Product Group"]]
    except:
        print "\n!!!Error: no '%s' product group" % df["Product Group"]
        sys.exit(0)
    return Par[Company]["discount"][new_grp]

def minus(Series):
    """
    buy price part = 1 - discount
    """
    if (Series != "NA"): return 1 - Series
    
def check_buy(df):
    global log
    report = []
    if ("USD" not in Company):
        if (df[Company] < df["Ref."]):
            report = [Company, df["Part Number"], df["Designation EN"], str(df["Ref."]), str(df[Company])]
            log.append(" ".join(report))
    else:
        if (df[Company]/cross < df["Ref."]): 
            report = [Company, df["Part Number"], df["Designation EN"], str(df["Ref."]), str(df[Company])]
            log.append(" ".join(report))


buy = lmsrp["Russia"][["Part Number", "Designation EN"]] # summary table for approval
buy.loc[:, "MSRP"] = msrp["partmsrp"]
buy.loc[:, "Ref."] = msrp["partrefp"]
buy.loc[:, "Trans."] = msrp["partxferbasep"]
for Company in Partners:
    print Company
    Pcode = Par[Company]["p_code"]
    Disc = msrp.apply(disc_calc, axis = 1)
    Disc = Disc.apply(minus)
    cur  = Par[Company]["cur"]
    if (cur == "EUR"):
        col4 = "Price [EUR]"
    else:
        Disc = Disc * cross
        col4 = "Price [USD]"
    a = lmsrp["Russia"][["Part Number", "Designation EN", "Product Group"]]
    jde = lmsrp["Russia"][["Part Number", "Designation EN"]]
    rus = prices["Russia"] * Disc # multiply before round 
    rus = rus.map(f) # round to 2 digits
    a.loc[:, col4] = rus
    jde.loc[:, col4] = rus
    jde.loc[:, "JDE"] = Pcode
    jde = jde.reindex(columns = ["JDE", "Part Number", "Designation EN", col4])
    
    # save buy price xls
    fname = str(Par[Company]["country"]) + "_" + cur + "_" + str(Company) + "_" + str(Pcode) + "_" + time.strftime("%Y%m%d") + ".xls"
    b_price = a[ a[col4] >= 0 ] # delete empty price items
    b_price.to_excel(fname, index=False)
    
    # save jde xls
    fname2 = "jde_" + fname
    jde2 = jde[ jde[col4] >= 0 ] # delete empty price items
    jde2.to_excel(fname2, index=False)
    
    buy.loc[:, Company] = a[col4]
    buy.apply(check_buy, axis = 1) # - check after round to 2 digits
    
    # calc difference with previous buy price
    last_file = Conf["price_dir"] + str(Par[Company]["country"]) + "_" + cur + "_" + str(Company) + "_" + str(Pcode) + "_" + Conf["last_date"] + ".xls"
    last_price = ExcelFile(last_file).parse(Conf["price_sheet"])
    last_price.set_index("Part Number", inplace = True)
    diff_price = buy[["Part Number", "Designation EN", str(Company)]]
    diff_price.set_index("Part Number", inplace = True)
    diff_price.loc[:, "last"] = last_price[col4]
    
    a.set_index("Part Number", inplace = True)
    diff_price.loc[:, "diff"] = a[col4] - last_price[col4]
    diff_file = "./diff_" + str(Company) + "_" + cur + "_" + time.strftime("%Y%m%d") + ".xls"
    diff_price.reset_index(inplace = True)
    diff_price.to_excel(diff_file, index=False)
    
buy.loc[:, "LMSRP_RU"] = lmsrp["Russia"]["Price [EUR]"]
#print buy
try:
    buy.to_excel("./Russia_EUR_USD_" + time.strftime("%Y%m%d") + ".xls", index=False)
except:
    print "\n!!!Error: '.xls' is busy!"

# LMSRP difference
for count in countries:
    lmsrp_file = Conf["price_dir"] + count + "_EUR_LMSRP_" + Conf["last_date"] + ".xls"
    last_lmsrp = ExcelFile(lmsrp_file).parse(Conf["price_sheet"])
    last_lmsrp.set_index("Part Number", inplace = True)
    diff_lmsrp = lmsrp[count][["Part Number", "Designation EN", "Price [EUR]"]]
    diff_lmsrp.set_index("Part Number", inplace = True)
    diff_lmsrp.loc[:, "last"] = last_lmsrp["Price [EUR]"]

    lmsrp[count].set_index("Part Number", inplace = True)
    diff_lmsrp.loc[:, "diff"] = lmsrp[count]["Price [EUR]"] - last_lmsrp["Price [EUR]"]
    diff_lmsrp_file = "./diff_" + count + "_EUR_LMSRP_" + time.strftime("%Y%m%d") + ".xls"
    diff_lmsrp.reset_index(inplace = True)
    diff_lmsrp.to_excel(diff_lmsrp_file, index=False)

# Log price issues
if log:
    print "\nCheck buy price: \n"
    for item in log:
        print str(item)
        log_file.writelines(str(item))
        log_file.write("\n")
log_file.close()

print "\nEx-rate EUR/USD:", cross
print "Done."
raw_input()