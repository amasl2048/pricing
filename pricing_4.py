#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Расчёт скидок и генерация файлов buy*.xls и jde*.xls
на основе Product groups и "lmsrp_ru" column
2016 May
'''
import sys
import yaml
#from pandas import ExelFile, ExcelWriter
import pandas as pd
import time
import pprint

print "Starting..."
log_file = open("pricing.log", "w")
log = []

config_file = "./base/pricing_conf2.yml"
print "\nConfig file: ", config_file
try:
    Conf = yaml.load(open(config_file))
except: 
    print "Error: no %s file" % config_file
    sys.exit(0)
pprint.pprint(Conf)
cross = Conf["cross"]   # ex-rate eur/usd

# choice lmsrp source
if Conf["exported"] == True:
    print "\nExported price list!\n"
    msrp = pd.ExcelFile(Conf["exported_msrp_ru"]).parse(Conf["exported_sheet"])
else:
    msrp = pd.ExcelFile(Conf["msrp_ru"]).parse(Conf["sheet"])

def change_u(Series):
    return unicode(Series)
msrp["partnum"] = msrp["partnum"].apply(change_u)

def is_duplicate(Series):
    pn = Series.values.tolist()
    full_len = len(pn)
    set_len = len( set(pn) )
    if full_len != set_len:
        print "msrp full len: %s, uniq len: %s, diff: %s" % (full_len, set_len, full_len - set_len)
        # http://stackoverflow.com/questions/9835762/find-and-list-duplicates-in-python-list
        import collections
        print [item for item, count in collections.Counter(pn).items() if count > 1]
        sys.exit(0)
is_duplicate(msrp["partnum"])

Par = yaml.load(open(Conf["partners"]))
#print "Parthner's file: ", Conf["partners"]
Partners = Par.keys()

f = lambda x: round(x, 2)

# Read prepared product groups discounts
d = pd.ExcelFile(Conf["prod_groups"]).parse(Conf["prod_sheet"])
d = d.set_index("Product Group")
d = d["Disc. group"]

if not "lmsrp_ru" in msrp.columns.values.tolist():
    print "Error: no 'lmsrp_ru' column price"
    sys.exit(0)
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
    prices[count] = msrp["Price [EUR]"] * Conf["countries"][count] # before round
    lmsrp[count]["Price [EUR]"] = prices[count].map(f) # round to 2 digits
    lmsrp[count].to_excel(count + "_EUR_LMSRP_" + time.strftime("%Y%m%d") + ".xls", index=False)
lmsrp = pd.Panel(lmsrp)

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
    '''If the price below Ref.?'''
    global log
    report = []
    if (Par[Company]["cur"] == "EUR"): 
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

def p_price(Company):
    Pcode = Par[Company]["p_code"]
    Disc = msrp.apply(disc_calc, axis = 1)
    Disc = Disc.apply(minus)
    cur  = Par[Company]["cur"]
    country = Par[Company]["country"]
    if (cur == "EUR"):
        col4 = "Price [EUR]"
    else:
        Disc = Disc * cross
        col4 = "Price [USD]"
    lmsrp_ru = lmsrp["Russia"][["Part Number", "Designation EN", "Product Group"]]
    jde = lmsrp["Russia"][["Part Number", "Designation EN"]]
    pr = prices[country] * Disc # multiply before round -> price depends on country
    pr = pr.map(f) # round to 2 digits
    lmsrp_ru.loc[:, col4] = pr
    jde.loc[:, col4] = pr
    jde.loc[:, "JDE"] = Pcode
    jde = jde.reindex(columns = ["JDE", "Part Number", "Designation EN", col4])
    
    # save buy price xls
    fname = str(Par[Company]["country"]) + "_" + cur + "_" + str(Company)+\
    "(" + str(Par[Company]["internal"]) + ")" + "_" +\
    str(Pcode) + "_" + time.strftime("%Y%m%d") + ".xls"
    b_price = lmsrp_ru[ lmsrp_ru[col4] >= 0 ] # delete empty price items
    b_price.to_excel(fname, index=False)
    
    # save jde xls
    fname2 = "jde_" + fname
    jde2 = jde[ jde[col4] >= 0 ] # delete empty price items
    jde2.to_excel(fname2, index=False)
    
    buy.loc[:, Company] = lmsrp_ru[col4]
    buy.apply(check_buy, axis = 1) # - check after round to 2 digits
    
    # calc difference with previous buy price
    last_file = Conf["price_dir"] + str(Company) + ".xls"
    try:
        last_price = pd.ExcelFile(last_file).parse(Conf["price_sheet"])
    except: 
        print "Error: %s price sheet name" % Conf["price_sheet"]
        sys.exit(0)
    last_price.set_index("Part Number", inplace = True)
    diff_price = buy[["Part Number", "Designation EN", Company]]
    diff_price.set_index("Part Number", inplace = True)
    diff_price.loc[:, "last"] = last_price[col4]
    
    lmsrp_ru.set_index("Part Number", inplace = True)
    diff_price.loc[:, "diff"] = lmsrp_ru[col4] - last_price[col4]
    diff_file = "./diff_" + str(Company) + "_" + cur + "_" + time.strftime("%Y%m%d") + ".xls"
    diff_price.reset_index(inplace = True)
    diff_price.to_excel(diff_file, index=False)

for Company in Partners:
    print Company
    p_price(Company)
    
buy.loc[:, "LMSRP_RU"] = lmsrp["Russia"]["Price [EUR]"]
buy.loc[:, "LMSRP_KZ"] = lmsrp["Kazakhstan"]["Price [EUR]"]
buy.loc[:, "LMSRP_UA"] = lmsrp["Ukraine"]["Price [EUR]"]

#print buy
writer = pd.ExcelWriter("./Prices_EUR_USD_" + time.strftime("%Y%m%d") + ".xls")
buy.to_excel(writer, "Prices", index=False)
buy_disc = lmsrp["Russia"][["Part Number", "Designation EN", "Product Group"]] # check discounts
buy_koef = lmsrp["Russia"][["Part Number", "Designation EN", "Product Group"]] # check k
for Company in Partners:
    cur  = Par[Company]["cur"]
    country = Par[Company]["country"]
    if (cur == "EUR"):
        buy_disc[Company] = (1 - buy[Company] / prices[country]).map(f) #   depends on country 
        buy_koef[Company] =  (buy[Company] / buy["Ref."]).map(f) 
    else:
        buy_disc[Company] = (1 - buy[Company] / prices[country] / cross).map(f) # depends on country
        buy_koef[Company] =  (buy[Company] / cross / buy["Ref."]).map(f) 

buy_disc.to_excel(writer, "Discounts", index=False)
buy_koef.to_excel(writer, "Coefficients", index=False)
try:
    writer.save()
except:
    print "\n!!!Error: '.xls' is busy!"

# LMSRP difference
for count in countries:
    lmsrp_file = Conf["price_dir"] + count + "_LMSRP" + ".xls"
    last_lmsrp = pd.ExcelFile(lmsrp_file).parse(Conf["price_sheet"])
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