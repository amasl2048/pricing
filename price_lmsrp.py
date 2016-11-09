#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Объединение цен из прайс-листов 
2016 July
'''
import sys
import yaml
from pandas import ExcelFile, ExcelWriter, isnull
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

msrp = ExcelFile(Conf["msrp_ru"]).parse(Conf["sheet"])
def change_u(Series):
    ''' Elsewhere some items from MS Excel are int or str '''
    return unicode(Series)
msrp["partnum"] = msrp["partnum"].apply(change_u)

Par = yaml.load(open(Conf["partners"]))
print "Parthner's file: ", Conf["partners"]
Partners = Par.keys()

f = lambda x: round(x, 2)

# Read prepared product groups discounts
d = ExcelFile(Conf["prod_groups"]).parse(Conf["prod_sheet"])
print "Product's groups: ", Conf["prod_groups"], "\n"
d = d.set_index("Product Group")
d = d["Disc. group"]

# LMSRP
countries = Conf["countries"]
lmsrp = {}
for count in countries:
    lmsrp[count] = msrp[["partnum", "partlabel"]]
    lmsrp[count].set_index("partnum", inplace=True)
    lmsrp_price = ExcelFile(Conf["price_dir"] + count + "_LMSRP.xls").parse(Conf["price_sheet"])
    lmsrp_price.set_index("Part Number", inplace=True)
    lmsrp[count].loc[:,"Product Group"] = lmsrp_price["Product Group"]
    lmsrp[count].loc[:,"MCN"] = lmsrp_price["Material Category Name"]
    lmsrp[count].loc[:,"Price [EUR]"] = lmsrp_price["Price [EUR]"]
    #lmsrp[count].to_excel(count + "_EUR_LMSRP_" + time.strftime("%Y%m%d") + ".xls", index=True)

buy = msrp[["partnum", "partlabel", "partmsrp", "partrefp", "partxferbasep", "partdisc"]] # summary table for approval

buy.set_index("partnum", inplace=True)

buy.loc[:,"Product Group"] = lmsrp["Russia"]["Product Group"]
buy.loc[:,"Material Category Name"] = lmsrp["Russia"]["MCN"]

def new_disc(Series):
    '''
    get new discount group from Product Group: "RC" -> "MBA"
    In : d.loc[part["Product Group"][0]]
    Out: u'MBA'
    '''
    if isnull(buy["Product Group"]).any(): # should be no empty items!
        print "Error: empty product group..."
        sys.exit(0)
    return d.loc[Series]

def check_buy(df):
    '''If the price below Ref.?'''
    global log
    report = []
    if (Par[Company]["cur"] == "EUR"): 
        if (df[Company] < df["partrefp"]):
            report = [Company, df.name, df["partlabel"], str(df["partrefp"]), str(df[Company])]
            log.append(" ".join(report))
    else:
        if (df[Company]/cross < df["partrefp"]): 
            report = [Company, df.name, df["partlabel"], str(df["partrefp"]), str(df[Company])]
            log.append(" ".join(report))

def check_disc(df):
    global log
    report = []
    if (df["LMSRP_RU"] == 0.02):
        return
    if not (df[Company] > 0):
        return
    par_disc = Par[Company]["discount"][df["New disc"]]
    if (par_disc == "NA"):
        return
    if (df[Company] < par_disc*100 - 0.5) or (df[Company] > par_disc*100 + 0.5):
        report = [df.name, df["partlabel"], str(par_disc*100), str(df[Company]), df["New disc"]]
        log.append(" ".join(report))
    return

# fill in partners prices
for Company in Partners:
    print Company
    cur  = Par[Company]["cur"]
    if (cur == "EUR"):
        col4 = "Price [EUR]"
    else:
        col4 = "Price [USD]"
     
    par_price = ExcelFile(Conf["price_dir"] + Company + ".xls").parse(Conf["price_sheet"])
    par_price.set_index("Part Number", inplace=True)
    buy.loc[:, Company] = par_price[col4]
    buy.apply(check_buy, axis = 1) # - check after round to 2 digits

buy.loc[:, "lmsrp_ru"] = lmsrp["Russia"]["Price [EUR]"]
buy.loc[:, "lmsrp_kz"] = lmsrp["Kazakhstan"]["Price [EUR]"]
buy.loc[:, "lmsrp_ua"] = lmsrp["Ukraine"]["Price [EUR]"]

#print buy
writer = ExcelWriter("./all_prices_EUR_USD_" + time.strftime("%Y%m%d") + ".xls")
buy.reset_index(inplace=True)
buy.to_excel(writer, "Prices", index=False)

try:
    writer.save()
except:
    print "\n!!!Error: '.xls' is busy!"

# Log price issues
if log:
    print "\nCheck buy price: \n"
    for item in log:
        print unicode(item)
        log_file.writelines(str(item))
        log_file.write("\n")
log_file.close()

print "\nEx-rate EUR/USD:", cross
print "Done."
#raw_input()