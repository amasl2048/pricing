#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Проверка коэф. и скидок партнерских прайс-листов 
2016 Sept
'''
import sys
import yaml
#from pandas import ExcelFile, ExcelWriter, isnull
import pandas as pd
import time
import pprint

print "Starting..."
log_file = open("pricing.log", "w")
log = []

config_file = "./base/pricing_items.yml"
print "\nConfig file: ", config_file
try:
    Conf = yaml.load(open(config_file))
except: 
    print "Error: no %s file" % config_file
    sys.exit(0)
pprint.pprint(Conf)
cross = Conf["cross"]   # ex-rate eur/usd

list_price = pd.ExcelFile(Conf["price_items"]).parse(Conf["items_sheet"])
mlp = pd.ExcelFile(Conf["mlp_ru"]).parse(Conf["sheet"])

mgroups = pd.ExcelFile(Conf["groups"]).parse(Conf["grp_sheet"])
mgroups.set_index("MPG", inplace = True)

def change_u(Series):
    ''' Elsewhere some items from MS Excel are int or str '''
    return unicode(Series)

mlp["partnum"] = mlp["partnum"].apply(change_u)
list_price["partnum"] = list_price["partnum"].apply(change_u)

Par = yaml.load(open(Conf["partners"]))
print "Parthner's file: ", Conf["partners"]
Partners = Par.keys()

f = lambda x: round(x, 2)

# Read prepared product groups discounts
d = pd.ExcelFile(Conf["prod_groups"]).parse(Conf["prod_sheet"])
print "Product's groups: ", Conf["prod_groups"], "\n"
d = d.set_index("Product Group")
d = d["Disc. group"]

# LMSRP
countries = Conf["countries"]
lmsrp = {}
for count in countries:
    lmsrp[count] = mlp[["partnum", "partlabel"]]
    lmsrp[count].set_index("partnum", inplace=True)
    lmsrp_price = pd.ExcelFile(Conf["price_dir"] + count + "_LMSRP.xls").parse(Conf["price_sheet"])
    lmsrp_price.set_index("Part Number", inplace=True)
    lmsrp[count].loc[:,"Product Group"] = lmsrp_price["Product Group"]
    lmsrp[count].loc[:,"MCN"] = lmsrp_price["Material Category Name"]
    lmsrp[count].loc[:,"Price [EUR]"] = lmsrp_price["Price [EUR]"]
    #lmsrp[count].to_excel(count + "_EUR_LMSRP_" + time.strftime("%Y%m%d") + ".xls", index=True)

# Buy prices
buy = mlp[["partnum", "partlabel", "partmsrp", "partrefp", "partxferbasep", "partdisc", "partsapmlp", "partsapmpg"]] # summary table for approval
buy.rename(columns={"partmsrp": "MSRP",
                    "partrefp": "Ref.",
                    "partxferbasep": "Trans.",
                    "partsapmlp": "MLP",
                    "partsapmpg": "MPG"
                    }, inplace=True)
                   
buy.set_index("partnum", inplace=True)
list_price.set_index("partnum", inplace=True)

buy.loc[:,"Product Group"] = lmsrp["Russia"]["Product Group"]
buy.loc[:,"Material Category Name"] = lmsrp["Russia"]["MCN"]

def new_disc(Series):
    '''
    get new discount group from Product Group: "RC" -> "MBA"
    In : d.loc[part["Product Group"][0]]
    Out: u'MBA'
    '''
    if pd.isnull(buy["Product Group"]).any(): # should be no empty items!
        print "Error: empty product group..." # will be NaN if item is new and not in exported LMSRP
        print buy[pd.isnull(buy["Product Group"])] 
        sys.exit(0)
    return d.loc[Series]

def check_buy(df):
    '''If the price below Ref.?'''
    global log
    report = []
    if (Par[Company]["cur"] == "EUR"): 
        if (df[Company] < df["Ref."]):
            report = [Company, df.name, df["partlabel"], str(df["Ref."]), str(df[Company])]
            log.append(" ".join(report))
    else:
        if (df[Company]/cross < df["Ref."]): 
            report = [Company, df.name, df["partlabel"], str(df["Ref."]), str(df[Company])]
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
     
    par_price = pd.ExcelFile(Conf["price_dir"] + Company + ".xls").parse(Conf["price_sheet"])
    par_price.set_index("Part Number", inplace=True)
    buy.loc[:, Company] = par_price[col4]
    buy.apply(check_buy, axis = 1) # - check after round to 2 digits

buy.loc[:, "LMSRP_RU"] = lmsrp["Russia"]["Price [EUR]"]
buy.loc[:, "LMSRP_KZ"] = lmsrp["Kazakhstan"]["Price [EUR]"]
buy.loc[:, "LMSRP_UA"] = lmsrp["Ukraine"]["Price [EUR]"]

# check dics & coef.
buy_disc = mlp[["partnum", "partlabel"]] # check discounts
buy_disc.set_index("partnum", inplace=True)
buy_disc.loc[:,"Product Group"] = buy["Product Group"] 
buy_disc.loc[:,"New disc"] = buy["Product Group"].apply(new_disc)
buy_disc.loc[:,"LMSRP_RU"] = buy["LMSRP_RU"] 

buy_koef = mlp[["partnum", "partlabel"]] # check k
buy_koef.set_index("partnum", inplace=True)
buy_koef.loc[:,"Material Category Name"] = buy["Material Category Name"] 

def change_zero(Series):
    '''prevent division by zero'''
    if Series == 0:
        return 0.01
    elif Series == "999999,999,99": # why?
        return 999999999.99 
    return float(Series)

# check MLP
buy_mlp = mlp[["partnum", "partlabel"]] 
buy_mlp.set_index("partnum", inplace=True)
buy_mlp.loc[:,"Product Group"] = buy["Product Group"]
buy_mlp.loc[:,"MPG"] = buy["MPG"]
buy_mlp.loc[:,"New disc"] = buy["Product Group"].apply(new_disc)
buy_mlp.loc[:,"MLP"] = buy["MLP"]
    
for Company in Partners:
    log.append(" ".join(Company))
    cur  = Par[Company]["cur"]
    country = Par[Company]["country"]
    if (cur == "EUR"):
        buy_disc[Company] = (100 - buy[Company] / lmsrp[country]["Price [EUR]"] * 100).map(f) # round to 2 digits
        buy_mlp[Company]  = (100 - buy[Company] / buy["MLP"].apply(change_zero) * cross * 100).map(f) # MLP in USD
        buy_koef[Company] = (buy[Company] / buy["Ref."]).map(f) 
    else: # in USD
        buy_disc[Company] = (100 - buy[Company] / lmsrp[country]["Price [EUR]"] * 100 / cross).map(f) # round to 2 digits
        buy_mlp[Company] =  (100 - buy[Company] / buy["MLP"].apply(change_zero) * 100).map(f)
        buy_koef[Company] = (buy[Company] / cross / buy["Ref."]).map(f) 
    buy_disc.apply(check_disc, axis = 1)

# Calculate new discounts
mpg = buy_mlp["MPG"].unique()
print "MPG: %s groups" % mpg.size
mpg.sort()
mlp1 = buy_mlp[buy_mlp["MLP"] != "999999,999,99"]
mlp2 = mlp1[mlp1["MLP"] != 0]
pdisc = pd.DataFrame(index = mpg, columns = Partners)
pdisc["Name"] = mgroups["Name"]
pdisc["Description"] = mgroups["Description"]
for grp in mpg:
    for Company in Partners:
        ndisc = mlp2[mlp2["MPG"] == grp][Company].mean()
        if not pd.isnull(ndisc):
            pdisc[Company][grp] = int(round(ndisc, 0))
print pdisc

#print buy
writer = pd.ExcelWriter("./ckeck_prices_EUR_USD_" + time.strftime("%Y%m%d") + ".xls")
buy.reset_index(inplace=True)
buy.to_excel(writer, "Prices", index=False)
#buy_disc.reset_index(inplace=True)
#buy_disc.to_excel(writer, "Discounts", index=False)
buy_mlp.reset_index(inplace=True)
buy_mlp.to_excel(writer, "Discounts_MLP", index=False)
#buy_koef.reset_index(inplace=True)
#buy_koef.to_excel(writer, "Coefficients", index=False)

pdisc.reset_index(inplace=True)
pdisc.to_excel(writer, "NEW_Discounts_MLP", index=False)

try:
    writer.save()
except:
    print "\n!!!Error: '.xls' is busy!"

# Log price issues
if log:
    print "\n### Check buy price: ref_p vs. buy_p in pricing.log file!\n"
    for item in log:
        #print unicode(item)
        log_file.writelines(unicode(item).encode('utf-8'))
        log_file.write("\n")
log_file.close()

print "\nEx-rate EUR/USD:", cross
print "Done."
#raw_input()