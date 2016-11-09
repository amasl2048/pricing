#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Расчёт скидок и генерация файлов buy*.xls и jde*.xls
на основе MPG и MLP column
2016 Oct.
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

#config_file = "./base/pricing_conf2.yml" # old
config_file = "./base/pricing_conf16.yml"
print "\nConfig file: ", config_file
try:
    Conf = yaml.load(open(config_file))
except: 
    print "Error: no %s file" % config_file
    sys.exit(0)
pprint.pprint(Conf)
cross = Conf["cross"]   # ex-rate eur/usd

# choice MLP source
if Conf["exported"] == True:
    print "\nExported price list!\n"
    mlp = pd.ExcelFile(Conf["exported_msrp_ru"]).parse(Conf["exported_sheet"])
else:
    mlp = pd.ExcelFile(Conf["mlp_file"]).parse(Conf["sheet"])

mlp = mlp[ mlp["partsapmlp"] != "999999,999,99" ] # delete '999' 
mlp = mlp[ mlp["partsapmpg"] != "GG" ] # delete 'GG' items

def change_u(Series):
    return unicode(Series)
mlp.loc[:,"partnum"] = mlp["partnum"].apply(change_u)

def is_duplicate(Series):
    pn = Series.values.tolist()
    full_len = len(pn)
    set_len = len( set(pn) )
    if full_len != set_len:
        print "mlp full len: %s, uniq len: %s, diff: %s" % (full_len, set_len, full_len - set_len)
        # http://stackoverflow.com/questions/9835762/find-and-list-duplicates-in-python-list
        import collections
        print [item for item, count in collections.Counter(pn).items() if count > 1]
        sys.exit(0)
is_duplicate(mlp["partnum"])

Par = yaml.load(open(Conf["partners"]))
Par_old = yaml.load(open(Conf["partners_old"]))
#print "Parthner's file: ", Conf["partners"]
Partners = Par.keys()

f = lambda x: round(x, 2)

# Read prepared product groups discounts
d = pd.ExcelFile(Conf["prod_groups"]).parse(Conf["prod_sheet"])
d = d.set_index("MPG")
#d = d["Disc. group"]

if not "partsapmlp" in mlp.columns.values.tolist():
    print "Error: no 'partsapmlp' column price"
    sys.exit(0)

mlp = mlp.rename(columns={"partnum": "Part Number",
                            "partlabel": "Designation EN", 
                            "partsapmpg": "Product Group",
                            "partsapmlp": "Price [USD]"})

buy = mlp[["Part Number", "Designation EN"]]
buy.loc[:, "MLP"] = mlp["Price [USD]"]
buy.loc[:, "MPG"] = mlp["Product Group"]

cpq = buy[["Part Number", "Designation EN", "MPG"]]
buy.set_index("Part Number", inplace = True)
cpq.set_index("Part Number", inplace = True)

def disc_calc(df):
    """
    #Price with discount calculated from Product Group
    """
    try:
        new_grp = d.loc[df["Product Group"]]
    except:
        print "\n!!!Error: no '%s' product group" % df["Product Group"]
        sys.exit(0)
    #return Par[Company]["discount"][new_grp]
    col_name = Conf["col_pre"] + " " + Par[Company]["status"]
    Disc = 1 - d[col_name][df["Product Group"]]
    if type(df["Price [USD]"]) == 'unicode':
        print "Error: ", df["Price [USD]"], type(df["Price [USD]"])
    #import pdb; pdb.set_trace()
    out = float(df["Price [USD]"] * Disc)
    if out < 0.01:
        out = 0.01
    return out

def p_price(Company, mlp):
    Pcode = Par[Company]["p_code"]
    cur_old  = Par_old[Company]["cur"]
    cur = "USD"
    country = Par[Company]["country"]
    col4 = "Price [USD]"
    if (cur_old == "EUR"):
        col4_old = "Price [EUR]"
    else:
        #Disc = Disc * cross
        col4_old = "Price [USD]"
    
    pr = mlp.apply(disc_calc, axis = 1)

    #print Company, pr
    pr = pr.map(f) # round to 2 digits

    # save buy price xls
    mlp.reset_index(inplace = True)
    b_price = mlp[["Part Number", "Designation EN", "Product Group"]]
    b_price.set_index("Part Number", inplace = True)
    b_price.loc[:, col4] = pr
    fname = str(Par[Company]["country"]) + "_" + cur + "_" + str(Company)+\
    "(" + str(Par[Company]["internal"]) + ")" + "_" +\
    str(Pcode) + "_" + time.strftime("%Y%m%d") + ".xls"
    b_price = b_price[ b_price[col4] >= 0 ] # delete empty price items
    b_price.reset_index(inplace = True)
    b_price.to_excel(fname, index=False)
    
    # save jde xls
    jde = mlp[["Part Number", "Designation EN"]]
    #import pdb; pdb.set_trace()
    jde.set_index("Part Number", inplace = True)
    jde.loc[:, col4] = pr
    jde.loc[:, "JDE"] = Pcode
    jde = jde[ jde[col4] >= 0 ] # delete empty price items
    jde.reset_index(inplace = True)
    jde = jde.reindex(columns = ["JDE", "Part Number", "Designation EN", col4])
    fname2 = "jde_" + fname
    jde.to_excel(fname2, index=False)
    
    mlp.set_index("Part Number", inplace = True)
    #buy.set_index("Part Number", inplace = True)
    buy.loc[:, Company] = pr
    #buy.apply(check_buy, axis = 1) # - check after round to 2 digits
    
    # calc difference with previous buy price
    last_file = Conf["price_dir"] + str(Company) + ".xls"
    try:
        last_price = pd.ExcelFile(last_file).parse(Conf["price_sheet"])
    except: 
        print "Error: %s price sheet name" % Conf["price_sheet"]
        sys.exit(0)
    last_price.set_index("Part Number", inplace = True)
    buy.reset_index(inplace = True)
    diff_price = buy[["Part Number", "Designation EN", "MPG", Company]]
    diff_price.set_index("Part Number", inplace = True)
    
    buy.set_index("Part Number", inplace = True)
    if cur_old == "EUR":
        cpq.loc[:, Company] = last_price[col4_old] * cross # all in USD
    else:
        cpq.loc[:, Company] = last_price[col4_old]
    diff_price.loc[:, "last"] = cpq[Company]
    diff_price.loc[:, "diff"] = buy[Company] - cpq[Company]
    #import pdb; pdb.set_trace()
    diff_file = "./diff_" + str(Company) + "_" + cur + "_" + time.strftime("%Y%m%d") + ".xls"
    diff_price.reset_index(inplace = True)
    diff_price.to_excel(diff_file, index=False)

mlp.set_index("Part Number", inplace = True)
for Company in Partners:
    print Company
    p_price(Company, mlp)

mlp.reset_index(inplace = True)    
buy_disc = mlp[["Part Number", "Designation EN", "Product Group"]] # check discounts
#buy_koef = lmsrp["Russia"][["Part Number", "Designation EN", "Product Group"]] # check k

mlp.set_index("Part Number", inplace = True)
buy_disc.set_index("Part Number", inplace = True)

def change_zero(Series):
    if Series == 0:
        return None
    return Series
mlp.loc[:,"Price [USD]"] = mlp["Price [USD]"].apply(change_zero) # Prevent divide by 0

for Company in Partners:
    buy_disc[Company] = (1 - buy[Company] / mlp["Price [USD]"]).map(f)
    #import pdb; pdb.set_trace()

# Check old price discounts level
mpg = buy["MPG"].unique()
print "MPG: %s groups" % mpg.size
mpg.sort()
#mlp1 = buy_mlp[buy_mlp["MLP"] != "999999,999,99"]
#buy2 = buy[buy["MLP"] != 0]
pdisc = pd.DataFrame(index = mpg, columns = Partners)
pdisc["Name"] = d["Name"]
pdisc["Description"] = d["Description"]
print "\nCheck discounts level..."
for grp in mpg:
    print grp
    for Company in Partners:
        mlp_avg = buy[buy["MPG"] == grp]["MLP"].mean()
        cpq_avg = cpq[cpq["MPG"] == grp][Company].mean()
        disc_avg = 1 - cpq_avg / mlp_avg 
        #print Company, mlp_avg, cpq_avg, disc_avg
        #import pdb; pdb.set_trace()
        if not pd.isnull(disc_avg):
            pdisc[Company][grp] = round(disc_avg, 2)
#print pdisc

# save buy
writer = pd.ExcelWriter("./Prices_EUR_USD_" + time.strftime("%Y%m%d") + ".xls")
buy.reset_index(inplace = True) 
buy_disc.reset_index(inplace = True)
buy.to_excel(writer, "Prices", index=False)
buy_disc.to_excel(writer, "Discounts", index=False)
pdisc.reset_index(inplace=True)
pdisc.to_excel(writer, "NEW_Discounts_MLP", index=False)
#buy_koef.to_excel(writer, "Coefficients", index=False)
try:
    writer.save()
except:
    print "\n!!!Error: '.xls' is busy!"

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