set FILE=new_141128.txt
set PRICE=2014-10-30_items.csv
pricing_1a.py %FILE% %PRICE% msrp.csv
awk95 -f comma_1a.awk msrp.csv > msrp_co.csv
echo "# Please add lmsrp_ru column (#8 or 'H') to msrp_co.csv"
