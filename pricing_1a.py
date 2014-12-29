#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
Выборка по заданному списку с part.num. соответствующих строк из общего файла с ценами
1 аргумент - текстовой файл с нужными именами (part.num.) в одну колонку
2 аргумент - исходный csv файл c ценами msrp ref trans
3 аргумент - имя выходного csv файла

2014 Dec
'''
import csv
import sys
from numpy import loadtxt, size

desccol = 2 # колонка с описаниями
catalog = 3
msrpcol = 4 # колонка с ценой msrp
grpcol = 7  # колонка ценовой группой

a = loadtxt(sys.argv[1], dtype=str)
n = size(a)

ifile  = open(sys.argv[2], 'rb')
reader = csv.reader(ifile, delimiter=';', quotechar='"')
ofile  = open(sys.argv[3], 'wb')
writer = csv.writer(ofile, delimiter=';', quotechar='"', quoting=csv.QUOTE_ALL)

rownum = 0
for row in reader:
    # Save header row.
    if rownum == 0:
        header = row
        writer.writerow(header)
    else:
        colnum = 0
        findit = False
        for col in row:
            if (colnum == 0): # ищем в первой колонке нужный part.num.
                if (n == 1):
                    partnum = a.item()
                    if (col == partnum.strip()): 
                        writer.writerow(row)
                else:
                    for partnum in a: 
                        if (col == partnum.strip()): 
                            writer.writerow(row)
                            #findit = True
            colnum += 1
    rownum += 1

ifile.close()
ofile.close()
