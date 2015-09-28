#!/usr/bin/env python
# -*- coding: utf-8 -*-
import csv
import sys
from numpy import loadtxt, size
'''
первый аргумент - текстовой файл с нужными именами в одну колонку
второй аргумент - исходный csv файл c Country of Origin
третий аргумент - имя выходного csv файла
'''
partcol = 1 # 
desccol = 2 # колонка с описаниями
coocol = 3  # колонка CoO

print "Start..."

a = loadtxt(sys.argv[1], dtype=str)
n = size(a)

ifile  = open(sys.argv[2], 'rb')
reader = csv.reader(ifile, delimiter=';')

ofile  = open(sys.argv[3], 'wb')
writer = csv.writer(ofile, delimiter=';')

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

print "Done."
raw_input()
