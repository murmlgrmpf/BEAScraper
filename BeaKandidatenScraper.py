# -*- coding: utf-8 -*-
"""
Created on Tue Dec 22 22:16:18 2015

@author: thomas
"""
#!/usr/bin/env python
 
#import sys
from subprocess import call
import os
import xlwt
import codecs
#import string

def compare(s1, s2):
    return "".join(s1.split()).lower() == "".join(s2.split()).lower()
    
folder = "/home/thomas/Downloads/6/"
folder = os.getcwd()
files = []

for file in os.listdir(folder):
    if file.endswith(".pdf"):
        filetxt = file[:-3]+"txt"
#        arguments =  "-l 2 "+folder+file+" "+folder+filetxt
#        call("pdftotext "+ arguments, shell=True)

        files.append(filetxt)

files = sorted(files)

nrs = ['Nr.']
Name = ['Name']
Geschl = ['m/w']
Abi = ['Abi']
Uni=['Uni']
FH = ['FH']
Fach = ['Fach']
Sem = ['Sem']
HSNote = ['HS-Note']
Schule =['Schule']
Hochschs=['Hochsch']
Stipendium = ['Stipendium']
Preis = ['Preis']
Praktikum = ['Praktikum']
Ausland = ['Ausland']
Kirche = ['Kirche']
Sozial = ['Sozial']
Musik = ['Musik']
Sport = ['Sport']
Motivationsschreiben = ['Motivationsschreiben']
Essay = ['Essay']

for idx,file in enumerate(files):
    nrs.append(idx+1)
    
    with codecs.open(file, 'rb','utf-8') as f:
        lines = f.readlines()
    
    # 0: Name, Geburtstag
    line = lines[0].split(sep=',')
    name = line[0].replace('\t',' ')
    if 'Frau' in name:
        Geschl.append('w')
        Name.append(name.replace('Frau ',''))
    else:
        Geschl.append('m')
        Name.append(name.replace('Herr ',''))
    # 1: Uni, Fakultät
    line = lines[1].split(sep='/')
    uni = line[0].strip().replace('\t',' ')
    if 'Uni' in uni:
        Uni.append(uni)
        FH.append('')
    else:
        FH.append(uni)
        Uni.append('')
    Fach.append(line[1].strip().replace('\t',' '))
    
    # Abschnitte identifizieren
    n= 0
    sects = [0,0,0,0,0,0]
    print(file)
    for line in lines:
        if compare(line, 'Hochschulreife'):
            sects[0]=n
        if compare(line, 'Studium'):
            sects[1]=n
        if compare(line, 'Stipendien / Auszeichnungen'):
            sects[2]=n
        if compare(line, 'Praktische Erfahrung'):
            sects[3]=n
        if compare(line, 'Fremdsprachen'):
            sects[4]=n
        if compare(line, 'Besonderes Engagement'):
            sects[5]=n
        n+=1
    
    # Zerlegen der Abschnitte
    Hochschulreife = lines[sects[0]+1:sects[1]]
    Studium = lines[sects[1]+1:sects[2]]
    Stipendien = lines[sects[2]+1:sects[3]]
    Praktisch = lines[sects[3]+1:sects[4]]
    Fremdsprachen = lines[sects[4]+1:sects[5]]
    Engagement = lines[sects[5]+1:]
    
    # Auswerten der Abschnitte
    line = Hochschulreife[0].split(sep='\t')
    if 'Abitur' in Hochschulreife[0]:
        Abi.append(line[-1].strip())
        Schule.append('')
    else:
        Abi.append('')
        Schule.append(line[-1].strip())
    
    Ausland.append('')
    HSNote.append('')
    Sem.append('')
    HSNoteFound = False
    for line in Studium:
        if compare(line,'Ausland:'):
            Ausland[-1]='x'
        if 'master' in line.lower():
            print("Master")
        if 'Fortschritt' in line:
            Sem[-1] = (line.split(sep='\t')[-1].strip())
        if (('Notendurchschnitt' in line)&(not(HSNoteFound))):
            HSNote[-1] = (line.split(sep='\t')[-1].strip())
            HSNoteFound = True
    
    Hochschs.append('')
    
    Preis.append(0)
    Stipendium.append(0)
    for line in Stipendien:
        Stipendium[-1]+=1
        if 'preis' in line.lower():
            Preis[-1]+=1
    if Stipendium[-1]==0:
        Stipendium[-1] = ''
    if Preis[-1]==0:
        Preis[-1] = ''
    
    Praktikum.append(0)
    for line in Praktisch:
        Praktikum[-1]+=1
    
    Kirche.append(0)
    Musik.append(0)
    Sport.append(0)
    Sozial.append(0)
    for line in Engagement:
        if any(x in line.lower() for x in ['kirche','gott','kathol','evangel']):
            Kirche[-1]+=1
        if any(x in line.lower() for x in ['musik','instrument','geige','gitarre']):
            Musik[-1]+=1
        if any(x in line.lower() for x in ['sport','ball']):
            Sport[-1]+=1
        if any(x in line.lower() for x in ['kinder','sozial','altersheim','aritativ','hilfe']):
            Sozial[-1]+=1
    Motivationsschreiben.append('')
    Essay.append('')
        
#        
#print(names)
#print(genders)
#print(abis)    
#print(unis)
#print(fachhochs)
#print(facs)
#print(ausland)
#print(Fortschritts)
#print(Notendurchschnitts)
#print(Stipendiens)

book = xlwt.Workbook(encoding="utf-8")

sh = book.add_sheet("Bewertungen")

style = xlwt.XFStyle()
# font
font = xlwt.Font()
font.bold = True
style.font = font

sh.write(0,1,"BEA Auswahlverfahren 2016 für den 18. Jahrgang", style=style)

z=2
for idx,name in enumerate(Name):
    
    sh.write(z,0,nrs[idx])
    sh.write(z,1,name)
    sh.write(z,2,Geschl[idx])
    sh.write(z,3,Abi[idx])
    sh.write(z,4,Uni[idx])
    sh.write(z,5,FH[idx])
    sh.write(z,6,Fach[idx])
    sh.write(z,7,Sem[idx])
    sh.write(z,8,HSNote[idx])
    sh.write(z,9,Schule[idx])
    sh.write(z,10,Hochschs[idx])
    sh.write(z,11,Stipendium[idx])
    sh.write(z,12,Preis[idx])
    sh.write(z,13,Praktikum[idx])
    sh.write(z,14,Ausland[idx])
    sh.write(z,15,Kirche[idx])
    sh.write(z,16,Sozial[idx])
    sh.write(z,17,Musik[idx])
    sh.write(z,18,Sport[idx])
    sh.write(z,19,Motivationsschreiben[idx])
    sh.write(z,20,Essay[idx])
    z+=1

book.save("Test.xls")

#f = open('Emmer+Bombe+14.pdf', 'rb')
#test = call(["pdftotext", "-l 1 "+folder+file])

#with open(file, 'rb') as f:
#    doc = PdfFileReader(f)


#for file in files:
#    with open(file, 'rb') as f:
#        doc = PdfFileReader(f)
#        print(doc[0])