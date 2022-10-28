# gerar relatório completo de presença de série de reuniões
import os
from datetime import datetime
import re
import csv
import codecs
import pandas as pd
import argparse

desc_msg = "Programa que junta varios relatórios de presença do teams em um unico excel"

parser = argparse.ArgumentParser(description = desc_msg)
parser.add_argument('-r', '--Reports', type=str, help = "Diretório que está localizado os relatórios de presença do teams")
parser.add_argument('-p', '--Pattern', type=str, help = "Regex usado para localizar os arquivos de relatórios no diretório especificado")
parser.add_argument('-o', '--Output', type=str, help = "Nome do arquivo xlsx gerado ao final")
args = parser.parse_args()
print(args)

if args.Reports == None:
    print('Parametro "-r" faltando por favor informe-o.')
    exit()
elif args.Pattern == None:
    print('Parametro "-p" faltando por favor informe-o.')
    exit()
elif args.Output == None:
    print('Parametro "-o" faltando por favor informe-o.')
    exit()

if not os.path.isdir(args.Reports):
    print('Não foi possível encontrar o caminho:',args.Reports)
    exit()

file_name, file_extension = os.path.splitext(args.Output)
if file_extension != '.xlsx' :
    args.Output = file_name + '.xlsx'
    print('Alterando extensão do OutPut para .xlsx:', args.Output)

args.Pattern = args.Pattern.replace('*', '.*')

rels = []
for path, dirs, files in os.walk(args.Reports):
    for file in files:
        if re.match(args.Pattern, file):
            rels.append(os.path.join(path, file))

tempDf = []
for csvPath in rels:
    with codecs.open(csvPath, 'rU', 'utf-16') as csvFile:
        reader = csv.reader(csvFile, delimiter='	')
        
        saveData = False
        
        for row in reader:
            if re.match('Nome.*', ' '.join(row)):
                if not saveData:
                    saveData = True
                    continue

            if re.match('3.*', ' '.join(row)):
                break
            
            if not saveData:
                continue
            
            if len(row) > 0:
                tempDf.append({'Participant': row[0], 'Date': datetime.strptime(row[1], '%d/%m/%y %H:%M:%S').date()})

df = pd.DataFrame.from_records(tempDf)

dfPartPerDate = df.groupby('Date')['Participant'].apply(list).reset_index(name='Participants').set_index('Date')

dfRel = pd.DataFrame(columns= sorted(df['Date'].unique()), index= sorted(df['Participant'].unique()))

for date in dfRel.columns:
    for participant in dfRel.index:
        dfRel[date][participant] = participant in dfPartPerDate['Participants'][date]
print(dfRel)
dfRel.to_excel(args.Output)