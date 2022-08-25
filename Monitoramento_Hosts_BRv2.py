#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Library
from pathlib import Path
import pandas as pd
from datetime import datetime
import subprocess
import sys
import time
import os
import win32com.client as win32

#Variaveis de ambiente
data = datetime.today().strftime('%d-%m-%Y').replace('-','/')
data_dia_mes = datetime.today().strftime('%d-%m').replace('-','/')
data_dia = datetime.today().strftime('%d')
hora = datetime.today().strftime('%H:%M')
dir_base = Path('C:/Data/Projeto_ADHosts')
list_dir_base = dir_base.iterdir()
powershell_cmd = 'powershell.exe'

# #listas
bases = ['SP','BH','FO','BS','PA','RE','SV','RJ','CA','JV','CB']
novos_hosts = []

#Pegar relatorios diarios
print('Downloading Base')
subprocess.run([powershell_cmd, 'C:\\data\\Projeto_ADHosts\\ReportADBR.ps1'])
print('Download Completed')
time.sleep(5)

espera = 'True'
while espera == 'True':
    if Path(r'C:/Data/Projeto_ADHosts/ReportADBR.csv').exists():
        #Carregar Base
        print('Carregando Base')
        base_nova = pd.read_csv(Path('c:/data/Projeto_ADHosts/ReportADBR.csv'))
        base_nova = base_nova[['Name', 'OperatingSystem', 'OperatingSystemVersion']]
        base_antiga = pd.read_csv(Path(r'c:/data/Projeto_ADHosts/ReportADBRantigo.csv'))
        base_antiga = base_antiga[['Name', 'OperatingSystem', 'OperatingSystemVersion']]
        arquivo_base = pd.ExcelWriter(Path(r'C:/data/Projeto_ADHosts/ReportADBR.xlsx', engine='xlsxwriter'))
        #Tratar Builds
        print('Gerando relatórios das Builds')
        base_builds_w10 = base_nova[base_nova['OperatingSystem'] == 'Windows 10 Enterprise']
        base_builds_w10 = base_builds_w10[['Name', 'OperatingSystemVersion']]
        base_builds_w10.rename(columns={'Name': 'Hosts', 'OperatingSystemVersion': 'Build'}, inplace = True)
        base_builds_w10_report = base_builds_w10
        base_builds_w10 = base_builds_w10.groupby(['Build']).count()
        base_builds_w10.to_excel(arquivo_base, sheet_name='Build Windows 10')
        base_builds_w11 = base_nova[base_nova['OperatingSystem'] == 'Windows 11 Enterprise']
        base_builds_w11 = base_builds_w11[['Name', 'OperatingSystemVersion']]
        base_builds_w11.rename(columns={'Name': 'Hosts', 'OperatingSystemVersion': 'Build'}, inplace = True)
        base_builds_w11_report = base_builds_w11
        base_builds_w11 = base_builds_w11.groupby(['Build']).count()
        base_builds_w11.to_excel(arquivo_base, sheet_name='Build Windows 11')
        base_builds_w111 = base_nova[base_nova['OperatingSystem'] == 'Windows 11 Enterprise Insider Preview']
        base_builds_w111 = base_builds_w111[['Name', 'OperatingSystemVersion']]
        base_builds_w111.rename(columns={'Name': 'Hosts', 'OperatingSystemVersion': 'Build'}, inplace = True)
        base_builds_w111_report = base_builds_w111
        base_builds_w111 = base_builds_w111.groupby(['Build']).count()
        base_builds_w111.to_excel(arquivo_base, sheet_name='Build Windows 11 Preview')
        base_builds_w10pro = base_nova[base_nova['OperatingSystem'] == 'Windows 10 Pro']
        base_builds_w10pro = base_builds_w10pro[['Name', 'OperatingSystemVersion']]
        base_builds_w10pro.rename(columns={'Name': 'Hosts', 'OperatingSystemVersion': 'Build'}, inplace = True)
        base_builds_w10pro_report = base_builds_w10pro
        base_builds_w10pro = base_builds_w10pro.groupby(['Build']).count()
        base_builds_w10pro.to_excel(arquivo_base, sheet_name='Build Windows 10 Pro')
        base_builds_w81 = base_nova[base_nova['OperatingSystem'] == 'Windows 8.1 Enterprise']
        base_builds_w81 = base_builds_w81[['Name', 'OperatingSystemVersion']]
        base_builds_w81.rename(columns={'Name': 'Hosts', 'OperatingSystemVersion': 'Build'}, inplace = True)
        base_builds_w81_report = base_builds_w81
        base_builds_w81 = base_builds_w81.groupby(['Build']).count()
        base_builds_w81.to_excel(arquivo_base, sheet_name='Build Windows 8.1 Enterprise')
        base_builds_w11_report.to_excel(arquivo_base, sheet_name='Windows 11 Hosts', index = False)
        base_builds_w111_report.to_excel(arquivo_base, sheet_name='Windows 11 Preview Hosts', index = False)
        base_builds_w10pro_report.to_excel(arquivo_base, sheet_name='Windows 10 Pro', index = False)
        base_builds_w81_report.to_excel(arquivo_base, sheet_name='Windows 8.1 Enterprise Hosts', index = False)
        
        
        #Importar/tratar bases
        print('Gerando relatório Hosts')
        arquivo_novo = base_nova['Name'].str.lower()
        arquivo_antigo = base_antiga['Name'].str.lower()
        arquivo_novo_lista = arquivo_novo.to_list()
        arquivo_antigo_lista = arquivo_antigo.to_list()
        for linha, n in enumerate(arquivo_novo_lista):
            if arquivo_novo_lista[linha] in arquivo_antigo_lista:
                continue
            else:
                novos_hosts.append(f'Novo: {arquivo_novo_lista[linha]}') 
        for linha2, l in enumerate(arquivo_antigo_lista):
            if arquivo_antigo_lista[linha2] in arquivo_novo_lista:
                continue
            else:
                novos_hosts.append(f'Removed: {arquivo_antigo_lista[linha2]}')
        #Gerando Reports
        for i, rel in enumerate(bases):
            Rel_final = pd.DataFrame(novos_hosts, columns=['Name'])
            procura = 'br' + bases[i].lower()
            Rel_final = Rel_final[Rel_final['Name'].str.contains(f'{procura}', na = False)]
            Rel_final.to_excel(arquivo_base, sheet_name=f'Hosts {bases[i]}', index=False)
        arquivo_base.save()
        os.remove('c:\Data\Projeto_ADHosts\ReportADBRantigo.csv')
        time.sleep(3)
        Path(dir_base / Path(r'ReportADBR.csv')).rename(dir_base / Path(r'ReportADBRantigo.csv'))
        print('Relatório Gerado')
        break

#Informação Email
print('Enviando Email')
relsp = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts SP')
relspnovo = len(relsp[relsp['Name'].str.contains('Novo', na = False)])
relspremovido = len(relsp[relsp['Name'].str.contains('Removed', na = False)])
relbh = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts BH')
relbhnovo = len(relbh[relbh['Name'].str.contains('Novo', na = False)])
relbhremovido = len(relbh[relbh['Name'].str.contains('Removed', na = False)])
relfo = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts FO')
relfonovo = len(relfo[relfo['Name'].str.contains('Novo', na = False)])
relforemovido = len(relfo[relfo['Name'].str.contains('Removed', na = False)])
relbs = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts BS')
relbsnovo = len(relbs[relbs['Name'].str.contains('Novo', na = False)])
relbsremovido = len(relbs[relbs['Name'].str.contains('Removed', na = False)])
relpa = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts PA')
relpanovo = len(relpa[relpa['Name'].str.contains('Novo', na = False)])
relparemovido = len(relpa[relpa['Name'].str.contains('Removed', na = False)])
relre = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts RE')
relrenovo = len(relre[relre['Name'].str.contains('Novo', na = False)])
relreremovido = len(relre[relre['Name'].str.contains('Removed', na = False)])
relsv = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts SV')
relsvnovo = len(relsv[relsv['Name'].str.contains('Novo', na = False)])
relsvremovido = len(relsv[relsv['Name'].str.contains('Removed', na = False)])
relrj = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts RJ')
relrjnovo = len(relrj[relrj['Name'].str.contains('Novo', na = False)])
relrjremovido = len(relrj[relrj['Name'].str.contains('Removed', na = False)])
relca = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts CA')
relcanovo = len(relca[relca['Name'].str.contains('Novo', na = False)])
relcaremovido = len(relca[relca['Name'].str.contains('Removed', na = False)])
reljv = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts JV')
reljvnovo = len(reljv[reljv['Name'].str.contains('Novo', na = False)])
reljvremovido = len(reljv[reljv['Name'].str.contains('Removed', na = False)])
relcb = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Hosts CB')
relcbnovo = len(relcb[relcb['Name'].str.contains('Novo', na = False)])
relcbremovido = len(relcb[relcb['Name'].str.contains('Removed', na = False)])

buildw10 = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Build Windows 10')
buildw11 = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Build Windows 11')
buildw111 = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Build Windows 11 Preview')    
buildw10pro = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Build Windows 10 Pro')
buildw81 = pd.read_excel(Path(r'c:/Data/Projeto_ADHosts/ReportADBR.xlsx'), sheet_name='Build Windows 8.1 Enterprise')    

#Email
base_contatoTI = pd.read_csv(dir_base / Path(r'C:/Data/Projeto_ADHosts/ContatosIT.csv'), delimiter=';')
outlook = win32.Dispatch('Outlook.application')
for i, contato in enumerate(base_contatoTI['Email']):
    mail = outlook.CreateItem(0)
    mail.To = base_contatoTI['Email'][i]
    mail.Subject = f'{data} - E-mail Automático - Monitoramento Hosts DTT-BR(Não Responder)'
    mail.Body = f''' Olá {base_contatoTI['Nome'][i]},
 
 Quantidade de Host Por Build
     
 Windows 10 Enterprise                 Windows 11 Enterprise                                     Windows 10 Pro: {buildw10pro['Hosts'][0]}               

 Build                                                   Build                                                                    Windows 8.1 Enterprise: {buildw81['Hosts'][0]}                                                   
 1709 (16299)         {buildw10['Hosts'][0]}                         10.0 (22000)      {buildw11['Hosts'][0]}
 1809 (17763)         {buildw10['Hosts'][1]}   
 1909 (18363)         {buildw10['Hosts'][2]}                    Windows 11 Enterprise Inside Preview
 20h2 (19042)         {buildw10['Hosts'][3]}                  10.0 (25131)      {buildw111['Hosts'][0]}
 21h1 (19043)         {buildw10['Hosts'][4]}
 21h2 (19044)         {buildw10['Hosts'][5]}   
 
 Movimentações Active Directory
 
 SP                            BH                         FO                              SV
 Adicionados: {relspnovo}     Adicionados: {relbhnovo}    Adicionados: {relfonovo}       Adicionados: {relsvnovo}
 Removidos: {relspremovido}       Removidos: {relbhremovido}      Removidos: {relforemovido}         Removidos: {relsvremovido}
 
 BS                            PA                         RE                              RJ
 Adicionados: {relbsnovo}     Adicionados: {relpanovo}    Adicionados: {relrenovo}       Adicionados: {relrjnovo}
 Removidos: {relbsremovido}       Removidos: {relparemovido}      Removidos: {relreremovido}         Removidos: {relrjremovido}
 
 CA                            JV                         CB
 Adicionados: {relcanovo}     Adicionados: {reljvnovo}    Adicionados: {relcbnovo}
 Removidos: {relcaremovido}       Removidos: {reljvremovido}      Removidos: {relcbremovido}
 

 '''
    attachment = r'C:\Data\Projeto_ADHosts\ReportADBR.xlsx'
    mail.Attachments.Add(attachment)
    mail.Send()
print('Email Enviado!')

# In[ ]:




