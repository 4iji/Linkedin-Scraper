import pandas
from bs4 import BeautifulSoup
import re
import requests
from sys import *
from googleapiclient.discovery import build
from openpyxl import load_workbook     #VERSION 1.8.5 ONLY
import xlsxwriter
import fuzzy



Start_Index=0
Last_Index=201


username=raw_input('Login:')
password=raw_input('Password:')


writer=pandas.ExcelWriter("Linkedin search output.xlsx",engine='openpyxl')
book=load_workbook("Linkedin search output.xlsx")
writer.book=book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

input_file=pandas.ExcelFile("Linkedin search input.xlsx")
sheet=input_file.parse(0)

def writeDataFrame(index,conm,namedes,midname,id,firstname,lastname,profile_link,maidenname='NA',formattedname='NA',phoneticfirstname='NA',phoneticlastname='NA',
            fphoneticname='NA',headline='NA',location='NA',industry='NA',currentshare='NA',numconnection='NA',summary='NA',specialities='NA',positions='NA',
            picurl='NA'):
    global writer
    df = pandas.DataFrame({'A_ID':[index+1],'B_Company':[conm],'C_name':[namedes],'D_firstname':[firstname],'E_midName':[midname],'F_lastName':[lastname],
                           'G_Linkedin Profile Link':[profile_link],'H_Linkedin-Id':[id],'I_maidenname':[maidenname],'J_formatted-name':[formattedname],
                           'K_phonetic-firstname':[phoneticfirstname],'L_phonetic-lastname':[phoneticlastname],'M_formatted-phonetic-name':[fphoneticname],
                           'N_Headline':[headline],'O_Location':[location],'P_Industry':[industry],'Q_CurrentShare':[currentshare],'R_Connections':[numconnection],
                           'S_Summary':[summary],'T_Specialities':[specialities],'U_Positions':[positions],'V_Picture-Url':[picurl]})
    if index==0:
        df.to_excel(writer,sheet_name="Sheet1",startrow=index,index=False)
    else:
        df.to_excel(writer,sheet_name="Sheet1",index=False,startrow=index+1,header=False)
    writer.save()
    return

def readExcel(index):
    global sheet
    row=sheet.irow(index).real
    if str(row[4])=='nan':
        row[4]=""
    return row[3]+" "+row[4]+" "+row[5],row[1],row[4],row[3]+" "+row[5],row[2],row[1]

session = requests.Session()
LINKEDIN_URL = 'https://www.linkedin.com'
LOGIN_URL = 'https://www.linkedin.com/uas/login-submit'
html = session.get(LINKEDIN_URL).content
soup = BeautifulSoup(html,'html5lib')
csrf = soup.find(id="loginCsrfParam-login")['value']

login_information = {
    'session_key':username,
    'session_password':password,
    'loginCsrfParam': csrf,
}

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.8; rv:31.0) Gecko/20100101 Firefox/31.0'}
session.post(LOGIN_URL, headers=headers, data=login_information)
device= build("customsearch","v1",developerKey="AIzaSyCVlwAUM9QvlrClxQXOl51WcfOFKxITsh8")    #API Keys:   AIzaSyCz56VeNPTb7BNbAwUXJvrLcBNssHa0WDU     ,  AIzaSyDY37hOk_lEmIUg_T-goXFYzL1qYm7zIaA

for j in range(Start_Index,Last_Index):
    print j
    r,c,m,o,nd,com=readExcel(j)

    try:
        results=device.cse().list(q=r+" "+c,num=1,cx='009259153963367271982:fkxfnawlcn4').execute()   # cx: 018365894794576624310:ubawqf09zme      ,       005187750176123224094:h3ylihk-rmi
        #print results
        link=results['items'][0]['formattedUrl']
        #print link
        if j==37:
            link="http://"+link
        html=session.get(link).content
    except  Exception as e:
        print 'Error:',
        print e
        writeDataFrame(j,com,nd,m,"NA","NA","NA","NA")
        continue

    soup=BeautifulSoup(html,'html5lib')


    names=soup.find_all('span')
    fullnames=''
    for i in names:
        #print i
        s=re.findall('.*class="full-name".*>(.*)</span>',str(i))
        if s!=[]:
            if re.match('.*',s[0]):
                fullnames=(s[0])
    if len(fullnames)==0:
        writeDataFrame(j,com,nd,m,"NA","NA","NA","NA")
        continue
    fullnames=fullnames.split()

    con1=soup.find_all('div')
    connections="NA"
    for i in con1:
        #print i
        s=[]
        s=re.findall('class="member-connections"><strong>(.*)</strong>',str(i))
        if s!=[]:
                connections=s[0]
                break


    summary="NA"
    sum1=soup.find_all('p')
    for i in sum1:
        s=[]
        #print i
        s=re.findall('<p class="description" dir="ltr">(.*)<br/>',str(i))
        #print s
        if s!=[]:
            summary=s[0]
            break


    headline=""
    for i in sum1:
        s=[]
        s=re.findall('class="title" dir="ltr">(.*)</p>',str(i))
        if s!=[]:
            headline=s[0]
            break

    loc1=soup.find_all('a')
    location="NA"
    industry="NA"
    for i in loc1:
        s=[]
        s=re.findall('name="location".*>(.*)</a>',str(i))
        if s!=[]:
            location=s[0]
            break
    for i in loc1:
        s=[]
        s=re.findall('name="industry".*>(.*)</a>',str(i))
        if s!=[]:
            industry=s[0]
            break


    img1=soup.find_all('img')
    pic_url="NA"
    for i in img1:
        #print i
        s=[]
        s=re.findall('<img alt=.*height="200".*src="(https://media.licdn.com.*.jpg)" width.*/>',str(i))
        if s!=[]:
            pic_url=s[0]
            break

    positions="NA"
    for i in loc1:
        s=[]
        s=re.findall('<a href=.* title="Learn more about this title">(.*)</a>',str(i))
        if s!=[]:
            positions=s[0]
            break

    skills="NA"
    for i in loc1:
        #print i
        s=[]
        s=re.findall('<a class="endorse-item-name-text" href=.* title="Learn more about this skill">(.*)</a>',str(i))
        if s!=[]:
            skills=s[0]
            break

    id1="NA"
    #print soup
    nom=soup.find_all('div')
    for i in nom:
        #print i
        s=[]
        s=re.findall('<div class=".*" id="member-([0-9]*)">',str(i))
        if s!=[]:
                id1=s[0]
                break


    #print id1
    phonetic_first_name=fuzzy.DMetaphone()(fullnames[0])[0]
    phonetic_last_name=fuzzy.nysiis(fullnames[1])
    formatted_phonetic=fuzzy.nysiis(r)
    writeDataFrame(j,com,nd,m,id1,fullnames[0],fullnames[1],link,formattedname=o,phoneticfirstname=phonetic_first_name,
                   phoneticlastname=phonetic_last_name,fphoneticname=formatted_phonetic,numconnection=connections,summary=summary,headline=headline,
                   location=location,industry=industry,positions=positions,specialities=skills,picurl=pic_url)


