# -*- coding: UTF-8 -*-
from __future__ import print_function
import httplib2
import os
import re
import json

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from apiclient.discovery import build

import datetime



try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
CLIENT_SECRET_FILE = 'client_secret.json'
SCOPES = ['https://www.googleapis.com/auth/drive','https://www.googleapis.com/auth/calendar.readonly']

def get_credentials(APPLICATION_NAME,jsonname):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   jsonname+'-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME+' Python Quickstart'
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials
def googleauth(kind):
    if kind == 'sheets':
        spcredentials = get_credentials('Google Sheets API','sheets.googleapis.com-python')
        sphttp = spcredentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?''version=v4')
        return discovery.build('sheets', 'v4', http=sphttp,discoveryServiceUrl=discoveryUrl)
    elif kind == 'calender':
        credentials = get_credentials('Google Calendar API','calendar-python')
        http = credentials.authorize(httplib2.Http())
        return discovery.build('calendar', 'v3', http=http)        
def spreadin():
    print('GoogleCalender～GoogleSpreadsheet')
    def jsondel(jsonname):
        jsonin = open(jsonname,'w')
        jsonin.write('')

        return(spservice)
    def spreaddel(sheet_id):
        spservice=googleauth(kind='sheets')
        spreadsheet_id = '1icjJD1kxiqYhhOzcNSpZ2FdBj9jTCMxliRCzuvNcgog'
        rows = []
        spvalues = []
        requests = []
        for i in range(0,500):
            a = ['','','','','','','','','','','','','','','','']
            for b in a:
                spvalues.append({'userEnteredValue': {'stringValue':b}})
            rows.append({'values': spvalues})
            spvalues = []
        requests.append(
                    {
                        'updateCells': {
                        'start': {
                            'sheetId': sheet_id,
                            'rowIndex': 0,
                            'columnIndex': 0
                        },
                        'rows': rows,
                        'fields': 'userEnteredValue'}})
        batch_update_spreadsheet_request_body = {'requests': requests}
        result = spservice.spreadsheets() \
            .batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_spreadsheet_request_body) \
            .execute()
    def spreadwrite(rows,row_index,sheet_id):
        spservice = googleauth(kind='sheets')
        requests = []
        spreadsheet_id = '1icjJD1kxiqYhhOzcNSpZ2FdBj9jTCMxliRCzuvNcgog'
        requests.append(
            {
                'updateCells': {
                'start': {
                    'sheetId': sheet_id,
                    'rowIndex': row_index,
                    'columnIndex': 0
                },
                'rows': rows,
                'fields': 'userEnteredValue'}})
        batch_update_spreadsheet_request_body = {'requests': requests}
        result = spservice.spreadsheets() \
                    .batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_spreadsheet_request_body) \
                    .execute()
    def spreadcreate(spreadsheet_id,title):
        requests=[]
        requests.append(
            { # シートの作成
            'addSheet': {
                'properties': {
                'title': title,
                'index': 0}}})
        batch_update_spreadsheet_request_body = {'requests': requests}
        result = spservice.spreadsheets() \
            .batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_spreadsheet_request_body) \
            .execute()
    import csv
    import codecs
    yearinput = int(input('取得する年を入力 例※18 : '))
    monthinput = int(input('取得する月を入力 例※1 : '))
    if monthinput == 12:
        monthMin = 12
        monthMin = str(monthMin)
        monthMax = '%02d' % 1
        monthMax = str(monthMax)
        yearMin = '20'+str(yearinput)
        yearMax = yearinput + 1
        yearMax = '20'+str(yearMax)
    else:
        monthMin = '%02d' % monthinput
        monthMin = str(monthMin)
        monthMax = monthinput + 1
        monthMax = '%02d' % monthMax
        monthMax = str(monthMax)
        yearMin = '20'+str(yearinput)
        yearMax = '20'+str(yearinput)
    service=googleauth(kind='calender')

    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    
    eventsResult = service.events().list(
        calendarId='40a2n1ueud081u9smar5n95ka8@group.calendar.google.com', timeMin=yearMin+'-'+monthMin+'-01T00:00:00.000000Z',timeMax=yearMax+'-'+monthMax+'-01T00:00:00.000000Z', maxResults=1000, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])
    if not events:
        print('No upcoming events found.')


    spservice=googleauth(kind='sheets')
    spreadsheet_id = '1icjJD1kxiqYhhOzcNSpZ2FdBj9jTCMxliRCzuvNcgog'
    sheets = spservice.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheet = sheets.get('sheets','')
    count = 0
    for i in range(0,len(sheet)):
        if re.search (yearMin+'-'+monthMin, sheet[i].get('properties').get('title')):
            sheetId = sheet[i].get('properties').get('sheetId')
            count=1
    if count == 0:
        title = '実績表'+yearMin+'-'+monthMin
        spreadcreate(spreadsheet_id,title)
        sheets = spservice.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet = sheets.get('sheets','')
        for i in range(0,len(sheet)):
            if re.search (yearMin+'-'+monthMin, sheet[i].get('properties').get('title')):
                sheetId = sheet[i].get('properties').get('sheetId')
    else:
        exitinput = input('既にこの月の実績表が作成されています、続行する場合は 1 を入力してEnter、終了する場合はそのままEnterを入力して下さい\n')
        if exitinput == '1':
            pass
        else:
            exit()

    row1='施主番号','案件番号','月日','施主名','出発予定地','到着予定地','総台数','誘導員名','前払い(円)','給与(円)','請求金額(円)','経費(請求)','経費(明細)','業者','備考欄(請求用)','備考欄(社内用)'
    spvalues=[]
    rows=[]
    for i in row1:
        spvalues.append({'userEnteredValue': {'stringValue':i}})
    rows.append({'values': spvalues})
    spreaddel(sheetId)
    spreadwrite(rows,0,sheetId)
    count = 1#スプレッドシート入力時に名前の配列を取ってくる番号
    spcount = 0 #スプレッドシート入力時にデータをｘ数まで取る変数
    datecount = 0
    datekeep = ''
    eventnum = 1
    row_index = 1#googlespread入力時に書き込み場所を定義
    onetoten = [1,2,3,4,5,6,7,8,9]
    tentonine = [10,11,12,13,14,15,16,17,18,19,20,21,23,24,25,26,27,28,29,30,31,32,33,34,34,35,36,37]
    spvalues = []
    rows = []
    requests = []
    suplyin = open('suplylist.json', 'a')
    suplyload = open('suplylist.json','r')
    for event in events:
        start = event['start'].get('dateTime')#建設と終日タグのものを消す dateTimeキーを持っていない物をNoneに変換
        if start == None:
            continue
        elif re.search('建設',event['summary']):#建設タグ所有の物を削除
            continue
        else:#不要な日付データを削除
            pattern = "(.*)T(.*)"
            date = re.search(pattern,start)
        if 'description' not in event.keys():#Description処理
            event['description'] = ''
        description = event['description'].replace('\n','')
        description = description.replace('\r','')#ここまで
        dateinfo = date.group(1).replace('-','')#日付をやりやすいよう変更
        month = dateinfo[4:6]#月日を詳細に取得
        day = dateinfo[6:8]#ここまで
        eventinfo = event['summary'].replace('〜','--')#~の種類を変換
        eventinfo = eventinfo.replace('～','--')
        eventinfo = eventinfo.replace('~','--')
        eventinfo = eventinfo.replace('→','--')#ここまで
        if re.search('※',eventinfo):#米印、アスタリスクをdescriptionへ
            eventinfo = eventinfo[1:]
            description = '※'+description
        elif re.search('＊',eventinfo):
            eventinfo = eventinfo[1:]
            description = '＊'+description
        elif re.search('#',eventinfo):
            eventinfo = eventinfo[1:]
            description = '#'+description#ここまで
        if datekeep == '':
            datekeep = dateinfo[6:8]
        if datekeep == dateinfo[6:8]:#
            datecount+=1
        elif datekeep != dateinfo[6:8]:
            datecount = 1
            datekeep = dateinfo[6:8]
        eventnumzeropad = '%03d' % datecount
        dateinfo = dateinfo+eventnumzeropad

        eventinfo = eventinfo.replace('　',' ')
        eventsplit = re.split(" +",eventinfo)#イベントの現場情報を取得から分割
        arr = ''
        dep = ''
        if re.search('--',eventsplit[0]):
            i = re.search('(.*)--(.*)',eventsplit[0])
            dep = i.group(1)
            arr = i.group(2)
        else:
            dep = eventsplit[0]
        if len(eventsplit) == 1:
            eventsplit.append('--')
            eventsplit.append('--')
            eventsplit.append('--')
        if len(eventsplit) == 2:
            eventsplit.append('--')
            eventsplit.append('--')
        if len(eventsplit) == 3:
            eventsplit.append('--')
        company = eventsplit[1]
        names = re.split('、',eventsplit[2])
        names.append(eventsplit[3])
        for i in range(0,20):
            try:
                if re.search('--',names[i]):
                    del names[i]
            except:
                pass
        cars = len(names)
        for i in range(1,10):
            try:
                if int(company[:1]) == i:
                    company = company[1:]
                    arr = (str(i)+arr)
            except:
                pass
        for i in range(10,100):
            try:
                if int(company[:2]) == i:
                    company = company[2:]
                    arr = (str(i)+arr)
            except:
                pass
        key_in = open('keys.json' , 'r')
        keys = json.load(key_in)
        company_list = keys['company']
        company_keys = company_list.keys()
        if company not in company_keys:
            company_name = company
            company_id = '不明'
        else:
            company = company_list.get(company)
            company_id = company.get('id')
            company_name = company.get('name')
        name_list = keys['name']
        name_keys = name_list.keys()
        fullname = []
        for i in names:
            if re.search('[0-9]+$',i):
                cars_in = re.search('[0-9]+$',i)
                cars = int(cars)+ int(cars_in.group(0))
                cars -= 1
                i=re.sub('[0-9]+$','',i)
                for b in range(1,int(cars_in.group(0))):
                    fullname.append(i)

            if i not in name_keys:
                fullname.append(i)
            else:
                fullname.append(name_list.get(i))

        count = 0
        monthday =month+'月'+day+'日'

        for i in range(0,cars):
            if i == 0:
                c = [company_id,dateinfo,monthday,company_name,dep,arr,str(cars),fullname[count],'','','','','','','',description]
            elif i != 0:
                c = ['',dateinfo,monthday,company_name,dep,arr,'',fullname[count],'','','','','','','']
            for d in c:
                spvalues.append({'userEnteredValue': {'stringValue':d}})
            rows.append({'values': spvalues})
            count += 1
            spcount += 1 


            spvalues = []
    row_index = row_index
    spreadwrite(rows,row_index,sheetId)
def createotherfile(name,parents,http,sheetid):
    service = build('drive', 'v3', http=http)
    new_file_body = {
    'name': name,  # 新しいファイルのファイル名. 省略も可能
    'parents': [parents],  # Copy先のFolder ID. 省略も可能
    }
    new_file = service.files().copy(
        fileId=sheetid, body=new_file_body
    ).execute()
def spalign(rows,row_index,columnIndex):
    requests=[]
    requests.append(
        {
            'updateCells': {
            'start': {
                'sheetId': 0,
                'rowIndex': row_index,
                'columnIndex': columnIndex
            },
            'rows': rows,
            'fields': 'userEnteredValue'}})
    return requests
def deleterows(rowcount,Sheet_id,rowindex,columnindex):
    spservice=googleauth(kind='sheets')
    invalues = []
    rows=[]
    for i in range(rowcount):
        for d in range(9):
            invalues.append({'userEnteredValue': {'stringValue':''}})
        rows.append({'values': invalues})
        invalues=[]
    requests=spalign(rows,rowindex,columnindex)
    batch_update_spreadsheet_request_body = {'requests': requests}
    print(batch_update_spreadsheet_request_body)
    result = spservice.spreadsheets() \
                .batchUpdate(spreadsheetId=Sheet_id, body=batch_update_spreadsheet_request_body) \
                .execute()
def statement():
    print('GoogleSpreadsheet 実績表～明細書')
    rows = []

    def general():
        rows = []
        rowcount = 0
        '----個人用----'
        yearinput = int(input('取得する年を入力 ※例 18:'))
        year = '20'+str(yearinput)
        monthinput = int(input('取得する月を入力 例 1:'))
        month = '%02d' % monthinput
        month = str(month)
        name = input('明細書を発行する誘導員名を入力 ※例 植田 博幸: ')

        spservice=googleauth(kind='sheets')
        sheetsuply = '1icjJD1kxiqYhhOzcNSpZ2FdBj9jTCMxliRCzuvNcgog'
        sheetdetail = '1lY_qDJlXlWdoTdIGAaYsSJukmLMTd36RhcGRLd_MQpI'
        rangeName = '実績表'+year+'-'+month+'!A2:P'
        result = spservice.spreadsheets().values().get(
            spreadsheetId=sheetsuply, range=rangeName).execute()
        values = result.get('values', [])
        rows.append({'values':{'userEnteredValue': {'stringValue':name}}})
        spalign(rows,6,2)
        rows = []
        rows.append({'values':{'userEnteredValue': {'stringValue':month+'月度支払い明細'}}})
        spalign(rows,0,2)
        rows = []
        salaytotal = 0
        prepaytotal = 0
        costtotal = 0
        invalues = []

        for i in values:
            lenminus = 16 - len(i)
            for a in range(0,lenminus):
                i.append('')
            if re.search(name,i[7]):
                rowcount += 1
                pass
            else:
                continue
            dateinfo = i[2]#日付
            dep = i[4]#出発地
            arr = i[5]#到着地
            prepay = i[8]#前払い
            salay = i[9]#給与
            cost = i[12]#経費(明細)
            company = i[13]#業者
            descriptionforus = i[15]#備考欄 社内用
            inlist = [dateinfo,dep,arr,salay,prepay,cost,'',descriptionforus]
            for d in inlist:
                invalues.append({'userEnteredValue': {'stringValue':d}})
            rows.append({'values': invalues})
            invalues=[]

            if salay == '':
                salay = 0
            if cost == '':
                cost = 0
            if prepay == '':
                prepay = 0
            salaytotal = salaytotal + int(salay)
            prepaytotal = prepaytotal + int(prepay)
            costtotal = costtotal + int(cost)
        print('カウント--'+str(rowcount))
        requests=spalign(rows,15,0)
        rows=[]
        batch_update_spreadsheet_request_body = {'requests': requests}
        print('給与:'+str(salaytotal))
        print('前払い'+str(prepaytotal))
        print('経費:'+str(costtotal))

        result = spservice.spreadsheets() \
                    .batchUpdate(spreadsheetId=sheetdetail, body=batch_update_spreadsheet_request_body) \
                    .execute()
        name = year+month+' 明細 '+name
        createotherfile(name,'1dZZwUlvEMIN2A7P2Avs5wHIVw6IoX7vN',sphttp,sheetdetail)
        deleterows(rowcount,sheetdetail,15,0)
    def company():
        rows = []
        rowcount = 0
        print('----企業用----')
        yearinput = int(input('取得する年を入力 ※例 18:'))
        year = '20'+str(yearinput)
        monthinput = int(input('取得する月を入力 例 1:'))
        month = '%02d' % monthinput
        month = str(month)
        name = input('明細書を発行する企業名を入力 ※例 エーブル: ')
        key_in = open('keys.json' , 'r')
        keys = json.load(key_in)
        company_list = keys['name']
        company_keys = company_list.keys()
        if name not in company_keys:
            company_name = name
            company_id = '不明'
        else:
            company_name = company_list.get(name)
        rows.append({'values':{'userEnteredValue': {'stringValue':company_name}}})
        spalign(rows,6,2)
        rows = []
        spservice=googleauth(kind='sheets')
        sheetsuply = '1icjJD1kxiqYhhOzcNSpZ2FdBj9jTCMxliRCzuvNcgog'
        sheetdetail = '1oB32QjnpCn2DHc8XcUF7TXSjLNCs1XdSckjUiZiPUj0'
        rangeName = '実績表'+year+'-'+month+'!A2:P'
        result = spservice.spreadsheets().values().get(
            spreadsheetId=sheetsuply, range=rangeName).execute()
        values = result.get('values', [])
        rows.append({'values':{'userEnteredValue': {'stringValue':month+'月度支払明細'}}})
        spalign(rows,0,2)
        rows = []
        salaytotal = 0
        prepaytotal = 0
        costtotal = 0
        invalues = []

        for i in values:
            lenminus = 16 - len(i)
            for a in range(0,lenminus):
                i.append('')
            if re.search(name,i[7]):
                rowcount += 1
                pass
            else:
                continue
            cars = i[6]#総台数
            prepay = i[8]#前払い
            salay = i[9]#給与
            cost = i[12]#経費(明細)
            dateinfo = i[2]#日付
            dep = i[4]#出発地
            arr = i[5]#到着地
            prepay = i[8]#前払い
            salay = i[9]#給与
            cost = i[12]#経費(明細)
            company = i[13]#業者
            descriptionforus = i[15]#備考欄 社内用                
            inlist = [dateinfo,dep,arr,cars,salay,prepay,cost,'',descriptionforus]
            for d in inlist:
                invalues.append({'userEnteredValue': {'stringValue':d}})
            rows.append({'values': invalues})
            invalues=[]

        spalign(rows,21,0)
        rows=[]

        batch_update_spreadsheet_request_body = {'requests': requests}
        result = spservice.spreadsheets() \
                    .batchUpdate(spreadsheetId=sheetdetail, body=batch_update_spreadsheet_request_body) \
                    .execute()
        name = year+month+' 明細 '+name
        createotherfile(name,'1dZZwUlvEMIN2A7P2Avs5wHIVw6IoX7vN',sphttp,sheetdetail)
        deleterows(rowcount,sheetdetail,21,0)
    if __name__ == '__main__':
        statementinput = int(input('個人用明細.1 業者用明細.2: '))
        if statementinput == 1:
            general()
        elif statementinput == 2:
            company()
def requests():
    requests=[]
    def spalign(rows,row_index,columnIndex):
        requests.append(
            {
                'updateCells': {
                'start': {
                    'sheetId': 0,
                    'rowIndex': row_index,
                    'columnIndex': columnIndex
                },
                'rows': rows,
                'fields': 'userEnteredValue'}})
    
    def general():
        rows=[]
        rowcount = 0
        print('一般請求書')
        yearinput = int(input('取得する年を入力 ※例 18:'))
        year = '20'+str(yearinput)
        monthinput = int(input('取得する月を入力 例 1:'))
        month = '%02d' % monthinput
        month = str(month)
        name = input('請求書を発行する企業名を入力 ※例 エーブル: ')

        key_in = open('keys.json','r')
        keys = json.load(key_in)
        company_list = keys['company']
        company_keys = company_list.keys()
        if name not in company_keys:
            company_name = name
            company_id = '不明'
        else:
            namelist = company_list.get(name)
            company_name = namelist.get('name')
        key_in.close()
        rows.append({'values':{'userEnteredValue': {'stringValue':company_name}}})
        spalign(rows,7,0)
        rows = []
        spservice=googleauth(kind='sheets')
        sheetsuply = '1icjJD1kxiqYhhOzcNSpZ2FdBj9jTCMxliRCzuvNcgog'
        sheetrequest = '1lznWRKoB29eCnFUdd2K-Cy8gYc--AhUui9_uc3T05rg'
        rangeName = '実績表'+year+'-'+month+'!A2:P'
        result = spservice.spreadsheets().values().get(
            spreadsheetId=sheetsuply, range=rangeName).execute()
        values = result.get('values', [])


        rows.append({'values':{'userEnteredValue': {'stringValue':year+'年'+str(monthinput)+'月末'}}})
        spalign(rows,4,6)
        rows = []
        totalcost = 0
        invalues = []

        for i in values:
            lenminus = 16 - len(i)
            for a in range(0,lenminus):
                i.append('')
            if i[0] == '':
                continue
            else:
                pass
            if re.search (name, i[3]):
                rowcount += 1
                pass
            else:
                continue
            dateinfo = i[2]#日付
            dep = i[4]#出発地
            arr = i[5]#到着地
            cars = i[6]#総台数
            prepay = i[8]#前払い
            salay = i[9]#給与
            unitcost = i[10]#請求金額(単価)
            outlay = i[11]#経費
            cost = i[12]#経費(明細)
            company = i[13]#業者
            descriptionforcompanys = i[14]#備考欄 請求用
            descriptionforus = i[15]#備考欄 社内用
            if unitcost == '':
                unitcost = '0'
            if cars == '':
                askcontinue = input(dateinfo+'--'+dep+'--'+arr+'の台数が未指定です。続行した場合この行は削除されます\n続行する場合はEnter')
                continue

            totalunitcost = int(unitcost)*int(cars)
            totalcost = totalcost + totalunitcost
            inlist = [dateinfo,dep,arr,cars,unitcost,outlay,str(totalunitcost),descriptionforcompanys]
            for d in inlist:
                invalues.append({'userEnteredValue': {'stringValue':d}})
            rows.append({'values': invalues})
            invalues=[]
        spalign(rows,21,0)
        rows=[]

        rows.append({'values':{'userEnteredValue': {'stringValue':'￥'+str(totalcost)+'-'}}})
        spalign(rows,16,1)
        rows=[]
        batch_update_spreadsheet_request_body = {'requests': requests}

        result = spservice.spreadsheets() \
                    .batchUpdate(spreadsheetId=sheetrequest, body=batch_update_spreadsheet_request_body) \
                    .execute()
        name = year+month+' 請求書 '+name
        createotherfile(name,'1gDLT3nlYNA1sDNfPadzRE1hYaCpIK6_h',sphttp,sheetrequest)
        deleterows(rowcount,sheetrequest,16,1)
    general()
if __name__ == '__main__':
    enterinput = int(input('実績表作成.1 明細書.2 請求書.3: '))
    if enterinput == 1:
        spreadin()
    elif enterinput == 2:
        statement()
    elif enterinput == 3:
        requests()