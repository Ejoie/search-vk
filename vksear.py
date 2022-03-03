import xlrd
import time
import codecs
import requests
import json

# Инициализация файла конфигурации, который содержит путь к файлу exсel с исходными данными и токен авторизации в vk
config = codecs.open( "config.txt", "r", "utf_8_sig" )
f = config.read()
config.close()

# распределение по переменным
f = f.split('\n')
count1 = 10 # Количество людей выдаваемых в результате поиска по указанному в данных городу
count2 = 15 # Количество людей выдаваемых в результате поиска без привязки к городу
token = f[1] # Токен доступа для vk
path = f[0][:-1]
age = 25
v=5.101

# Получение id города по названию
def id_city(title):
    url = "https://api.vk.com/method/database.getCities?q="+str(title)+"&country_id=1&need_all=0&count=1&access_token="+token+"&v="+str(v)
    response = requests.get(url).text # Выполнение запроса
    data = json.loads(response) # Преобразование в json
    data = data['response']
    data = data['items']
    city = data[0]['id']
    return city

# Формирование полей запроса для вторичного поиска без привязки к городу
def s_query(FI):
    return {
             'fields':'photo_100,bdate,education,contacts,site,city,activities,career,occupation',
             'q': FI,
             'age_from' : age,
             'count': str(count2),             
             'access_token': token,
             'v': v
             }

# Формирование полей запроса для первичного поиска людей в городе, указанном в источнике
def f_query(FI, city):
    return {
             'fields':'photo_100,bdate,education,contacts,site,city,activities,career,occupation',
             'country': '1',
             'city': id_city(city),
             'q': FI,
             'age_from' : age,
             'count': str(count1),             
             'access_token': token,
             'v': v
             }    

# Создание списка результатов поиска
def make_list(query):
    l = list() # Для результатов поиска
    time.sleep(0.4) # задержка в посылке запросов
    url = "https://api.vk.com/method/users.search"
    response = requests.get(url,params=query).text # Получение результатов запроса
    data = json.loads(response) # Формирование json ответа
    # "Выгрузка" данных в полученом ответе и подсчет информативности данных каждого результата для сортировки
    val = data['response']
    val = val['items']
    for i in val:
        score = 13

        # Основные данные
        photo = i['photo_100']
        uid = str(i['id'])
        FI = i['last_name']+' '+i['first_name']

        # Дополнительные данные
        info = ''
        
        if 'bdate' in i:
            info += "</br>Дата рождения: "+i['bdate']
            score -= 1
            
        if 'city' in i:
            city=i['city']
            info += ("</br>Город: "+city['title'])
            score -= 1
            
        if ('mobile_phone' in i) and (i['mobile_phone'] != ''):
            info += ("</br>Телефон: "+i['mobile_phone'])
            
        if 'site' in i and i['site'] != '':
            info += ("</br>Сайт: "+i['site']) 
            score -= 1
            
        if 'university_name' in i and 'university_name'!='':
            info += ("</br>Университет: "+i['university_name'])
            score -= 2
            
        if 'faculty_name' in i and 'faculty_name' !='':
            info += ("</br>  "+i['faculty_name'])
            score -= 2
            
        if 'education_status' in i and 'education_status' != '':
            info += ("</br>  "+i['education_status']) 
            score -= 2
            
        if 'career' in i and i['career'] != []:
            score -= 4
            comp = i['career']
            for j in comp:
                if 'company' in j and 'company'!='':
                    info += ("</br>Компания: "+j['company']) 
                if 'position' in j and j['position'] != '':
                    info += (" Должность: "+j['position'])
        
        l.append([score,photo,uid,FI,info]) # Добавление результатов в список
        
    l = sorted(l, key=lambda e: e[0]) # Сортировка списка по подробности полученных данных
    return l

# Формирование отчета в html
def make_report(FI,info,l):
    fhtml.write("""
                </br>
                <table border='1'> 
                <caption><b>"""+FI+"</b></br>"+info+"""</caption>
                """)    
    for i in l:
            fhtml.write("""
                        <tr><td><img src="""+i[1]+""" alt="-"></td><td>
                        <a href ="https://vk.com/id"""+i[2]+"""">"""+i[3]+"""<a>""")
            fhtml.write(i[4]+"</td></tr>")
            
    fhtml.write("</table>")
    return 0

# открыие excel файла с исходными данными
f = xlrd.open_workbook(path,formatting_info=True)
sheet = f.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]

f=list()
i=1

# Полученних данных с определенных ячеек файла excel в список l
while i < len(vals):
    f.append([ vals[i][1], vals[i][8] , vals[i][11] ])
    i+=1
    
# Начало html файла    
fhtml = codecs.open( "report.html", "w", "utf_8_sig" )
fhtml.write("""
                <!DOCTYPE HTML>
                <html>
                <head>
                <meta charset="utf-8">
                <title>Поиск по запросу</title>
                </head>
                <body>
            """)

# Для каждого человека в поиске формировать результаты
for a in f:
    a[0] = a[0].split(' ')
    FI = a[0][1]+' '+a[0][0]  # Потому что так заработало
    city = a[1][a[1].find(' ')+1:a[1].find(',')]
    info = a[2]
    l = make_list(f_query(FI,city)) # Создание списка поиска отдельного человека по городу
    l.extend( make_list(s_query(FI)) ) # Добавление результатов поиска без привязки к городу
    make_report(FI,info,l) # Формирование итогового списка результатов в html

# Закрытие html файла
fhtml.write("""
            </body>
            </html>          
            """)

fhtml.close()