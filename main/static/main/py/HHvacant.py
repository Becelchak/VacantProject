import datetime
import requests

def get_HH_vacants():
    needDay = datetime.datetime.now().day
    nowMonth = datetime.datetime.now().month
    nowYear = datetime.datetime.now().year

    if nowMonth < 10:
        nowMonth = '0' + str(nowMonth)

    per_page = 100
    hour = '00'
    min_sec = '00'
    answer = []
    for part_day in range(4):
        date_from = "{0}-{1}-{2}T{3}:00:00".format(nowYear, nowMonth, needDay, hour)
        hour = (part_day + 1) * 6
        if hour < 10:
            hour = '0' + str(hour)
        elif hour == 24:
            hour = 23
            min_sec = 59
        date_to = "{0}-{1}-{2}T{3}:{4}:{4}".format(nowYear, nowMonth, needDay, hour, min_sec)
        for page in range(20):
            if page == 20:
                per_page = 99
            url_HH = "https://api.hh.ru/vacancies?specialization=1&per_page={0}&page={1}&date_from={2}&date_to={3}&professional_role=96".format(per_page,page,date_from,date_to)
            res = requests.get(url_HH).json()
            for vac in res['items']:
                dict_temp = {}
                if (vac['name'].__contains__("нженер") and vac['name'].__contains__("рограммист")):
                    x = requests.get('https://api.hh.ru/vacancies/{0}'.format(vac['id'])).json()
                    dict_temp['name'] = vac['name']
                    try:
                        dict_temp['salary_from'] = vac['salary']['from']
                        dict_temp['salary_to'] = vac['salary']['to']
                        dict_temp['salary_currency'] = vac['salary']['currency']
                        skill_list = []
                        for skill in x['key_skills']:
                            skill_list.append(skill['name'])
                        dict_temp['skills'] = ','.join(skill_list)
                        # if len(x['description']) > 100:
                        #     x['description'] = x['description'][0:101] + '...'
                        dict_temp['description'] = x['description']
                    except:
                        dict_temp['salary_from'] = "0"
                        dict_temp['salary_to'] = "0"
                        dict_temp['salary_currency'] = ""
                    dict_temp['area_name'] = vac['area']['name']
                    dict_temp['employer_name'] = vac['employer']['name']
                    dict_temp['published_at'] = vac['published_at']
                    if len(answer) < 10:
                        answer.append(dict_temp)
                    else:
                        break
    return answer


f = get_HH_vacants()
#
# for item in f.items():
#     a = item[0]
#     b = item[1]