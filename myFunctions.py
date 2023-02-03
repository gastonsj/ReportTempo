import requests
from requests.auth import HTTPBasicAuth
############### INICIO DE FUNCIONES ######################################
def query_team(startDate,endDate,token_tempo):
    # Query para tareas por equipo de trabajo. ID Cloud Delivery = 23.
    url="https://api.tempo.io/4/worklogs/team/23?from="+startDate+"&to="+endDate+"&limit=5000"
    headers = {                                             
        "Authorization": "Bearer "+token_tempo
    }
    response = requests.request(
        "GET",
        url,
        headers=headers
    )
    #print(response.json())
    jsonData = response.json()
    return jsonData

def query_author(id,email_jira,token_jira):
    # Query para obtener autor
    url="https://bghtechpartner.atlassian.net/rest/api/2/user?accountId="+id       
    auth = HTTPBasicAuth(email_jira, token_jira)
    headers = {                                             
        "Accept": "application/json"
    }
    response = requests.request(
        "GET",
        url,
        headers=headers,
        auth=auth
    )
    jsonData = response.json()
    return jsonData['displayName']

def query_issue(id,email_jira,token_jira):
    # Query para obtener autor
    url="https://bghtechpartner.atlassian.net/rest/api/2/issue/"+str(id)
    auth = HTTPBasicAuth(email_jira, token_jira)
    headers = {                                             
        "Accept": "application/json"
    }
    response = requests.request(
        "GET",
        url,
        headers=headers,
        auth=auth
    )
    jsonData = response.json()
    
    return {
        'key':jsonData['fields']['project']['key'],
        'name':jsonData['fields']['project']['name']
    }

def distinctList(lst):
    distinct_list=[]
    for i in lst:
        if i not in distinct_list:
            distinct_list.append(i)
    return distinct_list

def sum_list(lst):
    suma=0
    for i in lst:
        suma+=i
    return suma

def dic_count_hour_empty(n,author):
    dic={}
    zeros=[0 for i in range(n)]
    for i in author:
        dic.update({i:zeros[:]})
    return dic

############3#### FINAL DE FUNCIONES #######################################