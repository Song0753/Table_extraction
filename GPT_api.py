import pandas as pd
import glob
from tabledataextractor import Table
import os
import openai


# 폴더 경로 설정
folder_path = 'D:/SHY/KIST/1.Research/OER project/Table 추출/data/performance table100'  # 실제 폴더 경로로 변경해주세요

# 폴더 내의 html 파일 가져오기
html_files = glob.glob(folder_path + "/*.html")

# tabledataextractor 사용해서 table 데이터를 dataframe으로 변환해서 리스트에 저장
dfs = []
for file in html_files:
    try:
        table = Table(file)
        df = table.to_pandas()
        dfs.append(df)
    except Exception as e:
        print(f"에러 발생: {file} - {e}")  

print(len(dfs))

# GPT prompt에 넣을 수 있게 dataframe을 string으로 변환
string_dfs=[]
for i in dfs:
    df_str = i.to_string(index=True, header=True, index_names=True)
    string_dfs.append(df_str)

# openai key
My_OpenAI_key = '<open api key>'
openai.api_key = My_OpenAI_key

# GPT3 API를 사용해서 completion 얻고, list에 저장
def UsingGPT3():
    total_completion = []
    for i in string_dfs:
        response = openai.Completion.create(
        model="text-davinci-003",
        prompt="I want to retrieve table data from a reported paper and save it in JSON format.\nThe data I want to extract is electrolyte, overpotential, Tafel slope, TOF, mass activity, specific activity, ECSA, Cdl, and stability.\nIf the mentioned nine items are not present in the table data, you can leave them blank.\nHere is an example of JSON format.\n\n[\n  {\n    \"catalyst\": \"Fe-NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(250, 50)],\n    \"Tafel slope (mVdec-1)\": 44,\n    \"TOF (s-1)\": [(200, 400)], \n    \"mass activity (Ag-1)\": [(300, 1)], \n    \"specific activity (Acm-2)\": [(300, 10)], \n    \"ECSA (cm2)\": 100,\n    \"Cdl (Fm-2)\": 100,\n    \"stability\": 5000, \n    \"Electrolyte\": \"1 M KOH\",\n  },\n  {\n    \"catalyst\": \"Fe-NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(250, 50)],\n    \"Tafel slope (mVdec-1)\": 44,\n    \"TOF (s-1)\": [(200, 400)], \n    \"mass activity (Ag-1)\": [(300, 1)], \n    \"specific activity (Acm-2)\": [(300, 10)], \n    \"ECSA (cm2)\": 100,\n    \"Cdl (Fm-2)\": 100,\n    \"stability\": 5000, \n    \"Electrolyte\": \"1 M KOH\",\n  },\n]\n\nIf I had a table like this.\ncatalyst\tη at 10 mA cm–2 (mV)\tη at 50 mA cm–2 (mV)\tTafel slope (mV decade–1)\tref\n0\tFe-NiS-(BO)/NF\t–\t250\t54.5\tthis work\n1\tNiS-(BO)/NF\t–\t290\t65.9\tthis work\n2\tNiS/NF\t–\t368\t88.0\tthis work\n\nI want to revise the table like this.\n[\n  {\n    \"catalyst name\": \"Fe-NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(250, 50)],\n    \"Tafel slope\": 54.5,\n    \"TOF (s-1)\": \"\", \n    \"mass activity (Ag-1)\": \"\", \n    \"specific activity (Acm-2)\": \"\", \n    \"ECSA (cm2)\": \"\",\n    \"Cdl (Fm-2)\": \"\",\n    \"stability\": \"\", \n    \"Electrolyte\": \"\",\n  },\n  {\n    \"catalyst name\": \"NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(\"–\", 10), (290, 50)],\n    \"Tafel slope\": 65.9,\n    \"TOF (s-1)\": \"\", \n    \"mass activity (Ag-1)\": \"\", \n    \"specific activity (Acm-2)\": \"\", \n    \"ECSA (cm2)\": \"\",\n    \"Cdl (Fm-2)\": \"\",\n    \"stability\": \"\", \n    \"Electrolyte\": \"\",\n  },\n  {\n    \"catalyst name\": \"NiS/NF\",\n    \"overpotential (mv/mA cm–2)\": [( 368, 50)],\n    \"Tafel slope\": 88.0,\n    \"TOF (s-1)\": \"\", \n    \"mass activity (Ag-1)\": \"\", \n    \"specific activity (Acm-2)\": \"\", \n    \"ECSA (cm2)\": \"\",\n    \"Cdl (Fm-2)\": \"\",\n    \"stability\": \"\", \n    \"Electrolyte\": \"\",\n  }\n]\n\nCan you revise the below table like above case?:\n"+i+ 'Please write just json format.',
        temperature=1,
        max_tokens=2000,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
        completion = response['choices'][0]['text']
        total_completion.append(completion)
    return(total_completion)

def UsingChatGPT():
    total_completion_chat = []
    for i in string_dfs:
        response =openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": "I want to retrieve table data from a reported paper and save it in JSON format.\nThe data I want to extract is electrolyte, overpotential, Tafel slope, TOF, mass activity, specific activity, ECSA, Cdl, and stability.\nIf the mentioned nine items are not present in the table data, you can leave them blank.\nHere is an example of JSON format.\n\n[\n  {\n    \"catalyst\": \"Fe-NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(250, 50)],\n    \"Tafel slope (mVdec-1)\": 44,\n    \"TOF (s-1)\": [(200, 400)], \n    \"mass activity (Ag-1)\": [(300, 1)], \n    \"specific activity (Acm-2)\": [(300, 10)], \n    \"ECSA (cm2)\": 100,\n    \"Cdl (Fm-2)\": 100,\n    \"stability\": 5000, \n    \"Electrolyte\": \"1 M KOH\",\n  },\n  {\n    \"catalyst\": \"Fe-NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(250, 50)],\n    \"Tafel slope (mVdec-1)\": 44,\n    \"TOF (s-1)\": [(200, 400)], \n    \"mass activity (Ag-1)\": [(300, 1)], \n    \"specific activity (Acm-2)\": [(300, 10)], \n    \"ECSA (cm2)\": 100,\n    \"Cdl (Fm-2)\": 100,\n    \"stability\": 5000, \n    \"Electrolyte\": \"1 M KOH\",\n  },\n]\n\nIf I had a table like this.\ncatalyst\tη at 10 mA cm–2 (mV)\tη at 50 mA cm–2 (mV)\tTafel slope (mV decade–1)\tref\n0\tFe-NiS-(BO)/NF\t–\t250\t54.5\tthis work\n1\tNiS-(BO)/NF\t–\t290\t65.9\tthis work\n2\tNiS/NF\t–\t368\t88.0\tthis work\n\nI want to revise the table like this.\n[\n  {\n    \"catalyst name\": \"Fe-NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(250, 50)],\n    \"Tafel slope\": 54.5,\n    \"TOF (s-1)\": \"\", \n    \"mass activity (Ag-1)\": \"\", \n    \"specific activity (Acm-2)\": \"\", \n    \"ECSA (cm2)\": \"\",\n    \"Cdl (Fm-2)\": \"\",\n    \"stability\": \"\", \n    \"Electrolyte\": \"\",\n  },\n  {\n    \"catalyst name\": \"NiS-(BO)/NF\",\n    \"overpotential (mv/mA cm–2)\": [(\"–\", 10), (290, 50)],\n    \"Tafel slope\": 65.9,\n    \"TOF (s-1)\": \"\", \n    \"mass activity (Ag-1)\": \"\", \n    \"specific activity (Acm-2)\": \"\", \n    \"ECSA (cm2)\": \"\",\n    \"Cdl (Fm-2)\": \"\",\n    \"stability\": \"\", \n    \"Electrolyte\": \"\",\n  },\n  {\n    \"catalyst name\": \"NiS/NF\",\n    \"overpotential (mv/mA cm–2)\": [( 368, 50)],\n    \"Tafel slope\": 88.0,\n    \"TOF (s-1)\": \"\", \n    \"mass activity (Ag-1)\": \"\", \n    \"specific activity (Acm-2)\": \"\", \n    \"ECSA (cm2)\": \"\",\n    \"Cdl (Fm-2)\": \"\",\n    \"stability\": \"\", \n    \"Electrolyte\": \"\",\n  }\n]\n\nCan you revise the below table like above case?:\n"+i+ 'Please write just json format.'}]
        )
        completion = response['choices'][0]['message']['content']
        total_completion_chat.append(completion)
    return(total_completion_chat)