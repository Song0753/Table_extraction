import pandas as pd
import os
import openai
import pickle

# 폴더 경로 설정
folder_path = 'D:/SHY/KIST/1.Research/OER project/Table 추출/data/test_5_excel'  # 실제 폴더 경로로 변경해주세요

# table excel 데이터를 dataframe으로 변환해서 리스트에 저장
file_list = os.listdir(folder_path)
dfs = []
for file_name in file_list:
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        df = pd.read_excel(file_path)
        dfs.append(df)

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
        prompt="I want to extract the necessary parameters from the table and remove the rest. The parameters I need are as follows: catalyst name, overpotential, Tafel slope, TOF, mass activity, specific activity, ECSA, Cdl, stability, electrolyte, and ref.\n\nPlease perform the following two tasks:\n1. Check if the given table contains the mentioned parameters.\n2. If they exist, keep the corresponding columns and remove the others.\n\n\nI'll provide a few examples.\nIf I had a table like this.\n\tcatalyst\tη at 10 mA cm–2 (mV)\tη at 50 mA cm–2 (mV)\tTafel slope (mV decade–1)\tref\n0\tFe-NiS-(BO)/NF\t–\t250\t54.5\tthis work\n1\tNiS-(BO)/NF\t–\t290\t65.9\tthis work\n2\tNiS/NF\t–\t368\t88.0\tthis work\n3\tIrO2/NF\t–\t317\t63.2\tthis work\n\nAll columns are relevant, so they should be kept as they are.\n->\n\tcatalyst\tη at 10 mA cm–2 (mV)\tη at 50 mA cm–2 (mV)\tTafel slope (mV decade–1)\tref\n0\tFe-NiS-(BO)/NF\t–\t250\t54.5\tthis work\n1\tNiS-(BO)/NF\t–\t290\t65.9\tthis work\n2\tNiS/NF\t–\t368\t88.0\tthis work\n3\tIrO2/NF\t–\t317\t63.2\tthis work\n\nIf I had a table like this.\n\tsample\tη50 (mV)\tη100 (mV)\tTafel slope (mV dec–1)\tRct (ohm)\tiex (mA cm–2)\tTOF (s–1)\tECSA (cm2)\troughness factor\tOER stability (h)\n0\tFeCoP (FCP350)\t230\t260\t54\t0.5\t142.66\t7.43\t37.0\t411.11\t60\n\n\nThe table should be revised as follows:\nThe Rct and iex parameters should be removed from the table as it is not the data I desire.\n->\n\tsample\tη50 (mV)\tη100 (mV)\tTafel slope (mV dec–1)\tTOF (s–1)\tECSA (cm2)\tOER stability (h)\n0\tFeCoP (FCP350)\t230\t260\t54\t7.43\t37.0\t60\n\nIf I had a table like this.\ncatalyst\tη (mV) at 10 mA cm–2\tTafel slope (mV/decade)\tCdl(μF cm–2)\tECSA (cm–2)\n0\tNi@NBC-800\t310\t77\t12\t105.0\n1\tNi@NBC-700\t330\t136\t14\t48.5\n2\tNi@NBC-600\t600\t140\t112\t18.3\n\nAll columns are relevant, so they should be kept as they are.\n->\ncatalyst\tη (mV) at 10 mA cm–2\tTafel slope (mV/decade)\tCdl(μF cm–2)\tECSA (cm–2)\n0\tNi@NBC-800\t310\t77\t12\t105.0\n1\tNi@NBC-700\t330\t136\t14\t48.5\n2\tNi@NBC-600\t600\t140\t112\t18.3\n\nIf I had a table like this.\nCatalyst\tMass loading (µg cm−2)\tη at J = 10 mA cm−2 (mV)\tTafel slope (mV dec−1)\tMass activity at η = 0.4 V (A g−1)\tTOF at η = 0.4 V (s−1)\tRef.\n0\tCo3O4/Fe3O4 (4:1) composites\t500\t322\t78.0\t34.4\t0.0215\tOur work\n1\tCo3O4\t500\t353\t84.0\t21.5\t0.0314\tOur work\n2\tCo3O4-MNTs\t531\t353\t80.0\t—\t—\t28\n3\tCo3O4@CoO SC\t25\t430\t89.0\t234.0\t0.0487\t29\n4\tCo0.54Fe0.46OOH\t—\t390\t47.0\t200 at η = 0.39 V\t0.0225\t30\n5\tCoOx-4-h nanoplates\t500\t306\t61.0\t—\t0.0030\t31\n6\tFe1Co1-ONS\t360\t308\t36.8\t54.9 at η = 0.35 V\t0.0214 at η = 0.35 V\t32\n\nThe table should be revised as follows:\nThe mass loading parameter should be removed from the table as it is not the data I desire.\n->\nCatalyst\tη at J = 10 mA cm−2 (mV)\tTafel slope (mV dec−1)\tMass activity at η = 0.4 V (A g−1)\tTOF at η = 0.4 V (s−1)\tRef.\n0\tCo3O4/Fe3O4 (4:1) composites\t322\t78.0\t34.4\t0.0215\tOur work\n1\tCo3O4\t353\t84.0\t21.5\t0.0314\tOur work\n2\tCo3O4-MNTs\t353\t80.0\t—\t—\t28\n3\tCo3O4@CoO SC\t430\t89.0\t234.0\t0.0487\t29\n4\tCo0.54Fe0.46OOH\t390\t47.0\t200\t0.0225\t30\n5\tCoOx-4-h nanoplates\t306\t61.0\t—\t0.0030\t31\n6\tFe1Co1-ONS\t 308\t36.8\t54.9\t0.0214\t32\n\n\nIf I had a table like this.\nAnode\tPotential (V vs. RHE) required to generate 20 mA cm−2\tOnset potential (V vs. RHE)\tMass activity (A g−1) at 1.65 V vs. RHE\tTafel slope (mV per decade)\tSpecific activity (mA cmECSA−2)\n0\tCo3O4\t1.724\t1.581\t16.12\t93\t1.045\n1\tCo3O4@Co3S4-1\t1.658\t1.369\t44.30\t—\t—\n2\tCo3O4@Co3S4-2\t1.647\t1.510\t50.02\t80\t1.156\n3\tRuO2\t1.723\t1.493\t—\t—\t—\n\nThe table should be revised as follows:\nThe Potential (V vs. RHE) required to generate 20 mA cm−2 data can be converted to overpotential, so please keep it and only remove the onset potential data.\n->\nAnode\tPotential (V vs. RHE) required to generate 20 mA cm−2\tMass activity (A g−1) at 1.65 V vs. RHE\tTafel slope (mV per decade)\tSpecific activity (mA cmECSA−2)\n0\tCo3O4\t1.724\t16.12\t93\t1.045\n1\tCo3O4@Co3S4-1\t1.658\t44.30\t—\t—\n2\tCo3O4@Co3S4-2\t1.647\t50.02\t80\t1.156\n3\tRuO2\t1.723\t—\t—\t—\n\nIf I had a table like this.\nelement\tNiS (atom %)\tNiS-(BO) (atom %)\tFe-NiS-(BO) (atom %)\n0\tNi\t62.9\t64.5\t55.9\n1\tS\t37.1\t25.8\t24.0\n2\tB\t–\t9.7\t9.5\n3\tNaN\t–\t(B/S: 37.6)\t(B/S: 39.6)\n4\tFe\t–\t–\t10.6\n5\tNaN\t–\t–\t(Fe/Ni: 18.9)\n\n->\n-1\n\nIf I had a table like this.\nCatalyst\tTafel slope (mV·dec−1)\tExchange current density, jo (mA·cm−2)\tη@0.1 A·cm−2 (V)\tSpecific activity @1.58 V (mA‧cm−2ECSA)\tECSA (cm2)\tRct (Ω‧cm2)\tRmt (Ω‧cm2)\n0\tNi3S2@NF-0 mL\t124\t1.573\t319\t3.198\t17.4\t1.132\t2.928\n1\tNi3S2@NF-1 mL\t86\t1.710\t281\t3.695\t16.8\t1.135\t1.653\n2\tNi3S2@NF-2 mL\t72\t2.582\t263\t9.428\t23.0\t0.908\t1.192\n3\tNi3S2@NF-3 mL\t115\t1.888\t294\t2.703\t18.6\t1.128\t1.288\n\nThe table should be revised as follows:\nThe Exchange current density, jo (mA·cm−2) data and Rct (Ω‧cm2) data should be removed from the table as it is not the data I desire.\n->\nCatalyst\tTafel slope (mV·dec−1)\tη@0.1 A·cm−2 (V)\tSpecific activity @1.58 V (mA‧cm−2ECSA)\tECSA (cm2)\n0\tNi3S2@NF-0 mL\t124\t319\t3.198\n1\tNi3S2@NF-1 mL\t86\t281\t3.695\t16.8\n2\tNi3S2@NF-2 mL\t72\t263\t9.428\t23.0\n3\tNi3S2@NF-3 mL\t115\t294\t2.703\t18.6\n\nCan you revise below table like above examples?\n\nRemember that the data I desire includes only catalyst name, overpotential, Tafel slope, TOF, mass activity, specific activity, ECSA, Cdl, stability, electrolyte, and ref. The rest should be deleted from the table.\nIt should be made solely from the data within the table.\nIf the given table does not contain any of the necessary parameters, print -1.\nHere is an table:"+i+ "\nPlease show me just revised table.\n->",
        temperature=0,
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
            messages=[{"role": "user", "content": "I want to extract the necessary parameters from the table and remove the rest. The parameters I need are as follows: catalyst name, overpotential, Tafel slope, TOF, mass activity, specific activity, ECSA, Cdl, stability, electrolyte, and ref.\n\nPlease perform the following two tasks:\n1. Check if the given table contains the mentioned parameters.\n2. If they exist, keep the corresponding columns and remove the others.\n\n\nI'll provide a few examples.\nIf I had a table like this.\n\tcatalyst\tη at 10 mA cm–2 (mV)\tη at 50 mA cm–2 (mV)\tTafel slope (mV decade–1)\tref\n0\tFe-NiS-(BO)/NF\t–\t250\t54.5\tthis work\n1\tNiS-(BO)/NF\t–\t290\t65.9\tthis work\n2\tNiS/NF\t–\t368\t88.0\tthis work\n3\tIrO2/NF\t–\t317\t63.2\tthis work\n\nAll columns are relevant, so they should be kept as they are.\n->\n\tcatalyst\tη at 10 mA cm–2 (mV)\tη at 50 mA cm–2 (mV)\tTafel slope (mV decade–1)\tref\n0\tFe-NiS-(BO)/NF\t–\t250\t54.5\tthis work\n1\tNiS-(BO)/NF\t–\t290\t65.9\tthis work\n2\tNiS/NF\t–\t368\t88.0\tthis work\n3\tIrO2/NF\t–\t317\t63.2\tthis work\n\nIf I had a table like this.\n\tsample\tη50 (mV)\tη100 (mV)\tTafel slope (mV dec–1)\tRct (ohm)\tiex (mA cm–2)\tTOF (s–1)\tECSA (cm2)\troughness factor\tOER stability (h)\n0\tFeCoP (FCP350)\t230\t260\t54\t0.5\t142.66\t7.43\t37.0\t411.11\t60\n\n\nThe table should be revised as follows:\nThe Rct and iex parameters should be removed from the table as it is not the data I desire.\n->\n\tsample\tη50 (mV)\tη100 (mV)\tTafel slope (mV dec–1)\tTOF (s–1)\tECSA (cm2)\tOER stability (h)\n0\tFeCoP (FCP350)\t230\t260\t54\t7.43\t37.0\t60\n\nIf I had a table like this.\ncatalyst\tη (mV) at 10 mA cm–2\tTafel slope (mV/decade)\tCdl(μF cm–2)\tECSA (cm–2)\n0\tNi@NBC-800\t310\t77\t12\t105.0\n1\tNi@NBC-700\t330\t136\t14\t48.5\n2\tNi@NBC-600\t600\t140\t112\t18.3\n\nAll columns are relevant, so they should be kept as they are.\n->\ncatalyst\tη (mV) at 10 mA cm–2\tTafel slope (mV/decade)\tCdl(μF cm–2)\tECSA (cm–2)\n0\tNi@NBC-800\t310\t77\t12\t105.0\n1\tNi@NBC-700\t330\t136\t14\t48.5\n2\tNi@NBC-600\t600\t140\t112\t18.3\n\nIf I had a table like this.\nCatalyst\tMass loading (µg cm−2)\tη at J = 10 mA cm−2 (mV)\tTafel slope (mV dec−1)\tMass activity at η = 0.4 V (A g−1)\tTOF at η = 0.4 V (s−1)\tRef.\n0\tCo3O4/Fe3O4 (4:1) composites\t500\t322\t78.0\t34.4\t0.0215\tOur work\n1\tCo3O4\t500\t353\t84.0\t21.5\t0.0314\tOur work\n2\tCo3O4-MNTs\t531\t353\t80.0\t—\t—\t28\n3\tCo3O4@CoO SC\t25\t430\t89.0\t234.0\t0.0487\t29\n4\tCo0.54Fe0.46OOH\t—\t390\t47.0\t200 at η = 0.39 V\t0.0225\t30\n5\tCoOx-4-h nanoplates\t500\t306\t61.0\t—\t0.0030\t31\n6\tFe1Co1-ONS\t360\t308\t36.8\t54.9 at η = 0.35 V\t0.0214 at η = 0.35 V\t32\n\nThe table should be revised as follows:\nThe mass loading parameter should be removed from the table as it is not the data I desire.\n->\nCatalyst\tη at J = 10 mA cm−2 (mV)\tTafel slope (mV dec−1)\tMass activity at η = 0.4 V (A g−1)\tTOF at η = 0.4 V (s−1)\tRef.\n0\tCo3O4/Fe3O4 (4:1) composites\t322\t78.0\t34.4\t0.0215\tOur work\n1\tCo3O4\t353\t84.0\t21.5\t0.0314\tOur work\n2\tCo3O4-MNTs\t353\t80.0\t—\t—\t28\n3\tCo3O4@CoO SC\t430\t89.0\t234.0\t0.0487\t29\n4\tCo0.54Fe0.46OOH\t390\t47.0\t200\t0.0225\t30\n5\tCoOx-4-h nanoplates\t306\t61.0\t—\t0.0030\t31\n6\tFe1Co1-ONS\t 308\t36.8\t54.9\t0.0214\t32\n\n\nIf I had a table like this.\nAnode\tPotential (V vs. RHE) required to generate 20 mA cm−2\tOnset potential (V vs. RHE)\tMass activity (A g−1) at 1.65 V vs. RHE\tTafel slope (mV per decade)\tSpecific activity (mA cmECSA−2)\n0\tCo3O4\t1.724\t1.581\t16.12\t93\t1.045\n1\tCo3O4@Co3S4-1\t1.658\t1.369\t44.30\t—\t—\n2\tCo3O4@Co3S4-2\t1.647\t1.510\t50.02\t80\t1.156\n3\tRuO2\t1.723\t1.493\t—\t—\t—\n\nThe table should be revised as follows:\nThe Potential (V vs. RHE) required to generate 20 mA cm−2 data can be converted to overpotential, so please keep it and only remove the onset potential data.\n->\nAnode\tPotential (V vs. RHE) required to generate 20 mA cm−2\tMass activity (A g−1) at 1.65 V vs. RHE\tTafel slope (mV per decade)\tSpecific activity (mA cmECSA−2)\n0\tCo3O4\t1.724\t16.12\t93\t1.045\n1\tCo3O4@Co3S4-1\t1.658\t44.30\t—\t—\n2\tCo3O4@Co3S4-2\t1.647\t50.02\t80\t1.156\n3\tRuO2\t1.723\t—\t—\t—\n\nIf I had a table like this.\nelement\tNiS (atom %)\tNiS-(BO) (atom %)\tFe-NiS-(BO) (atom %)\n0\tNi\t62.9\t64.5\t55.9\n1\tS\t37.1\t25.8\t24.0\n2\tB\t–\t9.7\t9.5\n3\tNaN\t–\t(B/S: 37.6)\t(B/S: 39.6)\n4\tFe\t–\t–\t10.6\n5\tNaN\t–\t–\t(Fe/Ni: 18.9)\n\n->\n-1\n\nIf I had a table like this.\nCatalyst\tTafel slope (mV·dec−1)\tExchange current density, jo (mA·cm−2)\tη@0.1 A·cm−2 (V)\tSpecific activity @1.58 V (mA‧cm−2ECSA)\tECSA (cm2)\tRct (Ω‧cm2)\tRmt (Ω‧cm2)\n0\tNi3S2@NF-0 mL\t124\t1.573\t319\t3.198\t17.4\t1.132\t2.928\n1\tNi3S2@NF-1 mL\t86\t1.710\t281\t3.695\t16.8\t1.135\t1.653\n2\tNi3S2@NF-2 mL\t72\t2.582\t263\t9.428\t23.0\t0.908\t1.192\n3\tNi3S2@NF-3 mL\t115\t1.888\t294\t2.703\t18.6\t1.128\t1.288\n\nThe table should be revised as follows:\nThe Exchange current density, jo (mA·cm−2) data and Rct (Ω‧cm2) data should be removed from the table as it is not the data I desire.\n->\nCatalyst\tTafel slope (mV·dec−1)\tη@0.1 A·cm−2 (V)\tSpecific activity @1.58 V (mA‧cm−2ECSA)\tECSA (cm2)\n0\tNi3S2@NF-0 mL\t124\t319\t3.198\n1\tNi3S2@NF-1 mL\t86\t281\t3.695\t16.8\n2\tNi3S2@NF-2 mL\t72\t263\t9.428\t23.0\n3\tNi3S2@NF-3 mL\t115\t294\t2.703\t18.6\n\nCan you revise below table like above examples?\n\nRemember that the data I desire includes only catalyst name, overpotential, Tafel slope, TOF, mass activity, specific activity, ECSA, Cdl, stability, electrolyte, and ref. The rest should be deleted from the table.\nIt should be made solely from the data within the table.\nIf the given table does not contain any of the necessary parameters, print -1.\nHere is an table:"+i+ "\nPlease show me just revised table.\n->"}]
        )
        completion = response['choices'][0]['message']['content']
        total_completion_chat.append(completion)
    return(total_completion_chat)


# 결과 pickle로 저장
def get_files_without_extension(folder_path):
    file_list = []
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            file_name_without_ext = os.path.splitext(file_name)[0]
            file_list.append(file_name_without_ext)
    return file_list

# 폴더 경로를 지정합니다.
folder_path = "D:/SHY/KIST/1.Research/OER project/Table 추출/data/test_5_excel"

# 파일 이름 리스트를 가져옵니다.
file_list = get_files_without_extension(folder_path)


pkl_folder_path = 'D:/SHY/KIST/1.Research/OER project/Table 추출/data/step1_complete'
def result_to_pickel(file_list, total_completion_chat):
    dictionaries = {key: value for key, value in zip(file_list, total_completion_chat)}
    for key in file_list:
        dictionary = {key: dictionaries[key]}
        file_name = key + '_step1.pkl'
        with open(pkl_folder_path+'/'+key + '_step1.pkl', 'wb') as file:
            pickle.dump(dictionary, file)