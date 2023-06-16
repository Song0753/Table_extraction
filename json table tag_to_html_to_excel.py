import os
import json
from html2excel import ExcelParser
import pandas as pd
import itertools
from nltk.tag import pos_tag
from nltk.corpus import stopwords
import nltk
from nltk import FreqDist
from nltk.tokenize import RegexpTokenizer
import pandas as pd

class JsontableToExcel:
    def __init__(self, folder_path):
        self.folder_path = folder_path

    # json 파일의 table tag를 가져와서 딕셔너리로 만들기
    def TableDict(self, folder_path, file_list):
        data = {}
        for file_name in file_list:
            if file_name.endswith('.json'):
                file_path = os.path.join(folder_path, file_name)

                with open(file_path, "r", encoding='utf-8') as json_file: #json 파일 열기
                    json_data = json.load(json_file)
                    table_list = list(json_data['tables'].keys())
                    for i in range(len(json_data['tables'])):
                        table_list[i] = {str(table_list[i]):json_data['tables']['tbl0'+str(i+1)]['tag']}
                    if len(json_data['tables']) == 1:
                        data[file_name[:-5]] = table_list[0]
                    else:
                        for j in range(len(json_data['tables'])-1):
                            data[file_name[:-5]] = table_list[0]
                            data[file_name[:-5]].update(table_list[j+1])
        return data 
    
    # dictionary로 만든 tag를 html로 변환해서 폴더 안에 저장하는 함수
    def TableDictToHTML(self, html_folder_path, file_list, table_dict): # file_list = json 파일 리스트, table_dict = TableDict 함수로 만든 dictionary
        for file_name in file_list:
            for i in range(len(table_dict[str(file_name[:-5])])):
                html_file = open(html_folder_path+'/'+str(file_name[:-5])+'_tbl0'+str(i+1)+'.html', 'w', encoding='utf-8') # html 파일 만들기
                html_text = """<!DOCTYPE html><meta charset="UTF-8"><html><head><meta charset="UTF-8"><title>table</title></head><body>"""+str(table_dict[str(file_name[:-5])]['tbl0'+str(i+1)])+"""</body></html>"""
                html_file.write(html_text)
                html_file.close()    

    # 폴더 내의 HTML 파일을 엑셀로 변환하는 함수
    def HtmlToExcel(self, excel_folder_path):
        html_folder_path = 'D:/SHY/KIST/1.Research/OER project/code/table html'
        file_list = os.listdir(html_folder_path)
        print(file_list)
        for file_name in file_list:
            input_file = 'D:/SHY/KIST/1.Research/OER project/code/table html/'+str(file_name)
            output_file = str(excel_folder_path)+'/'+str(file_name[:-5])+'.xlsx'
            parser = ExcelParser(input_file)
            parser.to_excel(output_file)

# 실행 코드
# folder_path = 'D:/SHY/KIST/1.Research/OER project/table json'
# html_folder_path = 'D:/SHY/KIST/1.Research/OER project/20230515_HTML_test'
# excel_folder_path = 'D:/SHY/KIST/1.Research/OER project/20230515_htmltoexcel_test'
# file_list = os.listdir(folder_path)
# test1 = JsontableToExcel(folder_path)
# tag_dict = test1.TableDict(folder_path=folder_path, file_list=file_list)
# dict_html = test1.TableDictToHTML(html_folder_path=html_folder_path, file_list=file_list, table_dict = tag_dict)
# html_excel = test1.HtmlToExcel(excel_folder_path=excel_folder_path)





excel_folder_path = 'D:/SHY/KIST/1.Research/OER project/code/table excel'
excel_file_list = os.listdir(excel_folder_path)
class TableTitlePlusHead:

# Table head 가져와서 리스트 만들기
    def TableHeadList(self, excel_folder_path):
        total_column_list = []
        for file_name in excel_file_list:
            if file_name.endswith('.xlsx'):
                file_path = os.path.join(excel_folder_path, file_name)

                df = pd.read_excel(io=file_path)
                column_name = ' '.join(list(df.columns))
                total_column_list.append(column_name)
        return(total_column_list)
# print(len(total_column_list))
# print(total_column_list)

    def TableTitleList(self, json_folder_path):
    
        total_table_title = []
        for file_name in json_file_list:
            if file_name.endswith('.json'):
                file_path = os.path.join(json_folder_path, file_name)

                with open(file_path, "r", encoding='utf-8') as json_file: #json 파일 열기
                    json_data = json.load(json_file)
                table_title = []
                for i in range(len(json_data['tables'])):
                    table_title.append(json_data['tables']['tbl0'+str(i+1)]['title'])
                total_table_title.append(table_title)     
        flattened_total = list(itertools.chain(*total_table_title))
        return(flattened_total)

json_folder_path = 'D:/SHY/KIST/1.Research/OER project/table json'
json_file_list = os.listdir(json_folder_path)


# table title만 가져와서 리스트 만들기
# def TableTitleList(json_folder_path):
    
#     total_table_title = []
#     for file_name in json_file_list:
#         if file_name.endswith('.json'):
#             file_path = os.path.join(json_folder_path, file_name)

#             with open(file_path, "r", encoding='utf-8') as json_file: #json 파일 열기
#                 json_data = json.load(json_file)
#             table_title = []
#             for i in range(len(json_data['tables'])):
#                 table_title.append(json_data['tables']['tbl0'+str(i+1)]['title'])
#             total_table_title.append(table_title)     
#     flattened_total = list(itertools.chain(*total_table_title))
#     return(flattened_total)

# 한 논문 안의 table list를 품사 태그 하는 함수
def POS_tag(joined_table):
    retokenize = RegexpTokenizer("[\w]+")
    sum_pos = pos_tag(retokenize.tokenize(joined_table))
    return sum_pos


a=TableTitlePlusHead()
table_title_list = a.TableTitleList(json_folder_path=json_folder_path)
table_head_list = a.TableHeadList(excel_folder_path=excel_folder_path)
# table_title_list = TableTitleList(json_folder_path=json_folder_path)
# table_head_list = TableHeadList(excel_folder_path=excel_folder_path)

title_plus_head=[]
for i in range(len(table_title_list)):
    titlePlusHead = table_title_list[i]+' '+table_head_list[i]
    title_plus_head.append(titlePlusHead)

# print(title_plus_head[0])
# print(len(title_plus_head))

sum_table = ' '.join(title_plus_head)
# print(sum_table)

stopwords = ['table', 'Table', 'Unnamed']

# 품사 tagging하기
POS1 = POS_tag(sum_table)
# print(POS1)

noun_list = [t[0] for t in POS1 if t[1] == "NNP" and t[0] not in stopwords] # 명사 태그만 추출
fd_nouns = FreqDist(noun_list) # 명사 태그 빈도수 추출
total_noun = fd_nouns.most_common(len(fd_nouns))




df = pd.DataFrame(total_noun)

df.to_excel('tokenized_noun.xlsx', index=False, header=False)

keywords = ['Tafel', 'Comparison', 'comparison', 'Catalyst', 'catalyst', 'Overpotential', 'overpotential', 'Sample', 'sample', 'η', 'Summary',
           'summary', 'ECSA', 'Samples', 'samples', 'TOF', 'η10', 'Catalysts', 'catalysts', 'Onset', 'onset', 'Slope', 'slope', 'Performance',
           'performance', 'Stability', 'stability', 'Eonset', 'Overpotentials', 'overpotentials', 'η100', 'ηOER', 'Slopes', 'slopes', 
           'η50', 'η20', 'Performances', 'performances', 'Catalystsa', 'ηonset', 'TOFs', 'Vtotal', 'TOFη', 'overpotentiala', 'Turnover',
           'turnover', 'TOFa', 'η10mA', 'Aecsa']

Fkeywords = ['calculated', 'Calculated', 'Computed', 'computed', 'XRF', 'XRD', 'EXAFS', 'HER']



performance_table_index=[]
for i in title_plus_head:
    if any(keyword in i for keyword in keywords):
        performance_table_index.append(title_plus_head.index(i))
print(len(performance_table_index))

for i in title_plus_head:
    if any(keyword in i for keyword in Fkeywords):
        if title_plus_head.index(i) in performance_table_index:
            performance_table_index.remove(title_plus_head.index(i))
print(len(performance_table_index))

def TableNumberList():
    folder_path = 'D:/SHY/KIST/1.Research/OER project/table json'
    file_list = os.listdir(folder_path)
    test1 = JsontableToExcel(folder_path)
    tag_dict = test1.TableDict(folder_path=folder_path, file_list=file_list)
    table_number_list = []
    for file_name in file_list:
        for i in range(len(tag_dict[str(file_name[:-5])])):
            table_number_list.append(str(file_name[:-5])+'_tbl0'+str(i+1))
    return(table_number_list)

table_number_l = TableNumberList()
table_df = pd.DataFrame({'table number': table_number_l, 'table title': table_title_list, 'table column name': table_head_list})
print(table_df)