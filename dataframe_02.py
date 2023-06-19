import os
import pandas as pd
import pickle

class TableExtraction:
    def __init__(self, folder_path):
        self.folder_path = folder_path
    
    def ExcelToPickle(self, folder_path):
        file_list = os.listdir(folder_path)
        for file_name in file_list:
            if file_name.endswith('.xlsx'):
                file_path = os.path.join(folder_path, file_name)
                df = pd.read_excel(file_path)        
                with open(folder_path+'/'+str(file_name[:-5])+'.pkl','wb') as f:
                    pickle.dump(df,f)  

    def change_catname(self, folder_path):
        file_list = os.listdir(folder_path)
        for file_name in file_list:
            if file_name.endswith('.pkl'):
                file_path = os.path.join(folder_path, file_name)
                df = pd.read_pickle(file_path)
                for i in df.columns:
                    if 'catalyst' in str(i):
                        column_to_rename = i
                        new_column_name = 'catalyst name'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Catalyst' in str(i):
                        column_to_rename = i
                        new_column_name = 'catalyst name'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'name' in str(i):
                        column_to_rename = i
                        new_column_name = 'catalyst name'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Name' in str(i):
                        column_to_rename = i
                        new_column_name = 'catalyst name'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'sample' in str(i):
                        column_to_rename = i
                        new_column_name = 'catalyst name'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Sample' in str(i):
                        column_to_rename = i
                        new_column_name = 'catalyst name'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                with open(folder_path+'/'+str(file_name[:-4])+'.pkl','wb') as f:
                    pickle.dump(df,f)
                    
    def change_overpotential(self, folder_path):
        file_list = os.listdir(folder_path)
        for file_name in file_list:
            if file_name.endswith('.pkl'):
                file_path = os.path.join(folder_path, file_name)
                df = pd.read_pickle(file_path)
                for i in df.columns:
                    if 'η' and '10' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 10 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'η' and '50' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 50 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'η' and '100' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 100 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'overpotential' and '10' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 10 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'overpotential' and '50' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 50 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'overpotential' and '100' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 100 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Overpotential' and '10' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 10 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Overpotential' and '50' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 50 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Overpotential' and '100' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 100 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'over-potential' and '10' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 10 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'over-potential' and '50' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 50 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'over-potential' and '100' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 100 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Over-potential' and '10' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 10 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Over-potential' and '50' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 50 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                    if 'Over-potential' and '100' in str(i):
                        column_to_rename = i
                        new_column_name = 'η (mV) at 100 mA·cm–2'
                        df.rename(columns={column_to_rename: new_column_name}, inplace=True)
                with open(folder_path+'/'+str(file_name[:-4])+'.pkl','wb') as f:
                    pickle.dump(df,f)

def dfToDict(folder_path, dict_folder_path):
    file_list = os.listdir(folder_path)
    for file_name in file_list:
        if file_name.endswith('.pkl'):
            file_path = os.path.join(folder_path, file_name)
            df = pd.read_pickle(file_path)  
            result_dict = {}
            for _, row in df.iterrows():
                if 'catalyst name' in df.columns:
                # df에 catalyst name이라는 column이 있다면이라는 조건문 추가
                    catalyst_name = str(row['catalyst name'])
                    result_dict[catalyst_name] = {}
                    if catalyst_name not in result_dict:
                        result_dict[catalyst_name] = {}
                else:
                    continue  # catalyst name 키가 없으면 다음 데이터프레임으로 넘어감

                if 'η (mV) at 10 mA·cm–2' in df.columns:
                    η10 = row['η (mV) at 10 mA·cm–2']
                    result_dict[catalyst_name].update({'η (mV) at 10 mA·cm–2':η10})
                if 'η (mV) at 50 mA·cm–2' in df.columns:
                    η50 = row['η (mV) at 50 mA·cm–2']
                    result_dict[catalyst_name].update({'η (mV) at 50 mA·cm–2':η50})
                if 'η (mV) at 100 mA·cm–2' in df.columns:
                    η100 = row['η (mV) at 100 mA·cm–2']
                    result_dict[catalyst_name].update({'η (mV) at 100 mA·cm–2':η100})
            with open(dict_folder_path+'/'+str(file_name[:-4])+'_.pkl','wb') as f:
                pickle.dump(result_dict,f)