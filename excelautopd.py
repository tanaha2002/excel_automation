import pandas as pd
import math
import argparse

def get_projects_name(sheet):
    projects_name = []
    for stt,(i,row) in enumerate(sheet.iterrows(),start=1):
        if stt < 20:
            continue
        if str(row[2]).isdigit():
            continue
        
        if row[0] == "TOTAL PAYMENT":
            break
        
        projects_name.append(str(row[2]).strip())
    
    while 'nan' in projects_name:
        projects_name.remove('nan')
    
    return projects_name

def get_start_end(projects_name, sheet):
    range_project = {i:{
        'start': 0,
        'end':0
        } for i in projects_name}
    end_end = 0
    project_name=None
    project_name_cur=None
    for stt,(i,row) in enumerate(sheet.iterrows(),start=1):
        if row[0] == "TOTAL PAYMENT":
            end_end = stt - 1
            break
        if str(row[2]).strip() in projects_name:
            # print(row[2])
            project_name_cur = str(row[2]).strip()
            if project_name_cur != project_name:
                range_project[project_name_cur]['start'] = stt + 1
                if project_name is not None:
                    range_project[project_name]['end'] = stt-1
                project_name = project_name_cur
    range_project[project_name_cur]['end'] = end_end
    return range_project

def get_data(path_workbook,sheet_name):
    sheet=pd.read_excel(path_workbook,sheet_name=sheet_name, header=None)
    projects_name= get_projects_name(sheet)
    range_project = get_start_end(projects_name, sheet)
    dict_data = {
        "Project":[],
        "Supplier":[],
        "Description":[],
        'Total':[],
        'Payment': [],
        'Remain':[],
        "By cash":[],
        "TP-Thu":[],
        "SHB-VND":[],
        "SHB-USD":[],
        "TP-Design":[],
        "Woori":[],
        "MrPark":[],
        "Remark":[]
    }
    for project in range_project:
        start = range_project[project]['start']
        end = range_project[project]['end']
        for stt,(i,row) in enumerate(sheet.iterrows(),start=1):
            # print(len(row))
            if start<=stt<=end:
                try:
                    if math.isnan(float(row[3])):
                        continue
                except:
                    pass
                dict_data['Project'].append(project)
                dict_data['Supplier'].append(row[3])
                dict_data['Description'].append(row[4])
                dict_data['Total'].append(row[7])
                dict_data['Payment'].append(row[8])
                dict_data['Remain'].append(row[9])
                dict_data['By cash'].append(row[10])
                dict_data['TP-Thu'].append(row[11])
                dict_data['SHB-VND'].append(row[12])
                dict_data['SHB-USD'].append(row[13])
                dict_data['TP-Design'].append(row[14])
                dict_data['Woori'].append(row[15])
                dict_data['MrPark'].append(row[16])
                dict_data['Remark'].append(row[17])

    df = pd.DataFrame(dict_data)
    df.fillna(0,inplace=True)
    return df

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument('--workbook_path', type=str, help='Path to excel file', required=True)
    parser.add_argument('--sheet_name', type=str, help='Sheet name in excel file', required=True)
    args = parser.parse_args()
    workbook_path = args.workbook_path
    sheet_name = args.sheet_name
    df = get_data(workbook_path,sheet_name)
    save_path = workbook_path.split('.')[0] + '.csv'
    df.to_csv(save_path,index=False)
    print("Done!")