from excelautopd import get_data
from logger import logfile
# from configs import xlsx_report, xlsx_data, date_sheets
import datetime
import openpyxl

map_insert = {
    'D{}': 'Supplier',
    'F{}': 'Description',
    'G{}': 'By cash',
    'H{}': 'TP-Thu',
    'I{}': 'SHB-VND',
    'J{}': 'SHB-USD',
    'K{}': 'TP-Design',
    'L{}': 'Woori',
    'M{}': 'MrPark',
    'N{}': 'Remark'
}
fmt_number = '#,##0'
fmt_date = 'dd-mmm'

def remove_blank_line(worksheet):
    rm_idx = []
    for i in range(len(tuple(worksheet.rows))):
        flag = False
        for cell in tuple(worksheet.rows)[i][1:]:
            if cell.value is not None:
                flag = True
                break
        if flag==False:
            rm_idx.append(i)
    for i in range(len(sorted(rm_idx, reverse=True))):
        worksheet.delete_rows(idx=rm_idx[i]+1-i, amount=1)

def insert(worksheet, df, date):
    m_row = worksheet.max_row
    for stt, data in df.iterrows():
        stt = stt+1
        idx = stt+m_row
        worksheet[f'A{idx}'] = m_row + stt - 3
        worksheet[f'B{idx}'] = datetime.datetime.strptime(date, '%d-%m-%Y') # ('01-Jan', '%d-%b')
        worksheet[f'C{idx}'] = '-'.join(data['Project'].split('-')[1:])
        worksheet[f'E{idx}'] = data['Description'].split('\n')[0]
        for pos in map_insert:
            worksheet[f'{pos.format(idx)}'] = data[map_insert[pos]]

        worksheet[f'B{idx}'].number_format = fmt_date
        worksheet[f'F{idx}'].number_format = fmt_number
        worksheet[f'G{idx}'].number_format = fmt_number
        worksheet[f'H{idx}'].number_format = fmt_number
        worksheet[f'I{idx}'].number_format = fmt_number
        worksheet[f'J{idx}'].number_format = fmt_number
        worksheet[f'K{idx}'].number_format = fmt_number
        worksheet[f'L{idx}'].number_format = fmt_number
    return worksheet

def daily(xlsx_report:str, xlsx_monthly:str, sheets: list):
    # load worksheet daily
    wb_report = openpyxl.load_workbook(filename=xlsx_report)
    ws_daily = wb_report['DAILY']
    # remove blank lines
    remove_blank_line(ws_daily)
    for date in sheets:
        try:
            data = get_data(xlsx_monthly, date)
            ws_daily = insert(ws_daily, data, date)
        except Exception as e:
            logfile(f'DAILY [{datetime.datetime.now()}] Error:{e}')
    wb_report.save(filename=xlsx_report)
    wb_report.close()

if __name__ == '__main__':
    # process('./report.xlsx', './empty/January.xlsx', ['01-01-2024', '02-01-2024', '03-01-2024'])
    daily(xlsx_report, xlsx_data, date_sheets)