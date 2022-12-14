import pandas as pd
import re
import easygui
import numpy as np


app_ads_xmlsx = 'app-ads.xlsx'
locate_tomerge = easygui.fileopenbox(filetypes = ['*.txt'])
data_ads = []

try:
    with open(locate_tomerge, 'r') as f:
        publisher_tomerge = f.readline().strip()
        adds_tomerge = f.readlines()
        for line in adds_tomerge:
            pattern = r"^([\w.+-]*)(\s+)*,(\s+)*([\w.+-]*)(\s+)*,(\s+)*(\w+)(\s+)?,?(\s+)?([\w.]*)?(\s+)?([\w#.])*$"
            data_ads += [re.sub(pattern, r"\1, \4, \7, \10", line).split(', ') + [publisher_tomerge]]
except Exception as error:
    print(error)
    print('Некорректный путь к файлу или название файла')

try:
    app_ads = pd.read_excel(app_ads_xmlsx)
    list_app_ads = app_ads.to_dict(orient='records')
    for line in list_app_ads:
        pattern = r"^([\w.+-]*)(\s+)*,(\s+)*([\w.+-]*)(\s+)*,(\s+)*(\w+)(\s+)?,?(\s+)?([\w.]*)?(\s+)?([\w#.])*$"
        data_ads += [re.sub(pattern, r"\1, \4, \7, \10", line['app-ads.txt']).split(', ') + [line['Publisher']]]
except Exception as error:
    print(error)
    print('Файл app-ads.xlsx не найден.')

colums = ['a', 'b', 'c', 'd', 'e']
pd.options.mode.chained_assignment = None
df = pd.DataFrame(data_ads, index=None, columns=colums)
df_format = pd.DataFrame({'a': df['a'].str.lower(), 'b': df['b'],
                        'c' : df['c'].str.upper(), 'd' : df['d'], 'e' : df['e']})
df_format = df_format.replace(to_replace=r'RESELLERS$', value='RESELLER', regex=True)
df_format.drop_duplicates(inplace=True)
df_format.drop(np.where(df['b'] == 'INSERT PUBLISHER ID')[0], inplace=True)
delete_caramel_at_first = df_format[~((df_format.duplicated(['a', 'b', 'c'], keep=False))&(df_format['e']=='Caramel Ads'))]
delete_duplicate = delete_caramel_at_first[~((delete_caramel_at_first.duplicated(['a', 'b', 'c'], keep='first'))&(delete_caramel_at_first['d']==''))]
delete_duplicate.drop_duplicates(subset=['a', 'b', 'c'], keep = 'first', inplace=True)
delete_duplicate.sort_values(by=['a', 'b'], inplace=True)
listing = [delete_duplicate.columns.values.tolist()] + delete_duplicate.values.tolist()

new_app_ads = pd.DataFrame({'app-ads.txt': [', '.join(i[:-1]).strip() if len(i[-2])>=9 else ', '.join(i[:-2]).strip() for i in listing[1:]],
                   'Publisher': [ i[-1] for i in listing[1:]]})

try:
    new_app_ads.to_excel('new_app-ads.xlsx', index= False )
    print('Файл new_app-ads.xlsx готов')
except PermissionError as error:
    print(error)
    print('''Невозможно записать в файл. 
Возможно, new_app-ads.xlsx открыт''')

try:
    with open ('ads.txt', 'w') as f:
        for line in listing[1:]:
            if len(line[-2])>=9:
                string = str(', '.join(line[:-1]).strip())
            else:
                string = str(', '.join(line[:-2]).strip())
            f.write(string +'\n')
    print('Файл ads.txt готов')
except PermissionError as error:
    print(error)
    print('''Невозможно записать в файл. 
Возможно, ads.txt открыт''')
