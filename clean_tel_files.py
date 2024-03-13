import os
import sys
import pandas as pd
    
    
def open_and_clean(file_type):
    '''
    Основная рабочая функция на открытие, обработку, сохранения файлов
    '''
    user_profile_path = os.environ['USERPROFILE']
    docs_code_path = '\\!Worker\\'
    scr_path = os.path.join(user_profile_path, docs_code_path)
    if not os.path.exists(scr_path):
        os.mkdir(scr_path)
    print('Сейчас откроется папка, куда необходимо положить файл для обработки...\n'
          'При копировании файла, скопируй в буфер обмена имя файла\n'
          'После того как переместишь файл в необходимую папку, вернись в эту консоль')
    input('Для продолжения нажми любую клавишу...')
    os.system(f'explorer {scr_path}')

    print('Для того, чтобы не париться с переписыванием имени файла и не удивляться тому, что тут не работает Ctrl+V\n'
          'Нужно нажимать Shift+Insert(на некоторых клавиатурах Ins)')
    file_name = input('Введите имя файла: ')
    if file_type == '1':
        data = pd.read_csv(f"{scr_path}\\{file_name}.csv", sep=';', low_memory=False) 
    else:
        data = pd.read_excel(f"{scr_path}\\{file_name}.xlsx", dtype=None)

    if ((not 'Рабочий телефон') or (not 'Контакт: Рабочий телефон') or (not 'телефон') or (not 'тел')) in data.columns:
        print('File incorrect')
        input('Для выхода нажмите любую клавишу...')
        sys.exit()

    if 'Рабочий телефон' in data.columns:
        private_tel = data['Рабочий телефон'].dropna()
        if 'Мобильный телефон' not in data.columns:
            work_tel = data['Рабочий телефон'].dropna()
        else:
            work_tel = data['Мобильный телефон'].dropna()
        if 'Другой телефон' not in data.columns:
            other_tel = data['Рабочий телефон'].dropna()
        else:
            other_tel = data['Другой телефон'].dropna()  

    if 'Контакт: Рабочий телефон' in data.columns:
        private_tel = data['Контакт: Рабочий телефон'].dropna()
        if 'Контакт: Мобильный телефон' not in data.columns:
            work_tel = data['Контакт: Рабочий телефон'].dropna()
        else:
            work_tel = data['Контакт: Мобильный телефон'].dropna()
        if 'Контакт: Другой телефон' not in data.columns:
            other_tel = data['Контакт: Рабочий телефон'].dropna()
        else:
            other_tel = data['Контакт: Другой телефон'].dropna()

    if 'телефон' in data.columns:
        private_tel = data['телефон'].dropna()
        if 'Контакт: Частный телефон' not in data.columns:
            work_tel = data['телефон'].dropna()
        else:
            work_tel = data['телефон'].dropna()
        if 'Контакт: Другой телефон' not in data.columns:
            other_tel = data['телефон'].dropna()
        else:
            other_tel = data['телефон'].dropna()   

    if 'телефон' in data.columns:
        private_tel = data['телефон'].dropna()
        if 'Контакт: Частный телефон' not in data.columns:
            work_tel = data['телефон'].dropna()
        else:
            work_tel = data['телефон'].dropna()
        if 'Контакт: Другой телефон' not in data.columns:
            other_tel = data['телефон'].dropna()
        else:
            other_tel = data['телефон'].dropna() 

    all_tel = pd.concat( 
                    [private_tel, work_tel,  
                    other_tel],
                    ignore_index=True, sort=False
                    ).drop_duplicates() 
    print(all_tel)
    result = []
    new_result = []
    all_tel = all_tel.drop_duplicates()


    symbols_for_rep = ' +( )-+ '
    symbol_8 = '8'
    for tel in (all_tel.tolist()): 
        if ',' in str(tel): 
            a = tel.split(", ")
            for i in a:
                result.append(i)
        elif (' ' or '+' or '-' or '(' or ')') in str(tel):
            for symbol in symbols_for_rep:
                rep = tel.replace(symbol,'')
            result.append(rep)
        elif '8' in str(tel):
            rep = tel.replace('8','7')
            result.append(rep)
        else:           
            result.append(tel)  
    

    print(f'result = {result}')
    a = input('___enter___')

    new_result = []
    for tel in result:
        print(tel)
        if len(str(tel)) <= 10:
            continue
        elif len(str(tel)) > 11:
            continue
        else:
            new_result.append(tel)

    new_result.sort()
    print(f'result = {new_result}')
    a = input('___enter___')

    final_lisе = pd.DataFrame(new_result).drop_duplicates()  
    writer = pd.ExcelWriter(f'{scr_path}\\eml_{file_name}.xlsx') 
    final_lisе.to_excel(writer, index=False) 
    writer.close()
    

if __name__ == "__main__":
    print('Добро пожаловать в скрипт автоматической чистки .CSV и .XLSX файлов\n'
    'Содержащие в себе телефонные номера\n'
    'На вход принимаются файлы CSV формата с разделителем ";"\n'
    'В данной версии принимаются файлы экспорта в первой строке которых\n'
    'Содержатся пометки вида:\n'
    'Рабочий телефон\n'
    'Контакт: Рабочий телефон\n'
    'телефон\n'
    'тел\n'
    )
    print('Выберите тип файла: *нажимая 1 или 2, Марина, 1 или 2*\n'
          '1. csv\n'
          '2. xlsx\n')
    file_type = input()
    if ((file_type == '1') or (file_type == '2')):
        open_and_clean(file_type)
    else:
        print('Введён неверный ключ')
        input('Для выхода нажмите любую клавишу...')