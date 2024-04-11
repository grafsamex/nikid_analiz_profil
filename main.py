import pandas as pd
import os


def filecsvlist(file: list) -> list:
    """Функция принимает все названия файлов и папок, находящихся в корневой с программой и возвращает список из файлов,
    относящихся к файлам баз данных с расширением .CSV"""
    listdbffiles = list()
    for i in file:
        if i.endswith('.CSV') or i.endswith('.csv') or i.endswith('.xlsx') or i.endswith('.XLSX'):
            listdbffiles.append(i)
    return listdbffiles

def org_tabl(org, profil):
    df_orgfilter = df.loc[(df['4'] == org)]
    org_data = pd.DataFrame(columns=['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15'])
    for i in range(len(profil)):
        org_profil = df_orgfilter.loc[(df_orgfilter['10'] == profil[i])]
        new_values1 = [org, profil[i], '', org_profil['15'].sum(), org_profil['16'].sum(),
                       org_profil['17'].sum(), org_profil['18'].sum(), org_profil['19'].sum(), org_profil['20'].sum(),
                       org_profil['15'].sum() + org_profil['18'].sum(), org_profil['16'].sum() + org_profil['19'].sum(),
                       org_profil['17'].sum() + org_profil['20'].sum(), 335, 0,
                       (org_profil['16'].sum() + org_profil['19'].sum()) / 335]
        org_data.loc[len(org_data.index)] = new_values1
    itog_values = ['Итого, ' + org, '', '', org_data['4'].sum(), org_data['5'].sum(), org_data['6'].sum(),
                   org_data['7'].sum(), org_data['8'].sum(), org_data['9'].sum(), org_data['10'].sum(), org_data['11'].sum(),
                   org_data['12'].sum(), org_data['13'].sum(), org_data['14'].sum(), org_data['15'].sum()]
    org_data.loc[len(org_data.index)] = itog_values
    org_data.loc[len(org_data.index)] = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
    org_data.columns = ['Наименование медицинская организация', 'Профиль койки', 'Количество коек',
                                 'По МО без МТР Случаи', 'По МО без МТР Дни. посещ.', 'По МО без МТР Стоимость(руб)',
                                 'По межтерр помощи МТР Случаи', 'По межтерр помощи МТР Дни. посещ.',
                                 'По межтерр помощи МТР Стоимость(руб)', 'Всего Случаи', 'Всего Дни. посещ.',
                                 'Всего Стоимость(руб)', 'Нормативный показатель работы койки', 'Работа койки',
                                 'Расчетное значение коек']
    return org_data



ROOT_DIR = os.path.dirname(os.path.abspath(__file__)) + '\Данные'
file = filecsvlist(os.listdir(ROOT_DIR))
if len(file) > 1:
    input(f'В папке {ROOT_DIR} находится более одного файла. Удалите файлы и перезапустите программу'
          f'Для выхода нажмите Enter')
    exit()
elif len(file) < 1:
    input(f'В папке {ROOT_DIR} отсутствует файл для анализа. Перезапустите выполнение. Для выхода нажмите Enter')
    exit()
filedir = ROOT_DIR +'\\' + file[0]
print(filedir)
df = pd.read_excel(filedir)
df = df.loc[(df['1'] == 'ГУЗ') & (df['5'] == 1)]
print(df)
medorglist = df['4'].tolist()
medorglist = list(set(medorglist))
print(medorglist)
svodtabl1 = pd.DataFrame(columns=['Наименование медицинская организация', 'Профиль койки', 'Количество коек',
                                 'По МО без МТР Случаи', 'По МО без МТР Дни. посещ.', 'По МО без МТР Стоимость(руб)',
                                 'По межтерр помощи МТР Случаи', 'По межтерр помощи МТР Дни. посещ.',
                                 'По межтерр помощи МТР Стоимость(руб)', 'Всего Случаи', 'Всего Дни. посещ.',
                                 'Всего Стоимость(руб)', 'Нормативный показатель работы койки', 'Работа койки',
                                 'Расчетное значение коек'])
svodtabl2 = pd.DataFrame(columns=['Наименование медицинская организация', 'Профиль койки', 'Количество коек',
                                 'По МО без МТР Случаи', 'По МО без МТР Дни. посещ.', 'По МО без МТР Стоимость(руб)',
                                 'По межтерр помощи МТР Случаи', 'По межтерр помощи МТР Дни. посещ.',
                                 'По межтерр помощи МТР Стоимость(руб)', 'Всего Случаи', 'Всего Дни. посещ.',
                                 'Всего Стоимость(руб)', 'Нормативный показатель работы койки', 'Работа койки',
                                 'Расчетное значение коек'])
profil_mp1 = ['дерматовенерологии', 'инфекционным болезням', 'педиатрии', 'неврологии', 'медицинской реабилитации',
              'пульмонологии', 'гастроэнтерологии', 'нефрологии', 'гематологии', 'ревматологии', 'детской эндокринологии',
              'детской кардиологии']
profil_mp2 = ['детской хирургии', 'офтальмологии', 'травматологии и ортопедии', 'хирургии',
              'оториноларингологии (за исключением кохлеарной имплантации)',
              'акушерству и гинекологии (за исключением использования вспомогательных репродуктивных технологий и искусственного прерывания беременности)',
              'детской урологии-андрологии', 'нейрохирургии', 'колопроктологии', 'акушерству и гинекологии (использованию вспомогательных репродуктивных технологий)',
              'челюстно-лицевой хирургии', 'сердечно-сосудистой хирургии', 'акушерству и гинекологии (искусственному прерыванию беременности)',
              'детской онкологии']



for i in range(len(medorglist)):
    new_frames1 = (org_tabl(medorglist[i], profil_mp1))
    frames_list = [svodtabl1, new_frames1]
    svodtabl1 = pd.concat(frames_list)
svodtabl1.to_excel('Анализ МО сводная 1 часть.xlsx', index=False)

for i in range(len(medorglist)):
    new_frames2 = (org_tabl(medorglist[i], profil_mp2))
    frames_list = [svodtabl2, new_frames2]
    svodtabl2 = pd.concat(frames_list)
svodtabl2.to_excel('Анализ МО сводная 2 часть.xlsx', index=False)