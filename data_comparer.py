import pandas as pd

def process_balance_columns(df):
    # Обработка столбца "Входящий остаток, в н\\в"
    df['Входящий отрицательный'] = df['Входящий остаток, в н\\в'].apply(lambda x: x if x < 0 else 0)
    df['Входящий положительный'] = df['Входящий остаток, в н\\в'].apply(lambda x: x if x >= 0 else 0)

    # Обработка столбца "Исходящий остаток, в нац. валюте"
    df['Исходящий отрицательный'] = df['Исходящий остаток, в нац. валюте'].apply(lambda x: x if x < 0 else 0)
    df['Исходящий положительный'] = df['Исходящий остаток, в нац. валюте'].apply(lambda x: x if x >= 0 else 0)

    return df

def compare_data(ekp_data, gbk_data):
    ekp_accounts = set()
    gbk_accounts = set()
    
    for df in ekp_data:
        ekp_accounts.update(df['№ Счета'].dropna().astype(str).unique())
    
    for df in gbk_data:
        gbk_accounts.update(df['Номер счета'].dropna().astype(str).unique())
    
    common_accounts = ekp_accounts & gbk_accounts
    ekp_only = ekp_accounts - gbk_accounts
    gbk_only = gbk_accounts - ekp_accounts
    
    discrepancies = {
        'common': list(common_accounts),
        'ekp_only': list(ekp_only),
        'gbk_only': list(gbk_only)
    }
    
    # Объединяем данные по столбцам "№ Счета" и "Номер счета" с использованием outer join
    ekp_df = pd.concat(ekp_data)
    gbk_df = pd.concat(gbk_data)
    gbk_df.rename(columns={"Номер счета": "№ Счета"}, inplace=True)
    merged_df = pd.merge(ekp_df, gbk_df, on="№ Счета", how="outer", suffixes=('_ekp', '_gbk'))

    # Удаляем строки, у которых значение ключевого столбца пустое или равно 0
    merged_df = merged_df[merged_df['№ Счета'].notna() & (merged_df['№ Счета'] != 0)]

    return discrepancies, merged_df

#Считаем разницу
def calculate_differences(merged_df):
    merged_df.fillna(0, inplace=True)

    # Преобразование данных в числовой формат
    cols_to_convert = [ 
        'Входящий отрицательный', 'в функциональной валюте', 'Входящий положительный', 'в функциональной валюте_1',
        'Дебетовый оборот', 'в функциональной валюте_2', 'Кредитовый оборот', 'в функциональной валюте_3',
        'Исходящий отрицательный', 'в функциональной валюте_4', 'Исходящий положительный', 'в функциональной валюте_5' 
    ] 
    for col in cols_to_convert:
        merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce')

    merged_df['difference_2_8'] = merged_df.apply(
        lambda row: abs(row['Входящий отрицательный']) - abs(row['в функциональной валюте']), axis = 1
    )
    merged_df['difference_3_9'] = merged_df.apply(
        lambda row: abs(row['Входящий положительный'])- abs(row['в функциональной валюте_1']), axis = 1
    )  
    merged_df['difference_4_10'] = merged_df.apply(
        lambda row: abs(row['Дебетовый оборот']) - abs(row['в функциональной валюте_2']), axis = 1
    )
    merged_df['difference_5_11'] = merged_df.apply(
        lambda row: abs(row['Кредитовый оборот']) - abs(row['в функциональной валюте_3']), axis = 1
    )
    merged_df['difference_6_12'] = merged_df.apply(
        lambda row: abs(row['Исходящий отрицательный']) - abs(row['в функциональной валюте_4']), axis = 1
    )
    merged_df['difference_7_13'] = merged_df.apply(
        lambda row: abs(row['Исходящий положительный']) - abs(row['в функциональной валюте_5']), axis = 1
    )
    return merged_df