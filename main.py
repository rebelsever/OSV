import os
import pandas as pd
from create_report import create_report, save_report
from file_reader import read_excel_files
from data_comparer import compare_data, calculate_differences, process_balance_columns

def main(ekp_files, gbk_files, save_path):
    ekp_columns = [
        'Балансовый счет', '№ Счета', 'Валюта', 'Входящий остаток', 'Входящий остаток, в н\\в', 
        'Дебетовый оборот', 'Дебетовый оборот н\\в', 'Кредитовый оборот', 'Кредитовый оборот н\\в', 
        'Исходящий остаток', 'Исходящий остаток, в нац. валюте'
    ]

    gbk_columns = [
        'Номер счета', 'Наименование счета', 'Код валюты', 'в валюте счета', 'в функциональной валюте', 
        'в валюте счета', 'в функциональной валюте', 'в валюте счета', 'в функциональной валюте', 
        'в валюте счета', 'в функциональной валюте', 'в валюте счета', 'в функциональной валюте', 
        'в валюте счета', 'в функциональной валюте'
    ]

    # Чтение файлов EKP
    print("Чтение данных EKP...")
    ekp_data = read_excel_files(ekp_files, ekp_columns)
    
    # Чтение файлов GBK
    print("Чтение данных GBK...")
    gbk_data = read_excel_files(gbk_files, gbk_columns)
    
    # Сравнение данных
    print("Сравнение данных...")
    discrepancies, merged_df = compare_data(ekp_data, gbk_data)
    
    # Обработка данных
    print("Обработка столбцов с балансами...")
    merged_df = process_balance_columns(merged_df)
    
    print("Вычисление разниц по модулю...")
    merged_df = calculate_differences(merged_df)
    
    # Создание отчета
    print("Создание отчета...")
    wb = create_report(discrepancies, merged_df)

    # Сохранение отчета 
    save_report(wb, save_path)
    
    print(f"Отчет успешно создан: {save_path}")

if __name__ == "__main__":
    main([], [], "")
