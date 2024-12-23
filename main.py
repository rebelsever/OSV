import os
import pandas as pd
from create_report import create_report
from file_reader import read_excel_files
from data_comparer import compare_data, calculate_differences, process_balance_columns

# Определение путей к файлам и целевых столбцов
ekp_files = [r"C:\Users\Компьютер\Desktop\ReconcilinationOfMigrationResults\input_ekp\Тестовый от ЕКП.xlsx",
             r"C:\Users\Компьютер\Desktop\ReconcilinationOfMigrationResults\input_ekp\корпорат.xlsx"]
gbk_files = [r"C:\Users\Компьютер\Desktop\ReconcilinationOfMigrationResults\input_gbk\Тестовый от гбк.xlsx"]

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

def main():
    # Определение пути для сохранения отчета на рабочем столе
    save_path = os.path.join(os.path.expanduser("~"), "Desktop")

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
    create_report(discrepancies, merged_df, save_path)
    
    print("Отчет успешно создан.")

if __name__ == "__main__":
    main()
