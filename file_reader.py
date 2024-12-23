import pandas as pd
import os

def read_excel_files(files, target_columns):
    data_frames = []
    for file_path in files:
        try:
            print(f"Найден файл: {file_path}")
            if file_path.endswith('.xml'):
                df = pd.read_xml(file_path)
                df = find_and_select_columns(df, target_columns)
                data_frames.append(df)
            elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                xl = pd.ExcelFile(file_path)
                sheet_processed = False
                for sheet_name in xl.sheet_names:
                    df = xl.parse(sheet_name, header=None)
                    df = find_and_select_columns(df, target_columns)
                    if df is not None:
                        data_frames.append(df)
                        sheet_processed = True
                        break
                if not sheet_processed:
                    raise ValueError(f"Файл {os.path.basename(file_path)} не содержит ожидаемых данных.")
            else:
                raise ValueError(f"Формат файла {os.path.basename(file_path)} не соответствует ожидаемому.")
        except ValueError as e:
            print(e)
    return data_frames



def find_and_select_columns(df, target_columns):
    combined_columns = [None] * len(target_columns)
    for idx, col in enumerate(target_columns):
        found = False
        for row in range(min(15, len(df))):  # Проверяем только первые 15 строк
            for col_index, cell in enumerate(df.iloc[row]):
                if cell and col in str(cell):
                    combined_columns[idx] = df.iloc[row, col_index]
                    found = True
                    break
            if found:
                break
    if all(col is not None for col in combined_columns):
        header_row_index = df[df.iloc[:, 0].str.contains(target_columns[0], na=False)].index[0]
        new_df = df.iloc[header_row_index+1:].reset_index(drop=True)
        new_df.columns = combined_columns
        new_df = rename_duplicate_columns(new_df)
        return new_df
    return None

def rename_duplicate_columns(df):
    columns = df.columns
    new_columns = []
    column_count = {}
    
    for column in columns:
        if column in column_count:
            column_count[column] += 1
            new_columns.append(f"{column}_{column_count[column]}")
        else:
            column_count[column] = 0
            new_columns.append(column)

    df.columns = new_columns
    return df

def print_values(files, target_columns, column_name):
    data_frames = read_excel_files(files, target_columns)
    values = set()
    for df in data_frames:
        if column_name in df.columns:
            values.update(df[column_name].dropna().astype(str).unique())
        else:
            print(f"Столбец '{column_name}' не найден в файле.")
    return values
