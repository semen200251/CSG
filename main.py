import datetime
import os
import pandas as pd
import openpyxl as opx
import win32com.client

import config


def get_project(path):
    """Открывает файл проекта и возвращает объект проекта"""
    try:
        msp = win32com.client.gencache.EnsureDispatch("MSProject.Application")
        msp.FileOpen(path)
        project = msp.ActiveProject
    except Exception as e:
        print(f"Такого файла(project) не существует: {e}")
        return None
    return project


def get_excel_pd(path):
    """Создает DataFrame из ОФ и заменяет пустые строки на значение, указанное в config"""
    try:
        data = pd.read_excel(path)
        data.fillna(value=config.errors.get('None'), inplace=True)
        for key, value in config.errors.items():
            if key != 'None':
                data.replace(key, value, inplace=True)
    except Exception as e:
        print(f"Такого файла(excel) не существует: {e}")
        return None
    return data


def get_project_pd(project, columns):
    """Создает DataFrame из столбцов объекта проекта."""
    task_collection = project.Tasks
    data = pd.DataFrame(columns=columns)
    data[columns[1]] = pd.to_datetime(data[columns[1]], dayfirst=True)
    data[columns[2]] = pd.to_datetime(data[columns[2]], dayfirst=True)
    for t in task_collection:
        if t.ActualStart != 'НД':
            actual_start = datetime.datetime.date(t.ActualStart)
        else:
            actual_start = 'НД'
        if t.ActualFinish != 'НД':
            actual_finish = datetime.datetime.date(t.ActualFinish)
        else:
            actual_finish = 'НД'
        data.loc[len(data.index)] = [t.Text4, actual_start, actual_finish]
    return data


def change_project(project, changes):
    """Вносит изменения в проект"""
    task_collection = project.Tasks
    for i, t in enumerate(task_collection):
        if i in changes.keys():
            t.Name = t.Name.replace('п', '1')


def check_str(excel_str, project_str, columns):
    """Сравнивает значения task ОФ и проекта"""
    for col in columns:
        if not pd.isnull(excel_str[col]) and excel_str[col] != 'НД':
            excel_str[col] = datetime.datetime.date(excel_str[col])
        if excel_str[col] != project_str[col]:
            return False, col
    return True, None


def paint_excel(ws, i, column):
    """Закрашивает ячейку в ОФ, которая не совпадает с проектом"""
    letter = chr(65 + column)
    work_sheet_a1 = ws[f"{letter}{i}"]
    work_sheet_a1.fill = opx.styles.PatternFill(fill_type='solid', start_color=config.color_fill)


def check_form(data_project, data_excel, columns):
    """Находит несоотвествия между ОФ и проектом и сохраняет их в словарь"""
    wb = opx.load_workbook(r"051-2000260_оф_ф_(13.04).xlsx")
    ws = wb.active
    changes = {}
    for i, excel_str in data_excel.iterrows():
        for j, project_str in data_project.iterrows():
            if excel_str[columns[0]] == project_str[columns[0]]:
                status, column = check_str(excel_str, project_str, columns[1:])
                if not status:
                    paint_excel(ws, i + 2, data_excel.columns.get_loc(column))
                    changes[j] = [column, excel_str[column]]
                break
    wb.save(r"Изменения в ОФ.xlsx")
    return changes


def main():
    try:
        os.chdir(config.path_directory)
    except Exception as e:
        print("Что-то не так с директорией: {0}".format(config.path_directory), e)
        return None
    path_to_project = config.path_directory + r"\051-2000260_2022_нг_ф_(06.04).mpp"
    path_to_excel = config.path_directory + r"\051-2000260_оф_ф_(13.04).xlsx"
    project = get_project(path_to_project)
    if project is None:
        return None
    data_project = get_project_pd(project, config.columns)
    data_excel = get_excel_pd(path_to_excel)
    if data_excel is None:
        return None
    changes = check_form(data_project, data_excel, config.columns)
    change_project(project, changes)


main()
