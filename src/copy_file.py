import os
import shutil

def copy_file_to_folder(source_file, target_folder):
    """
    Копирует файл в указанную папку и возвращает путь к скопированному файлу.

    :param source_file: Путь к исходному файлу.
    :param target_folder: Папка, в которую нужно скопировать файл.
    :return: Путь к скопированному файлу.
    """
    # Получаем абсолютный путь к корню проекта
    project_root = os.path.dirname(os.path.abspath(__file__))
    
    # Создаем полный путь к целевой папке
    target_folder_path = os.path.join(project_root, target_folder)

    # Проверяем, существует ли целевая папка, если нет - создаем ее
    os.makedirs(target_folder_path, exist_ok=True)

    # Получаем имя файла из исходного пути
    file_name = os.path.basename(source_file)

    # Полный путь к новому файлу
    target_file_path = os.path.join(target_folder_path, file_name)

    # Копируем файл
    shutil.copy2(source_file, target_file_path)

    return target_file_path



