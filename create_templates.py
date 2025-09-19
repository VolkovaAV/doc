import os
import config
def create_doc_templates(path_name):
    return  1

def checking_doc_templates(file_name):
    if os.path.exists(f'{config.TEMP_FOLDER_NAME}/{file_name}'):
        print(f"⚠️ Файл '{file_name}' уже существует. Создание отменено.")
        return True
    else:
        return False


print(checking_doc_templates('act.doc'))