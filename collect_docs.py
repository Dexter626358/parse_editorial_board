import os
import shutil

def collect_docs(source_dir, target_dir='collected_docs'):
    # Создаём папку для сбора файлов, если её ещё нет
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)

    # Счётчик для переименования файлов с одинаковым именем
    name_counts = {}

    for root, dirs, files in os.walk(source_dir):
        for file in files:
            # Проверяем расширение файла
            if file.lower().endswith(('.doc', '.docx')):
                src_path = os.path.join(root, file)
                # Проверка на уникальность имени
                if file not in name_counts:
                    name_counts[file] = 1
                    dst_file = file
                else:
                    name_counts[file] += 1
                    name_part, ext = os.path.splitext(file)
                    dst_file = f"{name_part}_{name_counts[file]}{ext}"

                dst_path = os.path.join(target_dir, dst_file)
                shutil.copy2(src_path, dst_path)
                print(f"Скопирован: {src_path} → {dst_path}")

if __name__ == "__main__":
    # Замените на путь к вашей исходной папке
    source_folder = r"C:\Users\gusev\Downloads\Анкеты\Анкеты"
    collect_docs(source_folder)
