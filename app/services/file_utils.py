# app/services/file_utils.py
import os
import sys
import shutil
from datetime import datetime


def get_available_drives():
    """Получает список доступных дисков в Windows"""
    drives = []
    if sys.platform == 'win32':
        import string
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                drives.append(drive)
    else:
        drives = ['/']
    return drives


def get_subdirectories(path):
    """Получает список подпапок в указанном пути"""
    subdirs = []
    try:
        if os.path.exists(path):
            for item in os.listdir(path):
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    subdirs.append(item)
    except:
        pass
    return sorted(subdirs)


def migrate_data_from_old_location(old_dir, new_dir, files_to_migrate, folders_to_migrate=None):
    """Переносит данные из старой папки в новую"""
    migrated_flag = os.path.join(new_dir, '.migrated')

    if os.path.exists(migrated_flag):
        return

    print("Checking for old data to migrate...")
    migrated_count = 0

    for old_name, new_path in files_to_migrate:
        old_path = os.path.join(old_dir, old_name)
        if os.path.exists(old_path) and not os.path.exists(new_path):
            try:
                shutil.copy2(old_path, new_path)
                print(f"  Migrated: {old_name}")
                migrated_count += 1
            except Exception as e:
                print(f"  Error migrating {old_name}: {e}")

    if folders_to_migrate:
        for old_name, new_path in folders_to_migrate:
            old_path = os.path.join(old_dir, old_name)
            if os.path.exists(old_path) and not os.path.exists(new_path):
                try:
                    shutil.copytree(old_path, new_path)
                    print(f"  Migrated folder: {old_name}")
                    migrated_count += 1
                except Exception as e:
                    print(f"  Error migrating {old_name}: {e}")

    if migrated_count > 0:
        print(f"Migration completed: {migrated_count} files/folders migrated")
    else:
        print("No old data found for migration")

    try:
        with open(migrated_flag, 'w') as f:
            f.write(f"Migration completed at {datetime.now().isoformat()}\n")
            f.write(f"Source: {old_dir}\n")
            f.write(f"Destination: {new_dir}\n")
    except:
        pass