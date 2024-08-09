import shutil
import os
import datetime
import openpyxl

class BackupManager:
    backup_folder = 'backup'

    @staticmethod
    def create_backup():
        if not os.path.exists(BackupManager.backup_folder):
            os.makedirs(BackupManager.backup_folder)

        if BackupManager.is_file_valid('data.xlsx'):
            date_prefix = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            destination_path = os.path.join(BackupManager.backup_folder, f'{date_prefix}_data.xlsx')

            try:
                shutil.copyfile('data.xlsx', destination_path)
                print(f"Backup created successfully at {destination_path}.")
            except IOError as e:
                print("Unable to copy file. %s" % e)
            except Exception as e:
                print("Unexpected error:", e)
        else:
            print("File 'data.xlsx' is invalid or corrupted. Backup not created.")

    @staticmethod
    def restore_latest_backup():
        try:
            if not os.path.exists(BackupManager.backup_folder):
                print("Backup folder does not exist.")
                return

            backups = sorted([f for f in os.listdir(BackupManager.backup_folder) if f.endswith('.xlsx')],
                             reverse=True)

            if not backups:
                print("No backups found.")
                return

            latest_backup = os.path.join(BackupManager.backup_folder, backups[0])
            if BackupManager.is_file_valid(latest_backup):
                shutil.copyfile(latest_backup, 'data.xlsx')
                print(f"Restored latest backup from {latest_backup}.")
            else:
                print("Latest backup file is corrupted. Restoration failed.")

        except Exception as e:
            print("Unexpected error during restoration:", e)

    @staticmethod
    def is_file_valid(file_path):
        try:
            openpyxl.load_workbook(file_path)
            return True
        except Exception as e:
            print(f"File validation error: {e}")
            return False
