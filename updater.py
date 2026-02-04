import sys
import subprocess
import os
from PySide6.QtCore import QThread, Signal

class GitUpdateWorker(QThread):
    update_available = Signal(bool, str) # has_update, latest_version_tag
    error_occurred = Signal(str)
    update_finished = Signal(bool, str) # success, message

    def __init__(self, repo_path=None):
        super().__init__()
        self.repo_path = repo_path or os.getcwd()
        self.mode = "check" # check or pull

    def set_mode(self, mode):
        self.mode = mode

    def run(self):
        if self.mode == "check":
            self.check_updates()
        elif self.mode == "pull":
            self.perform_update()

    def check_updates(self):
        try:
            # Fetch remote info
            subprocess.run(["git", "fetch"], check=True, cwd=self.repo_path, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            
            # Get current hash
            current_hash = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=self.repo_path, creationflags=subprocess.CREATE_NO_WINDOW).strip().decode()
            
            # Get upstream hash (assuming main branch)
            # Try main or master
            try:
                 upstream_hash = subprocess.check_output(["git", "rev-parse", "origin/main"], cwd=self.repo_path, creationflags=subprocess.CREATE_NO_WINDOW).strip().decode()
            except:
                 upstream_hash = subprocess.check_output(["git", "rev-parse", "origin/master"], cwd=self.repo_path, creationflags=subprocess.CREATE_NO_WINDOW).strip().decode()

            if current_hash != upstream_hash:
                self.update_available.emit(True, "Yeni sürüm mevcut!")
            else:
                self.update_available.emit(False, "Sürüm güncel.")
                
        except Exception as e:
            self.error_occurred.emit(f"Güncelleme kontrolü hatası: {str(e)}")

    def perform_update(self):
        try:
            # Pull changes
            result = subprocess.run(["git", "pull"], check=True, cwd=self.repo_path, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            self.update_finished.emit(True, "Güncelleme başarıyla tamamlandı.\nDeğişikliklerin etkili olması için uygulamayı yeniden başlatın.")
        except Exception as e:
            self.update_finished.emit(False, f"Güncelleme başarısız: {str(e)}")
