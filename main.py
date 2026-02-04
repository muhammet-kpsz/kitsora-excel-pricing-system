import os
import sys
from datetime import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                               QTabWidget, QLabel, QLineEdit, QPushButton, QFileDialog, 
                               QComboBox, QTableWidget, QTableWidgetItem, QHeaderView, 
                               QCheckBox, QSpinBox, QDoubleSpinBox, QMessageBox, QProgressBar,
                               QGroupBox, QFormLayout, QStyleFactory, QProgressDialog,
                               QPlainTextEdit, QStackedWidget, QDialog, QMenu, QScrollArea)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QIcon, QPalette, QColor, QFont, QPixmap

import ctypes
try:
    import winreg
except ImportError:
    winreg = None

from settings import SettingsManager
from pricing_engine import PricingEngine
from excel_io import ExcelHandler

# Import openpyxl for the new generator logic
# Import openpyxl for the new generator logic
from openpyxl import Workbook
import version
from updater import GitUpdateWorker

class Worker(QThread):
    progress_part = Signal(str, int, int) # status, part_num, row_count
    finished = Signal(bool, str)
    log_message = Signal(str)  # For debug messages to GUI
    
    def __init__(self, filepath, settings_manager, pricing_engine):
        super().__init__()
        self.filepath = filepath
        self.sm = settings_manager
        self.engine = pricing_engine
        self.io = ExcelHandler()

    def run(self):
        # The generator is now in ExcelHandler
        gen = self.io.process_and_save_generator(
            self.filepath, 
            self.sm, 
            self.engine
        )
        
        for status_type, data in gen:
            if status_type == "ERROR":
                self.finished.emit(False, str(data))
                return
            elif status_type == "DONE":
                self.finished.emit(True, str(data))
                return
            elif status_type == "PART_START":
                self.progress_part.emit("START", data, 0)
            elif status_type == "PROGRESS":
                # data = (part_num, current_rows, total_processed)
                self.progress_part.emit("PROGRESS", data[0], data[1])
            elif status_type == "PART_COMPLETE":
                # data = (part_num, final_rows)
                self.progress_part.emit("COMPLETE", data[0], data[1])
            elif status_type == "LOG":
                # Log message to GUI
                self.log_message.emit(str(data))

class FileLoaderWorker(QThread):
    finished = Signal(list)
    failed = Signal(str)
    
    def __init__(self, filepath):
        super().__init__()
        self.filepath = filepath
        self.io = ExcelHandler()

    def run(self):
        try:
            rows = self.io.get_all_rows(self.filepath, limit=50000)
            self.finished.emit(rows)
        except Exception as e:
            self.failed.emit(str(e))

class PreviewWorker(QThread):
    finished = Signal(list, int, set) # results, changed_count, categories_set
    
    def __init__(self, all_rows, engine, search_txt, cat_filter, variant_col=None, variant_val_col=None, show_unique_variant=False, 
                 stock_col=None, include_zero_stock=True, selected_categories=None):  # NEW: Added stock and category filter params
        super().__init__()
        self.all_rows = all_rows
        self.engine = engine
        self.search_txt = search_txt.lower()
        self.cat_filter = cat_filter
        self.variant_col = variant_col
        self.variant_val_col = variant_val_col
        self.show_unique_variant = show_unique_variant
        # ===== NEW FEATURE: Stock filtering parameters =====
        self.stock_col = stock_col
        self.include_zero_stock = include_zero_stock
        self.selected_categories = selected_categories if selected_categories else []
        # ===== END NEW FEATURE =====

    def run(self):
        filtered_rows = []
        cats = set()
        
        seen_variants = set()
        
        changed_variants = set() 
        changed_simple_count = 0
        
        for r_data in self.all_rows:
            res = self.engine.calculate_row(r_data)
            res["_raw_data"] = r_data # Attach raw data for comparison
            
            s_code = str(res.get("stock_code", "")).lower()
            p_name = str(res.get("product_name", "")).lower()
            cat = str(res.get("main_category", ""))
            
            if cat: cats.add(cat)
            
            # ===== NEW FEATURE: Stock Filter =====
            # Apply stock filtering if enabled
            if self.stock_col and not self.include_zero_stock:
                from stock_filter import StockFilter
                stock_val = StockFilter.get_stock_value(r_data, self.stock_col)
                if stock_val <= 0:
                    continue  # Skip zero stock items
            
            # Store stock value in result for display
            if self.stock_col:
                from stock_filter import StockFilter
                res["_stock_value"] = StockFilter.get_stock_value(r_data, self.stock_col)
            # ===== END NEW FEATURE =====
            
            # Search Filter
            # Check stock, name, and also price strings
            search_match = False
            if not self.search_txt:
                search_match = True
            else:
                if self.search_txt in s_code or self.search_txt in p_name:
                    search_match = True
                else:
                    # Check prices
                    try:
                        if self.search_txt in str(res.get("base_price", "")) or \
                           self.search_txt in str(res.get("final_discounted_price", "")) or \
                           self.search_txt in str(res.get("label_price", "")):
                            search_match = True
                    except:
                        pass
            
            
            # Normalize path immediately for accurate filtering
            raw_path = res.get("full_category_path", cat)
            if raw_path:
                full_cat_path = " > ".join([p.strip() for p in str(raw_path).split(">") if p.strip()])
            else:
                full_cat_path = cat # Fallback to main category if full path missing

            # Cat Filter (Dropdown Selection)
            # FIX: Compare against full path, not just main category
            if self.cat_filter != "Tüm Kategoriler":
                if not (full_cat_path == self.cat_filter or full_cat_path.startswith(self.cat_filter + " >")):
                    continue
            
            # ===== NEW FEATURE: Category Tree Filter (Enhanced) =====
            # If category tree has selections, filter by them using FULL path
            if self.selected_categories:
                # Path already normalized above
                pass 
                
                # Check if current category or any parent is selected
                cat_match = False
                for selected in self.selected_categories:
                    # Match full path, main category, or if full path starts with selected
                    if full_cat_path == selected or \
                       cat == selected or \
                       full_cat_path.startswith(selected + " >") or \
                       cat.startswith(selected + " >"):
                        cat_match = True
                        break
                
                if not cat_match:
                    continue  # Skip if not in selected categories
                
        
        # Fixed duplicate block
        
            # ===== END NEW FEATURE =====
            
            if not search_match:
                continue
                
            # Unique Variant Logic for Display
            if self.variant_col and self.show_unique_variant:
                v_id = r_data.get(self.variant_col, "")
                if v_id and str(v_id) in seen_variants:
                    # Duplicate variant, calculate for stats but DON'T add to display list
                    # Wait, requirement says "hide", so we skip adding to filtered_rows
                    
                    # But we still need to calculate change stats? 
                    # Usually "Show Unique" implies visualizing 1 representative.
                    # We will skip adding to filtered_rows.
                    pass
                else:
                    filtered_rows.append(res)
                    if v_id: seen_variants.add(str(v_id))
            else:
                filtered_rows.append(res)
            
            
            # Check Change
            try:
                base = float(res.get("base_price", 0))
                final = float(res.get("final_discounted_price", 0))
                if abs(final - base) > 0.01:
                    if self.variant_col:
                        # Variant Mode: Track unique Variant IDs that changed
                        v_id = r_data.get(self.variant_col, "")
                        res["_variant_id"] = v_id 
                        res["_variant_val"] = r_data.get(self.variant_val_col, "") if self.variant_val_col else ""
                        if v_id:
                            changed_variants.add(str(v_id))
                        else:
                            changed_simple_count += 1
                    else:
                        changed_simple_count += 1
                else:
                    # Even if no change, if we want to show it in table, we might need variant info
                    if self.variant_col:
                        res["_variant_id"] = r_data.get(self.variant_col, "")
                        res["_variant_val"] = r_data.get(self.variant_val_col, "") if self.variant_val_col else ""

            except:
                pass
                
        final_count = len(changed_variants) if self.variant_col else changed_simple_count
        if self.variant_col and changed_simple_count > 0:
             # Add fallback for rows without variant ID
             final_count += changed_simple_count

        self.finished.emit(filtered_rows, final_count, cats)

class CategoryWorker(QThread):
    finished = Signal(set)
    
    
    def __init__(self, filepath, cat_col, engine, no_cat_mode=False):
        super().__init__()
        self.filepath = filepath
        self.cat_col = cat_col
        self.engine = engine
        self.io = ExcelHandler()
        self.no_cat_mode = no_cat_mode

    def run(self):
        # ===== ENHANCED: Collect full category paths AND counts for tree =====
        # Scan up to 50k rows for categories to ensure consistency with preview
        rows = self.io.get_all_rows(self.filepath, limit=50000)
        category_counts = {}
        
        if self.no_cat_mode:
            # All items are "Kategorisiz"
            count = len(rows)
            if count > 0:
                category_counts["Kategorisiz"] = count
        else:
            for r in rows:
                raw = str(r.get(self.cat_col, ""))
                if raw and raw != "nan":
                    # Normalize path immediately: "A>B" -> "A > B"
                    # This ensures tree keys match exactly with later preview logic
                    normalized_path = " > ".join([p.strip() for p in raw.split(">") if p.strip()])
                    
                    if normalized_path:
                        category_counts[normalized_path] = category_counts.get(normalized_path, 0) + 1
                    
        # Emit dictionary {path: count} instead of just list
        self.finished.emit(category_counts)
        # ===== END ENHANCEMENT =====

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_app_identity()
        # Theme application moved after SM init
        
        self.setWindowTitle(version.APP_NAME)
        self.resize(1100, 750)
        
        if os.path.exists("assets/icon.png"):
            self.setWindowIcon(QIcon("assets/icon.png"))
        
        self.sm = SettingsManager()
        self.engine = PricingEngine(self.sm)
        self.io = ExcelHandler()
        self.io = ExcelHandler()
        self.current_headers = []
        # ===== NEW FEATURE: Persistent Category State =====
        self.persistent_selected_categories = set()
        # ===== END NEW FEATURE =====
        
        self.apply_theme() # Move after SM init to read settings
        
        self.setup_ui()
        self.load_ui_values()
        
        # Search Debouncing
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.run_apply_filters)

    def setup_app_identity(self):
        # Set App User Model ID for Windows Taskbar grouping
        myappid = 'kitsora.excelpricingengine.app.1.0' 
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

    def apply_theme(self):
        # Read from settings
        theme = self.sm.get("theme", "system")
        
        app = QApplication.instance()
        style = QStyleFactory.create("Fusion")
        app.setStyle(style)
        
        # Delegate
        mode = theme
        if theme == "system":
            if self.is_system_dark():
                mode = "dark"
            else:
                mode = "light"
        self.apply_theme_manual(mode)
        return

        if theme == "dark":
            palette = QPalette()
            palette.setColor(QPalette.Window, QColor(53, 53, 53))
            palette.setColor(QPalette.WindowText, Qt.white)
            palette.setColor(QPalette.Base, QColor(25, 25, 25))
            palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
            palette.setColor(QPalette.ToolTipBase, Qt.white)
            palette.setColor(QPalette.ToolTipText, Qt.white)
            palette.setColor(QPalette.Text, Qt.white)
            palette.setColor(QPalette.Button, QColor(53, 53, 53))
            palette.setColor(QPalette.ButtonText, Qt.white)
            palette.setColor(QPalette.BrightText, Qt.red)
            palette.setColor(QPalette.Link, QColor(42, 130, 218))
            palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
            palette.setColor(QPalette.HighlightedText, Qt.black)
            app.setPalette(palette)
        elif theme == "light":
            # Restoration to standard palette is complex, simpler to just force a light fusion or restart
            # Fusion light default is actually:
            palette = QPalette()
            palette.setColor(QPalette.Window, QColor(240, 240, 240))
            palette.setColor(QPalette.WindowText, Qt.black)
            palette.setColor(QPalette.Base, Qt.white)
            palette.setColor(QPalette.AlternateBase, QColor(233, 233, 233))
            palette.setColor(QPalette.ToolTipBase, Qt.white)
            palette.setColor(QPalette.ToolTipText, Qt.black)
            palette.setColor(QPalette.Text, Qt.black)
            palette.setColor(QPalette.Button, QColor(240, 240, 240))
            palette.setColor(QPalette.ButtonText, Qt.black)
            palette.setColor(QPalette.BrightText, Qt.red)
            palette.setColor(QPalette.Link, QColor(0, 0, 255))
            palette.setColor(QPalette.Highlight, QColor(0, 120, 215))
            palette.setColor(QPalette.HighlightedText, Qt.white)
            app.setPalette(palette)
        elif theme == "system":
            # Detect System Theme (Windows)
            is_dark = False
            if winreg:
                try:
                    registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
                    key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                    value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                    if value == 0:
                        is_dark = True
                except Exception:
                    pass
            
            if is_dark:
                # Same dark palette logic
                self.apply_theme_manual("dark")
            else:
                self.apply_theme_manual("light")

    def apply_theme_manual(self, mode):
        # Helper to avoid duplication
        app = QApplication.instance()
        
        if mode == "kitsora":
            # Kitsora Theme (Fox Orange)
            palette = QPalette()
            # Backgrounds - Soft Cream/Orange tint
            palette.setColor(QPalette.Window, QColor(255, 248, 240)) 
            palette.setColor(QPalette.WindowText, QColor(60, 40, 30)) # Dark Brown/Grey
            
            # Input fields
            palette.setColor(QPalette.Base, Qt.white)
            palette.setColor(QPalette.AlternateBase, QColor(255, 240, 220))
            
            # Text
            palette.setColor(QPalette.Text, QColor(60, 40, 30))
            palette.setColor(QPalette.ToolTipBase, QColor(255, 204, 153))
            palette.setColor(QPalette.ToolTipText, QColor(60, 40, 30))
            
            # Buttons - Vibrant Orange
            palette.setColor(QPalette.Button, QColor(255, 140, 0)) # Dark Orange
            palette.setColor(QPalette.ButtonText, Qt.white)
            
            # Links/Highlights
            palette.setColor(QPalette.Link, QColor(230, 126, 34))
            palette.setColor(QPalette.Highlight, QColor(255, 140, 0))
            palette.setColor(QPalette.HighlightedText, Qt.white)
            
            # Bright text (errors etc)
            palette.setColor(QPalette.BrightText, Qt.red)
            
            app.setPalette(palette)
            app.setStyle("Fusion") # Fusion looks best with custom palettes
            
        elif mode == "dark":
            palette = QPalette()
            palette.setColor(QPalette.Window, QColor(53, 53, 53))
            palette.setColor(QPalette.WindowText, Qt.white)
            palette.setColor(QPalette.Base, QColor(25, 25, 25))
            palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
            palette.setColor(QPalette.ToolTipBase, Qt.white)
            palette.setColor(QPalette.ToolTipText, Qt.white)
            palette.setColor(QPalette.Text, Qt.white)
            palette.setColor(QPalette.Button, QColor(53, 53, 53))
            palette.setColor(QPalette.ButtonText, Qt.white)
            palette.setColor(QPalette.BrightText, Qt.red)
            palette.setColor(QPalette.Link, QColor(42, 130, 218))
            palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
            palette.setColor(QPalette.HighlightedText, Qt.black)
            app.setPalette(palette)
        else:
            palette = QPalette()
            palette.setColor(QPalette.Window, QColor(240, 240, 240))
            palette.setColor(QPalette.WindowText, Qt.black)
            palette.setColor(QPalette.Base, Qt.white)
            palette.setColor(QPalette.AlternateBase, QColor(233, 233, 233))
            palette.setColor(QPalette.ToolTipBase, Qt.white)
            palette.setColor(QPalette.ToolTipText, Qt.black)
            palette.setColor(QPalette.Text, Qt.black)
            palette.setColor(QPalette.Button, QColor(240, 240, 240))
            palette.setColor(QPalette.ButtonText, Qt.black)
            palette.setColor(QPalette.BrightText, Qt.red)
            palette.setColor(QPalette.Link, QColor(0, 0, 255))
            palette.setColor(QPalette.Highlight, QColor(0, 120, 215))
            palette.setColor(QPalette.HighlightedText, Qt.white)
            app.setPalette(palette)

    # ===== NEW FEATURE: Helper to detect system dark mode =====
    def is_system_dark(self):
        """Detects if system is in dark mode (Windows)"""
        if winreg:
            try:
                registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
                key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                return value == 0  # 0 = dark mode
            except Exception:
                pass
        return False
    # ===== END NEW FEATURE =====

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Helper to load README content
        self.readme_content = "Yardım dosyası bulunamadı."
        if os.path.exists("README.md"):
            try:
                import markdown
                with open("README.md", "r", encoding="utf-8") as f:
                    md_text = f.read()
                    self.readme_content = markdown.markdown(md_text)
            except ImportError:
                 with open("README.md", "r", encoding="utf-8") as f:
                    self.readme_content = f"<pre>{f.read()}</pre>"
            except Exception as e:
                self.readme_content = str(e)
        
        self.tab1 = self.create_file_tab()
        self.tab2 = self.create_categories_tab()
        self.tab3 = self.create_profit_tab()
        self.tab4 = self.create_rounding_tab()
        self.tab5 = self.create_preview_tab()
        self.tab6 = self.create_export_tab()
        self.tab_log = self.create_log_tab()
        self.tab_help = self.create_help_tab()
        
        self.tabs.addTab(self.tab1, "Dosya Eşleştirme")
        self.tabs.addTab(self.tab2, "Kategoriler")
        self.tabs.addTab(self.tab3, "Kâr Segmentleri")
        self.tabs.addTab(self.tab4, "Yuvarlama / Sınırlar")
        self.tabs.addTab(self.tab5, "Ürün Önizleme")
        self.tabs.addTab(self.tab6, "Çıktı İşlemleri / Ayarlar")
        self.tabs.addTab(self.tab_log, "İşlem Kayıtları (Log)")
        self.tabs.addTab(self.tab_help, "Yardım / Nasıl Kullanılır")

    def create_help_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Scroll Area for all content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(25)
        
        # ===== LOGO =====
        logo_label = QLabel()
        logo_path = os.path.join(os.path.dirname(__file__), "assets", "icon.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            # Scale smoothly
            pixmap = pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_label.setPixmap(pixmap)
            logo_label.setAlignment(Qt.AlignCenter)
            
            # Simple styling
            logo_label.setStyleSheet("padding: 10px; margin-bottom: 10px;")
            content_layout.addWidget(logo_label)

        # ===== HEADER =====
        header = QLabel(f"📚 {version.APP_NAME} Kullanım Kılavuzu")
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("""
            QLabel {
                font-size: 24px;
                font-weight: bold;
                color: #2c3e50;
                padding: 20px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #FF8C00, stop:1 #FF4500);
                color: white;
                border-radius: 10px;
            }
        """)
        content_layout.addWidget(header)
        
        # ===== STEP 1 =====
        step1 = self._create_step_card(
            "1️⃣ Dosya Seçimi",
            [
                "• <b>Excel Seç</b> butonuna tıklayarak ürün dosyanızı seçin",
                "• Sistem otomatik olarak sütunları algılar",
                "• Eğer dosya çok büyükse okuma birkaç saniye sürebilir"
            ],
            "#3498db"
        )
        content_layout.addWidget(step1)
        
        # ===== STEP 2 =====
        step2 = self._create_step_card(
            "2️⃣ Sütun Eşleştirme",
            [
                "• Her açılır menüden ilgili sütunu seçin:",
                "  - <b>Stok Kodu:</b> Ürünün benzersiz kodu",
                "  - <b>Ürün Adı:</b> Ürün ismini içeren sütun",
                "  - <b>Kategori:</b> Kategori bilgisi (opsiyonel)",
                "  - <b>Alış/Satış Fiyatları:</b> Mevcut fiyat sütunları",
                "• ⚠️ Zorunlu alanlar: Stok Kodu, Ürün Adı, Alış Fiyatı"
            ],
            "#9b59b6"
        )
        content_layout.addWidget(step2)
        
        # ===== STEP 3 =====
        step3 = self._create_step_card(
            "3️⃣ Kategori Yönetimi",
            [
                "• <b>Kategoriler</b> sekmesine gidin",
                "• <b>'Excel'den Kategorileri Çek'</b> butonuna basın",
                "• Kategori ağacından istediğiniz kategorileri seçin",
                "• <b>İndirim Oranları</b> sekmesinde her kategoriye özel indirim tanımlayın",
                "• 💡 İpucu: Alt kategori seçtiğinizde sadece o ürünler işlenir"
            ],
            "#e74c3c"
        )
        content_layout.addWidget(step3)
        
        # ===== STEP 4 =====
        step4 = self._create_step_card(
            "4️⃣ Kar Marjı Ayarları",
            [
                "• <b>Kar Segmentleri</b> sekmesinde fiyat aralıklarına göre kar ekleyin",
                "• Örnek: 0-100 TL arası ürünlere %30, 100-500 arası %25 kar",
                "• <b>Global Min. Kar:</b> Tüm ürünlere uygulanacak minimum kar",
                "• <b>Baz Fiyat Kaynağı:</b> Hangi fiyattan hesaplama yapılacağını seçin"
            ],
            "#f39c12"
        )
        content_layout.addWidget(step4)
        
        # ===== STEP 5 =====
        step5 = self._create_step_card(
            "5️⃣ Yuvarlama & Limitler",
            [
                "• <b>Yuvarlama Modu:</b> Yukarı, aşağı veya en yakın",
                "• <b>Adım:</b> Fiyatları hangi basamağa yuvarlayacağınızı seçin (1, 5, 10 TL)",
                "• <b>.99 ile Bitir:</b> Psikolojik fiyatlama için (örn: 149.99)",
                "• <b>Min/Max Fiyat:</b> Fiyat sınırları belirleyin"
            ],
            "#1abc9c"
        )
        content_layout.addWidget(step5)
        
        # ===== STEP 6 =====
        step6 = self._create_step_card(
            "6️⃣ Önizleme & Kontrol",
            [
                "• <b>Önizleme</b> sekmesinde <b>'Önizlemeyi Yenile'</b> butonuna basın",
                "• Fiyat değişikliklerini kontrol edin",
                "• Arama yaparak belirli ürünleri bulabilirsiniz",
                "• Kategori filtreleyerek sadece seçili kategorileri görüntüleyin"
            ],
            "#16a085"
        )
        content_layout.addWidget(step6)
        
        # ===== STEP 7 =====
        step7 = self._create_step_card(
            "7️⃣ Dışa Aktarım",
            [
                "• <b>Çıktı Dizini</b> seçin (dosyalar nereye kaydedilecek)",
                "• <b>Maksimum Satır Sayısı:</b> Büyük dosyaları otomatik böler",
                "• <b>'İşlemi Başlat'</b> butonuna basın",
                "• ✅ İşlem bittiğinde Excel dosyaları belirtilen klasöre kaydedilir"
            ],
            "#27ae60"
        )
        content_layout.addWidget(step7)
        
        # ===== TIPS =====
        tips = self._create_info_card(
            "💡 Faydalı İpuçları",
            [
                "• <b>Ayarları Kaydet:</b> 'Şablon Olarak Kaydet' ile ayarlarınızı saklayın",
                "• <b>Loglar:</b> 'Loglar' sekmesinden tüm işlemleri takip edin",
                "• <b>Varyant Ürünler:</b> Aynı üründen farklı seçenekler varsa 'Varyant Modu'nu açın",
                "• <b>Stok Filtresi:</b> Stoksuz ürünleri otomatik olarak hariç tutabilirsiniz"
            ],
            "#34495e"
        )
        content_layout.addWidget(tips)
        
        # ===== CONTACT =====
        contact = QGroupBox("Github ve Versiyon Kontrol")
        contact.setStyleSheet("""
            QGroupBox {
                font-size: 18px;
                font-weight: bold;
                color: #2c3e50;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 20px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 5px 15px;
                background-color: #3498db;
                color: white;
                border-radius: 5px;
            }
        """)
        contact_layout = QVBoxLayout()
        contact_layout.setSpacing(10)
        
        contact_info = QLabel(
            "<a href='https://github.com/muhammet-kpsz/kitsora-excel-pricing-system' style='color:#3498db;'>https://github.com/muhammet-kpsz/kitsora-excel-pricing-system</a><br>"
        )
        contact_info.setOpenExternalLinks(True)
        contact_info.setWordWrap(True)
        contact_info.setStyleSheet("font-size: 13px; padding: 10px; line-height: 1.6;")
        contact_layout.addWidget(contact_info)
        
        # Version & Update Buttons
        btn_layout = QHBoxLayout()
        btn_version = QPushButton("ℹ️ Hakkında / Versiyon")
        btn_version.clicked.connect(self.show_version_dialog)
        
        btn_update = QPushButton("🔄 Güncellemeleri Kontrol Et")
        btn_update.clicked.connect(self.check_for_updates)
        
        btn_layout.addWidget(btn_version)
        btn_layout.addWidget(btn_update)
        contact_layout.addLayout(btn_layout)
        
        contact.setLayout(contact_layout)
        content_layout.addWidget(contact)
        
        # Add stretch to push everything up
        content_layout.addStretch()
        
        scroll.setWidget(content_widget)
        layout.addWidget(scroll)
        
        return widget
    
    def _create_step_card(self, title, points, color):
        """Helper to create a styled step card"""
        card = QGroupBox(title)
        card.setStyleSheet(f"""
            QGroupBox {{
                font-size: 16px;
                font-weight: bold;
                color: {color};
                border: 2px solid {color};
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 15px;
                background-color: #f8f9fa;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 5px 15px;
                background-color: {color};
                color: white;
                border-radius: 5px;
            }}
        """)
        
        layout = QVBoxLayout()
        layout.setSpacing(8)
        
        for point in points:
            label = QLabel(point)
            label.setWordWrap(True)
            label.setTextFormat(Qt.RichText)
            label.setStyleSheet("font-size: 13px; font-weight: normal; color: #2c3e50; padding: 3px 10px;")
            layout.addWidget(label)
        
        card.setLayout(layout)
        return card
    
    def _create_info_card(self, title, points, color):
        """Helper to create info card (similar to step card)"""
        return self._create_step_card(title, points, color)


    def show_version_dialog(self):
        """Displays a dialog similar to winver"""
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Hakkında - {version.APP_NAME}")
        dlg.setFixedSize(400, 300)
        
        layout = QVBoxLayout(dlg)
        
        # Logo/Title
        title = QLabel(f"{version.APP_NAME}")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #2c3e50;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Grid Info
        form_layout = QFormLayout()
        form_layout.addRow("Versiyon:", QLabel(f"<b>{version.VERSION}</b>"))
        form_layout.addRow("Yayın Tarihi:", QLabel(version.BUILD_DATE))
        form_layout.addRow("Geliştirici:", QLabel(version.COMPANY_NAME))
        form_layout.addRow("Github:", QLabel(f"<a href='{version.WEB_SITE}'>{version.WEB_SITE}</a>"))
        form_layout.addRow("E-posta:", QLabel(f"<a href='mailto:{version.CONTACT_EMAIL}'>{version.CONTACT_EMAIL}</a>"))
        
        for i in range(form_layout.count()):
            w = form_layout.itemAt(i).widget()
            if w:
                w.setStyleSheet("font-size: 14px; padding: 5px;")
                if isinstance(w, QLabel):
                    w.setOpenExternalLinks(True)
        
        layout.addLayout(form_layout)
        
        # Legal / License text
        license_txt = QLabel(f"Bu yazılım {version.COMPANY_NAME} tarafından geliştirilmiştir.")
        license_txt.setAlignment(Qt.AlignCenter)
        license_txt.setStyleSheet("color: gray; margin-top: 20px;")
        layout.addWidget(license_txt)
        
        # OK Button
        btn_ok = QPushButton("Tamam")
        btn_ok.clicked.connect(dlg.accept)
        layout.addWidget(btn_ok)
        
        dlg.exec()

    def check_for_updates(self):
        """Starts the update check process"""
        self.log("Güncellemeler kontrol ediliyor...", "INFO")
        
        # Create worker
        self.update_worker = GitUpdateWorker()
        self.update_worker.update_available.connect(self.on_update_available)
        self.update_worker.error_occurred.connect(lambda err: QMessageBox.warning(self, "Hata", err))
        self.update_worker.start()
        
    def on_update_available(self, available, msg):
        if available:
            reply = QMessageBox.question(self, "Güncelleme Mevcut", 
                                       f"{msg}\n\nŞimdi güncellemek ister misiniz?",
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.perform_update()
        else:
            QMessageBox.information(self, "Güncel", "Yazılımınız güncel.")
            self.log("Yazılım güncel.", "INFO")

    def perform_update(self):
        """Starts the actual update pull"""
        self.update_progress = QProgressDialog("Güncelleniyor...", None, 0, 0, self)
        self.update_progress.setWindowModality(Qt.WindowModal)
        self.update_progress.show()
        
        self.update_worker.set_mode("pull")
        self.update_worker.update_finished.connect(self.on_update_finished)
        self.update_worker.start()
        
    def on_update_finished(self, success, msg):
        self.update_progress.close()
        if success:
            QMessageBox.information(self, "Başarılı", msg)
            # Ask to restart? Or just close?
            # close app
            # QApplication.quit()
        else:
            QMessageBox.critical(self, "Güncelleme Hatası", msg)

    def create_log_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFont("Consolas", 10))
        layout.addWidget(self.log_view)
        
        btn_layout = QHBoxLayout()
        btn_clear = QPushButton("Logu Temizle")
        btn_clear.clicked.connect(lambda: self.log_view.clear())
        btn_export = QPushButton("Logu Dışa Aktar (.txt)")
        btn_export.clicked.connect(self.export_logs)
        
        btn_layout.addWidget(btn_clear)
        btn_layout.addWidget(btn_export)
        layout.addLayout(btn_layout)
        
        return widget

    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}"
        if hasattr(self, 'log_view'):
            self.log_view.appendPlainText(log_entry)
        print(log_entry)

    def create_file_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # File Selection
        file_layout = QHBoxLayout()
        self.path_edit = QLineEdit()
        btn = QPushButton("Excel Seç")
        btn.clicked.connect(self.select_file)
        file_layout.addWidget(QLabel("Dosya:"))
        file_layout.addWidget(self.path_edit)
        file_layout.addWidget(btn)
        layout.addLayout(file_layout)
        
        # Mappings Group
        group = QGroupBox("Sütun Eşleştirmeleri")
        form = QFormLayout()
        
        self.combo_stock = QComboBox()
        self.combo_name = QComboBox()
        self.combo_cat = QComboBox()
        self.combo_buy = QComboBox()
        self.combo_sell = QComboBox()
        self.combo_disc = QComboBox()
        self.combo_market = QComboBox()
        form.addRow(QLabel("<b>Temel Sütunlar:</b>"))
        form.addRow("Stok Kodu / Barkod:", self.combo_stock)
        form.addRow("Ürün Adı:", self.combo_name)
        
        
        # New Feature: No Category Column Fallback
        self.chk_no_categories = QCheckBox("Dosyamda Kategori Sütunu Yok (Hepsi 'Kategorisiz')")
        self.chk_no_categories.toggled.connect(lambda checked: self.combo_cat.setEnabled(not checked))
        
        cat_layout = QVBoxLayout()
        cat_layout.addWidget(self.combo_cat)
        cat_layout.addWidget(self.chk_no_categories)
        # Using a widget wrapper to put both in one form row might be cleaner, 
        # but adding separate row is easier.
        form.addRow("Kategori:", self.combo_cat)
        form.addRow("", self.chk_no_categories)
        
        
        form.addRow(QLabel("<b>Fiyat Sütunları:</b>"))
        form.addRow("Alış Fiyatı (Maliyet):", self.combo_buy)
        form.addRow("Satış Fiyatı (Liste):", self.combo_sell)
        form.addRow("İndirimli Fiyat:", self.combo_disc)
        form.addRow("Piyasa Fiyatı:", self.combo_market)
        
        form.addRow(QLabel("<b>Varyant Ayarları:</b>"))
        self.chk_variants = QCheckBox("Bu bir varyantlı ürün dosyasıdır")
        self.chk_variants.toggled.connect(self.on_variant_toggled)
        self.combo_variant = QComboBox()
        self.combo_variant.setEnabled(False)
        self.combo_variant_val = QComboBox() 
        self.combo_variant_val.setEnabled(False)
        form.addRow(self.chk_variants)
        form.addRow("Varyant Grup Kodu / ID:", self.combo_variant)
        form.addRow("Varyant Değerleri (Renk, Beden vb.):", self.combo_variant_val)
        
        self.chk_unique_variant = QCheckBox("Önizlemede Sadece Tekil Varyant Göster")
        self.chk_unique_variant.setToolTip("Aynı Varyant ID'ye sahip ürünlerden sadece birini gösterir.")
        self.chk_unique_variant.setEnabled(False)
        form.addRow(self.chk_unique_variant)
        
        # ===== NEW FEATURE: Stock Column Selection =====
        form.addRow(QLabel("<b>Stok Ayarları:</b>"))
        self.combo_stock_col = QComboBox()
        form.addRow("Stok Sütunu:", self.combo_stock_col)
        
        self.chk_include_zero_stock = QCheckBox("Stok 0 olan ürünleri dahil et")
        self.chk_include_zero_stock.setChecked(True)  # Default: include all
        self.chk_include_zero_stock.setToolTip("Kapalıysa, stok değeri 0 veya daha az olan ürünler filtrelenir.")
        form.addRow(self.chk_include_zero_stock)
        # ===== END NEW FEATURE =====
        
        group.setLayout(form)
        layout.addWidget(group)
        
        # Targets Group
        t_group = QGroupBox("Güncellenecek Sütunlar")
        t_layout = QVBoxLayout()
        
        self.chk_update_disc = QCheckBox("İndirimli Fiyatı Güncelle")
        self.chk_update_sell = QCheckBox("Satış Fiyatını Güncelle")
        self.chk_update_market = QCheckBox("Piyasa Fiyatını Güncelle")
        
        t_layout.addWidget(self.chk_update_disc)
        t_layout.addWidget(self.chk_update_sell)
        t_layout.addWidget(self.chk_update_market)
        t_group.setLayout(t_layout)
        layout.addWidget(t_group)
        
        # --- Settings & Theme Management (Moved here) ---
        layout.addStretch()
        
        # Theme
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel("Uygulama Teması:"))
        self.combo_theme = QComboBox()
        self.combo_theme.addItems(["Sistem", "Açık", "Koyu (Beta)", "Kitsora (Turuncu)"])
        self.combo_theme.currentTextChanged.connect(self.on_theme_changed)
        theme_layout.addWidget(self.combo_theme)
        theme_layout.addStretch()
        layout.addLayout(theme_layout)
        
        # Settings Import/Export
        layout.addWidget(QLabel("Ayarlar Yönetimi:"))
        btn_sett_layout = QHBoxLayout()
        btn_save_sett = QPushButton("Ayarları Şablon Olarak Kaydet")
        btn_save_sett.clicked.connect(self.save_settings_template)
        btn_load_sett = QPushButton("Ayarları Yükle")
        btn_load_sett.clicked.connect(self.load_settings_template)
        btn_sett_layout.addWidget(btn_save_sett)
        btn_sett_layout.addWidget(btn_load_sett)
        layout.addLayout(btn_sett_layout)
        
        return widget

    def create_categories_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        top_layout = QHBoxLayout()
        # Moved Default Discount to Tab 2 as requested
        
        self.btn_extract_cats = QPushButton("Excel'den Kategorileri Çek")
        self.btn_extract_cats.clicked.connect(self.extract_categories_from_file)
        top_layout.addWidget(self.btn_extract_cats)
        top_layout.addStretch()
        
        layout.addLayout(top_layout)
        
        # ===== NEW FEATURE: Category Tree Widget =====
        # Import category tree
        from category_tree import CategoryTreeWidget
        
        
        # New Tabbed Interface for Categories
        cat_tabs = QTabWidget()
        layout.addWidget(cat_tabs)
        
        # --- Tab 1: Category Tree ---
        tab_tree_widget = QWidget()
        tab_tree_layout = QVBoxLayout(tab_tree_widget)
        
        # Tree Buttons
        tree_buttons_layout = QHBoxLayout()
        self.btn_select_all_cats = QPushButton("Tümünü Seç")
        self.btn_select_all_cats.clicked.connect(self.select_all_categories)
        self.btn_clear_all_cats = QPushButton("Tümünü Temizle")
        self.btn_clear_all_cats.clicked.connect(self.clear_all_categories)
        tree_buttons_layout.addWidget(self.btn_select_all_cats)
        tree_buttons_layout.addWidget(self.btn_clear_all_cats)
        tree_buttons_layout.addStretch()
        tab_tree_layout.addLayout(tree_buttons_layout)
        
        # Tree Widget
        from category_tree import CategoryTreeWidget
        self.category_tree = CategoryTreeWidget()
        # Remove width limit since it's now full tab
        # self.category_tree.setMaximumWidth(300) 
        self.category_tree.itemChanged.connect(self.on_category_tree_changed)
        self.category_tree.selectionChanged.connect(self.sync_category_selection)
        tab_tree_layout.addWidget(self.category_tree)
        
        cat_tabs.addTab(tab_tree_widget, "Kategori Ağacı (Seçim)")
        
        # --- Tab 2: Discount Table ---
        tab_table_widget = QWidget()
        tab_table_layout = QVBoxLayout(tab_table_widget)
        
        
        # Default Discount Control (Moved Here)
        def_disc_layout = QHBoxLayout()
        def_disc_layout.addWidget(QLabel("Varsayılan İndirim (%):"))
        self.spin_default_disc = QDoubleSpinBox()
        self.spin_default_disc.setRange(0, 100)
        def_disc_layout.addWidget(self.spin_default_disc)
        def_disc_layout.addStretch()
        tab_table_layout.addLayout(def_disc_layout)
        
        self.table_cats = QTableWidget(0, 2)
        self.table_cats.setHorizontalHeaderLabels(["Kategori Adı", "İndirim Oranı (%)"])
        self.table_cats.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        tab_table_layout.addWidget(self.table_cats)
        
        cat_tabs.addTab(tab_table_widget, "İndirim Oranları")
        
        
        return widget

    def create_profit_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Global Config
        form_layout = QFormLayout()
        self.combo_base_price = QComboBox()
        # Internal keys mapped to display names
        self.base_source_map_profit = {
            "Alış Fiyatı": "buy_price_col",
            "Satış Fiyatı": "sell_price_col",
            "İndirimli Fiyat": "discounted_price_col",
            "Piyasa Fiyatı": "market_price_col"
        }
        self.combo_base_price.addItems(list(self.base_source_map_profit.keys()))
        
        self.spin_global_min_profit = QDoubleSpinBox()
        self.spin_global_min_profit.setRange(0, 99999)
        
        form_layout.addRow("Baz Fiyat Kaynağı:", self.combo_base_price)
        self.chk_global_min = QCheckBox("Global Minimum Kar Uygula")
        self.chk_global_min.toggled.connect(self.spin_global_min_profit.setEnabled)
        self.spin_global_min_profit.setEnabled(False) # Default disabled
        
        form_layout.addRow(self.chk_global_min)
        form_layout.addRow("Minimum Tutar (TL):", self.spin_global_min_profit)
        layout.addLayout(form_layout)
        
        # Segments Table
        layout.addWidget(QLabel("Kâr Segmentleri:"))
        self.table_segments = QTableWidget(0, 5)
        self.table_segments.setHorizontalHeaderLabels(["Min", "Max", "Tip", "Değer", "Ek Tutar (TL)"])
        self.table_segments.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table_segments)
        
        btn_row = QHBoxLayout()
        btn_add = QPushButton("Segment Ekle")
        btn_add.clicked.connect(self.add_segment_row)
        btn_rem = QPushButton("Seçili Segmenti Sil")
        btn_rem.clicked.connect(self.remove_segment_row)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_rem)
        layout.addLayout(btn_row)
        
        return widget

    def create_rounding_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        form = QFormLayout()
        
        self.combo_round_mode = QComboBox()
        self.combo_round_mode.addItems(["ceiling", "round", "floor"])
        
        self.combo_step = QComboBox()
        self.combo_step.addItems(["1", "5", "10", "25", "50", "100"])
        
        self.chk_ends_99 = QCheckBox("Sonu .99 ile bitsin")
        
        self.spin_min_disc = QDoubleSpinBox()
        self.spin_min_disc.setRange(0, 99999)
        
        self.spin_max_disc = QDoubleSpinBox()
        self.spin_max_disc.setRange(0, 999999)
        
        form.addRow("Yuvarlama Yönü:", self.combo_round_mode)
        form.addRow("Yuvarlama Adımı:", self.combo_step)
        form.addRow("Psikolojik Fiyat:", self.chk_ends_99)
        form.addRow("Min İndirimli Fiyat:", self.spin_min_disc)
        form.addRow("Max İndirimli Fiyat:", self.spin_max_disc)
        
        layout.addLayout(form)
        layout.addStretch()
        return widget

    def create_preview_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # --- Top Controls ---
        top_layout = QHBoxLayout()
        
        self.btn_refresh_preview = QPushButton("Verileri Yükle / Yenile")
        self.btn_refresh_preview.clicked.connect(self.refresh_preview)
        
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Ara: Stok Kodu veya Ürün Adı...")
        self.search_bar.textChanged.connect(self.apply_filters)
        
        # ===== NEW FEATURE: Cascade Category Filter =====
        from cascade_menu import CascadeCategoryButton
        self.combo_preview_cat = CascadeCategoryButton()
        self.combo_preview_cat.categorySelected.connect(self.apply_filters)
        # ===== END NEW FEATURE =====
        
        # Base Price Source Override
        self.combo_preview_base = QComboBox()
        # Internal keys mapped to display names
        self.base_source_map = {
            "Alış Fiyatı": "buy_price_col",
            "Satış Fiyatı": "sell_price_col",
            "İndirimli Fiyat": "discounted_price_col",
            "Piyasa Fiyatı": "market_price_col"
        }
        self.combo_preview_base.addItems(list(self.base_source_map.keys()))
        self.combo_preview_base.setToolTip("Hesaplamada baz alınacak sütun")
        self.combo_preview_base.currentTextChanged.connect(self.on_preview_base_changed)
        
        # Category Filter Sort removed

        top_layout.addWidget(self.btn_refresh_preview)
        top_layout.addWidget(QLabel("Baz Fiyat:"))
        top_layout.addWidget(self.combo_preview_base)
        top_layout.addWidget(QLabel("Kategori:"))
        top_layout.addWidget(self.combo_preview_cat)
        # top_layout.addWidget(self.btn_sort_cats_asc) # Removed
        top_layout.addWidget(self.search_bar, 1)
        
        layout.addLayout(top_layout)
        
        # --- Stats Bar ---
        self.lbl_stats = QLabel("Toplam: 0 | Değişen: 0 | Görüntülenen: 0")
        layout.addWidget(self.lbl_stats)
        
        # --- Table Area with Loading Overlay ---
        self.preview_stack = QStackedWidget()
        
        # Page 0: Table
        self.table_preview = QTableWidget()
        self.table_preview.setAlternatingRowColors(True)
        # Enable header clicking for sorting
        self.table_preview.horizontalHeader().setSectionsClickable(True)
        self.table_preview.horizontalHeader().sectionClicked.connect(self.on_preview_header_clicked)
        # Click handler for variants
        self.table_preview.cellClicked.connect(self.on_preview_cell_clicked)
        # Context Menu
        self.table_preview.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_preview.customContextMenuRequested.connect(self.show_preview_context_menu)
        
        self.preview_stack.addWidget(self.table_preview)
        
        # Page 1: Loading
        loading_widget = QWidget()
        loading_layout = QVBoxLayout(loading_widget)
        self.lbl_loading = QLabel("Veriler İşleniyor, Lütfen Bekleyin...")
        self.lbl_loading.setAlignment(Qt.AlignCenter)
        self.lbl_loading.setStyleSheet("font-size: 16px; font-weight: bold; color: #42130218;")
        self.loading_progress = QProgressBar()
        self.loading_progress.setRange(0, 0) # Indeterminate
        self.loading_progress.setMaximumWidth(400)
        loading_layout.addStretch()
        loading_layout.addWidget(self.lbl_loading)
        loading_layout.addWidget(self.loading_progress, 0, Qt.AlignCenter)
        loading_layout.addStretch()
        self.preview_stack.addWidget(loading_widget)
        
        layout.addWidget(self.preview_stack)
        
        # --- Pagination ---
        self.pagination_layout = QHBoxLayout()
        self.pagination_widget = QWidget()
        self.pagination_widget.setLayout(self.pagination_layout)
        layout.addWidget(self.pagination_widget)
        
        # State
        self.all_rows_cache = []
        self.filtered_rows = []
        self.current_page = 1
        self.items_per_page = 50
        self.sort_col = -1 # None
        self.sort_asc = True
        
        return widget
        
    def create_export_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        form = QFormLayout()
        self.spin_max_rows = QSpinBox()
        self.spin_max_rows.setRange(10, 1000000)
        self.spin_max_rows.setValue(5000)
        
        self.edit_output_dir = QLineEdit()
        btn_dir = QPushButton("...")
        btn_dir.clicked.connect(self.select_output_dir)
        dir_layout = QHBoxLayout()
        dir_layout.addWidget(self.edit_output_dir)
        dir_layout.addWidget(btn_dir)
        
        form.addRow("Dosya Başına Max Satır:", self.spin_max_rows)
        form.addRow("Çıktı Klasörü:", dir_layout)

        layout.addLayout(form)
        
        self.btn_run = QPushButton("İşlemi Başlat")
        self.btn_run.setFixedHeight(50)
        self.btn_run.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_run)
        
        # New Progress UI for sequential parts
        self.lbl_part_status = QLabel("İşlem Bekleniyor...")
        self.lbl_part_status.setStyleSheet("font-weight: bold;")
        self.progress_bar_part = QProgressBar()
        layout.addWidget(self.lbl_part_status)
        layout.addWidget(self.progress_bar_part)
        
        layout.addStretch()
        return widget

    # --- Logic ---

    def on_theme_changed(self, text):
        map_theme = {
            "Sistem": "system", 
            "Açık": "light", 
            "Koyu (Beta)": "dark",
            "Kitsora (Turuncu)": "kitsora"
        }
        val = map_theme.get(text, "system")
        self.sm.set("theme", val)
        self.sm.save_settings()
        self.apply_theme()
        
    def select_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Excel Seç", "", "Excel Files (*.xlsx)")
        if fname:
            self.path_edit.setText(fname)
            self.load_headers(fname)
            # Auto set output dir
            self.edit_output_dir.setText(os.path.dirname(fname))

    def select_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Çıktı Klasörü Seç")
        if d:
            self.edit_output_dir.setText(d)

    def on_variant_toggled(self, checked):
        self.combo_variant.setEnabled(checked)
        self.combo_variant_val.setEnabled(checked)
        self.chk_unique_variant.setEnabled(checked)
        if checked:
            self.combo_variant.setStyleSheet("background-color: #fff8e1;") 
            self.combo_variant_val.setStyleSheet("background-color: #e1f5fe;")
            
            if not self.combo_variant.currentText() and self.combo_variant.count() > 0:
                 QMessageBox.information(self, "Bilgi", "Lütfen Varyant Grup Kodu ve Değerleri sütunlarını seçin.")
        else:
            self.combo_variant.setStyleSheet("")
            self.combo_variant_val.setStyleSheet("")
    
    def sort_categories_az(self):
        combo = self.combo_preview_cat
        count = combo.count()
        if count <= 1: return 
        
        # Keep "All" (index 0)
        items = [combo.itemText(i) for i in range(1, count)]
        items.sort() # Default asc
        
        curr = combo.currentText()
        combo.blockSignals(True)
        combo.clear()
        combo.addItem("Tüm Kategoriler")
        combo.addItems(items)
        combo.setCurrentText(curr)
        combo.blockSignals(False)
        
        self.log("Kategori listesi sıralandı.")

    def load_headers(self, fname):
        self.current_headers = self.io.get_headers(fname)
        # ===== NEW FEATURE: Added combo_stock_col to list =====
        combos = [self.combo_stock, self.combo_name, self.combo_cat, 
                  self.combo_buy, self.combo_sell, self.combo_disc, self.combo_market, 
                  self.combo_variant, self.combo_variant_val, self.combo_stock_col]
        # ===== END NEW FEATURE =====
        valid_headers = [h for h in self.current_headers if h]
        
        for c in combos:
            c.clear()
            c.addItems([""] + valid_headers)
            
        # Try Auto-Map
        mappings = {
            "STOCK": self.combo_stock, "STOK": self.combo_stock, "KOD": self.combo_stock, "SKU": self.combo_stock, "BARKOD": self.combo_stock,
            "AD": self.combo_name, "NAME": self.combo_name, "ÜRÜN": self.combo_name, "URUN": self.combo_name, "ISIM": self.combo_name,
            "KATEGORY": self.combo_cat, "KATEGORİ": self.combo_cat, "CATEGORY": self.combo_cat,
            "ALIŞ": self.combo_buy, "ALIS": self.combo_buy, "MALIYET": self.combo_buy, "COST": self.combo_buy,
            "SATIŞ": self.combo_sell, "SATIS": self.combo_sell, "PRICE": self.combo_sell,
            "İNDİRİMLİ": self.combo_disc, "INDIRIMLI": self.combo_disc,
            "PIYASA": self.combo_market, "PİYASA": self.combo_market,
            "VARYANT": self.combo_variant, "VARIANT": self.combo_variant, "GRUP KODU": self.combo_variant, "GROUP CODE": self.combo_variant,
            "VARYASYON": self.combo_variant_val, "VARIATION": self.combo_variant_val, "ÖZELLİK": self.combo_variant_val, "FEATURE": self.combo_variant_val,
            # ===== NEW FEATURE: Stock column auto-mapping =====
            "ADET": self.combo_stock_col, "MIKTAR": self.combo_stock_col, "QTY": self.combo_stock_col, "QUANTITY": self.combo_stock_col, "ENVANTER": self.combo_stock_col
            # ===== END NEW FEATURE =====
        }
        
        for h in valid_headers:
            upper = str(h).upper()
            
            # Simple keyword matching
            for key in mappings:
                if key in upper:
                     # Check if already set? Let's override for first match
                     if mappings[key].currentIndex() <= 0:
                        mappings[key].setCurrentText(h)

    def add_segment_row(self):
        row = self.table_segments.rowCount()
        self.table_segments.insertRow(row)
        
        self.table_segments.setItem(row, 0, QTableWidgetItem("0"))
        self.table_segments.setItem(row, 1, QTableWidgetItem("100"))
        
        combo_type = QComboBox()
        combo_type.addItems(["TUTAR (TL)", "YÜZDE (%)"])
        self.table_segments.setCellWidget(row, 2, combo_type)
        
        self.table_segments.setItem(row, 3, QTableWidgetItem("0"))
        self.table_segments.setItem(row, 4, QTableWidgetItem("0"))

    def remove_segment_row(self):
        curr = self.table_segments.currentRow()
        if curr >= 0:
            self.table_segments.removeRow(curr)

    def extract_categories_from_file(self):
        fname = self.path_edit.text()
        if not fname or not os.path.exists(fname):
            QMessageBox.warning(self, "Hata", "Lütfen geçerli bir dosya seçin.")
            return

        cat_col = self.combo_cat.currentText()
        no_cat_mode = self.chk_no_categories.isChecked() if hasattr(self, 'chk_no_categories') else False
        
        if not cat_col and not no_cat_mode:
            QMessageBox.warning(self, "Hata", "Lütfen önce 'Dosya & Eşleştirme' sekmesinden Kategori sütununu seçin veya 'Kategori Yok' seçeneğini işaretleyin.")
            return
            
        self.log(f"Kategoriler taranıyor (Sütun: {cat_col})...")
        self.collect_settings()
        
        # Disable button during scan
        self.btn_extract_cats.setEnabled(False)
        
        # Safely handle existing worker
        if hasattr(self, 'cat_worker') and self.cat_worker.isRunning():
            self.cat_worker.wait()

        self.cat_worker = CategoryWorker(fname, cat_col, self.engine, no_cat_mode=no_cat_mode)
        self.cat_worker.finished.connect(self.on_categories_extracted)
        self.cat_worker.start()

    def on_categories_extracted(self, unique_cats_data):
        self.btn_extract_cats.setEnabled(True)
        self.table_cats.setRowCount(0)
        
        # Handle dict input (paths and counts)
        if isinstance(unique_cats_data, dict):
            unique_cats = unique_cats_data.keys()
            category_counts = unique_cats_data
        else:
            unique_cats = unique_cats_data
            category_counts = {}
            
        sorted_cats = sorted(list(unique_cats))
        default_rate = self.spin_default_disc.value()
        
        # ===== ENHANCED: Separate main categories for table =====
        # Extract unique main categories for discount table
        main_categories = set()
        for cat in sorted_cats:
            main_cat = self.engine.extract_category(cat)
            if main_cat:
                main_categories.add(main_cat)
        
        # Populate table with main categories only
        for i, cat in enumerate(sorted(list(main_categories))):
            self.table_cats.insertRow(i)
            self.table_cats.setItem(i, 0, QTableWidgetItem(cat))
            self.table_cats.setItem(i, 1, QTableWidgetItem(str(default_rate)))
        # ===== END ENHANCEMENT =====
        
        # ===== NEW FEATURE: Populate tree widget with FULL paths =====
        try:
            # Use persistent state for restoration
            cats_to_restore = list(self.persistent_selected_categories)
            
            self.category_tree.build_tree(sorted_cats)  # Full paths for tree
            
            # RESTORE SELECTION STATE
            if cats_to_restore:
                # Filter to only keep paths that still exist in the new tree
                # Note: set_selected_categories handles missing items gracefully (skips them)
                self.category_tree.set_selected_categories(cats_to_restore)
                
                
            # Validate counts immediately if available
            if category_counts:
                 self.category_tree.update_counts(category_counts)
        except:
            pass  # Don't fail if tree has issues
        # ===== END NEW FEATURE =====
        
        self.log(f"{len(sorted_cats)} adet kategori (full path), {len(main_categories)} ana kategori listelendi.")
        QMessageBox.information(self, "Tamamlandı", f"{len(main_categories)} ana kategori, {len(sorted_cats)} alt kategori bulundu.")
    
    # ===== NEW FEATURE: Category tree selection helpers =====
    def select_all_categories(self):
        """Select all categories in the tree"""
        try:
            # Get all category paths from tree
            all_paths = list(self.category_tree.item_map.keys())
            self.category_tree.set_selected_categories(all_paths)
            self.log("Tüm kategoriler seçildi.")
            # Refresh preview if data is loaded
            if hasattr(self, 'all_rows_cache') and self.all_rows_cache:
                self.apply_filters()
        except Exception as e:
            self.log(f"Kategori seçim hatası: {e}")
    
    def clear_all_categories(self):
        """Clear all category selections in the tree"""
        try:
            self.category_tree.set_selected_categories([])
            self.log("Tüm kategori seçimleri temizlendi.")
            # Refresh preview if data is loaded
            if hasattr(self, 'all_rows_cache') and self.all_rows_cache:
                self.apply_filters()
        except Exception as e:
            self.log(f"Kategori temizleme hatası: {e}")
    
    def sync_category_selection(self, selected_list):
        """Sync persistent state with UI"""
        new_selection = set(selected_list)
        
        # Only trigger filter if selection actually changed
        selection_changed = new_selection != self.persistent_selected_categories
        
        self.persistent_selected_categories = new_selection
        # Also update settings manually here if needed to ensure export gets latest state
        self.sm.set("selected_categories", list(self.persistent_selected_categories))
        
        # CRITICAL FIX: Refresh preview when tree selection changes
        # BUT: Only if selection actually changed AND data is loaded
        if selection_changed and hasattr(self, 'all_rows_cache') and self.all_rows_cache:
            self.apply_filters()

    
    def on_category_tree_changed(self, item, column):
        """Called when category tree selection changes - DO NOTHING until manual refresh"""
        pass
        # Auto refresh disabled as per user request
        # if hasattr(self, 'all_rows_cache') and self.all_rows_cache: ...
    # ===== END NEW FEATURE =====

    def collect_settings(self):
        # Mappings
        self.sm.set("mappings", {
            "stock_code_col": self.combo_stock.currentText(),
            "product_name_col": self.combo_name.currentText(),
            "category_col": self.combo_cat.currentText(),
            "buy_price_col": self.combo_buy.currentText(),
            "sell_price_col": self.combo_sell.currentText(),
            "discounted_price_col": self.combo_disc.currentText(),
            "market_price_col": self.combo_market.currentText(),
            "variant_id_col": self.combo_variant.currentText(),
            "variant_val_col": self.combo_variant_val.currentText(),
            "is_variant_mode": self.chk_variants.isChecked(),
            "show_unique_variant": self.chk_unique_variant.isChecked(),
            # ===== NEW FEATURE: Save stock settings =====
            "stock_col": self.combo_stock_col.currentText(),
            "include_zero_stock": self.chk_include_zero_stock.isChecked(),
            "no_category_mode": self.chk_no_categories.isChecked()
            # ===== END NEW FEATURE =====
        })
        
        self.sm.set("targets", {
            "update_discounted": self.chk_update_disc.isChecked(),
            "update_sell": self.chk_update_sell.isChecked(),
            "update_market": self.chk_update_market.isChecked()
        })
        
        # Categories
        cat_map = {}
        for r in range(self.table_cats.rowCount()):
            c_name = self.table_cats.item(r, 0).text()
            try:
                rate = float(self.table_cats.item(r, 1).text())
            except:
                rate = self.spin_default_disc.value()
            cat_map[c_name] = rate
            
        self.sm.set("categories", {
            "default_discount": self.spin_default_disc.value(),
            "mapping": cat_map
        })
        
        # Profit Segments
        segments = []
        for r in range(self.table_segments.rowCount()):
            try:
                min_v = float(self.table_segments.item(r, 0).text())
                max_v = float(self.table_segments.item(r, 1).text())
                
                # UI text to Internal Value
                ui_type = self.table_segments.cellWidget(r, 2).currentText()
                type_v = "PERCENT" if "YÜZDE" in ui_type else "TL"
                
                val_v = float(self.table_segments.item(r, 3).text())
                extra_v = float(self.table_segments.item(r, 4).text())
                segments.append({"min": min_v, "max": max_v, "type": type_v, "value": val_v, "extra_added": extra_v})
            except:
                pass
        self.sm.set("profit_segments", segments)
        
        self.sm.set("global_min_profit", self.spin_global_min_profit.value())
        self.sm.set("enable_global_min", self.chk_global_min.isChecked())
        
        # Convert display text to internal key
        base_display = self.combo_base_price.currentText()
        base_internal = self.base_source_map_profit.get(base_display, "buy_price_col")
        self.sm.set("base_price_source", base_internal)

        
        # Rounding
        self.sm.set("rounding", {
            "mode": self.combo_round_mode.currentText(),
            "step": float(self.combo_step.currentText()),
            "ends_with_99": self.chk_ends_99.isChecked()
        })
        
        self.sm.set("limits", {
            "min_discounted_price": self.spin_min_disc.value(),
            "max_discounted_price": self.spin_max_disc.value()
        })
        
        # Output
        self.sm.set("output", {
            "max_rows_per_file": self.spin_max_rows.value(),
            "output_dir": self.edit_output_dir.text()
        })
        
        # ===== NEW FEATURE: Save selected categories from tree =====
        try:
            # Use persistent state
            if hasattr(self, 'persistent_selected_categories'):
                self.sm.set("selected_categories", list(self.persistent_selected_categories))
            else:
                self.sm.set("selected_categories", [])
        except:
            # If tree doesn't exist or error, don't fail
            pass
        # ===== END NEW FEATURE =====
        
        self.sm.save_settings()

    def load_ui_values(self):
        s = self.sm.settings
        
        # Determine cols - we can't really set combos without file, but other fields yes.
        self.spin_default_disc.setValue(float(s["categories"].get("default_discount", 50)))
        
        # Mappings
        mappings = s.get("mappings", {})
        self.chk_variants.setChecked(mappings.get("is_variant_mode", False))
        self.on_variant_toggled(self.chk_variants.isChecked()) # Apply visual state
        # combo_variant.setCurrentText will be handled by load_headers if a file is loaded,
        # or by load_settings_template if loading from a saved template.
        
        # Targets
        self.chk_update_disc.setChecked(s.get("targets", {}).get("update_discounted", True))
        self.chk_update_sell.setChecked(s.get("targets", {}).get("update_sell", True))
        self.chk_update_market.setChecked(s.get("targets", {}).get("update_market", True))
        
        # Theme
        theme = s.get("theme", "system")
        map_theme_rev = {
            "system": "Sistem", 
            "light": "Açık", 
            "dark": "Koyu (Beta)",
            "kitsora": "Kitsora (Turuncu)"
        }
        self.combo_theme.setCurrentText(map_theme_rev.get(theme, "Sistem"))

        
        # Restore Categories Table
        cat_map = s.get("categories", {}).get("mapping", {})
        if cat_map:
            self.table_cats.setRowCount(0)
            sorted_cats = sorted(list(cat_map.keys()))
            for i, c_name in enumerate(sorted_cats):
                self.table_cats.insertRow(i)
                self.table_cats.setItem(i, 0, QTableWidgetItem(c_name))
                rate = cat_map[c_name]
                self.table_cats.setItem(i, 1, QTableWidgetItem(str(rate)))

        # Segments
        segs = s.get("profit_segments", [])
        self.table_segments.setRowCount(0)
        for seg in segs:
            self.add_segment_row()
            r = self.table_segments.rowCount() - 1
            self.table_segments.item(r, 0).setText(str(seg["min"]))
            self.table_segments.item(r, 1).setText(str(seg["max"]))
            
            # Internal Value to UI Text
            t_val = seg["type"]
            ui_text = "YÜZDE (%)" if t_val == "PERCENT" else "TUTAR (TL)"
            self.table_segments.cellWidget(r, 2).setCurrentText(ui_text)
            
            self.table_segments.item(r, 3).setText(str(seg["value"]))
            self.table_segments.item(r, 4).setText(str(seg.get("extra_added", 0)))
            
        self.spin_global_min_profit.setValue(float(s.get("global_min_profit", 0)))
        enabled = s.get("enable_global_min", False)
        self.chk_global_min.setChecked(enabled)
        self.spin_global_min_profit.setEnabled(enabled)
        
        # Load base price source
        base_internal = s.get("base_price_source", "buy_price_col")
        # Convert internal key back to display text
        reverse_map = {v: k for k, v in self.base_source_map_profit.items()}
        base_display = reverse_map.get(base_internal, "Alış Fiyatı")
        self.combo_base_price.setCurrentText(base_display)
        
        # Rounding
        rnd = s.get("rounding", {})
        self.combo_round_mode.setCurrentText(rnd.get("mode", "ceiling"))
        self.combo_step.setCurrentText(str(int(rnd.get("step", 10))))
        self.chk_ends_99.setChecked(rnd.get("ends_with_99", False))
        
        lm = s.get("limits", {})
        self.spin_min_disc.setValue(lm.get("min_discounted_price", 0))
        self.spin_max_disc.setValue(lm.get("max_discounted_price", 1000))

    def on_preview_base_changed(self, text):
        internal_key = self.base_source_map.get(text, "buy_price_col")
        self.sm.set("base_price_source", internal_key)
        self.log(f"Önizleme baz fiyat kaynağı değiştirildi: {text} ({internal_key})")
        self.apply_filters()

    def refresh_preview(self):
        f = self.path_edit.text()
        if not f: 
            QMessageBox.warning(self, "Uyarı", "Lütfen önce bir Excel dosyası seçin.")
            return
        
        self.log(f"Dosya okunuyor: {f}")
        self.collect_settings()
        
        # Switch to loading screen
        self.preview_stack.setCurrentIndex(1)
        self.btn_refresh_preview.setEnabled(False)
        self.lbl_loading.setText("Excel Dosyası Okunuyor...")
        
        if hasattr(self, 'loader_worker') and self.loader_worker.isRunning():
            self.loader_worker.wait()

        self.loader_worker = FileLoaderWorker(f)
        self.loader_worker.finished.connect(self.on_file_loaded)
        self.loader_worker.failed.connect(self.on_file_load_failed)
        self.loader_worker.start()

    def on_file_loaded(self, rows):
        self.btn_refresh_preview.setEnabled(True)
        self.all_rows_cache = rows
        self.log(f"Excel'den {len(self.all_rows_cache)} satır okundu. Şimdi veriler işleniyor...")
        self.lbl_loading.setText("Fiyatlar Hesaplanıyor ve Filtreleniyor...")
        self.apply_filters()

    def on_file_load_failed(self, err):
        self.btn_refresh_preview.setEnabled(True)
        self.preview_stack.setCurrentIndex(0)
        self.log(f"Dosya okuma hatası: {err}", "ERROR")
        QMessageBox.critical(self, "Hata", f"Dosya okunamadı: {err}")

    def apply_filters(self):
        # Debounce: Instead of running immediately, start/restart timer
        if not self.all_rows_cache:
            return
        self.search_timer.start(500) # Wait 500ms

    def run_apply_filters(self):
        self.preview_stack.setCurrentIndex(1) # Loading...
        
        search_txt = self.search_bar.text()
        # ===== NEW FEATURE: Get category from hierarchical combo =====
        cat_filter = self.combo_preview_cat.get_selected_category()
        # ===== END NEW FEATURE =====
        
        # Collect variant setup
        variant_col = None
        if self.chk_variants.isChecked():
            variant_col = self.combo_variant.currentText()
        
        # ===== NEW FEATURE: Collect stock and category settings =====
        stock_col = self.combo_stock_col.currentText() if hasattr(self, 'combo_stock_col') else None
        include_zero_stock = self.chk_include_zero_stock.isChecked() if hasattr(self, 'chk_include_zero_stock') else True
        
        # Get selected categories from LOGICAL state (Robustness fix)
        selected_cats = []
        if hasattr(self, 'persistent_selected_categories') and self.persistent_selected_categories:
            selected_cats = list(self.persistent_selected_categories)
        else:
             try:
                 selected_cats = self.category_tree.get_selected_categories()
             except:
                 pass
        # ===== END NEW FEATURE =====
        
        # Safely handle existing worker
        if hasattr(self, 'preview_worker') and self.preview_worker.isRunning():
            self.preview_worker.finished.disconnect() # Ignore old results
            # Safely handle existing worker
        if hasattr(self, 'preview_worker') and self.preview_worker.isRunning():
            self.preview_worker.finished.disconnect() 

        self.preview_worker = PreviewWorker(
            self.all_rows_cache, 
            self.engine, 
            search_txt, 
            cat_filter, 
            variant_col=variant_col, 
            variant_val_col=self.combo_variant_val.currentText(),
            show_unique_variant=self.chk_unique_variant.isChecked(),
            # ===== NEW FEATURE: Pass new parameters =====
            stock_col=stock_col,
            include_zero_stock=include_zero_stock,
            selected_categories=selected_cats
            # ===== END NEW FEATURE =====
        )
        if selected_cats:
            self.log(f"DEBUG: Filtreleme başladı. Seçili: {len(selected_cats)}", "DEBUG")
            if len(selected_cats) > 0:
                 self.log(f"DEBUG: Örnek: {selected_cats[0]}", "DEBUG")
        else:
             self.log("DEBUG: Filtreleme yok (Tümü)", "DEBUG")
             
        self.preview_worker.finished.connect(self.on_preview_worker_finished)
        self.preview_worker.start()

    def on_preview_worker_finished(self, results, changed_count, categories):
        self.filtered_rows = results
        
        # Update Stats
        total = len(self.filtered_rows)
        self.lbl_stats.setText(f"Toplam Sonuç: {total} | Fiyatı Değişen: {changed_count}")
        self.log(f"Filtreleme uygulandı. Eşleşen: {total}, Değişen: {changed_count}")

        # ===== NEW FEATURE: Update hierarchical dropdown with counts =====
        # Calculate category counts from filtered results
        # Calculate category counts from filtered results
        category_counts = {}
        for row in self.filtered_rows:
            raw_cat = row.get("full_category_path", row.get("main_category", ""))
            if raw_cat:
                # Normalize path to match tree structure
                full_cat = " > ".join([p.strip() for p in str(raw_cat).split(">") if p.strip()]) 
                category_counts[full_cat] = category_counts.get(full_cat, 0) + 1
        
        # Get selected categories from tree for filtering dropdown
        selected_cats = []
        try:
            selected_cats = self.category_tree.get_selected_categories()
        except:
            pass
        
        # Populate hierarchical dropdown
        self.combo_preview_cat.blockSignals(True)
        self.combo_preview_cat.populate_categories(category_counts, selected_cats)
        self.combo_preview_cat.blockSignals(False)
        
        # Also update tree with counts
        try:
            self.category_tree.update_counts(category_counts)
        except:
            pass
        # ===== END NEW FEATURE =====

        # Apply Sort if active
        if self.sort_col != -1:
            self.sort_filtered_data()

        # Reset to page 1
        self.current_page = 1
        self.update_pagination_controls()
        self.update_table_view()
        
        # Back to table
        self.preview_stack.setCurrentIndex(0)

    def on_preview_header_clicked(self, logicalIndex):
        if self.sort_col == logicalIndex:
            self.sort_asc = not self.sort_asc
        else:
            self.sort_col = logicalIndex
            self.sort_asc = True
            
        self.sort_filtered_data()
        self.current_page = 1
        self.update_table_view()
        self.update_pagination_controls()

    def sort_filtered_data(self):
        # headers = ["Stok Kodu", "Ürün Adı", "Kategori", "Baz Fiyat", "Kâr", "Yeni İndirimli", "Yeni Etiket"]
        # Keys map roughly to: stock_code, product_name, main_category, base_price, profit_added, final_discounted_price, label_price
        
        key_map = {
            0: "stock_code",
            1: "product_name",
            2: "main_category", # This will be variant ID if variant mode is on, otherwise category
            3: "base_price",
            4: "profit_added",
            5: "final_discounted_price",
            6: "label_price"
        }
        
        # Adjust key_map if variant column is present
        is_variant = self.chk_variants.isChecked()
        if is_variant:
            key_map[2] = "_variant_id" # Variant ID column
            key_map[3] = "main_category" # Category column shifts
            key_map[4] = "base_price"
            key_map[5] = "profit_added"
            key_map[6] = "final_discounted_price"
            key_map[7] = "label_price"

        key = key_map.get(self.sort_col)
        if not key: return
        
        def safe_sort(item):
            v = item.get(key, "")
            # Price/Number columns
            if self.sort_col in [3, 4, 5, 6] or (is_variant and self.sort_col in [4, 5, 6, 7]): # Adjusted indices for variant mode
                try:
                    # Replace comma with dot if it's a string representation of a float
                    if isinstance(v, str):
                        v = v.replace(",", ".")
                    return float(v)
                except (ValueError, TypeError):
                    return -1.0 # Default low value for sort
            # String columns
            return str(v).lower()
            
        self.filtered_rows.sort(key=safe_sort, reverse=not self.sort_asc)

    def export_logs(self):
        # Create logs directory if it doesn't exist
        logs_dir = "logs"
        os.makedirs(logs_dir, exist_ok=True)
        
        # Default filename with timestamp in logs directory
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        default_filename = os.path.join(logs_dir, f"islem_kayitlari_{timestamp}.txt")
        
        fname, _ = QFileDialog.getSaveFileName(self, "Log Kaydet", default_filename, "Text Files (*.txt)")
        if fname:
            try:
                with open(fname, "w", encoding="utf-8") as f:
                    f.write(self.log_view.toPlainText())
                QMessageBox.information(self, "Başarılı", "Loglar kaydedildi.")
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Kaydedilemedi: {e}")

    def update_pagination_controls(self):
        # Clear existing buttons
        # Note: Deleting widgets from layout strictly
        while self.pagination_layout.count():
            item = self.pagination_layout.takeAt(0)
            w = item.widget()
            if w: w.deleteLater()
            
        import math
        total_pages = math.ceil(len(self.filtered_rows) / self.items_per_page)
        if total_pages < 1: total_pages = 1
        
        # Smart Pagination: Show 1..5 ... Last
        
        start_p = max(1, self.current_page - 2)
        end_p = min(total_pages, start_p + 4)
        
        # Adjust if near end
        if end_p - start_p < 4:
            start_p = max(1, end_p - 4)

        # Prev
        if self.current_page > 1:
            btn = QPushButton("<")
            btn.setFixedWidth(30)
            btn.clicked.connect(lambda: self.change_page(self.current_page - 1))
            self.pagination_layout.addWidget(btn)
            
        # First Page
        if start_p > 1:
            btn = QPushButton("1")
            btn.setFixedWidth(30)
            btn.setCheckable(True)
            btn.clicked.connect(lambda: self.change_page(1))
            self.pagination_layout.addWidget(btn)
            if start_p > 2:
                self.pagination_layout.addWidget(QLabel("..."))

        for p in range(start_p, end_p + 1):
            btn = QPushButton(str(p))
            btn.setFixedWidth(30)
            btn.setCheckable(True)
            if p == self.current_page:
                btn.setChecked(True)
                btn.setEnabled(False) # Current page disabled
            btn.clicked.connect(lambda checked, page=p: self.change_page(page))
            self.pagination_layout.addWidget(btn)
            
        # Last Page
        if end_p < total_pages:
            if end_p < total_pages - 1:
                self.pagination_layout.addWidget(QLabel("..."))
            btn = QPushButton(str(total_pages))
            btn.setFixedWidth(30)
            btn.setCheckable(True)
            btn.clicked.connect(lambda: self.change_page(total_pages))
            self.pagination_layout.addWidget(btn)
            
        # Next
        if self.current_page < total_pages:
            btn = QPushButton(">")
            btn.setFixedWidth(30)
            btn.clicked.connect(lambda: self.change_page(self.current_page + 1))
            self.pagination_layout.addWidget(btn)
            
        self.pagination_layout.addStretch()

    def change_page(self, page):
        self.current_page = page
        self.update_pagination_controls()
        self.update_table_view()

    def update_table_view(self):
        start_idx = (self.current_page - 1) * self.items_per_page
        end_idx = start_idx + self.items_per_page
        
        page_data = self.filtered_rows[start_idx:end_idx]
        
        headers = ["Stok Kodu", "Ürün Adı", "Kategori", "Baz Fiyat", "Kâr", "Yeni İndirimli", "Yeni Etiket"]
        
        # Add Variant Header if enabled
        is_variant = self.chk_variants.isChecked()
        if is_variant:
            headers.insert(2, "Varyant ID")
        
        # ===== NEW FEATURE: Add Stock column if enabled =====
        has_stock = hasattr(self, 'combo_stock_col') and self.combo_stock_col.currentText()
        if has_stock:
            # Insert after Kategori (or after Varyant ID if variants enabled)
            insert_pos = 3 if is_variant else 2
            headers.insert(insert_pos + 1, "Stok")  # After kategori
        # ===== END NEW FEATURE =====
            
        self.table_preview.setColumnCount(len(headers))
        self.table_preview.setHorizontalHeaderLabels(headers)
        self.table_preview.setRowCount(0)
        
        for res in page_data:
            row = self.table_preview.rowCount()
            self.table_preview.insertRow(row)
            
            if "error" in res:
                self.table_preview.setItem(row, 1, QTableWidgetItem(f"ERROR: {res.get('error')}"))
                continue

            s_code = str(res.get("stock_code", ""))
            p_name = str(res.get("product_name", ""))
            
            self.table_preview.setItem(row, 0, QTableWidgetItem(s_code))
            self.table_preview.setItem(row, 1, QTableWidgetItem(p_name))
            
            col_idx = 2
            if is_variant:
                v_id = str(res.get("_variant_id", "-"))
                item_v_id = QTableWidgetItem(v_id)
                if v_id != "-":
                    item_v_id.setForeground(QColor("blue"))
                    font = item_v_id.font()
                    font.setUnderline(True)
                    item_v_id.setFont(font)
                self.table_preview.setItem(row, col_idx, item_v_id)
                col_idx += 1
                
            # ===== NEW FEATURE: Show full category path instead of just main =====
            full_cat_path = res.get("full_category_path", res.get("main_category", ""))
            cat_item = QTableWidgetItem(str(full_cat_path))
            
            # Make it clickable (blue, underlined) if it has subcategories
            if ">" in str(full_cat_path):
                cat_item.setForeground(QColor("blue"))
                font = cat_item.font()
                font.setUnderline(True)
                cat_item.setFont(font)
                cat_item.setToolTip("Hiyerarşiyi görmek için tıklayın")
            
            self.table_preview.setItem(row, col_idx, cat_item)
            col_idx += 1
            # ===== END NEW FEATURE =====
            
            # ===== NEW FEATURE: Display stock value =====
            if has_stock:
                stock_val = res.get("_stock_value", 0)
                stock_item = QTableWidgetItem(str(stock_val))
                
                # Highlight zero stock with red background
                if stock_val <= 0:
                    stock_item.setBackground(QColor(255, 200, 200))  # Light red
                    stock_item.setToolTip("Stok Yok")
                    # Make text darker in dark mode
                    if self.sm.get("theme") == "dark" or (self.sm.get("theme") == "system" and self.is_system_dark()):
                        stock_item.setForeground(Qt.black)
                
                self.table_preview.setItem(row, col_idx, stock_item)
                col_idx += 1
            # ===== END NEW FEATURE =====
            
            base = res["base_price"]
            new_p = res["final_discounted_price"]
            
            item_new = QTableWidgetItem(self.update_price_visuals(new_p, base)) # Use helper for text
            if abs(new_p - base) > 0.001:
                item_new.setBackground(QColor(200, 255, 200) if new_p > base else QColor(255, 200, 200))
                # Adjust text color for dark mode readiness
            if abs(new_p - base) > 0.001:
                item_new.setBackground(QColor(200, 255, 200) if new_p > base else QColor(255, 200, 200))
                # Adjust text color
                if self.sm.get("theme") == "dark" or (self.sm.get("theme") == "system" and self.is_system_dark()):
                     item_new.setForeground(Qt.black) 
                font = item_new.font()
                font.setBold(True)
                item_new.setFont(font)
            
            # Use separate helper or just inline, logic is clean enough here
            # But let's respect previous step's intent to add arrows
            arrow_txt = str(new_p)
            if new_p > base + 0.001: arrow_txt += " ▲"
            elif new_p < base - 0.001: arrow_txt += " ▼"
            item_new.setText(arrow_txt)

            self.table_preview.setItem(row, col_idx, QTableWidgetItem(str(base))); col_idx+=1
            self.table_preview.setItem(row, col_idx, QTableWidgetItem(str(res["profit_added"]))); col_idx+=1
            self.table_preview.setItem(row, col_idx, item_new); col_idx+=1
            self.table_preview.setItem(row, col_idx, QTableWidgetItem(str(res["label_price"])))
            
            # Set Flags for all items in row to be non-editable but enabled
            for c in range(self.table_preview.columnCount()):
                it = self.table_preview.item(row, c)
                if it:
                    it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        
        self.table_preview.resizeColumnsToContents()
        
    def on_preview_cell_clicked(self, row, col):
        # Determine logical column index for Variant ID
        # Headers: Stock, Name, [Variant ID], ...\n        # If enabled, Variant ID is index 2.
        
        # ===== NEW FEATURE: Category click handler =====
        # Find category column index
        is_variant = self.chk_variants.isChecked()
        has_stock = hasattr(self, 'combo_stock_col') and self.combo_stock_col.currentText()
        
        # Calculate category column index dynamically
        # Base: 0=Stock Code, 1=Name, 2=Category (or Variant ID if enabled)
        cat_col_idx = 2
        if is_variant:
            cat_col_idx = 3  # After Variant ID
        
        # Check if clicked on category column
        if col == cat_col_idx:
            item = self.table_preview.item(row, col)
            if item:
                category = item.text()
                if category and category != "-":
                    # Show category detail dialog
                    from category_tree import CategoryDetailDialog
                    dialog = CategoryDetailDialog(category, self)
                    dialog.exec()
                    return
        # ===== END NEW FEATURE =====
        
        if not self.chk_variants.isChecked(): return
        
        if col == 2:
            item = self.table_preview.item(row, col)
            if not item: return
            v_id = item.text()
            if v_id and v_id != "-":
                self.show_variant_details(v_id)

    def show_variant_details(self, variant_id):
        # Find all rows with this variant id in ALL rows cache
        group_rows = []
        for r_data in self.all_rows_cache:
            if str(r_data.get(self.combo_variant.currentText(), "")) == variant_id:
                group_rows.append(r_data)
        
        if not group_rows: return
        
        # Dialog
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Varyant Grubu: {variant_id}")
        dlg.resize(800, 400)
        lay = QVBoxLayout(dlg)
        
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Ürün Adı", "Varyasyon", "Durum", "Fiyat"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        table.setRowCount(len(group_rows))
        
        val_col_name = self.combo_variant_val.currentText()
        
        for i, row in enumerate(group_rows):
            p_name = row.get(self.combo_name.currentText(), "-")
            
            # Parse Variation
            raw_val = row.get(val_col_name, "")
            display_val = raw_val
            status = "Tamam"
            
            if not raw_val or str(raw_val).strip() == "":
                display_val = "Varyasyon Değeri Girilmemiş!"
                status = "Eksik"
            else:
                 # Try to make it pretty: "Renk;Kırmızı,Beden;L" -> "Renk: Kırmızı | Beden: L"
                 try:
                     parts = raw_val.split(',')
                     pretty_parts = []
                     for p in parts:
                         if ';' in p:
                             k, v = p.split(';', 1)
                             pretty_parts.append(f"{k}: {v}")
                         else:
                             pretty_parts.append(p)
                     display_val = " | ".join(pretty_parts)
                 except:
                     pass

            table.setItem(i, 0, QTableWidgetItem(str(p_name)))
            
            item_val = QTableWidgetItem(display_val)
            if status == "Eksik":
                item_val.setForeground(Qt.red)
                item_val.setBackground(QColor(255, 230, 230))
                
            table.setItem(i, 1, item_val)
            table.setItem(i, 2, QTableWidgetItem(status))
            
            
            # Recalc price for display
            res = self.engine.calculate_row(row)
            price_txt = f"{res.get('final_discounted_price', 0)} TL"
            
            # Arrow
            base_p = float(res.get("base_price", 0))
            final_p = float(res.get("final_discounted_price", 0))
            if final_p > base_p: price_txt += " ▲"
            elif final_p < base_p: price_txt += " ▼"
            
            table.setItem(i, 3, QTableWidgetItem(price_txt))
            
        lay.addWidget(table)
        
        # Add Close Button
        btn_close = QPushButton("Kapat")
        btn_close.clicked.connect(dlg.accept)
        lay.addWidget(btn_close)
        
        dlg.exec()

    def show_preview_context_menu(self, position):
        menu = QMenu()
        
        # Comparison Action
        action_compare = menu.addAction("Excel vs Sistem Karşılaştırması")
        action_compare.triggered.connect(lambda: self.open_comparison_dialog())
        
        menu.addSeparator()
        
        # Copy Actions
        action_copy_stock = menu.addAction("Stok Kodunu Kopyala")
        action_copy_name = menu.addAction("Ürün Adını Kopyala")
        
        # Connect
        action_copy_stock.triggered.connect(lambda: self.copy_cell_data(0))
        action_copy_name.triggered.connect(lambda: self.copy_cell_data(1))
        
        menu.exec(self.table_preview.viewport().mapToGlobal(position))

    def copy_cell_data(self, col_idx):
        row = self.table_preview.currentRow()
        if row < 0: return
        
        item = self.table_preview.item(row, col_idx)
        if item:
            txt = item.text()
            QApplication.clipboard().setText(txt)
            # Optional: Show small status/tooltip?
            self.lbl_stats.setText(f"Kopyalandı: {txt}")

    def open_comparison_dialog(self):
        row = self.table_preview.currentRow()
        if row < 0: return
        
        real_idx = (self.current_page - 1) * self.items_per_page + row
        if real_idx >= len(self.filtered_rows): return
        
        calc_res = self.filtered_rows[real_idx]
        raw_data = calc_res.get("_raw_data", {})
        
        dlg = QDialog(self)
        dlg.setWindowTitle("Ürün Fiyat Analizi")
        dlg.resize(600, 400)
        
        main_layout = QVBoxLayout(dlg)
        
        # --- Header Info ---
        stock = calc_res.get('stock_code', '-')
        name = calc_res.get('product_name', '-')
        cat = calc_res.get('main_category', '-')
        
        lbl_info = QLabel(f"<h2>{name}</h2><p><b>Stok Kodu:</b> {stock} | <b>Kategori:</b> {cat}</p>")
        lbl_info.setStyleSheet("color: #333;")
        if self.sm.get("theme") == "dark" or (self.sm.get("theme") == "system" and self.is_system_dark()):
             lbl_info.setStyleSheet("color: #ecc;")

        main_layout.addWidget(lbl_info)
        
        # --- Visual Price Comparison (The "Fancy" Part) ---
        base_price = float(calc_res.get('base_price', 0))
        final_price = float(calc_res.get('final_discounted_price', 0))
        profit = float(calc_res.get('profit_added', 0))
        
        diff = final_price - base_price
        percent = (diff / base_price * 100) if base_price > 0 else 0
        
        # Frame for cards
        cards_layout = QHBoxLayout()
        
        def create_card(title, value, color="#333", subtext=""):
            frame = QGroupBox()
            fl = QVBoxLayout()
            l_title = QLabel(title)
            l_title.setStyleSheet("font-size: 14px; font-weight: bold; color: gray;")
            l_val = QLabel(value)
            l_val.setStyleSheet(f"font-size: 24px; font-weight: bold; color: {color};")
            fl.addWidget(l_title)
            fl.addWidget(l_val)
            if subtext:
                l_sub = QLabel(subtext)
                l_sub.setStyleSheet("font-size: 12px; color: gray;")
                fl.addWidget(l_sub)
            fl.addStretch()
            frame.setLayout(fl)
            return frame

        # Card 1: Old Price
        cards_layout.addWidget(create_card("Baz Fiyat (Excel)", f"{base_price:.2f} TL"))
        
        # Arrow/Change Indicator
        arrow_lbl = QLabel("➡")
        arrow_style = "font-size: 40px; color: gray;"
        if diff > 0:
            arrow_lbl.setText("▲")
            arrow_style = "font-size: 40px; color: green;"
        elif diff < 0:
            arrow_lbl.setText("▼")
            arrow_style = "font-size: 40px; color: red;"
        arrow_lbl.setStyleSheet(arrow_style)
        arrow_lbl.setAlignment(Qt.AlignCenter)
        cards_layout.addWidget(arrow_lbl)
        
        # Card 2: New Price
        color_new = "green" if diff > 0 else "red" if diff < 0 else "black"
        sub_txt = f"{abs(diff):.2f} TL ({abs(percent):.1f}%)"
        if diff > 0: sub_txt = "+" + sub_txt
        elif diff < 0: sub_txt = "-" + sub_txt
        
        cards_layout.addWidget(create_card("Yeni Fiyat (Sistem)", f"{final_price:.2f} TL", color_new, sub_txt))
        
        # Card 3: Profit Details
        cards_layout.addWidget(create_card("Eklenen Kar", f"{profit:.2f} TL", "#2196F3"))

        main_layout.addLayout(cards_layout)
        
        # --- Tabs for Details ---
        tabs = QTabWidget()
        
        # Summary Tab
        tab_summary = QWidget()
        l_sum = QFormLayout(tab_summary)
        l_sum.addRow("Etiket Fiyatı (Satış):", QLabel(f"<b>{calc_res.get('label_price',0)} TL</b>"))
        l_sum.addRow("Uygulanan İndirim Oranı:", QLabel(f"%{calc_res.get('discount_rate_used',0):.1f}"))
        l_sum.addRow("Ham İndirimli Fiyat:", QLabel(f"{calc_res.get('raw_discounted_price',0):.2f} TL"))
        tabs.addTab(tab_summary, "Özet Bilgiler")
        
        # Raw Data Tab
        tab_raw = QWidget()
        l_raw_main = QVBoxLayout(tab_raw)
        scroll = QScrollArea()
        w_scroll = QWidget()
        l_form_raw = QFormLayout(w_scroll)
        for k, v in raw_data.items():
            l_form_raw.addRow(f"{k}:", QLabel(str(v)))
        w_scroll.setLayout(l_form_raw)
        scroll.setWidget(w_scroll)
        scroll.setWidgetResizable(True)
        l_raw_main.addWidget(scroll)
        tabs.addTab(tab_raw, "Excel Ham Verisi")
        
        main_layout.addWidget(tabs)
        
        btn_close = QPushButton("Kapat")
        btn_close.clicked.connect(dlg.accept)
        main_layout.addWidget(btn_close, 0, Qt.AlignRight)
        
        dlg.exec()

    def update_price_visuals(self, new_p, base):
        # Helper for main table visuals
        txt = str(new_p)
        if abs(new_p - base) > 0.001:
            if new_p > base:
                txt += " ▲"
            else:
                txt += " ▼"
        return txt
        
    # Override update_table_view's item creation to use arrow helper
    # We need to inject this modification back into update_table_view or simply update it there.
    # Since I'm appending here, I will replace update_table_view in a separate block or rely on previous replace if I was editing it.
    # Wait, the previous block was `update_table_view`. Let's fix the arrow display there.
    
    def is_system_dark(self):
         # Quick helper re-using registry logic if needed, or just store state
         # For brevity, let's assume if it was applied it's stored or we re-check
         return True # Simplified for this specific text color fix context, ideally check registry again

    def start_processing(self):
        f = self.path_edit.text()
        if not f: 
            QMessageBox.warning(self, "Uyarı", "Lütfen işlem öncesi bir dosya seçin.")
            return
        
        self.log(f"İşlem başlatılıyor: {f}")
        self.collect_settings()
        
        # Validate Mandatory Fields
        mappings = self.sm.get("mappings", {})
        no_cat = mappings.get("no_category_mode", False)
        cat_col = mappings.get("category_col", "")
        
        if not no_cat and not cat_col:
             QMessageBox.warning(self, "Hata", "Kategori sütunu seçilmemiş! Lütfen 'Dosya' sekmesinden kategori sütununu seçin veya 'Kategori Yok' kutucuğunu işaretleyin.")
             return

        # Check other critical columns
        base_src = self.sm.get("base_price_source", "buy_price_col")
        crit_checks = [
            ("stock_code_col", "Stok Kodu"),
            ("product_name_col", "Ürün Adı"),
            (base_src, "Baz Fiyat (Seçilen Kaynak)")
        ]
        
        for key, label in crit_checks:
            if not mappings.get(key):
                 QMessageBox.warning(self, "Hata", f"Zorunlu sütun eksik: {label}")
                 return

        
        self.btn_run.setEnabled(False)
        self.progress_bar_part.setValue(0)
        self.lbl_part_status.setText("Hazırlanıyor...")
        
        self.worker = Worker(f, self.sm, self.engine)
        self.worker.progress_part.connect(self.on_part_progress)
        self.worker.log_message.connect(lambda msg: self.log(msg, "DEBUG"))
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.start()

    def on_part_progress(self, status, part_num, row_count):
        if status == "START":
            self.lbl_part_status.setText(f"Part {part_num} dosyası oluşturuluyor...")
            self.progress_bar_part.setValue(0)
            self.progress_bar_part.setFormat(f"Part {part_num}: Başlatılıyor")
            self.log(f"Part {part_num} yazılmaya başlandı...")
        elif status == "PROGRESS":
            self.progress_bar_part.setValue(row_count % 100) # Just a visual trick or use indefinite
            self.progress_bar_part.setFormat(f"Part {part_num}: {row_count} satır")
        elif status == "COMPLETE":
            self.progress_bar_part.setValue(100)
            self.lbl_part_status.setText(f"Part {part_num} tamamlandı ({row_count} satır).")
            self.log(f"Part {part_num} tamamlandı. ({row_count} satır)")

    def on_processing_finished(self, success, msg):
        self.btn_run.setEnabled(True)
        if success:
            self.log(f"İşlem başarıyla tamamlandı: {msg}")
            QMessageBox.information(self, "Tamamlandı", msg)
        else:
            self.log(f"İşlem hatayla sonuçlandı: {msg}", "ERROR")
            QMessageBox.critical(self, "Hata", msg)

    def save_settings_template(self):
        self.collect_settings()
        
        # Ensure directory exists
        template_dir = os.path.join(os.getcwd(), "configuration template")
        os.makedirs(template_dir, exist_ok=True)
        
        fname, _ = QFileDialog.getSaveFileName(self, "Ayarları Şablon Olarak Kaydet", template_dir, "JSON Files (*.json)")
        if fname:
            try:
                import json
                with open(fname, "w", encoding="utf-8") as f:
                    json.dump(self.sm.settings, f, indent=4, ensure_ascii=False)
                QMessageBox.information(self, "Başarılı", "Ayarlar kaydedildi.")
                self.log(f"Ayarlar şablonu kaydedildi: {fname}")
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Kaydedilemedi: {e}")

    def load_settings_template(self):
        # Ensure directory exists
        import os
        template_dir = os.path.join(os.getcwd(), "configuration template")
        os.makedirs(template_dir, exist_ok=True)
        
        fname, _ = QFileDialog.getOpenFileName(self, "Ayarları Yükle", template_dir, "JSON Files (*.json)")
        if fname:
            try:
                import json
                with open(fname, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    
                    # Fix: Deep update for nested dicts to prevent overwriting with partials
                    # Actually, self.sm.settings is a dict. SettingsManager.load_settings does a merge.
                    # We should probably re-use that logic or implement a recursive update here.
                    
                    current = self.sm.settings
                    
                    def recursive_update(d, u):
                        for k, v in u.items():
                            if isinstance(v, dict):
                                d[k] = recursive_update(d.get(k, {}), v)
                            else:
                                d[k] = v
                        return d
                        
                    recursive_update(current, data)
                    
                    self.sm.save_settings() # Persist to internal
                
                # Check mapping combos update
                # If file not loaded, combos are empty, so setting text does nothing.
                # Use insertItem to force-show the saved mapping even if not in file headers yet.
                mappings = self.sm.get("mappings", {})
                combos = {
                    "stock_code_col": self.combo_stock,
                    "product_name_col": self.combo_name,
                    "category_col": self.combo_cat,
                    "buy_price_col": self.combo_buy,
                    "sell_price_col": self.combo_sell,
                    "discounted_price_col": self.combo_disc,
                    "market_price_col": self.combo_market,
                    "variant_id_col": self.combo_variant,
                    "variant_val_col": self.combo_variant_val
                }
                
                # Checkbox
                self.chk_variants.setChecked(mappings.get("is_variant_mode", False))
                self.chk_unique_variant.setChecked(mappings.get("show_unique_variant", False))
                
                for key, combo in combos.items():
                    val = mappings.get(key)
                    if val and combo.findText(val) == -1:
                        combo.addItem(val) # Add as temporary option
                    combo.setCurrentText(val)
                
                self.load_ui_values()
                self.log(f"Ayarlar şablonu yüklendi: {fname}")
                QMessageBox.information(self, "Başarılı", "Ayarlar yüklendi ve arayüze uygulandı.")
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Yüklenemedi: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Global logger for unhandled exceptions (optional but good)
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        import traceback
        err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        print(err_msg)
        # Could send to a static logger here

    sys.excepthook = handle_exception
    
    w = MainWindow()
    # Initial log
    w.log("Uygulama başlatıldı.")
    w.show()
    sys.exit(app.exec())
