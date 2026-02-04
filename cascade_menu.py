from PySide6.QtWidgets import QPushButton, QMenu, QWidgetAction, QLabel
from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QAction

class CascadeCategoryButton(QPushButton):
    """
    A button that mimics a ComboBox but opens a cascading QMenu
    for hierarchical category selection (Desktop application style).
    """
    categorySelected = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setText("Tüm Kategoriler")
        self.setMenu(QMenu(self))
        self.current_category = "Tüm Kategoriler"
        
        # Style to look like a ComboBox
        self.setStyleSheet("""
            QPushButton {
                text-align: left;
                padding-left: 10px;
                background-color: white;
                border: 1px solid #ccc;
                border-radius: 4px;
                height: 25px;
            }
            QPushButton::menu-indicator {
                subcontrol-origin: padding;
                subcontrol-position: center right;
                image: none; /* We can add a down arrow icon if needed */
                width: 20px;
            }
        """)

    def get_selected_category(self):
        return self.current_category

    def set_selected_category(self, category_path):
        self.current_category = category_path
        self.setText(category_path.split(" > ")[-1] if " > " in category_path else category_path)

    def populate_categories(self, category_counts, selected_categories=None):
        """
        Populate the cascading menu.
        
        Args:
            category_counts: dict { "Alt Giyim > Şort": 45, ... }
            selected_categories: list of allowed categories (from tree selection)
        """
        menu = QMenu(self)
        self.setMenu(menu)
        
        # Add "Tüm Kategoriler"
        action_all = QAction("Tüm Kategoriler", menu)
        action_all.triggered.connect(lambda: self._on_category_triggered("Tüm Kategoriler"))
        menu.addAction(action_all)
        menu.addSeparator()

        # Filter categories if selection provided
        filtered_counts = {}
        if selected_categories:
            for cat, count in category_counts.items():
                for sel in selected_categories:
                    # Include if exact match or subcategory
                    if cat == sel or cat.startswith(sel + " >"):
                        filtered_counts[cat] = count
                        break
        else:
            filtered_counts = category_counts

        # Build Hierarchy
        from category_tree import CategoryParser
        hierarchy = CategoryParser.build_hierarchy(list(filtered_counts.keys()))

        # Build Menu Recursively
        self._build_menu_recursive(menu, hierarchy, filtered_counts, "")
        
        # ===== FIX: Restore/Validate Selection =====
        # If current selection is not valid in new data (and not "Tüm Kategoriler"), reset?
        # Actually, let's keep it if possible, or check if it exists in counts.
        if self.current_category != "Tüm Kategoriler":
            # Check if current category is available in the new filtered counts
            # We match strict path.
            if self.current_category in filtered_counts:
                 # It's valid, keep it. Text is already set.
                 pass
            else:
                 # It might be a parent of something valid? 
                 # If partially valid, we might keep it.
                 # For safety, if strictly missing, we traditionally reset to All, 
                 # BUT user wants persistence. 
                 # If the category exists in the Tree but not in counts (0 items), we should still arguably keep the filter active (which shows 0 results).
                 # So we rely on the fact that populate_categories is visual.
                 # We simply ensure the Button Text reflects self.current_category
                 leaf_name = self.current_category.split(" > ")[-1]
                 self.setText(leaf_name)
        else:
            self.setText("Tüm Kategoriler")
        # ===== END FIX =====
        
    def _build_menu_recursive(self, parent_menu, hierarchy, counts, parent_path):
        for key in sorted(hierarchy.keys()):
            full_path = f"{parent_path} > {key}" if parent_path else key
            
            # Calculate total count including children
            total_count = self._calculate_total_count(full_path, counts)
            display_text = f"{key} ({total_count})" if total_count > 0 else key

            if hierarchy[key]:
                # It has children -> Create Submenu
                submenu = parent_menu.addMenu(display_text)
                
                # Add action for the parent category itself (optional, allows selecting parent)
                parent_action = QAction(f"{key} (Tümü)", submenu)
                parent_action.triggered.connect(lambda checked=False, p=full_path: self._on_category_triggered(p))
                submenu.addAction(parent_action)
                submenu.addSeparator()
                
                # Recurse
                self._build_menu_recursive(submenu, hierarchy[key], counts, full_path)
            else:
                # No children -> Create Action
                action = QAction(display_text, parent_menu)
                action.triggered.connect(lambda checked=False, p=full_path: self._on_category_triggered(p))
                parent_menu.addAction(action)

    def _on_category_triggered(self, category_path):
        self.current_category = category_path
        # Show leaf name on button
        leaf_name = category_path.split(" > ")[-1] if " > " in category_path else category_path
        self.setText(leaf_name)
        self.categorySelected.emit(category_path)

    def _calculate_total_count(self, category_path, counts):
        total = 0
        for cat, count in counts.items():
            if cat == category_path or cat.startswith(category_path + " >"):
                total += count
        return total
