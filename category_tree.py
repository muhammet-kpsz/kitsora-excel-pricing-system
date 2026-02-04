"""
Category Tree Module
Provides hierarchical category tree widget with checkbox selection.
This is a new module that adds optional tree-based category viewing.
"""

from PySide6.QtWidgets import QTreeWidget, QTreeWidgetItem, QDialog, QVBoxLayout, QLabel, QPushButton
from PySide6.QtCore import Qt, Signal


class CategoryParser:
    """
    Parses category strings in hierarchical format (e.g., "Parent > Child > SubChild")
    """
    
    @staticmethod
    def parse_category_path(raw_category, delimiter=">"):
        """
        Splits a category string into hierarchy levels.
        
        Args:
            raw_category: String like "Alt Giyim > Sütyen > Dantelli"
            delimiter: Separator character
            
        Returns:
            list: ["Alt Giyim", "Sütyen", "Dantelli"]
        """
        if not raw_category:
            return []
        
        parts = [p.strip() for p in str(raw_category).split(delimiter)]
        return [p for p in parts if p]  # Remove empty parts
    
    @staticmethod
    def build_hierarchy(category_list):
        """
        Builds a nested dictionary structure from category paths.
        
        Args:
            category_list: List of category strings
            
        Returns:
            dict: Nested structure like {"Parent": {"Child": {}, ...}, ...}
        """
        hierarchy = {}
        
        for cat in category_list:
            parts = CategoryParser.parse_category_path(cat)
            
            if not parts:
                continue
            
            # Build nested structure
            current = hierarchy
            for part in parts:
                if part not in current:
                    current[part] = {}
                current = current[part]
        
        return hierarchy


class CategoryTreeWidget(QTreeWidget):
    """
    Custom tree widget for hierarchical category selection with checkboxes.
    """
    
    # Signal emitting the full list of selected paths
    selectionChanged = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderLabel("Kategoriler")
        self.setSelectionMode(QTreeWidget.NoSelection)  # Use checkboxes instead
        
        # Track items by full path for easy lookup
        self.item_map = {}  # "Parent > Child" -> QTreeWidgetItem
        
        # Connect signals
        self.itemChanged.connect(self._on_item_changed)
    
    def build_tree(self, categories):
        """
        Builds the tree structure from a list of category strings.
        
        Args:
            categories: List of category strings (can include ">" for hierarchy)
        """
        # CRITICAL: Block signals during rebuild to prevent spurious empty selections
        self.blockSignals(True)
        
        self.clear()
        self.item_map.clear()
        
        # Parse and build hierarchy
        hierarchy = CategoryParser.build_hierarchy(categories)
        
        # Recursively add items
        self._add_tree_items(hierarchy, None, "")
        
        # Expand all by default
        self.expandAll()
        
        # Re-enable signals
        self.blockSignals(False)
    
    def update_counts(self, category_counts):
        """
        Updates tree items to show product counts.
        
        Args:
            category_counts: dict like {"Alt Giyim > Sütyen": 45, ...}
        """
        # Block signals to prevent setText from triggering itemChanged
        self.blockSignals(True)
        
        # ===== NEW FEATURE: Show product counts in tree =====
        for full_path, item in self.item_map.items():
            # Calculate total count (including subcategories)
            total_count = self._calculate_total_count(full_path, category_counts)
            
            # Get category name (last part of path)
            cat_name = full_path.split(" > ")[-1] if " > " in full_path else full_path
            
            # Update display text
            if total_count > 0:
                item.setText(0, f"{cat_name} ({total_count})")
            else:
                item.setText(0, cat_name)
        # ===== END NEW FEATURE =====
        
        # Re-enable signals
        self.blockSignals(False)
    
    def _calculate_total_count(self, category_path, counts):
        """Calculate total count for category including all subcategories"""
        total = 0
        for cat, count in counts.items():
            # Match exact or if it's a subcategory
            if cat == category_path or cat.startswith(category_path + " >"):
                total += count
        return total
    
    def _add_tree_items(self, hierarchy_dict, parent_item, parent_path):
        """
        Recursively adds items to the tree.
        
        Args:
            hierarchy_dict: Nested dict of categories
            parent_item: Parent QTreeWidgetItem (None for root)
            parent_path: Full path of parent (e.g., "Parent > ")
        """
        for cat_name, children in sorted(hierarchy_dict.items()):
            # Create item
            item = QTreeWidgetItem()
            item.setText(0, cat_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(0, Qt.Unchecked)
            
            # Calculate full path
            full_path = parent_path + cat_name if parent_path else cat_name
            item.setData(0, Qt.UserRole, full_path)  # Store full path
            
            # Add to tree
            if parent_item:
                parent_item.addChild(item)
            else:
                self.addTopLevelItem(item)
            
            # Store in map
            self.item_map[full_path] = item
            
            # Add children recursively
            if children:
                self._add_tree_items(children, item, full_path + " > ")
    
    def _on_item_changed(self, item, column):
        """
        Handles checkbox state changes with parent-child logic.
        """
        if column != 0:
            return
        
        # Block signals to prevent recursion
        self.blockSignals(True)
        
        # Get new state
        new_state = item.checkState(0)
        
        # Update all children
        self._update_children(item, new_state)
        
        # Update parent state
        self._update_parent(item)
        
        self.blockSignals(False)
        
        # Emit new state
        self.selectionChanged.emit(self.get_selected_categories())
    
    def _update_children(self, item, state):
        """
        Recursively updates all children to match parent state.
        """
        for i in range(item.childCount()):
            child = item.child(i)
            child.setCheckState(0, state)
            self._update_children(child, state)
    
    def _update_parent(self, item):
        """
        Updates parent state based on children (checked/unchecked/partial).
        """
        parent = item.parent()
        if not parent:
            return
        
        # Count children states
        total = parent.childCount()
        checked = sum(1 for i in range(total) if parent.child(i).checkState(0) == Qt.Checked)
        partial = sum(1 for i in range(total) if parent.child(i).checkState(0) == Qt.PartiallyChecked)
        
        if checked == total:
            parent.setCheckState(0, Qt.Checked)
        elif checked == 0 and partial == 0:
            parent.setCheckState(0, Qt.Unchecked)
        else:
            parent.setCheckState(0, Qt.PartiallyChecked)
        
        # Recursively update grandparent
        self._update_parent(parent)
    
    def get_selected_categories(self):
        """
        Returns list of all selected category paths.
        
        Returns:
            list: Full paths of checked categories (e.g., ["Alt Giyim", "Alt Giyim > Sütyen"])
        """
        selected = []
        
        def collect_checked(item):
            if item.checkState(0) == Qt.Checked:
                path = item.data(0, Qt.UserRole)
                if path:
                    selected.append(path)
            
            for i in range(item.childCount()):
                collect_checked(item.child(i))
        
        # Iterate all top-level items
        for i in range(self.topLevelItemCount()):
            collect_checked(self.topLevelItem(i))
        
        return selected
    
    def set_selected_categories(self, selected_paths):
        """
        Sets the selection state based on a list of category paths.
        
        Args:
            selected_paths: List of full category paths to select
        """
        self.blockSignals(True)
        
        # First uncheck all
        for i in range(self.topLevelItemCount()):
            self._uncheck_all(self.topLevelItem(i))
        
        # Then check specified paths
        for path in selected_paths:
            item = self.item_map.get(path)
            if item:
                item.setCheckState(0, Qt.Checked)
                # Update children
                self._update_children(item, Qt.Checked)
        
        # Update all parent states
        for i in range(self.topLevelItemCount()):
            self._update_parent_recursive(self.topLevelItem(i))
        
        self.blockSignals(False)
        self.selectionChanged.emit(self.get_selected_categories())
    
    def _uncheck_all(self, item):
        """Recursively unchecks all items."""
        item.setCheckState(0, Qt.Unchecked)
        for i in range(item.childCount()):
            self._uncheck_all(item.child(i))
    
    def _update_parent_recursive(self, item):
        """Recursively updates parent states from bottom up."""
        for i in range(item.childCount()):
            self._update_parent_recursive(item.child(i))
        self._update_parent(item)


class CategoryDetailDialog(QDialog):
    """
    Dialog to show category hierarchy breakdown when clicked in preview.
    """
    
    def __init__(self, category_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Kategori Detayları")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        # Parse category
        parts = CategoryParser.parse_category_path(category_path)
        
        # Show hierarchy
        layout.addWidget(QLabel("<b>Kategori Hiyerarşisi:</b>"))
        
        if parts:
            for i, part in enumerate(parts):
                indent = "  " * i
                label = QLabel(f"{indent}{'└─ ' if i > 0 else ''}• {part}")
                layout.addWidget(label)
        else:
            layout.addWidget(QLabel(f"• {category_path}"))
        
        # Close button
        btn_close = QPushButton("Kapat")
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close)
