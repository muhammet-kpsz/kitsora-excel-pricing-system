"""
Stock Filter Module
Handles stock-based filtering for the Excel Pricing Engine.
This is a new module that integrates as an optional filter.
"""

import pandas as pd


class StockFilter:
    """
    Provides stock filtering capabilities without modifying existing data structures.
    """
    
    @staticmethod
    def validate_stock_column(df, column_name):
        """
        Validates that the stock column exists and contains numeric data.
        
        Args:
            df: pandas DataFrame
            column_name: Name of the stock column
            
        Returns:
            tuple: (is_valid, error_message)
        """
        if not column_name:
            return False, "Stok sütunu seçilmedi"
        
        if column_name not in df.columns:
            return False, f"'{column_name}' sütunu bulunamadı"
        
        # Check if column contains numeric data
        try:
            pd.to_numeric(df[column_name], errors='coerce')
            return True, ""
        except Exception as e:
            return False, f"Stok sütunu sayısal değer içermiyor: {e}"
    
    @staticmethod
    def filter_by_stock(rows, stock_col, include_zero_stock=True):
        """
        Filters row data based on stock values.
        
        Args:
            rows: List of row dictionaries
            stock_col: Name of the stock column
            include_zero_stock: If False, filters out rows with stock <= 0
            
        Returns:
            list: Filtered rows
        """
        if not stock_col or include_zero_stock:
            # No filtering needed
            return rows
        
        filtered = []
        for row in rows:
            stock_val = row.get(stock_col, None)
            
            # Try to convert to float
            try:
                stock_num = float(stock_val) if stock_val is not None else 0
                
                # Only include if stock > 0
                if stock_num > 0:
                    filtered.append(row)
            except (ValueError, TypeError):
                # If can't convert, skip this row (treat as 0 stock)
                continue
        
        return filtered
    
    @staticmethod
    def get_stock_value(row_data, stock_col):
        """
        Safely extracts stock value from a row.
        
        Args:
            row_data: Dictionary of row data
            stock_col: Name of stock column
            
        Returns:
            float: Stock value, or 0 if invalid
        """
        if not stock_col:
            return 0
        
        stock_val = row_data.get(stock_col, 0)
        
        try:
            return float(stock_val) if stock_val is not None else 0
        except (ValueError, TypeError):
            return 0
