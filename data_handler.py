import pandas as pd
import re

class DataHandler:
    def __init__(self):
        self.df = None
        
    def load_excel(self, filepath):
        try:
            # Read excel with openpyxl engine
            # Using dtype=object to preserve data types initially, then convert as needed
            self.df = pd.read_excel(filepath, engine='openpyxl')
            
            # Clean column names
            self.df.columns = [str(c).strip() for c in self.df.columns]
            return True, None
        except Exception as e:
            return False, str(e)
            
    def get_headers(self):
        if self.df is None: return []
        return list(self.df.columns)

    def get_row_count(self):
        if self.df is None: return 0
        return len(self.df)

    def get_category_tree_structure(self, cat_col, separator=">"):
        """
        Kategori sütununu tarar ve nested dictionary (agac) yapisi dondurur.
        Ornek: "Giyim > Ic Giyim" -> {"Giyim": {"Ic Giyim": {}}}
        """
        if self.df is None or cat_col not in self.df.columns:
            return {}
            
        # Get unique categories, ignore NaN
        unique_cats = self.df[cat_col].dropna().astype(str).unique()
        tree = {}
        
        for cat_str in unique_cats:
            # Split by separator (default > or ;)
            # Trying intelligent split if separator not provided or default fails? 
            # Let's stick to default defined in settings usually, but here fixed mainly >
            parts = [p.strip() for p in cat_str.replace(";", ">").split(">")]
            
            current_node = tree
            for part in parts:
                if not part: continue
                if part not in current_node:
                    current_node[part] = {}
                current_node = current_node[part]
                
        return tree

    def filter_data(self, 
                    stock_col=None, 
                    include_zero_stock=True, 
                    cat_col=None, 
                    selected_categories=None, 
                    search_query=""):
        """
        Pandas üzerinde vektörel filtreleme yapar.
        - stock_col: Stok sütun adi
        - include_zero_stock: True ise 0 stoklular dahil, False ise haric
        - selected_categories: Secili kategori listesi (list of str)
        """
        if self.df is None: return pd.DataFrame()
        
        filtered = self.df.copy()
        
        # 1. Stok Filtresi
        if stock_col and stock_col in filtered.columns:
            # Sayisala cevir, hata verenleri 0 yap (NaN -> 0)
            # Downcast to numeric
            filtered[stock_col] = pd.to_numeric(filtered[stock_col], errors='coerce').fillna(0)
            
            if not include_zero_stock:
                filtered = filtered[filtered[stock_col] > 0]
        
        # 2. Kategori Filtresi
        if cat_col and cat_col in filtered.columns and selected_categories:
            # Eger "Tüm Kategoriler" veya liste bossa hepsini getir (UI tarafinda kontrol edilmeli ama burada da guvenlik)
            # selected_categories hiyerarsik path degil, sadece kelime listesi mi? 
            # Hayir, tree yapisinda secilen basliklar.
            # "Alt Giyim" secilmisse, icinde "Alt Giyim" gecenler gelmeli.
            
            # Regex pattern olustur: (Kategori1|Kategori2|...)
            # Escape edilmeli ozel karakterler icin
            if len(selected_categories) > 0:
                pattern = '|'.join(map(re.escape, selected_categories))
                # Case insensitive search
                filtered = filtered[filtered[cat_col].astype(str).str.contains(pattern, case=False, na=False)]
        
        # 3. Arama (Search Query)
        if search_query:
            # Tüm sütunlarda aramak yerine performans icin sadece text olabileceklerde arayalim
            # Veya tum df'i stringe cevirip arayalim (maliyetli ama basit)
            
            # Daha akilli yol: Satiri tek stringe indirgeyip arama
            # axis=1 ile satirlari birlestir
            
            # Ancak kullanici "Stok Kodu veya Ürün Adı" dedi.
            # Eger UI'da bu kolonlar tanimliysa onlarda arayalim
            
            # Simdilik genel arama:
            mask = filtered.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
            filtered = filtered[mask]
            
        return filtered
