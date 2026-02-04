import json
import os

DEFAULT_SETTINGS = {
    "mappings": {
        "stock_code_col": "",
        "product_name_col": "",
        "category_col": "",
        "buy_price_col": "",
        "sell_price_col": "",
        "discounted_price_col": "",
        "market_price_col": ""
    },
    "targets": {
        "update_discounted": True,
        "update_sell": True,
        "update_market": False
    },
    "categories": {
        "default_discount": 50.0,
        "mapping": {}  # "CategoryName": 50.0
    },
    "profit_segments": [
        {
            "min": 0,
            "max": 499,
            "type": "TL",  # or "PERCENT"
            "value": 200,
            "profit_min": 0, # Optional per segment min
            "profit_max": 0  # Optional per segment max
        },
        {
            "min": 500,
            "max": 999,
            "type": "TL",
            "value": 300,
            "profit_min": 0,
            "profit_max": 0
        },
        {
             "min": 1000,
             "max": 999999,
             "type": "PERCENT",
             "value": 30,
             "profit_min": 0,
             "profit_max": 0
        }
    ],
    "global_min_profit": 200.0,
    "base_price_source": "buy_price_col", # or "sell_price_col", "discounted_price_col", "market_price_col"
    "rounding": {
        "mode": "ceiling", # "ceiling", "round", "floor"
        "step": 10,       # 1, 5, 10, 25, 50, 100
        "ends_with_99": True
    },
    "limits": {
        "min_discounted_price": 75.0,
        "max_discounted_price": 1000.0,
        "max_increase_tl": 0, # 0 = disabled
        "max_increase_percent": 0 # 0 = disabled
    },
    "output": {
        "max_rows_per_file": 5000,
        "output_dir": "",
        "filename_template": "output_part_{n}.xlsx"
    },
    "category_extraction": {
        "mode": "first_delimiter", # "first_delimiter", "regex"
        "delimiters": [";", ">", "|", ","]
    }

}

class SettingsManager:
    def __init__(self, filepath="settings.json"):
        self.filepath = filepath
        self.settings = self.load_settings()

    def load_settings(self):
        if not os.path.exists(self.filepath):
            return DEFAULT_SETTINGS.copy()
        try:
            with open(self.filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Merge with defaults to ensure new keys exist
                merged = DEFAULT_SETTINGS.copy()
                merged.update(data) 
                # Deep merge for nested dicts (simplified for now)
                for key, val in data.items():
                    if isinstance(val, dict) and key in merged and isinstance(merged[key], dict):
                        merged[key].update(val)
                    else:
                        merged[key] = val
                return merged
        except Exception as e:
            print(f"Error loading settings: {e}")
            return DEFAULT_SETTINGS.copy()

    def save_settings(self):
        try:
            with open(self.filepath, "w", encoding="utf-8") as f:
                json.dump(self.settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def get(self, key, default=None):
        return self.settings.get(key, default)

    def set(self, key, value):
        self.settings[key] = value

    def update_nested(self, parent, key, value):
        if parent in self.settings:
            self.settings[parent][key] = value
