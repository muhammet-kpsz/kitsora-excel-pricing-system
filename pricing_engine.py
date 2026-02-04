import math
import re

class PricingEngine:
    def __init__(self, settings_manager):
        self.sm = settings_manager

    def extract_category(self, raw_text):
        """
        Extracts the main category from a raw string based on settings.
        Example: "Fantezi;Fantezi>Jartiyer" -> "Fantezi"
        """
        if not raw_text or not isinstance(raw_text, str):
            return "Uncategorized"

        config = self.sm.get("category_extraction")
        delimiters = config.get("delimiters", [";", ">", "|", ","])
        
        # Create a regex pattern for split
        pattern = '|'.join(map(re.escape, delimiters))
        parts = re.split(pattern, raw_text)
        
        if parts:
            return parts[0].strip()
        return raw_text.strip()

    def get_discount_rate(self, category):
        """
        Returns the discount rate (0.0 - 1.0) for a given category.
        If mapped percentage is 50, returns 0.50.
        """
        cats = self.sm.get("categories")
        mapping = cats.get("mapping", {})
        default_rate = cats.get("default_discount", 50.0)

        # Check exact match
        if category in mapping:
            return float(mapping[category]) / 100.0
        
        return float(default_rate) / 100.0

    def calculate_profit(self, base_price):
        """
        Calculates profit amount based on segments and global min profit.
        """
        segments = self.sm.get("profit_segments", [])
        global_min = float(self.sm.get("global_min_profit", 0.0))
        
        applied_profit = 0.0
        
        # Find matching segment
        matched_segment = None
        for seg in segments:
            # Safely cast
            try:
                s_min = float(seg["min"])
                s_max = float(seg["max"])
            except:
                continue
                
            if s_min <= base_price <= s_max:
                matched_segment = seg
                break
        
        if matched_segment:
            try:
                val = float(matched_segment["value"])
                t = str(matched_segment["type"]).upper()
                
                # Check Type
                if "PERCENT" in t or "YÜZDE" in t:
                    # Percentage: Base * (Val/100)
                    applied_profit = base_price * (val / 100.0)
                else:
                    # Assume Amount (TL)
                    applied_profit = val
                
                # Add Extra Fixed Amount (if any)
                extra = float(matched_segment.get("extra_added", 0.0))
                applied_profit += extra
            except:
                pass
        
        # Apply Global Min Profit
        # Logic: Profit cannot be less than Global Min Profit (IF ENABLED)
        if self.sm.get("enable_global_min", False):
            if applied_profit < global_min:
                applied_profit = global_min
            
        return applied_profit

    def apply_rounding(self, price):
        """
        Applies rounding rules and .99 logic.
        """
        r_config = self.sm.get("rounding")
        mode = r_config.get("mode", "ceiling")
        step = float(r_config.get("step", 1.0))
        ends_99 = r_config.get("ends_with_99", False)

        if step <= 0: step = 1

        rounded_price = price

        # 1. Round to step
        if mode == "ceiling":
            rounded_price = math.ceil(price / step) * step
        elif mode == "floor":
            rounded_price = math.floor(price / step) * step
        else: # round
            rounded_price = round(price / step) * step

        # 2. Apply .99 logic
        # If ends_with_99 is True, we want the price to end in .99
        # usually this means subtracting 0.01 from an integer, or finding nearest .99
        # Common logic: If we rounded to integer (e.g. 150), make it 149.99 or 150.99?
        # Requirement: "xx.99"
        # Let's assume we simply enforce .99 on the integer part if step >= 1
        
        if ends_99:
            # If we have 150, and we want it to end in .99
            # Usually people want 149.99 (psychological under) or 159.99
            # The prompt says: "yuvarlandıktan sonra -0.01 veya +0.99 mantığıyla"
            
            # Simple approach: floor to int, then add 0.99? 
            # Or if we have 150 -> 149.99
            
            # Let's try: take the rounded price. 
            # If it's an integer like 150.0:
            # Option A: 149.99 (Loss of 0.01) - safer for customers?
            # Option B: 150.99 (Gain of 0.99)
            
            # Prompt says: "yuvarlandıktan sonra -0.01" implied.
            # Example: 153 -> round(10) -> 160 -> -0.01 -> 159.99
            # Example: 153 -> round(10) -> 150 -> -0.01 -> 149.99
            
            # However this depends on the Step. 
            # If step is 10, valid values are 10, 20, 30.
            # 20 -> 19.99
            
            rounded_price = rounded_price - 0.01

        return rounded_price

    def calculate_row(self, row_data):
        """
        Main pipeline for a single row.
        row_data: dict with keys matching settings mapping (e.g. 'KATEGORILER': '...', 'ALIS': 100)
        Returns: dict with new values
        """
        mappings = self.sm.get("mappings")
        
        # 1. Extract Category
        no_cat_mode = mappings.get("no_category_mode", False)
        
        if no_cat_mode:
            raw_cat = "Kategorisiz"
            main_cat = "Kategorisiz"
        else:
            cat_col = mappings.get("category_col")
            raw_cat = str(row_data.get(cat_col, ""))
            main_cat = self.extract_category(raw_cat)
        
        # 2. Get Discount Rate
        # Discount rate is what we show to customer. 
        # Label Price = Discounted Price / (1 - rate)
        discount_rate = self.get_discount_rate(main_cat)
        
        # 3. Base Price
        base_source_key = self.sm.get("base_price_source") # e.g. "buy_price_col"
        base_col_name = mappings.get(base_source_key)
        
        try:
            base_price = float(row_data.get(base_col_name, 0))
        except (ValueError, TypeError):
            return {"error": "Invalid base price", "main_category": main_cat}

        if base_price <= 0:
             return {"error": "Zero or negative base price", "main_category": main_cat}

        # 4. Calculate Profit
        profit = self.calculate_profit(base_price)
        
        # 5. Raw Discounted Price
        raw_discounted_price = base_price + profit
        
        # 6. Limits (Min/Max Discounted)
        limits = self.sm.get("limits")
        min_p = float(limits.get("min_discounted_price", 0))
        max_p = float(limits.get("max_discounted_price", 999999))
        
        if raw_discounted_price < min_p: raw_discounted_price = min_p
        if raw_discounted_price > max_p: raw_discounted_price = max_p
        
        # Inflation check (Optional in prompt, skipped for simplicity or add later)
        
        # 7. Rounding & .99
        final_discounted_price = self.apply_rounding(raw_discounted_price)
        
        # Re-check max limit after rounding (e.g. 1000 -> 999.99 is OK, but 1000.99 is not if max is 1000)
        if final_discounted_price > max_p:
            # If rounding pushed it over, we should probably clamp it back. 
            # But if .99 is strict, clamping to 1000 loses .99.
            # Prompt says: "Maksimum sınırla çakışma durumunda (örn. 1000 tavan) doğru davran: 999.99"
            final_discounted_price = math.floor(max_p) - 0.01 if self.sm.get("rounding").get("ends_with_99") else max_p

        # 8. Calculate Label (Sell) Price
        # label = discounted / (1 - rate)
        # Avoid div by zero
        if discount_rate >= 1.0: discount_rate = 0.99 # Safety
        
        label_price = final_discounted_price / (1.0 - discount_rate)
        
        # Optional: Round label price too? Prompt says optional. Let's doing simple rounding to 2 decimals.
        label_price = round(label_price, 2)
        
        # Pass through identity info
        p_name = row_data.get(mappings.get("product_name_col", ""), "")
        s_code = row_data.get(mappings.get("stock_code_col", ""), "")

        return {
            "stock_code": s_code,
            "product_name": p_name,
            "main_category": main_cat,
            # ===== NEW FEATURE: Store full category path =====
            "full_category_path": raw_cat,  # Preserves "Alt Giyim > Sütyen > Dantelli"
            # ===== END NEW FEATURE =====
            "base_price": base_price,
            "profit_added": profit,
            "raw_discounted_price": raw_discounted_price,
            "final_discounted_price": final_discounted_price,
            "label_price": label_price,
            "discount_rate_used": discount_rate * 100
        }
