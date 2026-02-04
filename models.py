from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from PySide6.QtGui import QColor
import pandas as pd

class PandasTableModel(QAbstractTableModel):
    """
    Pandas DataFrame tabanli, yuksek performansli QTableView modeli.
    Stok <= 0 durumunda satirlari renklendirir.
    """
    def __init__(self, data=None):
        super().__init__()
        self._data = data if data is not None else pd.DataFrame()
        self.stock_col_name = "Stok" # Varsayilan kolon adi

    def setDataFrame(self, df):
        self.beginResetModel()
        self._data = df.copy() if df is not None else pd.DataFrame()
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        if self._data is None: return 0
        return self._data.shape[0]

    def columnCount(self, parent=QModelIndex()):
        if self._data is None: return 0
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        
        row = index.row()
        col = index.column()
        
        if role == Qt.DisplayRole:
            val = self._data.iat[row, col]
            if pd.isna(val): return ""
            return str(val)

        elif role == Qt.BackgroundRole:
            # Stok kontrolu
            if self.stock_col_name in self._data.columns:
                try:
                    # Stok degerine gore satir rengi
                    # Performans icin iloc yerine onceden hesaplanmis bir maske kullanilabilir
                    # ama 10k satirda bu da hizli calisir.
                    s_val = self._data.iloc[row][self.stock_col_name]
                    if float(s_val) <= 0:
                        return QColor(255, 235, 238) # Light Red
                except (ValueError, TypeError):
                    pass
        
        elif role == Qt.ToolTipRole:
             if self.stock_col_name in self._data.columns:
                try:
                    s_val = self._data.iloc[row][self.stock_col_name]
                    if float(s_val) <= 0:
                        return "Stok Yok"
                except:
                    pass

        return None

    def headerData(self, section, orientation, role):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])
            if orientation == Qt.Vertical:
                return str(self._data.index[section] + 1)
        return None
