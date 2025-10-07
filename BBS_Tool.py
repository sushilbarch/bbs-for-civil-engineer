#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BBS Tool (PyQt5, Dark UI)
- Members: Ties/Stirrups, Beam (simple), Column (basic)
- Editable assumptions: hook length multiplier, bend deduction, lap/dev length
- Live BBS preview (pandas → QTableView)
- Export to Excel using a formatted template (if present)
"""
import sys, math, os
from typing import Dict, Any, List
import pandas as pd

from PyQt5.QtCore import Qt, QAbstractTableModel, QVariant, QSize, QSettings, QByteArray
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFormLayout, QHBoxLayout, QVBoxLayout, QLabel,
    QDoubleSpinBox, QSpinBox, QComboBox, QPushButton, QTableView, QFileDialog,
    QMessageBox, QToolBar, QStatusBar
)

# ---------------- Helpers ----------------
def steel_unit_weight_mm(dia_mm: float) -> float:
    # Unit weight (kg/m) approximated by dia^2 / 162
    return (dia_mm * dia_mm) / 162.0

# ---------------- Qt Model ----------------
class DataFrameModel(QAbstractTableModel):
    def __init__(self, df=pd.DataFrame(), parent=None):
        super().__init__(parent)
        self._df = df

    def set_dataframe(self, df: pd.DataFrame):
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def rowCount(self, parent=None):
        return 0 if self._df is None else len(self._df.index)

    def columnCount(self, parent=None):
        return 0 if self._df is None else len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or self._df is None:
            return QVariant()
        r, c = index.row(), index.column()
        if role in (Qt.DisplayRole, Qt.EditRole):
            val = self._df.iat[r, c]
            if pd.isna(val): return ""
            return str(val)
        return QVariant()

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole or self._df is None:
            return QVariant()
        if orientation == Qt.Horizontal:
            try:
                return str(self._df.columns[section])
            except Exception:
                return QVariant()
        else:
            try:
                return str(self._df.index[section])
            except Exception:
                return QVariant()

    def flags(self, index):
        return Qt.ItemIsSelectable | Qt.ItemIsEnabled | Qt.ItemIsEditable

# ---------------- Calculations ----------------
def ties_cutting_length(a_clear: float, b_clear: float, cover: float, dia: float,
                        hook_mult: float=10.0, bend_ded_mult: float=2.0) -> float:
    """
    Rectangular stirrup cutting length (mm), using simple rule of thumb:
    Effective sides = (a_clear + 2*cover - dia), (b_clear + 2*cover - dia)
    Base = 2*(a+b); hooks = 2*(hook_mult*dia); bends deduction = 4*(bend_ded_mult*dia)
    """
    a = a_clear + 2*cover - dia
    b = b_clear + 2*cover - dia
    base = 2*(a + b)
    L_hooks = 2.0 * hook_mult * dia
    L_bend_ded = 4.0 * bend_ded_mult * dia
    return base + L_hooks - L_bend_ded

def beam_main_length(clear_span: float, cover: float, dia: float, dev_len: float) -> float:
    # Very simple: L = clear_span + 2*(cover + dev_len) [mm]
    return clear_span + 2*(cover + dev_len)

def column_longitudinal_length(height_clear: float, cover: float, dia: float, lap_len: float) -> float:
    # Simple: L = height_clear + 2*cover + lap_len [mm]
    return height_clear + 2*cover + lap_len

# ---------------- Main Window ----------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BBS Tool (PyQt5, Dark)")
        self.resize(1200, 800)
        self.settings = QSettings("SushilTools", "BBS_Tool_GUI")
        self.last_dir = self.settings.value("last_dir", os.path.expanduser("~"))

        self.df = pd.DataFrame()
        self.model = DataFrameModel(self.df)

        self._build_ui()
        self._apply_dark_theme()
        self._restore_state()

    def _build_ui(self):
        tb = QToolBar("Main"); tb.setIconSize(QSize(18, 18)); self.addToolBar(tb)
        self.btn_export = QPushButton("Export to Excel"); tb.addWidget(self.btn_export)

        left = QWidget(self); lf = QFormLayout(left); lf.setLabelAlignment(Qt.AlignLeft)
        self.cmb_member = QComboBox(); self.cmb_member.addItems(["Ties/Stirrups", "Beam (simple)", "Column (basic)"])

        # Common
        self.sp_cover = QDoubleSpinBox(); self.sp_cover.setRange(0, 1000); self.sp_cover.setValue(25); self.sp_cover.setSuffix(" mm")
        self.sp_dia = QDoubleSpinBox(); self.sp_dia.setRange(4, 40); self.sp_dia.setValue(8); self.sp_dia.setSuffix(" mm")

        # Ties
        self.sp_a = QDoubleSpinBox(); self.sp_a.setRange(0, 100000); self.sp_a.setValue(230); self.sp_a.setSuffix(" mm")
        self.sp_b = QDoubleSpinBox(); self.sp_b.setRange(0, 100000); self.sp_b.setValue(300); self.sp_b.setSuffix(" mm")
        self.sp_pitch = QDoubleSpinBox(); self.sp_pitch.setRange(10, 1000); self.sp_pitch.setValue(150); self.sp_pitch.setSuffix(" mm")
        self.sp_height = QDoubleSpinBox(); self.sp_height.setRange(0, 100000); self.sp_height.setValue(3000); self.sp_height.setSuffix(" mm")
        self.sp_hook_mult = QDoubleSpinBox(); self.sp_hook_mult.setRange(4, 20); self.sp_hook_mult.setValue(10); self.sp_hook_mult.setSuffix(" ×φ")
        self.sp_bend_mult = QDoubleSpinBox(); self.sp_bend_mult.setRange(0, 10); self.sp_bend_mult.setValue(2); self.sp_bend_mult.setSuffix(" ×φ")

        # Beam
        self.sp_span = QDoubleSpinBox(); self.sp_span.setRange(0, 100000); self.sp_span.setValue(4000); self.sp_span.setSuffix(" mm")
        self.sp_dev = QDoubleSpinBox(); self.sp_dev.setRange(0, 5000); self.sp_dev.setValue(40*20/4); self.sp_dev.setSuffix(" mm")
        self.sp_n_main = QSpinBox(); self.sp_n_main.setRange(1, 20); self.sp_n_main.setValue(2)

        # Column
        self.sp_col_h = QDoubleSpinBox(); self.sp_col_h.setRange(0, 100000); self.sp_col_h.setValue(3000); self.sp_col_h.setSuffix(" mm")
        self.sp_lap = QDoubleSpinBox(); self.sp_lap.setRange(0, 5000); self.sp_lap.setValue(40*20/4); self.sp_lap.setSuffix(" mm")
        self.sp_n_vert = QSpinBox(); self.sp_n_vert.setRange(2, 40); self.sp_n_vert.setValue(8)

        self.btn_compute = QPushButton("Compute BBS")

        lf.addRow(QLabel("Member Type"), self.cmb_member)
        lf.addRow(QLabel("Cover"), self.sp_cover)
        lf.addRow(QLabel("Bar Diameter (φ)"), self.sp_dia)

        self._ties_widgets = [
            (QLabel("Rect. Clear Size a (short)"), self.sp_a),
            (QLabel("Rect. Clear Size b (long)"), self.sp_b),
            (QLabel("Column/Beam Height"), self.sp_height),
            (QLabel("Pitch (c/c)"), self.sp_pitch),
            (QLabel("Hook Length Mult (×φ)"), self.sp_hook_mult),
            (QLabel("Bend Deduction Mult (×φ)"), self.sp_bend_mult),
        ]
        for lab, w in self._ties_widgets: lf.addRow(lab, w)

        self._beam_widgets = [
            (QLabel("Clear Span"), self.sp_span),
            (QLabel("Dev. Length (each end)"), self.sp_dev),
            (QLabel("No. of main bars"), self.sp_n_main),
        ]
        for lab, w in self._beam_widgets: lf.addRow(lab, w)

        self._col_widgets = [
            (QLabel("Column Clear Height"), self.sp_col_h),
            (QLabel("Lap Length"), self.sp_lap),
            (QLabel("No. of vertical bars"), self.sp_n_vert),
        ]
        for lab, w in self._col_widgets: lf.addRow(lab, w)

        lf.addRow(self.btn_compute)

        right = QWidget(self); rv = QVBoxLayout(right)
        self.table = QTableView(); self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setModel(self.model)
        rv.addWidget(self.table, 1)

        central = QWidget(self); layout = QHBoxLayout(central); layout.setContentsMargins(8,8,8,8)
        layout.addWidget(left, 0); layout.addWidget(right, 1)
        self.setCentralWidget(central)

        self.setStatusBar(QStatusBar(self)); self.statusBar().showMessage("Ready")

        # Signals
        self.cmb_member.currentTextChanged.connect(self._toggle_groups)
        self.btn_compute.clicked.connect(self.compute_bbs)
        self.btn_export.clicked.connect(self.export_excel)
        self._toggle_groups()

    def _apply_dark_theme(self):
        qss = """
        QWidget { background: #0f1115; color: #e6e6e6; font-family: Segoe UI, 'Noto Sans', Arial; font-size: 11pt; }
        QToolBar, QStatusBar { background: #151821; border: 0; }
        QTableView, QDoubleSpinBox, QSpinBox, QComboBox {
            background: #12141b; border: 1px solid #2a2f3a; border-radius: 8px;
        }
        QTableView::item:selected { background: #2b3445; }
        QPushButton { background: #1e2633; border: 1px solid #2a2f3a; border-radius: 10px; padding: 6px 12px; }
        QPushButton:hover { background: #2b3445; } QPushButton:pressed { background: #222a39; }
        """
        self.setStyleSheet(qss)

    def _restore_state(self):
        geo = self.settings.value("geometry")
        if isinstance(geo, QByteArray): self.restoreGeometry(geo)
        st = self.settings.value("windowState")
        if isinstance(st, QByteArray): self.restoreState(st)

    def closeEvent(self, e):
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())
        self.settings.setValue("last_dir", self.last_dir)
        super().closeEvent(e)

    def _toggle_groups(self):
        m = self.cmb_member.currentText()
        ties = (m == "Ties/Stirrups"); beam = (m == "Beam (simple)"); col = (m == "Column (basic)")
        for lab, w in self._ties_widgets: lab.setVisible(ties); w.setVisible(ties)
        for lab, w in self._beam_widgets: lab.setVisible(beam); w.setVisible(beam)
        for lab, w in self._col_widgets: lab.setVisible(col); w.setVisible(col)

    # -------- Compute --------
    def compute_bbs(self):
        try:
            member = self.cmb_member.currentText()
            cover = self.sp_cover.value(); dia = self.sp_dia.value()
            unit_wt = steel_unit_weight_mm(dia)
            rows: List[Dict[str, Any]] = []

            if member == "Ties/Stirrups":
                a = self.sp_a.value(); b = self.sp_b.value()
                height = self.sp_height.value(); pitch = self.sp_pitch.value()
                hook_mult = self.sp_hook_mult.value(); bend_mult = self.sp_bend_mult.value()
                n = int(max(1, int(height // pitch) + 1))
                CL = ties_cutting_length(a, b, cover, dia, hook_mult, bend_mult)  # mm
                total_len_m = (CL/1000.0) * n; total_wt = total_len_m * unit_wt
                rows.append({"Mark":"ST","Member":"Stirrups/Ties","Bar φ (mm)":dia,"No.":n,"Spacing/Pitch (mm)":pitch,
                             "Cut Length (mm)":round(CL,0),"Total Length (m)":round(total_len_m,2),
                             "Unit Wt (kg/m)":round(unit_wt,3),"Total Wt (kg)":round(total_wt,2),
                             "Shape":f"Rect with hooks"})

            elif member == "Beam (simple)":
                span = self.sp_span.value(); dev = self.sp_dev.value(); n_main = self.sp_n_main.value()
                L = beam_main_length(span, cover, dia, dev)
                total_len_m = (L/1000.0) * n_main; total_wt = total_len_m * unit_wt
                rows.append({"Mark":"BM","Member":"Beam","Bar φ (mm)":dia,"No.":n_main,"Spacing/Pitch (mm)":"",
                             "Cut Length (mm)":round(L,0),"Total Length (m)":round(total_len_m,2),
                             "Unit Wt (kg/m)":round(unit_wt,3),"Total Wt (kg)":round(total_wt,2),
                             "Shape":"Straight + dev @ ends"})

            elif member == "Column (basic)":
                h = self.sp_col_h.value(); lap = self.sp_lap.value(); n_vert = self.sp_n_vert.value()
                L = column_longitudinal_length(h, cover, dia, lap)
                total_len_m = (L/1000.0) * n_vert; total_wt = total_len_m * unit_wt
                rows.append({"Mark":"CL","Member":"Column","Bar φ (mm)":dia,"No.":n_vert,"Spacing/Pitch (mm)":"",
                             "Cut Length (mm)":round(L,0),"Total Length (m)":round(total_len_m,2),
                             "Unit Wt (kg/m)":round(unit_wt,3),"Total Wt (kg)":round(total_wt,2),
                             "Shape":"Straight + lap"})

            self.df = pd.DataFrame(rows, columns=["Mark","Member","Bar φ (mm)","No.","Spacing/Pitch (mm)",
                                                  "Cut Length (mm)","Total Length (m)","Unit Wt (kg/m)","Total Wt (kg)","Shape"])
            self.model.set_dataframe(self.df)
            self.table.resizeColumnsToContents()
            self.statusBar().showMessage(f"Computed {len(self.df)} BBS row(s).")
        except Exception as e:
            QMessageBox.critical(self, "Compute error", str(e))

    # -------- Export --------
    def export_excel(self):
        if self.df is None or self.df.empty:
            QMessageBox.information(self, "Empty", "Nothing to export. Compute BBS first."); return
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "BBS_Output.xlsx", "Excel Files (*.xlsx)")
        if not path: return
        try:
            template = os.path.join(os.path.dirname(__file__), "templates", "BBS_Template.xlsx")
            if os.path.exists(template):
                import openpyxl
                from openpyxl.utils.dataframe import dataframe_to_rows
                wb = openpyxl.load_workbook(template); ws = wb["BBS"]
                max_row = ws.max_row
                if max_row >= 6: ws.delete_rows(idx=6, amount=max_row-5)
                for r in dataframe_to_rows(self.df, index=False, header=False): ws.append(r)
                wb.save(path)
            else:
                # Fallback: plain export
                self.df.to_excel(path, index=False)
            self.statusBar().showMessage(f"Saved: {path}")
            QMessageBox.information(self, "Export", "Excel exported successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Export error", str(e))

def main():
    app = QApplication(sys.argv)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    w = MainWindow(); w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
