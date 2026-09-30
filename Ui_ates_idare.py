# -*- coding: utf-8 -*-
"""Ateş İdare — modern arayüz. Widget adları atesidare.py ile uyumludur."""

from PyQt5 import QtCore, QtGui, QtWidgets


STYLESHEET = """
/* Katmanlar: zemin (gri) → panel (beyaz) → girdi (beyaz) → vurgu (yalnızca eylem) */
QMainWindow {
    background: #E6EAEF;
    color: #1A2332;
    font-family: "Segoe UI", "Arial";
    font-size: 13px;
}
QWidget {
    background: #E6EAEF;
    color: #1A2332;
    font-family: "Segoe UI", "Arial";
    font-size: 13px;
}
QScrollArea, QScrollArea > QWidget > QWidget {
    background: #E6EAEF;
    border: none;
}
QStatusBar {
    background: #F7F8FA;
    color: #5B6B7C;
    border-top: 1px solid #D0D7DE;
}

QLabel { background: transparent; color: #1A2332; }

QFrame#header {
    background: #1E3A5F;
    border: none;
    border-radius: 0;
}
QLabel#appTitle {
    color: #FFFFFF;
    font-size: 18px;
    font-weight: 700;
    letter-spacing: 1.2px;
    background: transparent;
}
QLabel#appSub {
    color: #C5D4E4;
    font-size: 11px;
    background: transparent;
}
QLabel#headerMeta {
    color: #A8BDD0;
    font-size: 11px;
    background: transparent;
}

QTabWidget {
    background: #F7F8FA;
}
QTabWidget::pane {
    border: 1px solid #D0D7DE;
    border-radius: 4px;
    background: #F7F8FA;
    top: -1px;
}
QTabWidget > QWidget {
    background: #F7F8FA;
}
QTabBar::tab {
    background: #DDE3EA;
    color: #3D4F5F;
    padding: 9px 16px;
    margin-right: 2px;
    border: 1px solid #C5CDD6;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    font-weight: 600;
    min-width: 108px;
}
QTabBar::tab:hover { background: #E8EDF2; color: #1A2332; }
QTabBar::tab:selected {
    background: #F7F8FA;
    color: #1E3A5F;
    border-color: #D0D7DE;
}

QFrame#card, QGroupBox {
    background: #FFFFFF;
    border: 1px solid #D0D7DE;
    border-radius: 6px;
}
QGroupBox {
    margin-top: 12px;
    padding: 14px 12px 12px 12px;
    font-weight: 700;
    color: #1E3A5F;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 6px;
    color: #1E3A5F;
    background: #FFFFFF;
}

QFrame#card[hazir="true"] {
    border: 1px solid #1B7A4E;
    border-left: 4px solid #1B7A4E;
    background: #FFFFFF;
}
QLabel#cardTitle {
    color: #1E3A5F;
    font-size: 13px;
    font-weight: 700;
    background: transparent;
}
QLabel#sectionTitle {
    color: #1E3A5F;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.6px;
    background: transparent;
    padding: 2px 0 4px 0;
}
QLabel#badge {
    background: #1B7A4E;
    color: #FFFFFF;
    border-radius: 3px;
    padding: 2px 8px;
    font-size: 10px;
    font-weight: 700;
}

QLineEdit {
    background: #FFFFFF;
    color: #1A2332;
    border: 1px solid #C5CDD6;
    border-radius: 4px;
    padding: 6px 8px;
    selection-background-color: #1E5AA8;
    selection-color: #FFFFFF;
}
QLineEdit:focus { border: 1px solid #1E5AA8; }
QLineEdit:disabled {
    background: #F0F3F6;
    color: #5B6B7C;
}

QLabel[role="value"] {
    background: #F4F6F8;
    color: #1A2332;
    border: 1px solid #D0D7DE;
    border-radius: 4px;
    padding: 8px 6px;
    font-size: 15px;
    font-weight: 700;
    font-family: "Cascadia Mono", "Consolas", "Segoe UI";
}
QLabel[role="kpi"] {
    background: #F4F6F8;
    color: #1A2332;
    border: 1px solid #D0D7DE;
    border-radius: 4px;
    padding: 8px 8px;
    font-size: 16px;
    font-weight: 700;
    font-family: "Cascadia Mono", "Consolas", "Segoe UI";
    min-height: 26px;
}
QLabel[role="hint"] {
    color: #5B6B7C;
    font-size: 11px;
    font-weight: 600;
    background: transparent;
}
QLabel[role="field"] {
    color: #3D4F5F;
    font-size: 12px;
    font-weight: 600;
    background: transparent;
}

QPushButton {
    background: #FFFFFF;
    color: #1E3A5F;
    border: 1px solid #C5CDD6;
    border-radius: 4px;
    padding: 8px 14px;
    font-weight: 600;
}
QPushButton:hover { background: #F0F3F6; border-color: #1E3A5F; }
QPushButton:pressed { background: #E2E8EE; }
QPushButton:disabled { color: #9AA6B2; background: #F0F3F6; }

QPushButton[role="primary"] {
    background: #1E5AA8;
    color: #FFFFFF;
    border: 1px solid #1E5AA8;
    font-weight: 700;
}
QPushButton[role="primary"]:hover { background: #174A8C; border-color: #174A8C; }
QPushButton[role="primary"]:pressed { background: #123A70; }

QPushButton[role="fire"] {
    background: #B42318;
    color: #FFFFFF;
    border: 1px solid #B42318;
    font-weight: 700;
}
QPushButton[role="fire"]:hover { background: #912018; }
QPushButton[role="danger"] {
    background: #FFFFFF;
    color: #B42318;
    border: 1px solid #E4B4B0;
}
QPushButton[role="danger"]:hover { background: #FEF3F2; }
QPushButton[role="ghost"] {
    background: #FFFFFF;
    color: #5B6B7C;
    border: 1px solid #D0D7DE;
}

QTableWidget {
    background: #FFFFFF;
    alternate-background-color: #F7F8FA;
    gridline-color: #E2E8EE;
    border: 1px solid #D0D7DE;
    border-radius: 4px;
    color: #1A2332;
    selection-background-color: #D6E4F5;
    selection-color: #1A2332;
}
QHeaderView::section {
    background: #1E3A5F;
    color: #FFFFFF;
    padding: 8px;
    border: none;
    border-right: 1px solid #2E4D73;
    font-weight: 700;
}
QTableCornerButton::section { background: #1E3A5F; border: none; }
QScrollBar:vertical, QScrollBar:horizontal {
    background: #E6EAEF;
    width: 10px;
    height: 10px;
    margin: 0;
    border: none;
}
QScrollBar::handle:vertical, QScrollBar::handle:horizontal {
    background: #B7C2CD;
    border-radius: 5px;
    min-height: 24px;
}
QScrollBar::handle:hover { background: #8A99A8; }
QScrollBar::add-line, QScrollBar::sub-line { height: 0; width: 0; }
"""


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1440, 900)
        MainWindow.setMinimumSize(1180, 720)
        MainWindow.setWindowTitle("Ateş İdare")
        MainWindow.setStyleSheet(STYLESHEET)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        root = QtWidgets.QVBoxLayout(self.centralwidget)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_header())

        body = QtWidgets.QWidget()
        body_lay = QtWidgets.QVBoxLayout(body)
        body_lay.setContentsMargins(14, 12, 14, 8)
        body_lay.setSpacing(10)

        self.tabWidget_4 = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget_4.setObjectName("tabWidget_4")
        self.tabWidget_4.setDocumentMode(True)
        body_lay.addWidget(self.tabWidget_4, 1)
        root.addWidget(body, 1)

        self._build_atis_gorevi()
        self._build_btsb()
        self._build_duzeltme()
        self._build_esaslar()
        self._build_hedefler()

        self.tabWidget_4.setCurrentIndex(0)

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setObjectName("statusBar")
        MainWindow.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Hazır")

        self.actionDeneme = QtWidgets.QAction(MainWindow)
        self.actionDeneme.setObjectName("actionDeneme")
        self.actionDeneme.setText("Deneme")

        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def _build_header(self):
        header = QtWidgets.QFrame()
        header.setObjectName("header")
        header.setFixedHeight(56)
        lay = QtWidgets.QHBoxLayout(header)
        lay.setContentsMargins(18, 8, 18, 8)

        titles = QtWidgets.QVBoxLayout()
        titles.setSpacing(0)
        title = QtWidgets.QLabel("ATEŞ İDARE")
        title.setObjectName("appTitle")
        sub = QtWidgets.QLabel("Batarya atış yönetim sistemi")
        sub.setObjectName("appSub")
        titles.addWidget(title)
        titles.addWidget(sub)
        lay.addLayout(titles)
        lay.addStretch()
        meta = QtWidgets.QLabel("1. ve 2. obüs  ·  A–D bölgeleri")
        meta.setObjectName("headerMeta")
        meta.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        lay.addWidget(meta)
        return header

    # ------------------------------------------------------------------ helpers
    def _line(self, name, placeholder="", text=""):
        w = QtWidgets.QLineEdit()
        w.setObjectName(name)
        if placeholder:
            w.setPlaceholderText(placeholder)
        if text:
            w.setText(text)
        setattr(self, name, w)
        return w

    def _btn(self, name, text, role=None):
        w = QtWidgets.QPushButton(text)
        w.setObjectName(name)
        w.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        if role:
            w.setProperty("role", role)
        setattr(self, name, w)
        return w

    def _value(self, name, role="value"):
        w = QtWidgets.QLabel("—")
        w.setObjectName(name)
        w.setProperty("role", role)
        w.setAlignment(QtCore.Qt.AlignCenter)
        w.setMinimumHeight(34)
        setattr(self, name, w)
        return w

    def _form_row(self, layout, label, widget):
        row = QtWidgets.QHBoxLayout()
        lab = QtWidgets.QLabel(label)
        lab.setMinimumWidth(96)
        lab.setProperty("role", "field")
        row.addWidget(lab)
        row.addWidget(widget, 1)
        layout.addLayout(row)

    def _metric(self, parent_lay, caption, widget):
        col = QtWidgets.QVBoxLayout()
        col.setSpacing(3)
        cap = QtWidgets.QLabel(caption)
        cap.setProperty("role", "hint")
        cap.setAlignment(QtCore.Qt.AlignCenter)
        col.addWidget(cap)
        col.addWidget(widget)
        parent_lay.addLayout(col)

    def _table(self, name, rows, headers, row_labels=None):
        t = QtWidgets.QTableWidget(rows, len(headers))
        t.setObjectName(name)
        t.setHorizontalHeaderLabels(headers)
        t.horizontalHeader().setStretchLastSection(True)
        t.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        t.verticalHeader().setDefaultSectionSize(28)
        t.setAlternatingRowColors(True)
        t.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        if row_labels:
            t.setVerticalHeaderLabels(row_labels)
        setattr(self, name, t)
        return t

    # ----------------------------------------------------------- Atış Görevi
    def _build_atis_gorevi(self):
        page = QtWidgets.QWidget()
        page.setObjectName("tabWidget_4Page2")
        self.tabWidget_4.addTab(page, "Atış Görevi")

        grid = QtWidgets.QGridLayout(page)
        grid.setContentsMargins(8, 10, 8, 8)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(10)
        grid.setColumnStretch(0, 0)
        grid.setColumnStretch(1, 1)

        left_wrap = QtWidgets.QWidget()
        left_wrap.setFixedWidth(312)
        left = QtWidgets.QVBoxLayout(left_wrap)
        left.setContentsMargins(0, 0, 0, 0)
        left.setSpacing(10)
        left.addWidget(self._hedef_karti())
        left.addWidget(self._barut_karti())
        self.btn_ilk_hesaplama = self._btn("btn_ilk_hesaplama", "HESAPLA", "primary")
        self.btn_ilk_hesaplama.setMinimumHeight(44)
        left.addWidget(self.btn_ilk_hesaplama)
        left.addStretch()
        grid.addWidget(left_wrap, 0, 0)

        right = QtWidgets.QScrollArea()
        right.setWidgetResizable(True)
        right.setFrameShape(QtWidgets.QFrame.NoFrame)
        wrap = QtWidgets.QWidget()
        v = QtWidgets.QVBoxLayout(wrap)
        v.setContentsMargins(0, 0, 4, 0)
        v.setSpacing(8)

        hazir_title = QtWidgets.QLabel("HAZIR OBÜSLER")
        hazir_title.setObjectName("sectionTitle")
        v.addWidget(hazir_title)
        hazir = QtWidgets.QHBoxLayout()
        hazir.setSpacing(8)
        hazir.addWidget(self._atis_karti("1B", "1. Obüs · B bölgesi", True))
        hazir.addWidget(self._atis_karti("2B", "2. Obüs · B bölgesi", True))
        v.addLayout(hazir)

        diger_title = QtWidgets.QLabel("DİĞER OBÜSLER")
        diger_title.setObjectName("sectionTitle")
        v.addWidget(diger_title)
        diger = QtWidgets.QGridLayout()
        diger.setSpacing(8)
        diger.addWidget(self._atis_karti("1A", "1. Obüs · A"), 0, 0)
        diger.addWidget(self._atis_karti("1C", "1. Obüs · C"), 0, 1)
        diger.addWidget(self._atis_karti("1D", "1. Obüs · D"), 0, 2)
        diger.addWidget(self._atis_karti("2A", "2. Obüs · A"), 1, 0)
        diger.addWidget(self._atis_karti("2C", "2. Obüs · C"), 1, 1)
        diger.addWidget(self._atis_karti("2D", "2. Obüs · D"), 1, 2)
        v.addLayout(diger)
        v.addStretch()
        right.setWidget(wrap)
        grid.addWidget(right, 0, 1)

        # Eski toolbox/tab adları — uyumluluk için boş nesneler
        self.toolBox = QtWidgets.QToolBox()
        self.toolBox.setVisible(False)
        self.toolBoxPage1 = QtWidgets.QWidget()
        self.toolBoxPage2 = QtWidgets.QWidget()

    def _hedef_karti(self):
        box = QtWidgets.QGroupBox("Hedef bilgileri")
        box.setObjectName("groupBox_17")
        lay = QtWidgets.QVBoxLayout(box)
        lay.setSpacing(8)

        bolge = QtWidgets.QHBoxLayout()
        self._form_like = None
        lab = QtWidgets.QLabel("Bölge")
        lab.setMinimumWidth(96)
        lab.setProperty("role", "field")
        bolge.addWidget(lab)
        bolge.addWidget(self._line("lne_hedef_bolge_nu", "37", "37"), 1)
        bolge.addWidget(self._line("lne_hedef_bolge_sayisi", "SFA", "SFA"), 1)
        lay.addLayout(bolge)

        self._form_row(lay, "Sağ", self._line("lne_hedef_sag_deger", "19072"))
        self._form_row(lay, "Yukarı", self._line("lne_hedef_yukari_deger", "60022"))
        self._form_row(lay, "Rakım", self._line("lne_hedef_rakim", "boşsa API"))
        self.sil_btn = self._btn("sil_btn", "Hedefi sil", "ghost")
        lay.addWidget(self.sil_btn)
        return box

    def _barut_karti(self):
        box = QtWidgets.QGroupBox("Barut hakkı")
        box.setObjectName("groupBox_18")
        lay = QtWidgets.QVBoxLayout(box)
        lay.setSpacing(8)
        self._form_row(lay, "Barut ısısı", self._line("lne_barut_isisi", "örn. 75"))
        self._form_row(lay, "Önerilen BH", self._line("lne_barut_hakki", "hesaplanır"))
        row = QtWidgets.QHBoxLayout()
        row.addWidget(self._btn("btn_barut_hakki_oner", "Barut hakkı öner"))
        row.addWidget(self._btn("btn_barut_hakki_degistir", "Değiştir"))
        lay.addLayout(row)
        return box

    def _atis_karti(self, kod, baslik, hazir=False):
        card = QtWidgets.QFrame()
        card.setObjectName("card")
        if hazir:
            card.setProperty("hazir", "true")
        lay = QtWidgets.QVBoxLayout(card)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(8)

        head = QtWidgets.QHBoxLayout()
        title = QtWidgets.QLabel(baslik)
        title.setObjectName("cardTitle")
        head.addWidget(title)
        head.addStretch()
        if hazir:
            badge = QtWidgets.QLabel("HAZIR")
            badge.setObjectName("badge")
            badge.setAlignment(QtCore.Qt.AlignCenter)
            head.addWidget(badge)
        lay.addLayout(head)

        bh_name = "sonuc_batur_hakki_1A" if kod == "1A" else "sonuc_barut_hakki_%s" % kod
        grid = QtWidgets.QGridLayout()
        grid.setSpacing(6)
        pairs = [
            (0, 0, "YAN", "sonuc_yan_%s" % kod),
            (0, 1, "YÜKSELİŞ", "sonuc_yukselis_%s" % kod),
            (1, 0, "BARUT HAKKI", bh_name),
            (1, 1, "TAPA", "sonuc_tapa_%s" % kod),
            (2, 0, "İSTİKAMET AÇISI", "sonuc_istikamet_acisi_%s" % kod),
            (2, 1, "MESAFE", "sonuc_mesafe_%s" % kod),
        ]
        for r, c, cap, name in pairs:
            cell = QtWidgets.QVBoxLayout()
            cell.setSpacing(2)
            lab = QtWidgets.QLabel(cap)
            lab.setProperty("role", "hint")
            lab.setAlignment(QtCore.Qt.AlignCenter)
            cell.addWidget(lab)
            cell.addWidget(self._value(name))
            grid.addLayout(cell, r, c)
        lay.addLayout(grid)
        return card

    # ----------------------------------------------------- Bt.Sb / mevzi / müh
    def _build_btsb(self):
        page = QtWidgets.QWidget()
        page.setObjectName("tabWidget_4Page1")
        self.tabWidget_4.addTab(page, "Bt.Sb. Raporu")
        v = QtWidgets.QVBoxLayout(page)
        v.setContentsMargins(6, 8, 6, 6)

        self.tabWidget = QtWidgets.QTabWidget()
        self.tabWidget.setObjectName("tabWidget")
        v.addWidget(self.tabWidget)

        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.tabWidget.addTab(self.tab, "Mevzi bilgileri")
        mevzi = QtWidgets.QScrollArea(self.tab)
        mevzi.setWidgetResizable(True)
        mevzi.setFrameShape(QtWidgets.QFrame.NoFrame)
        outer = QtWidgets.QVBoxLayout(self.tab)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(mevzi)
        host = QtWidgets.QWidget()
        grid = QtWidgets.QGridLayout(host)
        grid.setSpacing(10)
        specs = [
            ("1A", 0, 0, True), ("1B", 0, 1, False), ("1C", 0, 2, False), ("1D", 0, 3, False),
            ("2A", 1, 0, False), ("2B", 1, 1, False), ("2C", 1, 2, False), ("2D", 1, 3, False),
        ]
        defaults = {
            "1A": ("19034", "65603", "364", "2400", "10.0", "2600"),
            "1B": ("19034", "65603", "364", "3000", "10.0", "2600"),
            "1C": ("19034", "65603", "364", "3450", "10.0", "2600"),
            "1D": ("19034", "65603", "364", "3700", "10.0", "2600"),
            "2A": ("19054", "65616", "364", "1000", "9.0", "2600"),
            "2B": ("19054", "65616", "364", "1800", "9.0", "2600"),
            "2C": ("19054", "65616", "364", "1800", "9.0", "2600"),
            "2D": ("19054", "65616", "364", "200", "9.0", "2600"),
        }
        for kod, r, c, first in specs:
            obus = "1. Obüs" if kod[0] == "1" else "2. Obüs"
            grid.addWidget(self._mevzi_karti(kod, "%s · %s bölgesi" % (obus, kod[1]), defaults[kod], first), r, c)
        mevzi.setWidget(host)

        self.toolBox_2 = QtWidgets.QToolBox()
        self.toolBox_2.setVisible(False)

        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.tabWidget.addTab(self.tab_2, "Mühimmat bilgileri")
        muh = QtWidgets.QVBoxLayout(self.tab_2)
        ammo_tabs = QtWidgets.QTabWidget()
        muh.addWidget(ammo_tabs)

        mermi_page = QtWidgets.QWidget()
        ammo_tabs.addTab(mermi_page, "Mermi")
        QtWidgets.QVBoxLayout(mermi_page).addWidget(self._table(
            "tableWidget", 8,
            ["MENŞEİ", "MODELİ", "KAFİLESİ", "TİPİ", "KARE AĞIRLIĞI", "KALAN MİKTARI", "TARİH"],
            [str(i) for i in range(1, 9)],
        ))

        tapa_page = QtWidgets.QWidget()
        ammo_tabs.addTab(tapa_page, "Tapa")
        QtWidgets.QVBoxLayout(tapa_page).addWidget(self._table(
            "tableWidget_2", 8,
            ["MENŞEİ", "MODELİ", "KAFİLESİ", "TİPİ", "KALAN MİKTARI", "TARİH"],
            [str(i) for i in range(1, 9)],
        ))

        barut_page = QtWidgets.QWidget()
        ammo_tabs.addTab(barut_page, "Barut")
        QtWidgets.QVBoxLayout(barut_page).addWidget(self._table(
            "tableWidget_3", 8,
            ["MENŞEİ", "MODELİ", "KAFİLESİ", "KESE RENGİ", "KALAN MİKTARI", "TARİH"],
            [str(i) for i in range(1, 9)],
        ))

        self.toolBox_3 = QtWidgets.QToolBox()
        self.toolBox_3.setVisible(False)

    def _mevzi_karti(self, kod, baslik, placeholders, is_1a):
        box = QtWidgets.QGroupBox(baslik)
        lay = QtWidgets.QVBoxLayout(box)
        sag, yukari, rakim, ahia, ihf, musyan = placeholders
        if is_1a:
            names = (
                "lne_1A_obus_sag", "lne_1A_obus_yukari", "lne_1A_obus_rakim",
                "lne_1A_obus_AHIA", "lne_1A_obus_IHF", "lne_1A_obus_MUSYAN",
            )
        else:
            names = (
                "lne_%s_obus_sag" % kod, "lne_%s_obus_yukari" % kod, "lne_%s_obus_rakim" % kod,
                "lne_%s_obus_ahia" % kod, "lne_%s_obus_ihf" % kod, "lne_%s_obus_musyan" % kod,
            )
        self._form_row(lay, "Sağ", self._line(names[0], sag))
        self._form_row(lay, "Yukarı", self._line(names[1], yukari))
        self._form_row(lay, "Rakım", self._line(names[2], rakim))
        self._form_row(lay, "AHİA", self._line(names[3], ahia))
        self._form_row(lay, "İHF", self._line(names[4], ihf))
        self._form_row(lay, "Müş. yan", self._line(names[5], musyan))
        lay.addWidget(self._btn("btn_mevzi_%s" % kod, "Mevzi kaydet"))
        return box

    # --------------------------------------------------------------- Düzeltme
    def _build_duzeltme(self):
        page = QtWidgets.QWidget()
        page.setObjectName("tabWidget_4Page3")
        self.tabWidget_4.addTab(page, "Düzeltme")
        grid = QtWidgets.QGridLayout(page)
        grid.setContentsMargins(8, 10, 8, 8)
        grid.setSpacing(10)

        grid.addWidget(self._dzl_atis_esas(), 0, 0)
        grid.addWidget(self._dzl_pad(), 0, 1)
        grid.addWidget(self._dzl_hedef(), 0, 2)
        grid.addWidget(self._dzl_bilgi(), 0, 3)

        actions = QtWidgets.QVBoxLayout()
        self.btn_hesapla = self._btn("btn_hesapla", "HESAPLA", "primary")
        self.btn_hesapla.setMinimumHeight(44)
        self.btn_atildi = self._btn("btn_atildi", "ATILDI", "fire")
        self.btn_atildi.setMinimumHeight(44)
        self.btn_yazdir = self._btn("btn_yazdir", "YAZDIR")
        self.btn_yazdir.setMinimumHeight(44)
        actions.addWidget(self.btn_hesapla)
        actions.addWidget(self.btn_atildi)
        actions.addWidget(self.btn_yazdir)
        actions.addStretch()
        grid.addLayout(actions, 0, 4)

        self.tablo_atislar = self._table(
            "tablo_atislar", 20,
            ["TARİH-SAAT", "YAN", "YÜKSELİŞ", "TAPA SANİYESİ",
             "SAĞ", "SOL", "UZAT", "KISALT", "KALDIR", "İNDİR"],
            [str(i) for i in range(1, 21)],
        )
        grid.addWidget(self.tablo_atislar, 1, 0, 1, 5)
        grid.setRowStretch(1, 1)
        grid.setColumnStretch(0, 2)
        grid.setColumnStretch(1, 2)
        grid.setColumnStretch(2, 2)
        grid.setColumnStretch(3, 2)
        grid.setColumnStretch(4, 0)

    def _dzl_atis_esas(self):
        box = QtWidgets.QGroupBox("Atış esasları")
        box.setObjectName("groupBox_32")
        lay = QtWidgets.QGridLayout(box)
        pairs = [
            (0, "YAN", "lbl_sonuc_yan_2"),
            (1, "YÜKSELİŞ", "lbl_sonuc_yukselis_2"),
            (2, "BARUT HAKKI", "lbl_sonuc_barut_hakki_2"),
            (3, "TAPA SANİYESİ", "lbl_sonuc_tapa_saniyesi"),
        ]
        for r, cap, name in pairs:
            lab = QtWidgets.QLabel(cap)
            lab.setProperty("role", "hint")
            lay.addWidget(lab, r, 0)
            w = self._value(name, "kpi")
            w.setText("")
            lay.addWidget(w, r, 1)
        return box

    def _dzl_pad(self):
        box = QtWidgets.QGroupBox("Düzeltmeler")
        box.setObjectName("groupBox_27")
        lay = QtWidgets.QGridLayout(box)
        fields = [
            (0, 0, "Sola", "lne_dzl_sola_2"),
            (0, 1, "Sağa", "lne_dzl_saga_2"),
            (1, 0, "Uzat", "lne_dzl_uzat_2"),
            (1, 1, "Kısalt", "lne_dzl_kisalt_2"),
            (2, 0, "Kaldır", "lne_dzl_kaldir_2"),
            (2, 1, "İndir", "lne_dzl_indir_2"),
        ]
        for r, c, cap, name in fields:
            cell = QtWidgets.QVBoxLayout()
            cell.setSpacing(3)
            lab = QtWidgets.QLabel(cap)
            lab.setProperty("role", "hint")
            lab.setAlignment(QtCore.Qt.AlignCenter)
            cell.addWidget(lab)
            cell.addWidget(self._line(name, "0"))
            lay.addLayout(cell, r, c)
        return box

    def _dzl_hedef(self):
        box = QtWidgets.QGroupBox("Hedef koordinatı")
        box.setObjectName("groupBox_31")
        lay = QtWidgets.QVBoxLayout(box)
        self._form_row(lay, "Sağ", self._line("lne_hedef_sag_deger_sonuc"))
        self._form_row(lay, "Yukarı", self._line("lne_hedef_yukari_deger_sonuc"))
        self._form_row(lay, "Rakım", self._line("lne_hedef_rakim_sonuc"))
        return box

    def _dzl_bilgi(self):
        box = QtWidgets.QGroupBox("Düzeltme bilgileri")
        box.setObjectName("groupBox_33")
        lay = QtWidgets.QVBoxLayout(box)
        self._form_row(lay, "Mesafe", self._line("lne_sonuc_mesafe"))
        self._form_row(lay, "BH", self._line("lne_sonuc_barut_hakki"))
        self._form_row(lay, "Yan", self._line("lne_sonuc_yan"))
        self._form_row(lay, "Yükseliş", self._line("lne_sonuc_yukselis"))
        self._form_row(lay, "Tapa sn", self._line("lne_sonuc_tapa_saniyesi"))
        self._form_row(lay, "1 mil değişim", self._line("lne_sonuc_bir_ml_degisiklil"))
        self._form_row(lay, "M / 100", self._line("lne_sonuc_yuz_m"))
        return box

    # ---------------------------------------------------------------- Esaslar
    def _build_esaslar(self):
        page = QtWidgets.QWidget()
        page.setObjectName("tabWidget_4Page4")
        self.tabWidget_4.addTab(page, "Esaslar")
        grid = QtWidgets.QGridLayout(page)
        grid.setContentsMargins(12, 12, 12, 12)
        grid.setHorizontalSpacing(28)
        grid.setVerticalSpacing(8)

        left = [
            ("Plan yanı", "lbl_sonuc_plan_yani"),
            ("Yükseliş", "lbl_sonuc_yukselis_6"),
            ("Mesafe", "lbl_sonuc_mesafe"),
            ("Barut hakkı", "lbl_sonuc_barut_hakki_20"),
            ("İstikamet açısı", "lbl_sonuc_istikamet_acisi"),
            ("Uçuş süresi", "lbl_sonuc_ucus_suresi"),
            ("Nişangah", "lbl_sonuc_nisangah"),
            ("Doğal yan düzeltmesi", "lbl_sonuc_dogal_yan_dzl"),
            ("Mesafede sapma olasılığı", "lbl_sonuc_mso"),
            ("Yanca sapma olasılığı", "lbl_sonuc_yso"),
            ("100 / M", "lbl_sonuc_100_M"),
            ("20 / M", "lbl_sonuc_20_M"),
            ("Düşüş açısı", "lbl_sonuc_dusus_acisi"),
            ("Tepe yüksekliği", "lbl_sonuc_tepe_yuksekligi"),
        ]
        right = [
            ("Doğru toprak açısı (DTAÇ)", "lbl_sonuc_dtac"),
            ("Toprak açısı (TAÇ)", "lbl_sonuc_tac"),
            ("GAC yan düzeltmesi", "lbl_sonuc_gac_yan_duzeltmesi"),
            ("Toplam yan düzeltmesi", "lbl_sonuc_toplam_yan_duzeltmesi"),
            ("Toplam mesafe düzeltmesi", "lbl_sonuc_toplam_mesafe_duzeltmesi"),
            ("Toplam tapa saniye düzeltmesi", "lbl_sonuc_toplam_tapa_saniye_duzeltmesi"),
            ("En küçük yükseliş", "lbl_sonuc_en_kucuk_yukselis"),
            ("Metro yan düzeltmesi", "lbl_sonuc_metro_yan_duzeltmesi"),
            ("Metro mesafe düzeltmesi", "lbl_sonuc_metro_mesafe_duzeltmesi"),
            ("GAC mesafesi", "lbl_gac_mesafesi"),
        ]
        self._kpi_column(grid, 0, left)
        self._kpi_column(grid, 1, right)
        grid.setColumnStretch(1, 1)
        grid.setColumnStretch(3, 1)

    def _kpi_column(self, grid, col, items):
        for i, (cap, name) in enumerate(items):
            lab = QtWidgets.QLabel(cap)
            lab.setProperty("role", "hint")
            w = self._value(name, "kpi")
            w.setText("")
            grid.addWidget(lab, i, col * 2)
            grid.addWidget(w, i, col * 2 + 1)

    # --------------------------------------------------------------- Hedefler
    def _build_hedefler(self):
        page = QtWidgets.QWidget()
        page.setObjectName("tabWidget_4Page5")
        self.tabWidget_4.addTab(page, "Hedefler")
        lay = QtWidgets.QVBoxLayout(page)
        t = self._table(
            "table_hedef_listesi", 2,
            ["HEDEFİN ADI", "HEDEFİN TANIMI", "BÖLGE", "BÖLGE NU",
             "SAĞ DEĞER", "YUKARI DEĞER", "RAKIM"],
            ["MAVİ1", "KAPAK9"],
        )
        data = [
            ["MAVİ 1", "GENEL HEDEF", "37", "SFA", "23480", "64750", "351"],
            ["KAPAK 9", "KAPAKLAR", "37", "SFA", "19072", "60022", "339"],
        ]
        for r, row in enumerate(data):
            for c, val in enumerate(row):
                t.setItem(r, c, QtWidgets.QTableWidgetItem(val))
        lay.addWidget(t)
