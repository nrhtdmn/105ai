import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox
from Ui_ates_idare import Ui_MainWindow
import math
import mgrs
import requests
import shutil
import os
import json
import traceback
from datetime import datetime
import time
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MEVZI_DOSYA = os.path.join(BASE_DIR, "mevzi_kayit.json")
METRO_DOSYA = os.path.join(BASE_DIR, "METRAP.txt")
ATIS_KAYIT_DIR = os.path.join(BASE_DIR, "atiskayitlari")
ATIS_SABLON = os.path.join(BASE_DIR, "Atis_Kayit_Formu.xlsx")
HATA_LOG = os.path.join(BASE_DIR, "hata.log")
_HATA_EKRANDA = False


def hata_goster(hata, ayrinti=None):
    """Hatayı kaydeder, uyarı gösterir; programı kapatmaz."""
    global _HATA_EKRANDA
    metin = str(hata) if hata is not None else "Bilinmeyen hata"
    if ayrinti:
        kayit = ayrinti
    else:
        kayit = traceback.format_exc() or metin
    try:
        with open(HATA_LOG, "a", encoding="utf-8") as f:
            f.write("\n%s\n%s\n" % (datetime.now().strftime("%d.%m.%Y %H:%M:%S"), kayit))
    except OSError:
        pass
    if _HATA_EKRANDA:
        return
    _HATA_EKRANDA = True
    try:
        kutu = QMessageBox()
        kutu.setIcon(QMessageBox.Warning)
        kutu.setWindowTitle("Hata")
        kutu.setText("İşlem tamamlanamadı. Program açık kalıyor.")
        kutu.setInformativeText(metin[:400])
        kutu.setDetailedText(kayit[-4000:])
        kutu.setStandardButtons(QMessageBox.Ok)
        kutu.exec_()
    except Exception:
        pass
    finally:
        _HATA_EKRANDA = False


def guvenli_slot(fn):
    def sarmal(self, *args, **kwargs):
        try:
            return fn(self, *args, **kwargs)
        except Exception as exc:
            hata_goster(exc, traceback.format_exc())
            return None
    sarmal.__name__ = fn.__name__
    return sarmal


def _genel_hata_yakala(exc_type, exc, tb):
    if issubclass(exc_type, (KeyboardInterrupt, SystemExit)):
        return
    hata_goster(exc, "".join(traceback.format_exception(exc_type, exc, tb)))


class SafeApplication(QtWidgets.QApplication):
    def notify(self, alici, olay):
        try:
            return super(SafeApplication, self).notify(alici, olay)
        except Exception as exc:
            hata_goster(exc, traceback.format_exc())
            return False

class myApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(myApp, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        
        self.ui.btn_mevzi_1A.clicked.connect(lambda: self._mevzi_kaydet_tikla("1A"))
        self.ui.btn_mevzi_1B.clicked.connect(lambda: self._mevzi_kaydet_tikla("1B"))
        self.ui.btn_mevzi_1C.clicked.connect(lambda: self._mevzi_kaydet_tikla("1C"))
        self.ui.btn_mevzi_1D.clicked.connect(lambda: self._mevzi_kaydet_tikla("1D"))
        self.ui.btn_mevzi_2A.clicked.connect(lambda: self._mevzi_kaydet_tikla("2A"))
        self.ui.btn_mevzi_2B.clicked.connect(lambda: self._mevzi_kaydet_tikla("2B"))
        self.ui.btn_mevzi_2C.clicked.connect(lambda: self._mevzi_kaydet_tikla("2C"))
        self.ui.btn_mevzi_2D.clicked.connect(lambda: self._mevzi_kaydet_tikla("2D"))
        self.ui.btn_barut_hakki_oner.clicked.connect(self.barut_hakki_oner)
        self.ui.btn_barut_hakki_degistir.clicked.connect(self.secilen_barut_hakki)
        self.ui.btn_ilk_hesaplama.clicked.connect(self.ilk_hesaplama)
        self.ui.btn_hesapla.clicked.connect(self.dzl_hesapla)
        self.ui.btn_atildi.clicked.connect(self.atildi)
        self.ui.btn_yazdir.clicked.connect(self.yazdir)
        self.ui.sil_btn.clicked.connect(self.sil)
        self.ui.table_hedef_listesi.cellClicked.connect(self._hedef_sec)

        self.satir_sayisi = 0
        self.isleme_alindi = False
        os.makedirs(ATIS_KAYIT_DIR, exist_ok=True)
        self._mevzi_yukle()
        if not self._tablo_hucre(self.ui.tableWidget, 0, 3):
            self.ui.tableWidget.setItem(0, 3, QtWidgets.QTableWidgetItem("TD"))
        if not self._tablo_hucre(self.ui.tableWidget, 0, 4):
            self.ui.tableWidget.setItem(0, 4, QtWidgets.QTableWidgetItem("2"))
        if not self._tablo_hucre(self.ui.tableWidget_2, 0, 3):
            self.ui.tableWidget_2.setItem(0, 3, QtWidgets.QTableWidgetItem("HS"))

    def _txt(self, name, value):
        w = getattr(self.ui, name, None)
        if w is None:
            return
        if value is None:
            w.setText("")
            return
        if isinstance(value, float):
            w.setText(str(round(value, 2)))
        else:
            w.setText(str(value))

    def _tablo_hucre(self, table, row, col):
        item = table.item(row, col)
        return item.text().strip() if item and item.text() else ""

    def _muhimmat_oku(self):
        mermi, tapa, kare = "TD", "HS", 2
        tip = self._tablo_hucre(self.ui.tableWidget, 0, 3)
        if tip:
            mermi = tip
        kare_txt = self._tablo_hucre(self.ui.tableWidget, 0, 4)
        if kare_txt:
            try:
                kare = int(float(kare_txt))
            except ValueError:
                pass
        tapa_tip = self._tablo_hucre(self.ui.tableWidget_2, 0, 3)
        if tapa_tip:
            tapa = tapa_tip
        return mermi, tapa, kare

    def _mevzi_alanlari(self, kod):
        if kod == "1A":
            return {
                "sag": self.ui.lne_1A_obus_sag,
                "yukari": self.ui.lne_1A_obus_yukari,
                "rakim": self.ui.lne_1A_obus_rakim,
                "ahia": self.ui.lne_1A_obus_AHIA,
                "ihf": self.ui.lne_1A_obus_IHF,
                "musyan": self.ui.lne_1A_obus_MUSYAN,
            }
        return {
            "sag": getattr(self.ui, "lne_%s_obus_sag" % kod),
            "yukari": getattr(self.ui, "lne_%s_obus_yukari" % kod),
            "rakim": getattr(self.ui, "lne_%s_obus_rakim" % kod),
            "ahia": getattr(self.ui, "lne_%s_obus_ahia" % kod),
            "ihf": getattr(self.ui, "lne_%s_obus_ihf" % kod),
            "musyan": getattr(self.ui, "lne_%s_obus_musyan" % kod),
        }

    def _mevzi_diske_yaz(self):
        data = {}
        for kod in ("1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D"):
            alan = self._mevzi_alanlari(kod)
            data[kod] = {k: w.text() for k, w in alan.items()}
        with open(MEVZI_DOSYA, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _mevzi_yukle(self):
        if not os.path.exists(MEVZI_DOSYA):
            return
        try:
            with open(MEVZI_DOSYA, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            return
        for kod, degerler in data.items():
            try:
                alan = self._mevzi_alanlari(kod)
            except AttributeError:
                continue
            for k, w in alan.items():
                if degerler.get(k):
                    w.setText(str(degerler[k]))

    @guvenli_slot
    def _mevzi_kaydet_tikla(self, kod):
        getattr(self, "mevzi_%s" % kod)()
        self._mevzi_diske_yaz()
        QMessageBox.information(self, "Mevzi", "%s mevzi bilgileri kaydedildi." % kod)

    @guvenli_slot
    def _hedef_sec(self, row, _col):
        t = self.ui.table_hedef_listesi
        def hucre(c):
            item = t.item(row, c)
            return item.text().strip() if item else ""
        if hucre(2):
            self.ui.lne_hedef_bolge_nu.setText(hucre(2))
        if hucre(3):
            self.ui.lne_hedef_bolge_sayisi.setText(hucre(3))
        if hucre(4):
            self.ui.lne_hedef_sag_deger.setText(hucre(4))
        if hucre(5):
            self.ui.lne_hedef_yukari_deger.setText(hucre(5))
        if hucre(6):
            self.ui.lne_hedef_rakim.setText(hucre(6))

    def _doldur_diger_ve_esaslar(self, grup, yan_duzeltme, yanlar, ref_mesafe, ref_yukselis, birmil, plan_mesafeler, esas):
        kodlar = ("1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D")
        try:
            birmil = float(birmil) if birmil else 1.0
        except (TypeError, ValueError):
            birmil = 1.0
        if birmil == 0:
            birmil = 1.0
        try:
            yan_duzeltme = float(yan_duzeltme)
        except (TypeError, ValueError):
            yan_duzeltme = 0.0
        for i, kod in enumerate(kodlar):
            try:
                yan = round(float(yanlar[i]) + yan_duzeltme)
            except (TypeError, ValueError, IndexError):
                continue
            self._txt("sonuc_yan_%s" % kod, yan)
            try:
                dy = (float(plan_mesafeler[i]) - float(ref_mesafe)) / birmil
                self._txt("sonuc_yukselis_%s" % kod, round(float(ref_yukselis) + dy, 1))
            except (TypeError, ValueError, IndexError):
                pass
        self._txt("lbl_sonuc_plan_yani", esas.get("plan_yan"))
        self._txt("lbl_sonuc_yukselis_6", esas.get("yukselis"))
        self._txt("lbl_sonuc_mesafe", esas.get("mesafe"))
        self._txt("lbl_sonuc_barut_hakki_20", esas.get("barut"))
        self._txt("lbl_sonuc_istikamet_acisi", esas.get("istikamet"))
        self._txt("lbl_sonuc_ucus_suresi", esas.get("ucus"))
        self._txt("lbl_sonuc_nisangah", esas.get("nisangah"))
        self._txt("lbl_sonuc_dogal_yan_dzl", esas.get("dogal_yan"))
        self._txt("lbl_sonuc_mso", esas.get("mso"))
        self._txt("lbl_sonuc_yso", esas.get("yso"))
        self._txt("lbl_sonuc_100_M", esas.get("yuz_m"))
        self._txt("lbl_sonuc_20_M", esas.get("yirmi_m"))
        self._txt("lbl_sonuc_dusus_acisi", esas.get("dusus"))
        self._txt("lbl_sonuc_tepe_yuksekligi", esas.get("tepe"))
        self._txt("lbl_sonuc_dtac", esas.get("dtac"))
        self._txt("lbl_sonuc_tac", esas.get("tac"))
        self._txt("lbl_sonuc_gac_yan_duzeltmesi", esas.get("gac_yan"))
        self._txt("lbl_sonuc_toplam_yan_duzeltmesi", esas.get("toplam_yan"))
        self._txt("lbl_sonuc_toplam_mesafe_duzeltmesi", esas.get("toplam_mesafe"))
        self._txt("lbl_sonuc_toplam_tapa_saniye_duzeltmesi", esas.get("toplam_tapa"))
        self._txt("lbl_sonuc_en_kucuk_yukselis", esas.get("en_kucuk"))
        self._txt("lbl_sonuc_metro_yan_duzeltmesi", esas.get("metro_yan"))
        self._txt("lbl_sonuc_metro_mesafe_duzeltmesi", esas.get("metro_mesafe"))
        self._txt("lbl_gac_mesafesi", esas.get("gac_mesafe"))

    def _doldur_sonuc(self, grup, yan_duzeltme, yukselis, birmil, loc):
        def g(*names):
            for name in names:
                if name in loc:
                    val = loc[name]
                    if isinstance(val, tuple):
                        continue
                    return val
            return None
        ref_mesafe = g("plan_mesafesi_1B_obüs") if grup == "1" else g("plan_mesafesi_2B_obüs")
        esas = dict(
            plan_yan=g("yan_1B") if grup == "1" else g("yan_2B"),
            yukselis=yukselis,
            mesafe=g("mesafe"),
            barut=g("secilen_barut_hakki"),
            istikamet=g("atis_istikameti_1", "atis_istikameti"),
            ucus=g("ucussuresi"),
            nisangah=g("nisangah5bh", "nisangah6bh"),
            dogal_yan=g("dogalyandz"),
            yuz_m=(float(g("mesafe")) / 1000.0) if g("mesafe") not in (None, "") else None,
            yirmi_m=birmil,
            dusus=g("dusus_acisi"),
            tepe=g("tepe_yuksekligi"),
            dtac=g("dtac"),
            tac=g("tac"),
            gac_yan=g("gac_yan_duzeltmesi2", "gac_yan_duzeltmesi"),
            toplam_yan=yan_duzeltme,
            toplam_mesafe=g("toplam_mesafe_duzeltmesi"),
            toplam_tapa=g("deltaTS"),
            metro_yan=g("metro_yan_duzeltmesi2", "metro_yan_duzeltmesi"),
            metro_mesafe=g("metro_mesafe_duzeltmesi"),
            gac_mesafe=g("gac_mesafe"),
        )
        self._doldur_diger_ve_esaslar(
            grup, yan_duzeltme,
            (g("yan_1A"), g("yan_1B"), g("yan_1C"), g("yan_1D"),
             g("yan_2A"), g("yan_2B"), g("yan_2C"), g("yan_2D")),
            ref_mesafe, yukselis, birmil,
            (g("plan_mesafesi_1A_obüs"), g("plan_mesafesi_1B_obüs"), g("plan_mesafesi_1C_obüs"), g("plan_mesafesi_1D_obüs"),
             g("plan_mesafesi_2A_obüs"), g("plan_mesafesi_2B_obüs"), g("plan_mesafesi_2C_obüs"), g("plan_mesafesi_2D_obüs")),
            esas,
        )
    def mevzi_1A(self):
        """BİRİNCİ OBÜSÜN 1A MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_1A_obus_sag = int(self.ui.lne_1A_obus_sag.text()) if self.ui.lne_1A_obus_sag.text() else 19034
        lne_1A_obus_yukari = int(self.ui.lne_1A_obus_yukari.text()) if self.ui.lne_1A_obus_yukari.text() else 65603
        lne_1A_obus_rakim = int(self.ui.lne_1A_obus_rakim.text()) if self.ui.lne_1A_obus_rakim.text() else 364
        lne_1A_obus_AHIA = int(self.ui.lne_1A_obus_AHIA.text()) if self.ui.lne_1A_obus_AHIA.text() else 2400
        lne_1A_obus_IHF = float(self.ui.lne_1A_obus_IHF.text()) if self.ui.lne_1A_obus_IHF.text() else 10.0
        lne_1A_obus_MUSYAN = int(self.ui.lne_1A_obus_MUSYAN.text()) if self.ui.lne_1A_obus_MUSYAN.text() else 2600
        return lne_1A_obus_sag,lne_1A_obus_yukari,lne_1A_obus_rakim,lne_1A_obus_AHIA,lne_1A_obus_IHF,lne_1A_obus_MUSYAN
    
    def mevzi_1B(self):
        """BİRİNCİ OBÜSÜN 1B MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_1B_obus_sag = int(self.ui.lne_1B_obus_sag.text()) if self.ui.lne_1B_obus_sag.text() else 19034
        lne_1B_obus_yukari = int(self.ui.lne_1B_obus_yukari.text()) if self.ui.lne_1B_obus_yukari.text() else 65603
        lne_1B_obus_rakim = int(self.ui.lne_1B_obus_rakim.text()) if self.ui.lne_1B_obus_rakim.text() else 364
        lne_1B_obus_ahia = int(self.ui.lne_1B_obus_ahia.text()) if self.ui.lne_1B_obus_ahia.text() else 3000
        lne_1B_obus_ihf = float(self.ui.lne_1B_obus_ihf.text()) if self.ui.lne_1B_obus_ihf.text() else 10.0
        lne_1B_obus_musyan = int(self.ui.lne_1B_obus_musyan.text()) if self.ui.lne_1B_obus_musyan.text() else 2600
        return lne_1B_obus_sag,lne_1B_obus_yukari,lne_1B_obus_rakim,lne_1B_obus_ahia,lne_1B_obus_ihf,lne_1B_obus_musyan
    
    def mevzi_1C(self):

        """BİRİNCİ OBÜSÜN 1C MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_1C_obus_sag = int(self.ui.lne_1C_obus_sag.text()) if self.ui.lne_1C_obus_sag.text() else 19034
        lne_1C_obus_yukari = int(self.ui.lne_1C_obus_yukari.text()) if self.ui.lne_1C_obus_yukari.text() else 65603
        lne_1C_obus_rakim = int(self.ui.lne_1C_obus_rakim.text()) if self.ui.lne_1C_obus_rakim.text() else 364
        lne_1C_obus_ahia = int(self.ui.lne_1C_obus_ahia.text()) if self.ui.lne_1C_obus_ahia.text() else 3450
        lne_1C_obus_ihf = float(self.ui.lne_1C_obus_ihf.text()) if self.ui.lne_1C_obus_ihf.text() else 10.0
        lne_1C_obus_musyan = int(self.ui.lne_1C_obus_musyan.text()) if self.ui.lne_1C_obus_musyan.text() else 2600
        return lne_1C_obus_sag, lne_1C_obus_yukari,lne_1C_obus_rakim,lne_1C_obus_ahia,lne_1C_obus_ihf,lne_1C_obus_musyan
    
    def mevzi_1D(self):
        """BİRİNCİ OBÜSÜN 1D MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_1D_obus_sag = int(self.ui.lne_1D_obus_sag.text()) if self.ui.lne_1D_obus_sag.text() else 19034
        lne_1D_obus_yukari = int(self.ui.lne_1D_obus_yukari.text()) if self.ui.lne_1D_obus_yukari.text() else 65603
        lne_1D_obus_rakim = int(self.ui.lne_1D_obus_rakim.text()) if self.ui.lne_1D_obus_rakim.text() else 364
        lne_1D_obus_ahia = int(self.ui.lne_1D_obus_ahia.text()) if self.ui.lne_1D_obus_ahia.text() else 3700
        lne_1D_obus_ihf = float(self.ui.lne_1D_obus_ihf.text()) if self.ui.lne_1D_obus_ihf.text() else 10.0
        lne_1D_obus_musyan = int(self.ui.lne_1D_obus_musyan.text()) if self.ui.lne_1D_obus_musyan.text() else 2600
        return lne_1D_obus_sag,lne_1D_obus_yukari,lne_1D_obus_rakim,lne_1D_obus_ahia,lne_1D_obus_ihf,lne_1D_obus_musyan

    def mevzi_2A(self):
        """İKİNCİ OBÜSÜN 2A MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_2A_obus_sag = int(self.ui.lne_2A_obus_sag.text()) if self.ui.lne_2A_obus_sag.text() else 19054
        lne_2A_obus_yukari = int(self.ui.lne_2A_obus_yukari.text()) if self.ui.lne_2A_obus_yukari.text() else 65616
        lne_2A_obus_rakim = int(self.ui.lne_2A_obus_rakim.text()) if self.ui.lne_2A_obus_rakim.text() else 364
        lne_2A_obus_ahia = int(self.ui.lne_2A_obus_ahia.text()) if self.ui.lne_2A_obus_ahia.text() else 1000
        lne_2A_obus_ihf = float(self.ui.lne_2A_obus_ihf.text()) if self.ui.lne_2A_obus_ihf.text() else 9.0
        lne_2A_obus_musyan = int(self.ui.lne_2A_obus_musyan.text()) if self.ui.lne_2A_obus_musyan.text() else 2600
        return lne_2A_obus_sag,lne_2A_obus_yukari,lne_2A_obus_rakim,lne_2A_obus_ahia,lne_2A_obus_ihf,lne_2A_obus_musyan
        
    def mevzi_2B(self):
        """İKİNCİ OBÜSÜN 2B MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_2B_obus_sag = int(self.ui.lne_2B_obus_sag.text()) if self.ui.lne_2B_obus_sag.text() else 19054
        lne_2B_obus_yukari = int(self.ui.lne_2B_obus_yukari.text()) if self.ui.lne_2B_obus_yukari.text() else 65616
        lne_2B_obus_rakim = int(self.ui.lne_2B_obus_rakim.text()) if self.ui.lne_2B_obus_rakim.text() else 364
        lne_2B_obus_ahia = int(self.ui.lne_2B_obus_ahia.text()) if self.ui.lne_2B_obus_ahia.text() else 1800
        lne_2B_obus_ihf = float(self.ui.lne_2B_obus_ihf.text()) if self.ui.lne_2B_obus_ihf.text() else 9.0
        lne_2B_obus_musyan = int(self.ui.lne_2B_obus_musyan.text()) if self.ui.lne_2B_obus_musyan.text() else 2600
        return lne_2B_obus_sag,lne_2B_obus_yukari,lne_2B_obus_rakim,lne_2B_obus_ahia,lne_2B_obus_ihf,lne_2B_obus_musyan

    def mevzi_2C(self):
        """İKİNCİ OBÜSÜN 2C MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_2C_obus_sag = int(self.ui.lne_2C_obus_sag.text()) if self.ui.lne_2C_obus_sag.text() else 19054
        lne_2C_obus_yukari = int(self.ui.lne_2C_obus_yukari.text()) if self.ui.lne_2C_obus_yukari.text() else 65616
        lne_2C_obus_rakim = int(self.ui.lne_2C_obus_rakim.text()) if self.ui.lne_2C_obus_rakim.text() else 364
        lne_2C_obus_ahia = int(self.ui.lne_2C_obus_ahia.text()) if self.ui.lne_2C_obus_ahia.text() else 1800
        lne_2C_obus_ihf = float(self.ui.lne_2C_obus_ihf.text()) if self.ui.lne_2C_obus_ihf.text() else 9.0
        lne_2C_obus_musyan = int(self.ui.lne_2C_obus_musyan.text()) if self.ui.lne_2C_obus_musyan.text() else 2600
        return lne_2C_obus_sag,lne_2C_obus_yukari,lne_2C_obus_rakim,lne_2C_obus_ahia,lne_2C_obus_ihf,lne_2C_obus_musyan
        
    def mevzi_2D(self):
        """İKİNCİ OBÜSÜN 2C MEVZİ BİLGİLERİNİ ALIYORUZ
                """
        lne_2D_obus_sag = int(self.ui.lne_2D_obus_sag.text()) if self.ui.lne_2D_obus_sag.text() else 19054
        lne_2D_obus_yukari = int(self.ui.lne_2D_obus_yukari.text()) if self.ui.lne_2D_obus_yukari.text() else 65616
        lne_2D_obus_rakim = int(self.ui.lne_2D_obus_rakim.text()) if self.ui.lne_2D_obus_rakim.text() else 364
        lne_2D_obus_ahia = int(self.ui.lne_2D_obus_ahia.text()) if self.ui.lne_2D_obus_ahia.text() else 200
        lne_2D_obus_ihf = float(self.ui.lne_2D_obus_ihf.text()) if self.ui.lne_2D_obus_ihf.text() else 9.0
        lne_2D_obus_musyan = int(self.ui.lne_2D_obus_musyan.text()) if self.ui.lne_2D_obus_musyan.text() else 2600
        return lne_2D_obus_sag,lne_2D_obus_yukari,lne_2D_obus_rakim,lne_2D_obus_ahia,lne_2D_obus_ihf,lne_2D_obus_musyan

    @guvenli_slot
    def barut_hakki_oner(self):
        """ÖNCELİKLE DİĞER DEF LERDEN RETURN ETTİĞİMİZ MEVZİ BİLGİLERİNİ ÇEKİYORUZ"""
        lne_1A_obus_sag,lne_1A_obus_yukari,lne_1A_obus_rakim,lne_1A_obus_AHIA,lne_1A_obus_IHF,lne_1A_obus_MUSYAN= self.mevzi_1A()
        lne_1B_obus_sag,lne_1B_obus_yukari,lne_1B_obus_rakim,lne_1B_obus_ahia,lne_1B_obus_ihf,lne_1B_obus_musyan = self.mevzi_1B()
        lne_1C_obus_sag, lne_1C_obus_yukari,lne_1C_obus_rakim,lne_1C_obus_ahia,lne_1C_obus_ihf,lne_1C_obus_musyan = self.mevzi_1C()
        lne_1D_obus_sag,lne_1D_obus_yukari,lne_1D_obus_rakim,lne_1D_obus_ahia,lne_1D_obus_ihf,lne_1D_obus_musyan = self.mevzi_1D()
        lne_2A_obus_sag,lne_2A_obus_yukari,lne_2A_obus_rakim,lne_2A_obus_ahia,lne_2A_obus_ihf,lne_2A_obus_musyan = self.mevzi_2A()
        lne_2B_obus_sag,lne_2B_obus_yukari,lne_2B_obus_rakim,lne_2B_obus_ahia,lne_2B_obus_ihf,lne_2B_obus_musyan = self.mevzi_2B()
        lne_2C_obus_sag,lne_2C_obus_yukari,lne_2C_obus_rakim,lne_2C_obus_ahia,lne_2C_obus_ihf,lne_2C_obus_musyan = self.mevzi_2C()
        lne_2D_obus_sag,lne_2D_obus_yukari,lne_2D_obus_rakim,lne_2D_obus_ahia,lne_2D_obus_ihf,lne_2D_obus_musyan = self.mevzi_2D()
 #################################################################################################################       
        """MÜHİMMAT BİLGİLERİNİ ALIYORUZ."""    
        mermi, tapa, mermi_kare_agirligi = self._muhimmat_oku()
        paralanma_yuksekliği = 0  
###################################################################################################################
        """HEDEFİN KOORDİNAT BİLGİLERİNİ ALIYORUZ
                """
        if not self.ui.lne_hedef_sag_deger.text().strip() or not self.ui.lne_hedef_yukari_deger.text().strip():
            QMessageBox.warning(self, "Hedef", "Hedef sağ ve yukarı değerlerini girin.")
            return None
        lne_hedef_sag_deger = int(self.ui.lne_hedef_sag_deger.text())
        lne_hedef_yukari_deger = int(self.ui.lne_hedef_yukari_deger.text())
        lne_hedef_bolge_sayisi= self.ui. lne_hedef_bolge_sayisi.text() if self.ui. lne_hedef_bolge_sayisi.text() else "SFA"
        lne_hedef_bolge_nu = str(self.ui.lne_hedef_bolge_nu.text()) if self.ui.lne_hedef_bolge_nu.text() else "37"
        bolgevekoordinat =str(lne_hedef_bolge_nu+lne_hedef_bolge_sayisi+str(lne_hedef_sag_deger)+str(lne_hedef_yukari_deger))
        rakim = int(self.ui.lne_hedef_rakim.text()) if self.ui.lne_hedef_rakim.text() else None
        print("Hedefin Rakimi:",rakim)
    ########################################################################################################################  
        
# HEDEFİN RAKIMINI HESAPLIYORUZ.
        if rakim is None:        
            try:
                m = mgrs.MGRS()
                c = bolgevekoordinat.encode()
                d = m.toLatLon(c)
                d = str(d)
                cog_element = d.replace("(","").replace(")","").replace(" ","").split(",")
                enlem = cog_element[0]
                boylam = cog_element[1]
                
                def get_elevation(lat, lon):
                    base_url = "https://api.open-elevation.com/api/v1/lookup"
                    params = {
                        "locations": f"{lat},{lon}",
                    }

                    try:
                        response = requests.get(base_url, params=params)
                        data = response.json()
                        elevation = data["results"][0]["elevation"]
                        elevation = int(elevation)
                        print(f"Yükseklik: {elevation} metre")
                        return elevation
                    except Exception as e:
                        print(f"Hata: {e}")
                        return None

                # Örnek koordinatlar (Ankara, Türkiye)
                latitude = enlem
                longitude = boylam
                rakim = get_elevation(latitude, longitude)

            except Exception as ex:
                print(f"Hata: {ex}")
                rakim = "-------"

        if rakim is not None:
            self.ui.lne_hedef_rakim.setText(str(rakim))

        if rakim is None or rakim == "-------":
            QMessageBox.warning(self, "Rakım", "Hedef rakımı alınamadı. Rakımı elle girin.")
            return None
        try:
            lne_hedef_rakim = int(self.ui.lne_hedef_rakim.text())
        except (TypeError, ValueError):
            QMessageBox.warning(self, "Rakım", "Hedef rakımı sayı olmalıdır.")
            return None


###############################################################################################################
        """BATARYA İLE HEDEF MESAFESİNİ BULUYORUZ.
               BİRİNCİ OBÜS ( 1A-1B-1C-1D) """
        plan_mesafesi_1A_obüs = round(math.sqrt((lne_1A_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_1A_obus_yukari - lne_hedef_yukari_deger) ** 2))
        plan_mesafesi_1B_obüs = round(math.sqrt((lne_1B_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_1B_obus_yukari - lne_hedef_yukari_deger) ** 2))
        plan_mesafesi_1C_obüs = round(math.sqrt((lne_1C_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_1C_obus_yukari - lne_hedef_yukari_deger) ** 2))
        plan_mesafesi_1D_obüs = round(math.sqrt((lne_1D_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_1D_obus_yukari - lne_hedef_yukari_deger) ** 2))

        """BATARYA İLE HEDEF MESAFESİNİ BULUYORUZ.
               İKİNCİ OBÜS ( 2A-2B-2C-2D) """
        plan_mesafesi_2A_obüs = round(math.sqrt((lne_2A_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_2A_obus_yukari - lne_hedef_yukari_deger) ** 2))
        plan_mesafesi_2B_obüs = round(math.sqrt((lne_2B_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_2B_obus_yukari - lne_hedef_yukari_deger) ** 2))
        plan_mesafesi_2C_obüs = round(math.sqrt((lne_2C_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_2C_obus_yukari - lne_hedef_yukari_deger) ** 2))
        plan_mesafesi_2D_obüs = round(math.sqrt((lne_2D_obus_sag - lne_hedef_sag_deger) ** 2 + (lne_2D_obus_yukari - lne_hedef_yukari_deger) ** 2))
    ##########################################################################################################    
                      
        """BİRİNCİ OBÜSÜN BARUT HAKKINI BULUYORUZ."""
        if plan_mesafesi_1A_obüs == 0:   
            sonuc_barut_hakki_1 = 0
        elif plan_mesafesi_1A_obüs <= 2348:
            sonuc_barut_hakki_1 = 1
        elif plan_mesafesi_1A_obüs <= 2766:
            sonuc_barut_hakki_1 = 2
        elif plan_mesafesi_1A_obüs <= 3354: 
            sonuc_barut_hakki_1 = 3
        elif plan_mesafesi_1A_obüs <= 4218:
            sonuc_barut_hakki_1 = 4
        elif plan_mesafesi_1A_obüs <= 5363:
            sonuc_barut_hakki_1 = 5
        elif plan_mesafesi_1A_obüs <= 6694:
            sonuc_barut_hakki_1 = 6
        elif plan_mesafesi_1A_obüs <= 11000:
            sonuc_barut_hakki_1 = 7
        else:
            sonuc_barut_hakki_1 = "Atış Yapılamaz..."
            QMessageBox.warning(self, "Uyarı", "Atış yapılamaz!")
            return
        
        """İKİNCİ OBÜSÜN BARUT HAKKINI BULUYORUZ."""
        if plan_mesafesi_2A_obüs == 0:
            sonuc_barut_hakki_2A = 0
        elif plan_mesafesi_2A_obüs <= 2348:
            sonuc_barut_hakki_2A = 1
        elif plan_mesafesi_2A_obüs <= 2766:
            sonuc_barut_hakki_2A = 2
        elif plan_mesafesi_2A_obüs <= 3354: 
            sonuc_barut_hakki_2A = 3
        elif plan_mesafesi_2A_obüs <= 4218:
            sonuc_barut_hakki_2A = 4
        elif plan_mesafesi_2A_obüs <= 5363:
            sonuc_barut_hakki_2A = 5
        elif plan_mesafesi_2A_obüs <= 6694:
            sonuc_barut_hakki_2A = 6
        elif plan_mesafesi_2A_obüs <= 11000:
            sonuc_barut_hakki_2A = 7
        else:
            sonuc_barut_hakki_2A = "Atış Yapılamaz..."
            QMessageBox.warning(self, "Uyarı", "Atış yapılamaz!")
            return
            
############################################################################################################
        """İSTİKAMET AÇISINI BULURUZ
                BİRİNCİ OBÜS İÇİN"""
        dx_1 = abs(lne_1B_obus_sag - lne_hedef_sag_deger)
        dy_1 = abs(lne_1B_obus_yukari - lne_hedef_yukari_deger)
        batarya_hedef_İA_1 = round(math.atan2(dx_1, dy_1) * 3200 / math.pi)

        if lne_1B_obus_sag > lne_hedef_sag_deger and lne_1B_obus_yukari > lne_hedef_yukari_deger:
            print("Üçüncü Bölge")
            batarya_hedef_İA_1 += 3200
        elif lne_1B_obus_sag > lne_hedef_sag_deger and lne_1B_obus_yukari < lne_hedef_yukari_deger:
            print("Dördüncü Bölge")
            batarya_hedef_İA_1 = 6400 - batarya_hedef_İA_1
        elif lne_1B_obus_sag < lne_hedef_sag_deger and lne_1B_obus_yukari < lne_hedef_yukari_deger:
            print("Birinci Bölge")
            batarya_hedef_İA_1 = batarya_hedef_İA_1
        elif lne_1B_obus_sag < lne_hedef_sag_deger and lne_1B_obus_yukari > lne_hedef_yukari_deger:
            print("İkinci Bölge")
            batarya_hedef_İA_1 = 3200 - batarya_hedef_İA_1

###########################################################################################################
        """YANLARI HESAPLIYORUZ.YALNIZCA 1B"""
        yan1A_fark = abs(batarya_hedef_İA_1 -lne_1A_obus_AHIA)  
        yan1B_fark = abs(batarya_hedef_İA_1 - lne_1B_obus_ahia)  
        yan1C_fark = abs(batarya_hedef_İA_1 - lne_1C_obus_ahia)  
        yan1D_fark = abs(batarya_hedef_İA_1 - lne_1D_obus_ahia)  
        """YAN İA FARKINDAN YAN HESAPLAMA BÖLGEYE GÖRE"""
        """1A YAN"""
        if lne_1A_obus_AHIA < batarya_hedef_İA_1:
            yan_1A = lne_1A_obus_MUSYAN - yan1A_fark
            if yan_1A <= 0:
                yan_1A = 6400 + yan_1A
        else:
            yan_1A = lne_1A_obus_MUSYAN + yan1A_fark
            if yan_1A <= 0:
                yan_1A = 6400 + yan_1A
        """1B YAN"""
        if lne_1B_obus_ahia < batarya_hedef_İA_1:
            yan_1B = lne_1B_obus_musyan - yan1B_fark
            if yan_1B <= 0:
                yan_1B = 6400 + yan_1B
        else:
            yan_1B = lne_1B_obus_musyan + yan1B_fark
            if yan_1B <= 0:
                yan_1B = 6400 + yan_1B
        """1C YAN"""
        if lne_1C_obus_ahia < batarya_hedef_İA_1:
            yan_1C = lne_1C_obus_musyan - yan1C_fark
            if yan_1C <= 0:
                yan_1C = 6400 + yan_1C
        else:
            yan_1C = lne_1C_obus_musyan + yan1C_fark
            if yan_1C <= 0:
                yan_1C = 6400 + yan_1C
        """1D YAN"""
        if lne_1D_obus_ahia < batarya_hedef_İA_1:
            yan_1D = lne_1D_obus_musyan - yan1D_fark
            if yan_1D <= 0:
                yan_1D = 6400 + yan_1D
        else:
            yan_1D = lne_1D_obus_musyan + yan1D_fark
            if yan_1D <= 0:
                yan_1D = 6400 + yan_1D

############################################################################################################
        """İSTİKAMET AÇISINI BULURUZ
                İKİNCİ OBÜS İÇİN"""
        dx_2 = abs(lne_2B_obus_sag - lne_hedef_sag_deger)
        dy_2 = abs(lne_2B_obus_yukari - lne_hedef_yukari_deger)
        batarya_hedef_İA_2 = round(math.atan2(dx_2, dy_2) * 3200 / math.pi)

        if lne_2B_obus_sag > lne_hedef_sag_deger and lne_2B_obus_yukari > lne_hedef_yukari_deger:
            print("Üçüncü Bölge")
            batarya_hedef_İA_2 += 3200
        elif lne_2B_obus_sag > lne_hedef_sag_deger and lne_2B_obus_yukari < lne_hedef_yukari_deger:
            print("Dördüncü Bölge")
            batarya_hedef_İA_2 = 6400 - batarya_hedef_İA_2
        elif lne_2B_obus_sag < lne_hedef_sag_deger and lne_2B_obus_yukari < lne_hedef_yukari_deger:
            print("Birinci Bölge")
            batarya_hedef_İA_2 = batarya_hedef_İA_2
        elif lne_2B_obus_sag < lne_hedef_sag_deger and lne_2B_obus_yukari > lne_hedef_yukari_deger:
            print("İkinci Bölge")
            batarya_hedef_İA_2 = 3200 - batarya_hedef_İA_2

###########################################################################################################
        """YANLARI HESAPLIYORUZ.YALNIZCA 1B"""
        yan2A_fark = abs(batarya_hedef_İA_2 -lne_2A_obus_ahia)  
        yan2B_fark = abs(batarya_hedef_İA_2 - lne_2B_obus_ahia)  
        yan2C_fark = abs(batarya_hedef_İA_2 - lne_2C_obus_ahia)  
        yan2D_fark = abs(batarya_hedef_İA_2 - lne_2D_obus_ahia)  
        """YAN İA FARKINDAN YAN HESAPLAMA BÖLGEYE GÖRE"""
        """2A YAN"""
        if lne_2A_obus_ahia < batarya_hedef_İA_2:
            yan_2A = lne_2A_obus_musyan - yan2A_fark
            if yan_2A <= 0:
                yan_2A = 6400 + yan_2A
        else:
            yan_2A = lne_2A_obus_musyan + yan2A_fark
            if yan_2A <= 0:
                yan_2A = 6400 + yan_2A
        """2B YAN"""
        if lne_2B_obus_ahia < batarya_hedef_İA_2:
            yan_2B = lne_2B_obus_musyan - yan2B_fark
            if yan_2B <= 0:
                yan_2B = 6400 + yan_2B
        else:
            yan_2B = lne_2B_obus_musyan + yan2B_fark
            if yan_2B <= 0:
                yan_2B = 6400 + yan_2B
        """2C YAN"""
        if lne_2C_obus_ahia < batarya_hedef_İA_2:
            yan_2C = lne_2C_obus_musyan - yan2C_fark
            if yan_2C <= 0:
                yan_2C = 6400 + yan_2C
        else:
            yan_2C = lne_2C_obus_musyan + yan2C_fark
            if yan_2C <= 0:
                yan_2C = 6400 + yan_2C
        """2D YAN"""
        if lne_2D_obus_ahia < batarya_hedef_İA_2:
            yan_2D = lne_2D_obus_musyan - yan2D_fark
            if yan_2D <= 0:
                yan_2D = 6400 + yan_2D
        else:
            yan_2D = lne_2D_obus_musyan + yan2D_fark
            if yan_2D <= 0:
                yan_2D = 6400 + yan_2D
        """İSTİKAMET AÇILARINI EKRANA YAZDIRIYORUZ
                BİRİNCİ OBÜS İÇİN"""
        self.ui.sonuc_istikamet_acisi_1A.setText(str(batarya_hedef_İA_1))
        self.ui.sonuc_istikamet_acisi_1B.setText(str(batarya_hedef_İA_1))
        self.ui.sonuc_istikamet_acisi_1C.setText(str(batarya_hedef_İA_1))
        self.ui.sonuc_istikamet_acisi_1D.setText(str(batarya_hedef_İA_1))
        self.ui.sonuc_istikamet_acisi_2A.setText(str(batarya_hedef_İA_2))
        self.ui.sonuc_istikamet_acisi_2B.setText(str(batarya_hedef_İA_2))
        self.ui.sonuc_istikamet_acisi_2C.setText(str(batarya_hedef_İA_2))
        self.ui.sonuc_istikamet_acisi_2D.setText(str(batarya_hedef_İA_2))

        """BARUT HAKLARINI EKRANA YAZDIRIYORUZ"""
        """BİRİNCİ OBÜS"""
        self.ui.lne_barut_hakki.setText(str(sonuc_barut_hakki_1))
        self.ui.sonuc_batur_hakki_1A.setText(str(sonuc_barut_hakki_1))
        self.ui.sonuc_barut_hakki_1B.setText(str(sonuc_barut_hakki_1))
        self.ui.sonuc_barut_hakki_1C.setText(str(sonuc_barut_hakki_1))
        self.ui.sonuc_barut_hakki_1D.setText(str(sonuc_barut_hakki_1))
        """İKİNCİ OBÜS"""
        self.ui.sonuc_barut_hakki_2A.setText(str(sonuc_barut_hakki_2A))
        self.ui.sonuc_barut_hakki_2B.setText(str(sonuc_barut_hakki_2A))
        self.ui.sonuc_barut_hakki_2C.setText(str(sonuc_barut_hakki_2A))
        self.ui.sonuc_barut_hakki_2D.setText(str(sonuc_barut_hakki_2A))
        """MÜHİMMAT BİLGİLERİNİ EKRANA YAZDIRIR"""
        """BİRİNCİ OBÜS"""
        self.ui.sonuc_tapa_1A.setText(str(tapa))
        self.ui.sonuc_tapa_1B.setText(str(tapa))
        self.ui.sonuc_tapa_1C.setText(str(tapa))
        self.ui.sonuc_tapa_1D.setText(str(tapa))
        """İKİNCİ OBÜS"""
        self.ui.sonuc_tapa_2A.setText(str(tapa))
        self.ui.sonuc_tapa_2B.setText(str(tapa))
        self.ui.sonuc_tapa_2C.setText(str(tapa))
        self.ui.sonuc_tapa_2D.setText(str(tapa))
        """RAKIMI YERİNE YAZDIRIRIZ"""
        self.ui.lne_hedef_rakim.setText(str(rakim))
        """BİRİNCİ OBÜS MESAFELERİ YAZDIRIR"""
        self.ui.sonuc_mesafe_1A.setText(str(plan_mesafesi_1A_obüs))
        self.ui.sonuc_mesafe_1B.setText(str(plan_mesafesi_1B_obüs))
        self.ui.sonuc_mesafe_1C.setText(str(plan_mesafesi_1C_obüs))
        self.ui.sonuc_mesafe_1D.setText(str(plan_mesafesi_1D_obüs))
        """İKİNCİ OBÜS MESAFELERİ YAZDIRIR"""
        self.ui.sonuc_mesafe_2A.setText(str(plan_mesafesi_2A_obüs))
        self.ui.sonuc_mesafe_2B.setText(str(plan_mesafesi_2B_obüs))
        self.ui.sonuc_mesafe_2C.setText(str(plan_mesafesi_2C_obüs))
        self.ui.sonuc_mesafe_2D.setText(str(plan_mesafesi_2D_obüs))
        self.ui.sonuc_yan_1A.setText(str(yan_1A))
        self.ui.sonuc_yan_1B.setText(str(yan_1B))
        self.ui.sonuc_yan_1C.setText(str(yan_1C))
        self.ui.sonuc_yan_1D.setText(str(yan_1D))
        self.ui.sonuc_yan_2A.setText(str(yan_2A))
        self.ui.sonuc_yan_2B.setText(str(yan_2B))
        self.ui.sonuc_yan_2C.setText(str(yan_2C))
        self.ui.sonuc_yan_2D.setText(str(yan_2D))
        self.ui.lne_hedef_sag_deger_sonuc.setText(str(lne_hedef_sag_deger))
        self.ui.lne_hedef_yukari_deger_sonuc.setText(str(lne_hedef_yukari_deger))
        self.ui.lne_hedef_rakim_sonuc.setText(str(lne_hedef_rakim))


        return paralanma_yuksekliği,mermi_kare_agirligi,lne_hedef_sag_deger,lne_hedef_yukari_deger,lne_hedef_bolge_sayisi,lne_hedef_bolge_nu,bolgevekoordinat,lne_hedef_rakim,plan_mesafesi_1A_obüs,plan_mesafesi_1B_obüs,plan_mesafesi_1C_obüs,plan_mesafesi_1D_obüs,plan_mesafesi_2A_obüs,plan_mesafesi_2B_obüs,plan_mesafesi_2C_obüs,plan_mesafesi_2D_obüs,sonuc_barut_hakki_1,sonuc_barut_hakki_2A,batarya_hedef_İA_1,yan_1A,yan_1B,yan_1C,yan_1D,batarya_hedef_İA_2,yan_2A,yan_2B,yan_2C,yan_2D,mermi,tapa,rakim

    def keyPressEvent(self, event):
        try:
            if event.key() in (QtCore.Qt.Key_Return, QtCore.Qt.Key_Enter):
                if isinstance(self.focusWidget(), QtWidgets.QLineEdit):
                    self.barut_hakki_oner()
                    self.ilk_hesaplama()
                    return
            super(myApp, self).keyPressEvent(event)
        except Exception as exc:
            hata_goster(exc, traceback.format_exc())
###########################################################################################################
    @guvenli_slot
    def secilen_barut_hakki(self):
        try:
            secilen_barut_hakki = int(self.ui.lne_barut_hakki.text())
        except ValueError:
            QMessageBox.warning(self, "Barut hakkı", "Barut hakkı sayı olmalıdır.")
            return None
        print("Secilen Barut Hakkı2", secilen_barut_hakki)
        self.ui.sonuc_barut_hakki_1B.setText(str(secilen_barut_hakki))
        self.ui.sonuc_barut_hakki_2B.setText(str(secilen_barut_hakki))

        return secilen_barut_hakki
    
    @guvenli_slot
    def ilk_hesaplama(self):
        if not self.ui.lne_barut_hakki.text().strip():
            QMessageBox.warning(self, "Barut hakkı", "Önce barut hakkı önerin veya girin.")
            return None
        try:
            secilen = self.secilen_barut_hakki()
            if secilen is None:
                return None
            secilen_barut_hakki = int(secilen)
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Barut hakkı", "Barut hakkı sayı olmalıdır.")
            return None
        oner = self.barut_hakki_oner()
        if oner is None:
            return None
        paralanma_yuksekliği,mermi_kare_agirligi,lne_hedef_sag_deger,lne_hedef_yukari_deger,lne_hedef_bolge_sayisi,lne_hedef_bolge_nu,bolgevekoordinat,lne_hedef_rakim,plan_mesafesi_1A_obüs,plan_mesafesi_1B_obüs,plan_mesafesi_1C_obüs,plan_mesafesi_1D_obüs,plan_mesafesi_2A_obüs,plan_mesafesi_2B_obüs,plan_mesafesi_2C_obüs,plan_mesafesi_2D_obüs,sonuc_barut_hakki_1,sonuc_barut_hakki_2A,batarya_hedef_İA_1,yan_1A,yan_1B,yan_1C,yan_1D,batarya_hedef_İA_2,yan_2A,yan_2B,yan_2C,yan_2D,mermi,tapa,rakim = oner
        """ÖNCELİKLE DİĞER DEF LERDEN RETURN ETTİĞİMİZ MEVZİ BİLGİLERİNİ ÇEKİYORUZ"""
        lne_1A_obus_sag,lne_1A_obus_yukari,lne_1A_obus_rakim,lne_1A_obus_AHIA,lne_1A_obus_IHF,lne_1A_obus_MUSYAN= self.mevzi_1A()
        lne_1B_obus_sag,lne_1B_obus_yukari,lne_1B_obus_rakim,lne_1B_obus_ahia,lne_1B_obus_ihf,lne_1B_obus_musyan = self.mevzi_1B()
        lne_1C_obus_sag, lne_1C_obus_yukari,lne_1C_obus_rakim,lne_1C_obus_ahia,lne_1C_obus_ihf,lne_1C_obus_musyan = self.mevzi_1C()
        lne_1D_obus_sag,lne_1D_obus_yukari,lne_1D_obus_rakim,lne_1D_obus_ahia,lne_1D_obus_ihf,lne_1D_obus_musyan = self.mevzi_1D()
        lne_2A_obus_sag,lne_2A_obus_yukari,lne_2A_obus_rakim,lne_2A_obus_ahia,lne_2A_obus_ihf,lne_2A_obus_musyan = self.mevzi_2A()
        lne_2B_obus_sag,lne_2B_obus_yukari,lne_2B_obus_rakim,lne_2B_obus_ahia,lne_2B_obus_ihf,lne_2B_obus_musyan = self.mevzi_2B()
        lne_2C_obus_sag,lne_2C_obus_yukari,lne_2C_obus_rakim,lne_2C_obus_ahia,lne_2C_obus_ihf,lne_2C_obus_musyan = self.mevzi_2C()
        lne_2D_obus_sag,lne_2D_obus_yukari,lne_2D_obus_rakim,lne_2D_obus_ahia,lne_2D_obus_ihf,lne_2D_obus_musyan = self.mevzi_2D()

        """BİRİNCİ OBÜS MEVZİLERİ İÇİN AYRI AYRI YAN DEĞERLERİNİ BULALIM"""
        """ÖCELİKLE DOĞAL YANLARI MESAFEYE GÖRE HER OBÜS İÇİN AYRI AYRI BULUYORUZ."""
        a3bh = [(0 ,0.0),(100 ,0.0),(200 ,0.1),(300 ,0.3),(400 ,0.4),(500 ,0.6),(600 ,0.7),(700 ,0.9),(800 ,1.1),
                (900 ,1.2),(1000 ,1.4),(1100 ,1.5),(1200 ,1.7),(1300 ,1.9),(1400 ,2.1),(1500 ,2.2),(1600 ,2.4),(1700 ,2.6),
                (1800 ,2.8),(1900 ,3.0),(2000 ,3.2),(2100 ,3.4),(2200 ,3.6),(2300 ,3.8),(2400 ,4.0),(2500 ,4.2),(2600 ,4.4),
                (2700 ,4.7),(2800 ,4.9),(2900 ,5.2),(3000 ,5.4),(3100 ,5.7),(3200 ,6.0),(3300 ,6.3),(3400 ,6.6),(3500 ,7.0),(3600 ,7.3),
                (3700 ,7.7),(3800 ,8.1),(3900 ,8.6),(4000 ,9.0),(4100 ,9.6),(4200 ,10.2),(4300 ,10.9),(4400 ,11.7),(4500 ,12.7),]
        data3bh = []
        for i in range(len(a3bh)-1):
            start = a3bh[i]
            end = a3bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data3bh.append((int(x), y))

        data3bh.append(a3bh[-1])

        liste4bh = [
    (0 ,0.0),
    (100 ,0.0),
    (200 ,0.1),
    (300 ,0.2),
    (400 ,0.3),
    (500 ,0.4),
    (600 ,0.5),
    (700 ,0.7),
    (800 ,0.8),
    (900 ,0.9),
    (1000 ,1.0),
    (1100 ,1.2),
    (1200 ,1.3),
    (1300 ,1.4),
    (1400 ,1.5),
    (1500 ,1.7),
    (1600 ,1.8),
    (1700 ,2.0),
    (1800 ,2.1),
    (1900 ,2.2),
    (2000 ,2.4),
    (2100 ,2.5),
    (2200 ,2.7),
    (2300 ,2.8),
    (2400 ,3.0),
    (2500 ,3.1),
    (2600 ,3.3),
    (2700 ,3.5),
    (2800 ,3.6),
    (2900 ,3.8),
    (3000 ,4.0),
    (3100 ,4.2),
    (3200 ,4.3),
    (3300 ,4.5),
    (3400 ,4.7),
    (3500 ,4.9),
    (3600 ,5.1),
    (3700 ,5.3),
    (3800 ,5.6),
    (3900 ,5.8),
    (4000 ,6.0),
    (4100 ,6.3),
    (4200 ,6.6),
    (4300 ,6.8),
    (4400 ,7.1),
    (4500 ,7.4),
    (4600 ,7.8),
    (4700 ,8.1),
    (4800 ,8.5),
    (4900 ,8.9),
    (5000 ,9.3),
    (5100 ,9.8),
    (5200 ,10.3),
    (5300 ,10.9),
    (5400 ,11.7),
    (5500 ,12.5),
    (5600 ,13.7),
    (5700 ,15.7)
    ]

        data4bh = []
        for i in range(len(liste4bh)-1):
            start = liste4bh[i]
            end = liste4bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data4bh.append((int(x), y))

        data4bh.append(liste4bh[-1])

        liste5bh = [
    (0 , 0.0),
    (100 , 0.0),
    (200 , 0.0),
    (300 , 0.1),
    (400 , 0.2),
    (500 , 0.3),
    (600 , 0.4),
    (700 , 0.4),
    (800 , 0.5),
    (900 , 0.6),
    (1000 , 0.7),
    (1100 , 0.8),
    (1200 , 0.9),
    (1300 , 1.0),
    (1400 , 1.1),
    (1500 , 1.2),
    (1600 , 1.3),
    (1700 , 1.3),
    (1800 , 1.4),
    (1900 , 1.5),
    (2000 , 1.6),
    (2100 , 1.7),
    (2200 , 1.8),
    (2300 , 1.9),
    (2400 , 2.1),
    (2500 , 2.2),
    (2600 , 2.3),
    (2700 , 2.4),
    (2800 , 2.5),
    (2900 , 2.6),
    (3000 , 2.7),
    (3100 , 2.8),
    (3200 , 3.0),
    (3300 , 3.1),
    (3400 , 3.2),
    (3500 , 3.3),
    (3600 , 3.5),
    (3700 , 3.6),
    (3800 , 3.7),
    (3900 , 3.9),
    (4000 , 4.0),
    (4100 , 4.1),
    (4200 , 4.3),
    (4300 , 4.4),
    (4400 , 4.6),
    (4500 , 4.8),
    (4600 , 4.9),
    (4700 , 5.1),
    (4800 , 5.3),
    (4900 , 5.4),
    (5000 , 5.6),
    (5100 , 5.8),
    (5200 , 6.0),
    (5300 , 6.2),
    (5400 , 6.4),
    (5500 , 6.6),
    (5600 , 6.9),
    (5700 , 7.1),
    (5800 , 7.3),
    (5900 , 7.6),
    (6000 , 7.9),
    (6100 , 8.2),
    (6200 , 8.5),
    (6300 , 8.8),
    (6400 , 9.2),
    (6500 , 9.6),
    (6600 , 10.0),
    (6700 , 10.5),
    (6800 , 11.0),
    (6900 , 11.6),
    (7000 , 12.4),
    (7100 , 13.3),
    (7200 , 17.7)]

        data5bh = []
        for i in range(len(liste5bh)-1):
            start = liste5bh[i]
            end = liste5bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data5bh.append((int(x), y))

        data5bh.append(liste5bh[-1])

        liste6bh = [
            (0 , 0.0),
        (100 , 0.0),
        (200 , 0.0),
        (300 , 0.1),
        (400 , 0.2),
        (500 , 0.2),
        (600 , 0.3),
        (700 , 0.3),
        (800 , 0.4),
        (900 , 0.5),
        (1000 , 0.5),
        (1100 , 0.6),
        (1200 , 0.7),
        (1300 , 0.7),
        (1400 , 0.8),
        (1500 , 0.9),
        (1600 , 0.9),
        (1700 , 1.0),
        (1800 , 1.1),
        (1900 , 1.2),
        (2000 , 1.2),
        (2100 , 1.3),
        (2200 , 1.4),
        (2300 , 1.5),
        (2400 , 1.5),
        (2500 , 1.6),
        (2600 , 1.7),
        (2700 , 1.8),
        (2800 , 1.9),
        (2900 , 2.0),
        (3000 , 2.0),
        (3100 , 2.1),
        (3200 , 2.2),
        (3300 , 2.3),
        (3400 , 2.4),
        (3500 , 2.5),
        (3600 , 2.6),
        (3700 , 2.7),
        (3800 , 2.8),
        (3900 , 2.9),
        (4000 , 3.0),
        (4100 , 3.1),
        (4200 , 3.2),
        (4300 , 3.3),
        (4400 , 3.4),
        (4500 , 3.5),
        (4600 , 3.6),
        (4700 , 3.8),
        (4800 , 3.9),
        (4900 , 4.0),
        (5000 , 4.1),
        (5100 , 4.2),
        (5200 , 4.4),
        (5300 , 4.5),
        (5400 , 4.6),
        (5500 , 4.7),
        (5600 , 4.9),
        (5700 , 5.0),
        (5800 , 5.2),
        (5900 , 5.3),
        (6000 , 5.4),
        (6100 , 5.6),
        (6200 , 5.8),
        (6300 , 5.9),
        (6400 , 6.1),
        (6500 , 6.2),
        (6600 , 6.4),
        (6700 , 6.6),
        (6800 , 6.8),
        (6900 , 7.0),
        (7000 , 7.2),
        (7100 , 7.4),
        (7200 , 7.6),
        (7300 , 7.8),
        (7400 , 8.0),
        (7500 , 8.3),
        (7600 , 8.5),
        (7700 , 8.8),
        (7800 , 9.1),
        (7900 , 9.4),
        (8000 , 9.7),
        (8100 , 10.0),
        (8200 , 10.4),
        (8300 , 10.7),
        (8400 , 11.2),
        (8500 , 11.6),
        (8600 , 12.2),
        (8700 , 12.8),
        (8800 , 13.5),
        (8900 , 14.4),
        (9000 , 15.7)]

        data6bh = []
        for i in range(len(liste6bh)-1):
            start = liste6bh[i]
            end = liste6bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data6bh.append((int(x), y))

        data6bh.append(liste6bh[-1])

        liste7bh = [
        (0 , 0.0),
    (100 , 0.0),
    (200 , 0.0),
    (300 , 0.1),
    (400 , 0.1),
    (500 , 0.2),
    (600 , 0.2),
    (700 , 0.3),
    (800 , 0.3),
    (900 , 0.4),
    (1000 , 0.4),
    (1100 , 0.5),
    (1200 , 0.5),
    (1300 , 0.6),
    (1400 , 0.6),
    (1500 , 0.7),
    (1600 , 0.7),
    (1700 , 0.8),
    (1800 , 0.9),
    (1900 , 0.9),
    (2000 , 1.0),
    (2100 , 1.0),
    (2200 , 1.1),
    (2300 , 1.2),
    (2400 , 1.2),
    (2500 , 1.3),
    (2600 , 1.4),
    (2700 , 1.4),
    (2800 , 1.5),
    (2900 , 1.6),
    (3000 , 1.7),
    (3100 , 1.7),
    (3200 , 1.8),
    (3300 , 1.9),
    (3400 , 2.0),
    (3500 , 2.0),
    (3600 , 2.1),
    (3700 , 2.2),
    (3800 , 2.3),
    (3900 , 2.4),
    (4000 , 2.4),
    (4100 , 2.5),
    (4200 , 2.6),
    (4300 , 2.7),
    (4400 , 2.8),
    (4500 , 2.9),
    (4600 , 3.0),
    (4700 , 3.1),
    (4800 , 3.2),
    (4900 , 3.3),
    (5000 , 3.4),
    (5100 , 3.5),
    (5200 , 3.6),
    (5300 , 3.7),
    (5400 , 3.8),
    (5500 , 3.9),
    (5600 , 4.0),
    (5700 , 4.1),
    (5800 , 4.2),
    (5900 , 4.4),
    (6000 , 4.5),
    (6100 , 4.6),
    (6200 , 4.7),
    (6300 , 4.8),
    (6400 , 4.9),
    (6500 , 5.1),
    (6600 , 5.2),
    (6700 , 5.3),
    (6800 , 5.5),
    (6900 , 5.6),
    (7000 , 5.7),
    (7100 , 5.9),
    (7200 , 6.0),
    (7300 , 6.2),
    (7400 , 6.3),
    (7500 , 6.5),
    (7600 , 6.6),
    (7700 , 6.8),
    (7800 , 6.9),
    (7900 , 7.1),
    (8000 , 7.3),
    (8100 , 7.4),
    (8200 , 7.6),
    (8300 , 7.8),
    (8400 , 8.0),
    (8500 , 8.2),
    (8600 , 8.4),
    (8700 , 8.6),
    (8800 , 8.8),
    (8900 , 9.0),
    (9000 , 9.2),
    (9100 , 9.5),
    (9200 , 9.7),
    (9300 , 10.0),
    (9400 , 10.2),
    (9500 , 10.5),
    (9600 , 10.8),
    (9700 , 11.1),
    (9800 , 11.4),
    (9900 , 11.7),
    (10000 , 12.1),
    (10100 , 12.5),
    (10200 , 12.9),
    (10300 , 13.3),
    (10400 , 13.8),
    (10500 , 14.4),
    (10600 , 15.0),
    (10700 , 15.7),
    (10800 , 16.6),
    (10900 , 17.8),
    (11000 , 20.3)]


        data7bh = []
        for i in range(len(liste7bh)-1):
            start = liste7bh[i]
            end = liste7bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data7bh.append((int(x), y))

        data7bh.append(liste7bh[-1])


        
        dogal_yan_1A = 0
        secilen_barut_hakki = int(secilen_barut_hakki)
        print("seçilen barut hakkı",secilen_barut_hakki)
        print("tipi",type(secilen_barut_hakki))
        print("plan mesafesi:",plan_mesafesi_1B_obüs)


        """BİRİNCİ OBÜS İÇİN DOĞAL YAN BULUYORUZ"""
        print("sevilen bh:", secilen_barut_hakki)
        if secilen_barut_hakki == 3:
            for item in data3bh:
                if item[0] == plan_mesafesi_1B_obüs:
                    dogal_yan_1A = item[1]
                    dogal_yan_1A = round(dogal_yan_1A,2)
            print("Üçüncü BH Doğal Yan:",dogal_yan_1A)
        if secilen_barut_hakki == 4:
            for item in data4bh:
                if item[0] == plan_mesafesi_1B_obüs:
                    dogal_yan_1A = item[1]
                    dogal_yan_1A = round(dogal_yan_1A,2)
            print("Dördüncü BH Doğal Yan:",dogal_yan_1A)
        if secilen_barut_hakki == 5:
            for item in data5bh:
                if item[0] == plan_mesafesi_1B_obüs:
                    dogal_yan_1A = item[1]
                    dogal_yan_1A = round(dogal_yan_1A,2)
            print("Beşinci BH Doğal Yan:",dogal_yan_1A)
                
        if secilen_barut_hakki == 6:
            for item in data6bh:
                if item[0] == plan_mesafesi_1B_obüs:
                    dogal_yan_1A = item[1]
                    dogal_yan_1A = round(dogal_yan_1A,2)
            print("Altıncı BH Doğal Yan:",dogal_yan_1A)
        if secilen_barut_hakki == 7:
            for item in data7bh:
                if item[0] == plan_mesafesi_1B_obüs:
                    dogal_yan_1A = item[1]
                    dogal_yan_1A = round(dogal_yan_1A,2)
            print("Yedinci BH Doğal Yan:",dogal_yan_1A)
        dogal_yan_1A = float(dogal_yan_1A) if float(dogal_yan_1A) else 1
        print("doğal yan",dogal_yan_1A)

 
 ##############################################################################################
        """İKİNCİ OBÜSÜN DOĞAL YANINI BULUYORUZ"""
        dogal_yan_2A = 0
        
        """BİRİNCİ OBÜS İÇİN DOĞAL YAN BULUYORUZ"""
        print("sevilen bh:", secilen_barut_hakki)
        if secilen_barut_hakki == 3:
            for item in data3bh:
                if item[0] == plan_mesafesi_2B_obüs:
                    dogal_yan_2A = item[1]
                    dogal_yan_2A = round(dogal_yan_2A,2)
            print("Üçüncü BH Doğal Yan:",dogal_yan_2A)
        if secilen_barut_hakki == 4:
            for item in data4bh:
                if item[0] == plan_mesafesi_2B_obüs:
                    dogal_yan_2A = item[1]
                    dogal_yan_2A = round(dogal_yan_2A,2)
            print("Dördüncü BH Doğal Yan:",dogal_yan_2A)
        if secilen_barut_hakki == 5:
            for item in data5bh:
                if item[0] == plan_mesafesi_2B_obüs:
                    dogal_yan_2A = item[1]
                    dogal_yan_2A = round(dogal_yan_2A,2)
            print("Beşinci BH Doğal Yan:",dogal_yan_2A)
                
        if secilen_barut_hakki == 6:
            for item in data6bh:
                if item[0] == plan_mesafesi_2B_obüs:
                    dogal_yan_2A = item[1]
                    dogal_yan_2A = round(dogal_yan_2A,2)
            print("Altıncı BH Doğal Yan:",dogal_yan_2A)
        if secilen_barut_hakki == 7:
            for item in data7bh:
                if item[0] == plan_mesafesi_2B_obüs:
                    dogal_yan_2A = item[1]
                    dogal_yan_2A = round(dogal_yan_2A,2)
            print("Yedinci BH Doğal Yan:",dogal_yan_2A)
        dogal_yan_2A = float(dogal_yan_2A) if float(dogal_yan_2A) else 1
        print("doğal yan2",dogal_yan_2A)
        #yan1 = round(dogal_yan_1A + yan_1B)
        yan2= round(dogal_yan_2A + yan_2B)

        tac = 0
        toplam_mesafe_duzeltmesi = 0
        atis_istikameti_1 = batarya_hedef_İA_1
        if secilen_barut_hakki not in (5, 6, 7):
            QMessageBox.information(
                self, "Barut hakkı",
                "Tam atış esası 5, 6 ve 7. barut hakkı için hesaplanır.\nPlan yanları ekrana yazıldı.")
            return None

        """YÜKSELİS HESAPLAMA - TAÇ HESAPLAMA NİŞANGAH + TAÇ"""  
        """BİRİNCİ OBÜS"""
        if secilen_barut_hakki == 5:
            plan_mesafesi = plan_mesafesi_1B_obüs
            batarya_rakimi = lne_1B_obus_rakim
            bt_rakimi_1 = lne_1B_obus_rakim
            hedef_rakimi = lne_hedef_rakim
            ahia = lne_1B_obus_ahia
            paralanma_yuksekliği = paralanma_yuksekliği
            mermi_kare_agirligi = mermi_kare_agirligi
            ilk_hiz_farki = lne_1B_obus_ihf
            atis_istikameti = batarya_hedef_İA_1
            atis_istikameti_1 = batarya_hedef_İA_1
            mevzi_hiz_degisikligi = 0
            mevzi_yan_duzeltmesi = 0
            mevzi_ts_duzeltmesi = 0
            tapasaniyesi = 1
            barut_isisi_str = self.ui.lne_barut_isisi.text()
            barut_isisi = float(barut_isisi_str) if barut_isisi_str else 75
            # barut_isisi = self.ui.lne_barut_isisi.text() if self.ui.lne_barut_isisi.text() else 75
            print("Barutttt İSİSİ",barut_isisi)
                # 1. GİRİŞ MESAFESİNİ BUL / MANUEL HASSASİYETİ

            # mesafe = 1866
            # en_yakin_10_metre = round(mesafe / 10) * 10
            # print("1836'nın en yakın 10 metreye çevrilmiş hali:", en_yakin_10_metre)

            # BATARYA RAKIMINI EN YAKIN 10 M YE ÇEVİR.
            batarya_rakimi = round(batarya_rakimi/10)*10
            # PARALANMA NOKTASININ RAKIMINI BUL.
            paralanma_noktasi_rakimi = hedef_rakimi + paralanma_yuksekliği
            # HEDEF-BATARYA YÜKSEKLİK FARKINI BUL
            yukseklik_farki = paralanma_noktasi_rakimi-batarya_rakimi
            yukseklik_farki_1 = paralanma_noktasi_rakimi-bt_rakimi_1
            # yukseklik_farki = round(yukseklik_farki/100)*100
            ################# B cetvelini yüklüyoruz.################
            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (100, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (200, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (300, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (400, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (500, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (600, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (700, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (800, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (900, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1000, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1100, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1200, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1300, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1400, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1500, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1600, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1700, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1800, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1900, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2000, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2100, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2200, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2300, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2400, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2500, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2600, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2700, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2800, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2900, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (3000, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (3100, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (3200, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3300, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3400, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3500, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3600, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3700, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3800, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3),
                (3900, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4000, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4100, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4200, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4300, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4400, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4500, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4600, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4700, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4800, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3),
                (4900, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3),
                (5000, 1, 1, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3),
                (5100, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3),
                (5200, 1, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4),
                (5300, 1, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4),
                (5400, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4),
                (5500, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4),
                (5600, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 4, 4),
                (5700, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4),
                (5800, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4),
                (5900, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4),
                (6000, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4),
                (6100, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4),
                (6200, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 5),
                (6300, 2, 2, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5),
                (6400, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5),
                (6500, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 5, 5, 5),
                (6600, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5),
                (6700, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5),
                (6800, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5),
                (6900, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5),
                (7000, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5),
                (7100, 3, 3, 3, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, 5),
                (7200, 3, 3, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5),
            ]

            data2 = [
                (100, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0),
                (200, 0, 0, 0, 0, 0, 1, 2, 3, 5, 7, 8, 10, 12, 14, 0),
                (300, 0, 0, 0, 0, 0, 1, 3, 5, 8, 10, 13, 16, 19, 22, 25),
                (400, 0, 0, 0, 0, 0, 2, 4, 7, 10, 14, 17, 21, 25, 28, 32),
                (500, 0, 0, 0, 0, 0, 3, 6, 9, 13, 17, 21, 26, 30, 35, 40),
                (600, 0, 0, 0, 0, 0, 3, 7, 11, 15, 20, 25, 30, 36, 41, 47),
                (700, 0, 0, 0, 0, 0, 4, 8, 13, 18, 23, 29, 35, 41, 47, 54),
                (800, 0, 0, 0, 0, 0, 4, 9, 14, 20, 26, 32, 39, 46, 53, 61),
                (900, 0, 0, 0, 0, 0, 5, 10, 16, 22, 29, 36, 43, 51, 59, 68),
                (1000, 0, 0, 0, 0, 0, 5, 11, 18, 24, 32, 39, 47, 56, 65, 74),
                (1100, 0, 0, 0, 0, 0, 6, 12, 19, 27, 35, 43, 52, 61, 71, 81),
                (1200, 0, 0, 0, 0, 0, 7, 14, 21, 29, 38, 46, 56, 66, 76, 87),
                (1300, 0, 0, 0, 0, 0, 7, 15, 23, 31, 40, 50, 60, 71, 82, 94),
                (1400, 0, 0, 0, -7, 0, 8, 16, 24, 34, 43, 54, 64, 76, 88, 100),
                (1500, 0, 0, 0, -8, 0, 8, 17, 26, 36, 46, 57, 69, 81, 93, 107),
                (1600, 0, 0, 0, -8, 0, 9, 18, 28, 38, 49, 61, 73, 86, 99, 113),
                (1700, 0, 0, 0, -9, 0, 9, 19, 30, 41, 52, 64, 77, 91, 105, 120),
                (1800, 0, 0, 0, -9, 0, 10, 20, 31, 43, 55, 68, 82, 96, 111, 126),
                (1900, 0, 0, -19, -10, 0, 10, 22, 33, 45, 58, 72, 86, 101, 116, 133),
                (2000, 0, 0, -20, -10, 0, 11, 23, 35, 48, 61, 75, 90, 106, 122, 139),
                (2100, 0, 0, -22, -11, 0, 12, 24, 37, 50, 64, 79, 95, 111, 128, 146),
                (2200, 0, 0, -23, -12, 0, 12, 25, 38, 53, 67, 83, 99, 116, 134, 153),
                (2300, 0, -35, -24, -12, 0, 13, 26, 40, 55, 71, 87, 104, 122, 140, 160),
                (2400, 0, -36, -25, -13, 0, 13, 27, 42, 58, 74, 91, 108, 127, 146, 167),
                (2500, 0, -38, -26, -13, 0, 14, 29, 44, 60, 77, 95, 113, 132, 152, 174),
                (2600, 0, -40, -27, -14, 0, 15, 30, 46, 63, 80, 98, 118, 138, 159, 181),
                (2700, -54, -42, -28, -15, 0, 15, 31, 48, 65, 83, 102, 122, 143, 165, 188),
                (2800, -57, -43, -30, -15, 0, 16, 32, 50, 68, 87, 106, 127, 149, 171, 195),
                (2900, -59, -45, -31, -16, 0, 16, 34, 52, 70, 90, 111, 132, 154, 178, 202),
                (3000, -61, -47, -32, -16, 0, 17, 35, 54, 73, 93, 115, 137, 160, 184, 210),
                (3100, -64, -49, -33, -17, 0, 18, 36, 56, 76, 97, 119, 142, 166, 191, 218),
                (3200, -69, -53, -36, -18, 0, 19, 39, 60, 79, 100, 123, 147, 172, 198, 225),
                (3300, -66, -51, -34, -18, 0, 18, 38, 58, 81, 104, 127, 152, 178, 205, 233),
                (3400, -71, -54, -37, -19, 0, 20, 40, 62, 84, 107, 132, 157, 184, 212, 241),
                (3500, -74, -56, -38, -20, 0, 20, 42, 64, 87, 111, 136, 163, 190, 219, 249),
                (3600, -76, -58, -40, -20, 0, 21, 43, 66, 90, 115, 141, 168, 197, 226, 258),
                (3700, -79, -60, -41, -21, 0, 22, 44, 68, 93, 119, 145, 174, 203, 234, 266),
                (3800, -81, -62, -42, -22, 0, 22, 46, 70, 96, 122, 150, 179, 210, 241, 275),
                (3900, -84, -64, -44, -22, 0, 23, 47, 73, 99, 126, 155, 185, 216, 249, 284),
                (4000, -87, -66, -45, -23, 0, 24, 49, 75, 102, 130, 160, 191, 223, 257, 293),
                (4100, -89, -68, -46, -24, 0, 25, 50, 77, 105, 134, 165, 197, 230, 265, 302),
                (4200, -92, -70, -48, -24, 0, 25, 52, 80, 108, 139, 170, 203, 238, 274, 312),
                (4300, -95, -73, -49, -25, 0, 26, 54, 82, 112, 143, 175, 209, 245, 282, 322),
                (4400, -98, -75, -51, -26, 0, 27, 55, 85, 115, 147, 181, 216, 253, 291, 332),
                (4500, -101, -77, -52, -27, 0, 28, 57, 87, 119, 152, 186, 222, 260, 300, 342),
                (4600, -104, -79, -54, -27, 0, 29, 59, 90, 122, 156, 192, 229, 269, 310, 353),
                (4700, -107, -82, -55, -28, 0, 29, 60, 92, 126, 161, 198, 236, 277, 320, 365),
                (4800, -110, -84, -57, -29, 0, 30, 62, 95, 130, 166, 204, 244, 286, 330, 376),
                (4900, -113, -86, -59, -30, 0, 31, 64, 98, 134, 171, 210, 251, 295, 340, 389),
                (5000, -116, -89, -60, -31, 0, 32, 66, 101, 138, 176, 217, 259, 304, 351, 401),
                (5100, -120, -91, -62, -32, 0, 33, 68, 104, 142, 182, 223, 267, 314, 363, 415),
                (5200, -123, -94, -64, -33, 0, 34, 70, 107, 146, 187, 230, 276, 324, 375, 429),
                (5300, -126, -97, -66, -34, 0, 35, 72, 110, 151, 193, 238, 285, 335, 387, 444),
                (5400, -130, -100, -68, -35, 0, 36, 74,
                114, 155, 199, 245, 294, 346, 401, 460),
                (5500, -134, -102, -70, -36, 0, 37, 76,
                117, 160, 205, 253, 304, 358, 415, 477),
                (5600, -138, -105, -72, -37, 0, 38, 79,
                121, 165, 212, 262, 314, 370, 430, 495),
                (5700, -142, -108, -74, -38, 0, 40, 81,
                125, 171, 219, 271, 325, 384, 447, 515),
                (5800, -146, -112, -76, -39, 0, 41, 84,
                129, 176, 227, 280, 337, 399, 465, 537),
                (5900, -150, -115, -78, -40, 0, 42, 86,
                133, 182, 235, 290, 350, 414, 485, 562),
                (6000, -154, -118, -81, -41, 0, 43, 89,
                137, 189, 243, 301, 364, 432, 507, 590),
                (6100, -159, -122, -83, -43, 0, 45, 92,
                142, 195, 252, 313, 379, 452, 532, 625),
                (6200, -164, -126, -86, -44, 0, 46, 95,
                147, 203, 262, 327, 397, 475, 563, 669),
                (6300, -169, -130, -89, -45, 0, 48, 99,
                153, 211, 273, 342, 417, 502, 603, 739),
                (6400, -174, -134, -91, -47, 0, 50, 103, 159, 220, 286, 359, 441, 538, 666, 0),
                (6500, -180, -138, -95, -49, 0, 52, 107, 166, 230, 301, 380, 472, 594, 0, 0),
                (6600, -185, -143, -98, -50, 0, 54, 111, 174, 242, 318, 407, 523, 0, 0, 0),
                (6700, -192, -148, -101, -52, 0, 56, 116, 183, 256, 341, 452, 0, 0, 0, 0),
                (6800, -199, -153, -105, -54, 0, 59, 123, 193, 275, 381, 0, 0, 0, 0, 0),
                (6900, -206, -159, -110, -57, 0, 62, 130, 208, 310, 0, 0, 0, 0, 0, 0),
                (7000, -214, -166, -115, -60, 0, 66, 140, 238, 0, 0, 0, 0, 0, 0, 0),
                (7100, -223, -173, -120, -63, 0, 71, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (7200, -233, -182, -127, -67, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),

            ]


            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            interpolasyonlu_veri2 = veri_kumesi_interpolasyonu(data2)
            # print(interpolasyonlu_veri2)

            hedefin_toptan_yuksekligi = yukseklik_farki
            hedefin_toptan_yuksekligi = round(hedefin_toptan_yuksekligi/100)*100
            print("Hedefin Toptan Yüksekliği: ", hedefin_toptan_yuksekligi)
            mesafe = plan_mesafesi
            print("mesafe = ", mesafe)

            if hedefin_toptan_yuksekligi == -400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[1])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[2])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[3])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[4])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -0:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[5])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[6])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[7])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[8])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[9])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 500:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[10])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 600:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[11])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 700:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 800:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 900:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[13])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 1000:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[14])
                        print("Seçilen Bölge:", bolge)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")


            # TXT dosyasının adı ve dosya yolu
            dosya_adı = METRO_DOSYA

            # İlk satırı atlayarak dosyanın geri kalanını yeni bir liste olarak alın
            metrorpr = []
            if not os.path.isfile(dosya_adı):
                QMessageBox.warning(self, "Metro", "METRAP.txt bulunamadı.")
                return None
            with open(dosya_adı, "r") as dosya:
                for satır in dosya:
                    metrorpr.append(satır.strip())

            # Satırları yan yana birleştirerek elde etmek için join() yöntemini kullanıyoruz
            birlesik_veri = "\n".join(metrorpr[1:])

            rapor_tipi = birlesik_veri[:5]
            oktant = birlesik_veri[5:6]
            metro_koor = birlesik_veri[6:9] + " " + birlesik_veri[9:12]
            tarih = birlesik_veri[13:15]
            saat = birlesik_veri[15:17]
            dakika = str(int(birlesik_veri[17:18]) * 6)
            if dakika == "0":
                dakika = str(int(birlesik_veri[17:18]) * 6) + "0"
            ara = ":"
            saat_dakika = saat + ara + str(dakika)

            gecerlilik = birlesik_veri[18:19]
            metro_istasyonu_rakimi = str(int(birlesik_veri[19:22]) * 10)
            yogunluk = int(birlesik_veri[22:25]) / 10

            print("Rapor Tipi:", rapor_tipi)
            print("Oktant:", oktant)
            print("Metro Koordinat:", metro_koor)
            print("Tarih:", tarih)
            print("Saat:", saat_dakika)
            print("Geçerlilik:", gecerlilik)
            print("Metro İstasyonu Rakımı:", metro_istasyonu_rakimi)
            print("Yoğunluk:", yogunluk)


            if bolge == 0:
                ruzgar_istikameti = int(birlesik_veri[28:30]) * 100
                ruzgar_hizi = int(birlesik_veri[30:32])
                if birlesik_veri[32] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[32:35]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[32:35]) / 10
                if birlesik_veri[35] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[35:38]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[35:38]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)  # 39-40-41 3 say

            elif bolge == 1:
                ruzgar_istikameti = int(birlesik_veri[41:43]) * 100
                ruzgar_hizi = int(birlesik_veri[43:45])
                if birlesik_veri[45] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[45:48]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[45:48]) / 10
                if birlesik_veri[48] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[48:51]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[48:51]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)

            elif bolge == 2:
                ruzgar_istikameti = int(birlesik_veri[54:56])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[56:58])  # 3
                if birlesik_veri[58] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[58:61])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[58:61])/10  # 4
                if birlesik_veri[61] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[61:64])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[61:64])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 3:
                ruzgar_istikameti = int(birlesik_veri[67:69])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[69:71])  # 3
                if birlesik_veri[71] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[71:74])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[71:74])/10  # 4
                if birlesik_veri[74] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[74:77])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[74:77])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 4:
                ruzgar_istikameti = int(birlesik_veri[80:82])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[82:84])  # 3
                if birlesik_veri[84] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[84:87])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[84:87])/10  # 4
                if birlesik_veri[87] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[87:90])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[87:90])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 5:
                ruzgar_istikameti = int(birlesik_veri[93:95])*100  # 3
                ruzgar_hizi = int(birlesik_veri[95:97])  # 3
                if birlesik_veri[97] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[97:100])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[97:100])/10  # 4
                if birlesik_veri[100] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[100:103])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[100:103])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 6:
                ruzgar_istikameti = int(birlesik_veri[106:108])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[108:110])  # 3
                if birlesik_veri[110] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[110:113])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[110:113])/10  # 4
                if birlesik_veri[113] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[113:116])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[113:116])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 7:
                ruzgar_istikameti = int(birlesik_veri[119:121])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[121:123])  # 3
                if birlesik_veri[123] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[123:126])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[123:126])/10  # 4
                if birlesik_veri[126] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[126:129])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[126:129])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 8:
                ruzgar_istikameti = int(birlesik_veri[132:134])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[134:136])  # 3
                if birlesik_veri[136] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[136:139])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[136:139])/10  # 4
                if birlesik_veri[139] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[139:142])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[139:142])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 9:
                ruzgar_istikameti = int(birlesik_veri[145:147])*100  # 3
                ruzgar_hizi = int(birlesik_veri[147:149])  # 3
                if birlesik_veri[149] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[149:152])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[149:152])/10  # 4
                if birlesik_veri[152] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[152:155])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[152:155])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 10:
                ruzgar_istikameti = int(birlesik_veri[158:160])*100  # 3
                ruzgar_hizi = int(birlesik_veri[160:162])  # 3
                if birlesik_veri[162] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[162:165])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[162:165])/10  # 4
                if birlesik_veri[165] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[165:168])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[165:168])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)


            # TAMAMLAYICI MESAFEYİ BULUYORUZ.
            if hedefin_toptan_yuksekligi == -400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[1])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[2])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[3])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[4])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -0:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[5])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[6])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[7])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[8])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[9])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 500:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[10])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 600:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[11])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 700:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 800:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 900:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[13])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 1000:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[14])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")

            bt_metro_ist_rakim_farki = batarya_rakimi - int(metro_istasyonu_rakimi)

            giris_mesafesi = round((int(mesafe) + int(tamamlayici_mesafe))/100)*100
            print("Giriş Mesafesi :", giris_mesafesi)

            ruzgarin_plan_istikameti_100 = ruzgar_istikameti - atis_istikameti

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            isi_dzl = [
                (-390, 0.9),
                (-380, 0.9),
                (-370, 0.9),
                (-360, 0.8),
                (-350, 0.8),
                (-340, 0.8),
                (-330, 0.8),
                (-320, 0.7),
                (-310, 0.7),
                (-300, 0.7),
                (-290, 0.7),
                (-280, 0.7),
                (-270, 0.7),
                (-260, 0.6),
                (-250, 0.6),
                (-240, 0.6),
                (-230, 0.6),
                (-220, 0.5),
                (-210, 0.5),
                (-200, 0.5),
                (-190, 0.4),
                (-180, 0.4),
                (-170, 0.4),
                (-160, 0.3),
                (-150, 0.3),
                (-140, 0.3),
                (-130, 0.3),
                (-120, 0.2),
                (-110, 0.2),
                (-100, 0.2),
                (-90, 0.2),
                (-80, 0.2),
                (-70, 0.2),
                (-60, 0.1),
                (-50, 0.1),
                (-40, 0.1),
                (-30, 0.1),
                (-20, 0),
                (-10, 0),
                (0, 0),
                (10, 0),
                (20, 0),
                (30, -0.1),
                (40, -0.1),
                (50, -0.1),
                (60, -0.1),
                (70, -0.2),
                (80, -0.2),
                (90, -0.2),
                (100, -0.2),
                (110, -0.2),
                (120, -0.2),
                (130, -0.3),
                (140, -0.3),
                (150, -0.3),
                (160, -0.3),
                (170, -0.4),
                (180, -0.4),
                (190, -0.4),
                (200, -0.5),
                (210, -0.5),
                (220, -0.5),
                (230, -0.6),
                (240, -0.6),
                (250, -0.6),
                (260, -0.6),
                (270, -0.7),
                (280, -0.7),
                (290, -0.7),
                (300, -0.7),
                (310, -0.7),
                (320, -0.7),
                (330, -0.8),
                (340, -0.8),
                (350, -0.8),
                (360, -0.8),
                (370, -0.9),
                (380, -0.9),
                (390, -0.9),

            ]


            isi_data5bh = []
            for i in range(len(isi_dzl)-1):
                start = isi_dzl[i]
                end = isi_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    isi_data5bh.append((int(x), y))

            isi_data5bh.append(isi_dzl[-1])


            yogunluk_dzl = [
                (-390, 3.9),
                (-380, 3.8),
                (-370, 3.7),
                (-360, 3.6),
                (-350, 3.5),
                (-340, 3.4),
                (-330, 3.3),
                (-320, 3.2),
                (-310, 3.1),
                (-300, 3),
                (-290, 2.9),
                (-280, 2.8),
                (-270, 2.7),
                (-260, 2.6),
                (-250, 2.5),
                (-240, 2.4),
                (-230, 2.3),
                (-220, 2.2),
                (-210, 2.1),
                (-200, 2),
                (-190, 1.0),
                (-180, 1.8),
                (-170, 1.7),
                (-160, 1.6),
                (-150, 1.5),
                (-140, 1.4),
                (-130, 1.3),
                (-120, 1.2),
                (-110, 1.1),
                (-100, 1),
                (-90, 0.9),
                (-80, 0.8),
                (-70, 0.7),
                (-60, 0.6),
                (-50, 0.5),
                (-40, 0.4),
                (-30, 0.3),
                (-20, 0.2),
                (-10, 0.1),
                (0, 0),
                (10, -0.1),
                (20, -0.2),
                (30, -0.3),
                (40, -0.4),
                (50, -0.5),
                (60, -0.6),
                (70, -0.7),
                (80, -0.8),
                (90, -0.9),
                (100, -1),
                (110, -1.1),
                (120, -1.2),
                (130, -1.3),
                (140, -1.4),
                (150, -1.5),
                (160, -1.6),
                (170, -1.7),
                (180, -1.8),
                (190, -1.9),
                (200, -2),
                (210, -2.1),
                (220, -2.2),
                (230, -2.3),
                (240, -2.4),
                (250, -2.5),
                (260, -2.6),
                (270, -2.7),
                (280, -2.8),
                (290, -2.9),
                (300, -3),
                (310, -3.1),
                (320, -3.2),
                (330, -3.3),
                (340, -3.4),
                (350, -3.5),
                (360, -3.6),
                (370, -3.7),
                (380, -3.8),
                (390, -3.9),
            ]

            yogunluk_data5bh = []

            for i in range(len(yogunluk_dzl)-1):
                start = yogunluk_dzl[i]
                end = yogunluk_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    yogunluk_data5bh.append((int(x), y))

            yogunluk_data5bh.append(yogunluk_dzl[-1])
            # print(yogunluk_data5bh)


            for i in isi_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    isi_duzeltmesi = i[1]
                    print("Isı Düzeltmesi :", isi_duzeltmesi)


            for i in yogunluk_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    yogunluk_duzeltmesi = i[1]
                    print("Yoğunluk Düzeltmesi", yogunluk_duzeltmesi)

            duzeltilmis_sıcaklik_degeri = isi_duzeltmesi + hava_sicakligi
            print("Düzeltilmiş Isı Değeri:", duzeltilmis_sıcaklik_degeri)

            duzeltilmis_yogunluk_degeri = yogunluk_duzeltmesi + hava_yogunlugu
            print("Düzeltilmiş Yoğunluk Değeri:", duzeltilmis_yogunluk_degeri)

            # RUZGAR BİLEŞENLERİ

            atis_istikameti = round(atis_istikameti/100)*100

            if ruzgar_istikameti < atis_istikameti:
                ruzgar_istikameti = ruzgar_istikameti + 6400
            else:
                ruzgar_istikameti = ruzgar_istikameti
            print("ruzgar istikameti", ruzgar_istikameti)

            print("atış istikameti", atis_istikameti)

            ruzgarin_plan_istikameti = ruzgar_istikameti - atis_istikameti

            print("Rüzgarın Plan İstikameti:", ruzgarin_plan_istikameti)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            data1 = [
                (0, 0, 1),
                (100, -0.1, 0.99),
                (200, -0.2, 0.98),
                (300, -0.29, 0.96),
                (400, -0.38, 0.92),
                (500, -0.47, 0.88),
                (600, -0.56, 0.83),
                (700, -0.63, 0.77),
                (800, -0.71, 0.71),
                (900, -0.77, 0.63),
                (1000, -0.83, 0.56),
                (1100, -0.88, 0.47),
                (1200, -0.92, 0.38),
                (1300, -0.96, 0.29),
                (1400, -0.98, 0.2),
                (1500, -0.99, 0.1),
                (1600, -1, 0),
                (1700, -0.99, -0.1),
                (1800, -0.98, -0.2),
                (1900, -0.96, -0.29),
                (2000, -0.92, -0.38),
                (2100, -0.88, -0.47),
                (2200, -0.83, -0.56),
                (2300, -0.77, -0.63),
                (2400, -0.71, -0.71),
                (2500, -0.63, -0.77),
                (2600, -0.56, -0.83),
                (2700, -0.47, -0.88),
                (2800, -0.38, -0.92),
                (2900, -0.29, -0.96),
                (3000, -0.2, -0.98),
                (3100, -0.1, -0.99),
                (3200, 0, -1),
                (3300, 0.1, -0.99),
                (3400, 0.2, -0.98),
                (3500, 0.29, -0.96),
                (3600, 0.38, -0.92),
                (3700, 0.47, -0.88),
                (3800, 0.56, -0.83),
                (3900, 0.63, -0.77),
                (4000, 0.71, -0.71),
                (4100, 0.77, -0.63),
                (4200, 0.83, -0.56),
                (4300, 0.88, -0.47),
                (4400, 0.92, -0.38),
                (4500, 0.96, -0.29),
                (4600, 0.98, -0.2),
                (4700, 0.99, -0.1),
                (4800, 1, 0),
                (4900, 0.99, 0.1),
                (5000, 0.98, 0.2),
                (5100, 0.96, 0.29),
                (5200, 0.92, 0.38),
                (5300, 0.88, 0.47),
                (5400, 0.83, 0.56),
                (5500, 0.77, 0.63),
                (5600, 0.71, 0.71),
                (5700, 0.63, 0.77),
                (5800, 0.56, 0.83),
                (5900, 0.47, 0.88),
                (6000, 0.38, 0.92),
                (6100, 0.29, 0.96),
                (6200, 0.2, 0.98),
                (6300, 0.1, 0.99),
                (6400, 0, 1),
            ]


            # veri_noktasi[0] = MESAFE
            # veri_noktasi[1] = YAN RÜZGARI
            # veri_noktasi[2] = MESAFE RÜZGARI

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Mesafe :", mesafe)

            for rzdlz in interpolasyonlu_veri1:
                if rzdlz[0] == ruzgarin_plan_istikameti:
                    yan_ruzgari_bileseni = rzdlz[1]
                    mesafe_ruzgari_bileseni = rzdlz[2]
                    print("Yan Rüzgarı Bileşeni:", yan_ruzgari_bileseni)
                    print("mesafe Rüzgarı Bileşeni:", mesafe_ruzgari_bileseni)


            # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (100, 5.6, 0, 00, 0.3, 0, 0, 0.7, -0.6, 0, 0, 0, 0, 0, 0, -1, 1),
                (200, 11.1, 0, 0, 0.7, 0, 0.01, 1.4, -1.2, 0, 0, 0, 0, 0, 0, -2, 2),
                (300, 16.6, 0, 0, 1, 0.1, 0.01, 2.1, -1.8, 0.1, 0, 0.1, 0, 0, 0, -4, 4),
                (400, 22.2, 0, 0, 1.3, 0.2, 0.02, 2.7, -2.4, 0.1, 0, 0.1, 0, -0.1, 0.1, -5, 5),
                (500, 27.8, 0, 0, 1.7, 0.3, 0.02, 3.4, -3, 0.2, 0, 0.2, 0, -0.1, 0.1, -6, 6),
                (600, 33.4, 2, 1.02, 2, 0.4, 0.02, 4.1, -3.6, 0.2, 0, 0.3, 0, -0.1, 0.1, -7, 7),
                (700, 39.1, 2.3, 0.87, 2.4, 0.4, 0.03, 4.7, -
                4.2, 0.3, -0.1, 0.4, -0.1, -0.1, 0.1, -8, 8),
                (800, 44.9, 2.6, 0.76, 2.7, 0.5, 0.03, 5.4, -
                4.8, 0.3, -0.1, 0.4, -0.1, -0.2, 0.2, -9, 9),
                (900, 50.6, 3, 0.68, 3, 0.6, 0.04, 6, -5.3,
                0.4, -0.1, 0.5, -0.1, -0.2, 0.2, -10, 10),
                (1000, 56.4, 3.3, 0.61, 3.4, 0.7, 0.04, 6.7, -
                5.9, 0.5, -0.1, 0.6, -0.1, -0.3, 0.3, -11, 11),
                (1100, 62.3, 3.7, 0.55, 3.7, 0.8, 0.04, 7.3, -
                6.4, 0.5, -0.1, 0.7, -0.1, -0.4, 0.4, -12, 13),
                (1200, 68.2, 4, 0.5, 4.1, 0.9, 0.05, 7.9, -
                7, 0.6, -0.2, 0.8, -0.1, -0.4, 0.4, -13, 14),
                (1300, 74.1, 4.4, 0.46, 4.4, 1, 0.05, 8.6, -
                7.5, 0.7, -0.2, 0.9, -0.1, -0.5, 0.5, -14, 15),
                (1400, 80.1, 4.8, 0.43, 4.8, 1.1, 0.06, 9.2, -
                8.1, 0.8, -0.2, 1, -0.1, -0.6, 0.6, -15, 16),
                (1500, 86.1, 5.1, 0.4, 5.2, 1.2, 0.06, 9.8, -
                8.6, 0.9, -0.2, 1.2, -0.2, -0.6, 0.6, -16, 17),
                (1600, 92.2, 5.5, 0.37, 5.5, 1.3, 0.06, 10.4, -
                9.1, 1, -0.3, 1.3, -0.2, -0.7, 0.7, -17, 17),
                (1700, 98.3, 5.8, 0.35, 5.9, 1.3, 0.07, 11.1, -
                9.7, 1.1, -0.3, 1.4, -0.2, -0.8, 0.8, -18, 18),
                (1800, 104.5, 6.2, 0.33, 6.2, 1.4, 0.07, 11.7, -
                10.2, 1.1, -0.3, 1.5, -0.2, -0.9, 0.9, -19, 19),
                (1900, 110.7, 6.6, 0.31, 6.6, 1.5, 0.08, 12.3, -
                10.7, 1.2, -0.3, 1.6, -0.2, -1, 1, -19, 20),
                (2000, 117, 6.9, 0.3, 7, 1.6, 0.08, 12.9, -
                11.2, 1.3, -0.4, 1.7, -0.2, -1.1, 1.1, -20, 21),
                (2100, 123.3, 7.3, 0.28, 7.3, 1.7, 0.08, 13.5, -
                11.8, 1.4, -0.4, 1.8, -0.2, -1.2, 1.3, -21, 22),
                (2200, 129.7, 7.7, 0.27, 7.7, 1.8, 0.09, 14.1, -
                12.3, 1.5, -0.4, 1.9, -0.2, -1.4, 1.4, -22, 23),
                (2300, 136.2, 8.1, 0.26, 8.1, 1.9, 0.09, 14.7, -
                12.8, 1.6, -0.5, 2, -0.2, -1.5, 1.5, -23, 24),
                (2400, 142.7, 8.4, 0.25, 8.5, 2.1, 0.1, 15.3, -
                13.3, 1.7, -0.5, 2.1, -0.2, -1.6, 1.6, -24, 25),
                (2500, 149.2, 8.8, 0.24, 8.9, 2.2, 0.1, 15.9, -
                13.8, 1.8, -0.6, 2.2, -0.2, -1.8, 1.8, -24, 26),
                (2600, 155.9, 9.2, 0.23, 9.2, 2.3, 0.11, 16.5, -
                14.3, 1.9, -0.6, 2.3, -0.2, -1.9, 1.9, -25, 26),
                (2700, 162.6, 9.6, 0.22, 9.6, 2.4, 0.11, 17.1, -
                14.9, 2, -0.6, 2.4, -0.2, -2, 2.1, -26, 27),
                (2800, 169.4, 10, 0.21, 10, 2.5, 0.11, 17.7, -
                15.4, 2.1, -0.7, 2.5, -0.2, -2.2, 2.2, -27, 28),
                (2900, 176.2, 10.4, 0.2, 10.4, 2.6, 0.12, 18.3, -
                15.9, 2.2, -0.7, 2.5, -0.2, -2.3, 2.4, -27, 29),
                (3000, 183.1, 10.8, 0.19, 10.8, 2.7, 0.12, 18.9, -
                16.4, 2.3, -0.8, 2.6, -0.2, -2.5, 2.5, -28, 29),
                (3100, 190.1, 11.2, 0.19, 11.2, 2.8, 0.13, 19.4, -
                16.9, 2.4, -0.8, 2.7, -0.2, -2.7, 2.7, -29, 30),
                (3200, 197.2, 11.6, 0.18, 11.6, 3, 0.13, 20, -
                17.4, 2.5, -0.9, 2.8, -0.2, -2.8, 2.9, -29, 31),
                (3300, 204.4, 12, 0.17, 12, 3.1, 0.14, 20.6, -
                17.9, 2.6, -0.9, 2.8, -0.1, -3, 3.1, -30, 32),
                (3400, 211.6, 12.4, 0.17, 12.4, 3.2, 0.14, 21.2, -
                18.4, 2.7, -1, 2.9, -0.1, -3.2, 3.3, -31, 32),
                (3500, 219, 12.8, 0.16, 12.8, 3.3, 0.14, 21.7, -
                18.9, 2.8, -1, 2.9, -0.1, -3.4, 3.4, -31, 33),
                (3600, 226.4, 13.2, 0.16, 13.2, 3.5, 0.15, 22.3, -
                19.4, 2.9, -1.1, 3, -0.1, -3.6, 3.6, -32, 34),
                (3700, 233.9, 13.6, 0.15, 13.7, 3.6, 0.15, 22.9, -
                19.9, 3, -1.1, 3, -0.1, -3.8, 3.8, -33, 34),
                (3800, 241.6, 14.1, 0.15, 14.1, 3.7, 0.16, 23.4, -
                20.4, 3.1, -1.2, 3.1, -0.1, -4, 4.1, -33, 35),
                (3900, 249.3, 14.5, 0.14, 14.5, 3.9, 0.16, 24, -
                20.9, 3.2, -1.2, 3.1, 0, -4.2, 4.3, -34, 36),
                (4000, 257.2, 14.9, 0.14, 15, 4, 0.17, 24.6, -
                21.4, 3.3, -1.3, 3.2, 0, -4.4, 4.5, -34, 36),
                (4100, 265.2, 15.4, 0.14, 15.4, 4.1, 0.17, 25.1, -
                21.9, 3.4, -1.4, 3.2, 0, -4.6, 4.7, -35, 37),
                (4200, 273.3, 15.8, 0.13, 15.9, 4.3, 0.18, 25.7, -
                22.3, 3.5, -1.4, 3.2, 0, -4.8, 4.9, -36, 38),
                (4300, 281.5, 16.3, 0.13, 16.3, 4.4, 0.18, 26.2, -
                22.8, 3.6, -1.5, 3.3, 0.1, -5.1, 5.2, -36, 38),
                (4400, 289.9, 16.7, 0.13, 16.8, 4.6, 0.19, 26.8, -
                23.3, 3.7, -1.6, 3.3, 0.1, -5.3, 5.4, -37, 39),
                (4500, 298.4, 17.2, 0.12, 17.2, 4.8, 0.19, 27.3, -
                23.8, 3.8, -1.6, 3.3, 0.1, -5.6, 5.7, -37, 39),
                (4600, 307, 17.7, 0.12, 17.7, 4.9, 0.2, 27.9, -
                24.3, 3.9, -1.7, 3.3, 0.1, -5.8, 5.9, -38, 40),
                (4700, 315.9, 18.2, 0.12, 18.2, 5.1, 0.2, 28.4, -
                24.8, 4, -1.8, 3.3, 0.2, -6, 6.2, -38, 40),
                (4800, 324.9, 18.6, 0.11, 18.7, 5.3, 0.21, 29, -
                25.3, 4.1, -1.8, 3.3, 0.2, -6.3, 6.5, -39, 41),
                (4900, 334, 19.1, 0.11, 19.1, 5.4, 0.21, 29.5, -
                25.7, 4.2, -1.9, 3.4, 0.2, -6.6, 6.7, -39, 41),
                (5000, 343.4, 19.6, 0.11, 19.6, 5.6, 0.22,
                30, -26.2, 4.3, -2, 3.4, 0.3, -6.8, 7, -40, 42),
                (5100, 353, 20.1, 0.11, 20.2, 5.8, 0.23, 30.6, -
                26.7, 4.4, -2.1, 3.4, 0.3, -7.1, 7.3, -40, 42),
                (5200, 362.8, 20.7, 0.1, 20.7, 6, 0.23, 31.1, -
                27.2, 4.6, -2.1, 3.3, 0.4, -7.4, 7.6, -40, 43),
                (5300, 372.8, 21.2, 0.1, 21.2, 6.2, 0.24, 31.6, -
                27.7, 4.7, -2.2, 3.3, 0.4, -7.7, 7.9, -41, 43),
                (5400, 383.1, 21.7, 0.1, 21.7, 6.4, 0.24, 32.2, -
                28.1, 4.8, -2.3, 3.3, 0.4, -8, 8.2, -41, 44),
                (5500, 393.7, 22.3, 0.1, 22.3, 6.6, 0.25, 32.7, -
                28.6, 4.9, -2.4, 3.3, 0.5, -8.3, 8.5, -42, 44),
                (5600, 404.6, 22.8, 0.09, 22.8, 6.9, 0.26, 33.2, -
                29.1, 5, -2.5, 3.3, 0.5, -8.6, 8.8, -42, 45),
                (5700, 415.8, 23.4, 0.09, 23.4, 7.1, 0.26, 33.7, -
                29.6, 5.1, -2.6, 3.3, 0.6, -8.9, 9.1, -42, 45),
                (5800, 427.3, 24, 0.09, 24, 7.3, 0.27, 34.2, -
                30, 5.2, -2.7, 3.2, 0.6, -9.2, 9.5, -43, 45),
                (5900, 439.3, 24.6, 0.09, 24.6, 7.6, 0.28, 34.7, -
                30.5, 5.3, -2.8, 3.2, 0.6, -9.5, 9.8, -43, 46),
                (6000, 451.7, 25.2, 0.08, 25.2, 7.9, 0.28, 35.2, -
                31, 5.4, -2.9, 3.2, 0.7, -9.8, 10.1, -43, 46),
                (6100, 464.6, 25.9, 0.08, 25.9, 8.2, 0.29, 35.7, -
                31.5, 5.5, -3, 3.1, 0.7, -10.2, 10.5, -44, 46),
                (6200, 478.1, 26.5, 0.08, 26.6, 8.5, 0.3, 36.2, -
                31.9, 5.6, -3.1, 3.1, 0.8, -10.5, 10.8, -44, 47),
                (6300, 492.2, 27.2, 0.08, 27.2, 8.8, 0.31, 36.7, -
                32.4, 5.8, -3.2, 3.1, 0.8, -10.9, 11.2, -44, 47),
                (6400, 507.1, 28, 0.08, 28, 9.2, 0.31, 37.2, -
                32.9, 5.9, -3.3, 3, 0.8, -11.2, 11.6, -44, 47),
                (6500, 522.9, 28.7, 0.07, 28.7, 9.6, 0.32,
                37.7, -33.3, 6, -3.4, 3, 0.8, -11.6, 12, -44, 47),
                (6600, 539.7, 29.5, 0.07, 29.5, 10, 0.33, 38.1, -
                33.8, 6.2, -3.6, 2.9, 0.9, -11.9, 12.4, -45, 48),
                (6700, 557.8, 30.4, 0.07, 30.4, 10.5, 0.34, 38.6, -
                34.2, 6.3, -3.7, 2.9, 0.9, -12.3, 12.8, -45, 48),
                (6800, 577.7, 31.3, 0.07, 31.3, 11, 0.35, 39, -
                34.7, 6.5, -3.8, 2.8, 0.9, -12.7, 13.2, -45, 48),
                (6900, 599.7, 32.3, 0.07, 32.3, 11.6, 0.37, 39, -
                35.2, 6.8, -4, 2.8, 0.9, -13.1, 13.6, -45, 48),
                (7000, 625, 33.5, 0.06, 33.5, 12.4, 0.38, 39, -
                35.6, 6.8, -4.1, 2.8, 0.9, -13.5, 14.1, -45, 48),
                (7100, 655.6, 34.8, 0.06, 34.8, 13.3, 0.4, 39, -
                36.1, 6.8, -4.3, 2.8, 0.8, -13.9, 14.7, -45, 48),
                (7200, 697.4, 36.6, 0.06, 36.6, 14.7, 0.42, 39, -
                36.5, 6.8, -4.5, 2.7, 0.8, -14.4, 14.7, -45, 48),


            ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == giris_mesafesi:
                    nisangah5bh = round(nsg[1], 3)
                    print("Nişangah:", nisangah5bh)

            for tsg in interpolasyonlu_veri1:
                if tsg[0] == giris_mesafesi:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            for ts in interpolasyonlu_veri1:
                if ts[0] == giris_mesafesi:
                    deltaTS = round(ts[3], 3)
                    # print("Delta Tapa Saniyesi:", deltaTS)


            for ucs in interpolasyonlu_veri1:
                if ucs[0] == giris_mesafesi:
                    ucussuresi = round(ucs[4], 3)
                    # print("Uçuş Süresi:", ucussuresi)

            for dgl in interpolasyonlu_veri1:
                if dgl[0] == giris_mesafesi:
                    dogalyandz = round(dgl[5], 3)
                    # print("Doğal Yan Düzeltmesi:", dogalyandz)

            for yrd in interpolasyonlu_veri1:
                if yrd[0] == giris_mesafesi:
                    yan_ruzgari_duzeltme_faktoru = round(yrd[6], 1)
                    # print("Yan Rüzgarı Düzeltmesi:", yanruzgaridz)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_eksilme = round(ihz[7], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimeksilme)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_martma = round(ihz[8], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimartma)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_bas = round(mrd[9], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (BAŞ):", mesaferuzgaridzl_bas)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_arka = round(mrd[10], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (ARKA):", mesaferuzgaridzl_arka)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_eksilme = round(hvs[11], 3)
                    # print("Hava Isısı Düzeltmesi (Eksilme):", havaisisi_eksilme)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_artma = round(hvs[12], 3)
                    # print("Hava Isısı Düzeltmesi (Artma):", havaisisi_artma)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_eksilme = round(hyd[13], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", havayogunlugu_eksilme)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_artma = round(hyd[14], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", havayogunlugu_artma)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_eksilme = round(mka[15], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", mermikareagirligi_eksilme)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_artma = round(mka[16], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", mermikareagirligi_artma)

            print("Yan Rüzgarı Düzeltme Faktörü:", yan_ruzgari_duzeltme_faktoru)

            mesafe_ruzgari = ruzgar_hizi * mesafe_ruzgari_bileseni
            print("mesafe ruzgari", mesafe_ruzgari)

            ruzgar_yan_duzeltmesi = round(
                (ruzgar_hizi * yan_ruzgari_bileseni * yan_ruzgari_duzeltme_faktoru), 1)
            print("Rüzgar Yan Düzeltmesi", ruzgar_yan_duzeltmesi)

            dogalyandz = round(dogalyandz, 1)
            print("Doğal Yan Düzeltmesi:", dogalyandz)

            atis_istikameti = round(atis_istikameti/400)*400

            # dünyanın dönmesi düzeltmesi


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (1000, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2),
                (1500, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2),
                (2000, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3),
                (2500, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4),
                (3000, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5),
                (3500, 0.5, 0.5, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6),
                (4000, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7),
                (4500, 0.7, 0.7, 0.7, 0.8, 0.8, 0.8, 0.8, 0.9, 0.9),
                (5000, 0.8, 0.8, 0.8, 0.8, 0.9, 0.9, 1, 1, 1),
                (5500, 0.9, 0.9, 0.9, 1, 1, 1.1, 1.1, 1.1, 1.2),
                (6000, 0.9, 1, 1, 1.1, 1.1, 1.2, 1.3, 1.3, 1.3),
                (6500, 1, 1, 1.1, 1.2, 1.3, 1.4, 1.5, 1.5, 1.6),
                (7000, 1.1, 1.1, 1.2, 1.3, 1.5, 1.6, 1.8, 1.9, 1.9),




            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            if atis_istikameti == 00 or atis_istikameti == 6400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[1], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 6000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[2], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 5600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[3], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 5200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[4], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 4800:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[5], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2000 or atis_istikameti == 4400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[6], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2400 or atis_istikameti == 4000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[7], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2800 or atis_istikameti == 3600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[8], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 3200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[9], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (1000, 0, -1, -3, -4, -5, -6, -6, -7, -7, 0, 1, 3, 4, 5, 6, 7, 7),
                (1500, 0, -2, -4, -6, -7, -8, -9, -10, -10, 0, 2, 4, 6, 7, 8, 10, 10),
                (2000, 0, -3, -5, -7, -9, -11, -12, -13, -13, 0, 3, 5, 7, 9, 11, 13, 13),
                (2500, 0, -3, -6, -9, -11, -13, -15, -16, -16, 0, 3, 6, 9, 11, 13, 16, 16),
                (3000, 0, -4, -7, -10, -13, -16, -17, -18, -19, 0, 4, 7, 10, 13, 16, 18, 19),
                (3500, 0, -4, -8, -12, -15, -18, -20, -21, -21, 0, 4, 8, 12, 15, 18, 21, 21),
                (4000, 0, -5, -9, -13, -17, -20, -22, -23, -24, 0, 5, 9, 13, 17, 20, 23, 24),
                (4500, 0, -5, -10, -14, -18, -21, -24, -25, -26, 0, 5, 10, 14, 18, 21, 25, 26),
                (5000, 0, -5, -11, -15, -19, -23, -25, -27, -28, 0, 5, 11, 15, 19, 23, 27, 28),
                (5500, 0, -6, -11, -16, -21, -24, -27, -29, -29, 0, 6, 11, 16, 21, 24, 29, 29),
                (6000, 0, -6, -12, -17, -21, -25, -28, -30, -30, 0, 6, 12, 17, 21, 25, 30, 30),
                (6500, 0, -6, -12, -17, -22, -26, -29, -30, -31, 0, 6, 12, 17, 22, 26, 30, 31),
                (7000, 0, -6, -12, -17, -22, -25, -28, -30, -31, 0, 6, 12, 17, 22, 25, 30, 31),



            ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)
            if atis_istikameti == 0 or atis_istikameti == 3200 or atis_istikameti == 6400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[1], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 200 or atis_istikameti == 3000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[2], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 2800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[3], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 600 or atis_istikameti == 2600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[4], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 2400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[5], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1000 or atis_istikameti == 2200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[6], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 2000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[7], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1400 or atis_istikameti == 1800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[8], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 1600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[9], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            if atis_istikameti == 3400 or atis_istikameti == 6200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[11], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3600 or atis_istikameti == 6000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[12], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3800 or atis_istikameti == 5800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[13], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4000 or atis_istikameti == 5600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[14], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4200 or atis_istikameti == 5400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[15], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4400 or atis_istikameti == 5200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[16], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4600 or atis_istikameti == 5000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[17], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4800 or atis_istikameti == 4800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[18], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            print("Dünyanın Dönmesi Düzeltmesi:", arzindonusuyanduzeltmesi)

            metro_yan_duzeltmesi = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            metro_yan_duzeltmesi_1 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi:", metro_yan_duzeltmesi)
            metro_yan_duzeltmesi = round(metro_yan_duzeltmesi/1)*1
            print("Metro Yan Düzeltmesi en yakın 1 milyem:", metro_yan_duzeltmesi)


            # METRO MESAFE DÜZELTMESİ
            mesafe_ruzgari_fark = round(mesafe_ruzgari-0, 1)
            if mesafe_ruzgari_fark > 0:
                mesafe_ruzgari_durum = "BAŞ"
            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_durum = "ARKA"
            else:
                mesafe_ruzgari_durum = 0
            print("Mesafe Ruzgari Fark :", mesafe_ruzgari_fark)
            hava_sicakligi_fark = round(duzeltilmis_sıcaklik_degeri - 100, 1)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_durum = "EKSİLME"
            if hava_sicakligi_fark > 0:
                hava_sicakligi_durum = "ARTMA"
            else:
                hava_sicakligi_durum = 0
            print("Hava Sıcaklığı Fark :", hava_sicakligi_fark)
            hava_yogunlugu_fark = round(duzeltilmis_yogunluk_degeri - 100, 1)
            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_durum = "EKSİLME"
            if hava_yogunlugu_fark > 0:
                hava_yogunlugu_durum = "ARTMA"
            else:
                hava_yogunlugu_durum = 0
            print("Hava Yoğunluğu Fark :", hava_yogunlugu_fark)
            mermi_kare_agirligi_fark = round(mermi_kare_agirligi-2, 1)
            if mermi_kare_agirligi_fark < 0:
                merkar_durum = "EKSİLME"
            if mermi_kare_agirligi_fark > 0:
                merkar_durum = "ARTMA"
            if mermi_kare_agirligi_fark == 0:
                merkar_durum = 0

            print("Mermi Kare Ağırlığı Fark :", mermi_kare_agirligi_fark)

            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_arka
            else:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_bas
            if mesafe_ruzgari_fark == 0:
                mesafe_ruzgari_duzeltme_birimi = 0

            print("Mesafe Rüzgarı Düzeltme Birimi:", mesafe_ruzgari_duzeltme_birimi)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_duzeltme_birimi = havaisisi_eksilme
            else:
                hava_sicakligi_duzeltme_birimi = havaisisi_artma
            if hava_sicakligi_fark == 0:
                hava_sicakligi_duzeltme_birimi = 0
            print("Hava Isısı Düzeltme Birimi:", hava_sicakligi_duzeltme_birimi)

            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_eksilme
            else:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_artma
            if hava_yogunlugu_fark == 0:
                hava_yogunlugu_duzeltme_birimi = 0
            print("Hava Yoğunluğu Düzeltme Birimi:", hava_yogunlugu_duzeltme_birimi)

            if mermi_kare_agirligi_fark < 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_eksilme
            elif mermi_kare_agirligi_fark > 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_artma
            else:
                mermi_kare_agirligi_duzeltme_birimi = 0
            mesafe_ruzgari_fark = abs(mesafe_ruzgari_fark)
            hava_sicakligi_fark = abs(hava_sicakligi_fark)
            hava_yogunlugu_fark = abs(hava_yogunlugu_fark)


            print("Mermi Kare Ağırlığı Düzeltme Birimi:",
                mermi_kare_agirligi_duzeltme_birimi)

            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi = round(metro_yan_duzeltmesi + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)
            gac_yan_duzeltmesi = round(toplam_yan_duzeltmesi - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi)

            # Lineer interpolasyon fonksiyonu

            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            brt_isisi_ilk_hiz_dzl = [
                (-40, -8.5),
                (-30, -8),
                (-20, -7.4),
                (-10, -6.7),
                (0, -6),
                (10, -5.3),
                (20, -4.5),
                (30, -3.7),
                (40, -2.8),
                (50, -1.9),
                (60, -1),
                (70, 0),
                (80, 1),
                (90, 2.1),
                (100, 3.2),
                (110, 4.4),
                (120, 5.6),
                (130, 6.9),

            ]
            barut__isisi_dzl2 = 3
            barut_isisi_dzl = []
            for i in range(len(brt_isisi_ilk_hiz_dzl)-1):
                start = brt_isisi_ilk_hiz_dzl[i]
                end = brt_isisi_ilk_hiz_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    barut_isisi_dzl.append((int(x), y))

            barut_isisi_dzl.append(brt_isisi_ilk_hiz_dzl[-1])
            print("Barut Isısı",barut_isisi)

            for i in barut_isisi_dzl:
                if i[0] == barut_isisi:
                    print("mesafe",i[0])
                    barut__isisi_dzl2 = round(i[1], 1)
                    print("İ1:",i[1])
                    print("Barurrrrrrrr",barut__isisi_dzl2)
            # HIZ DEĞİŞİKLİĞİNİN HESAPLANMASI
            hiz_degisikligi = round(ilk_hiz_farki + mevzi_hiz_degisikligi, 1)
            deltaV_hiz_farki = round(hiz_degisikligi + barut__isisi_dzl2, 1)
            print("Hız Değişikliği:", hiz_degisikligi)
            print("Barut Isısı Düzeltmesi2:", barut__isisi_dzl2)
            print("Delta V Hız Farkı:", deltaV_hiz_farki)

            if deltaV_hiz_farki > 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_martma
            elif deltaV_hiz_farki < 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_eksilme
            else:
                ilk_hiz_duzeltme_birimi = 0
            print("İlk Hız Düzeltme Birimi:", ilk_hiz_duzeltme_birimi)

            deltaV_mesafe_duzeltmesi = abs(
                round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            if deltaV_hiz_farki > 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))*-1
            elif deltaV_hiz_farki < 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            else:
                deltaV_mesafe_duzeltmesi = 0
            print("Delta V Mesafe Düzeltmesi:", deltaV_mesafe_duzeltmesi)

            #### TOPLAM MESAF DÜZELTMESİ ###
            toplam_mesafe_duzeltmesi = round(metro_mesafe_duzeltmesi + deltaV_mesafe_duzeltmesi)
            print("Toplam Mesafe Düzeltmesi:", toplam_mesafe_duzeltmesi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)

            # toplam yan düzeltmede 1 milyem fazla mesafe düzeltmesinde işaret + ve 335 çıktı
            # baiek de yan 4 sl mesafe -326 çıktı

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (2, -0.007, 0.007, 0, 0, 0, 0, 0, 0, 0.013, -0.013),
                (3, -0.01, 0.01, 0, 0, -0.001, 0, 0, 0, 0.019, -0.019),
                (4, -0.013, 0.013, -0.001, 0, -0.001, 0, 0, 0, 0.024, -0.025),
                (5, -0.017, 0.016, -0.001, 0, -0.001, 0, 0.001, -0.001, 0.03, -0.03),
                (6, -0.02, 0.019, -0.001, 0, -0.002, 0, 0.001, -0.001, 0.035, -0.036),
                (7, -0.023, 0.022, -0.002, 0, -0.002, 0, 0.001, -0.001, 0.041, -0.041),
                (8, -0.026, 0.025, -0.002, 0, -0.003, 0, 0.001, -0.001, 0.046, -0.047),
                (9, -0.029, 0.028, -0.002, 0.001, -0.004, 0.001, 0.002, -0.002, 0.051, -0.052),
                (10, -0.032, 0.031, -0.003, 0.001, -0.004, 0.001, 0.002, -0.002, 0.056, -0.057),
                (11, -0.036, 0.034, -0.003, 0.001, -0.005, 0.001, 0.003, -0.003, 0.061, -0.062),
                (12, -0.039, 0.037, -0.003, 0.001, -0.005, 0.001, 0.003, -0.003, 0.066, -0.067),
                (13, -0.042, 0.04, -0.003, 0.001, -0.005, 0.001, 0.003, -0.003, 0.07, -0.072),
                (14, -0.045, 0.043, -0.004, 0.001, -0.006, 0.001, 0.004, -0.004, 0.075, -0.077),
                (15, -0.048, 0.046, -0.004, 0.001, -0.006, 0.001, 0.005, -0.004, 0.08, -0.082),
                (16, -0.051, 0.048, -0.004, 0.001, -0.007, 0.001, 0.005, -0.005, 0.084, -0.086),
                (17, -0.054, 0.051, -0.005, 0.001, -0.007, 0.001, 0.006, -0.006, 0.089, -0.091),
                (18, -0.057, 0.054, -0.005, 0.001, -0.007, 0.001, 0.006, -0.006, 0.093, -0.096),
                (19, -0.06, 0.057, -0.005, 0.001, -0.007, 0.001, 0.007, -0.007, 0.097, -0.1),
                (20, -0.063, 0.06, -0.005, 0.002, -0.008, 0, 0.008, -0.008, 0.102, -0.105),
                (21, -0.066, 0.063, -0.005, 0.002, -0.008, 0, 0.008, -0.008, 0.106, -0.109),
                (22, -0.069, 0.066, -0.006, 0.002, -0.008, 0, 0.009, -0.009, 0.11, -0.113),
                (23, -0.072, 0.068, -0.006, 0.002, -0.008, 0, 0.01, -0.01, 0.114, -0.118),
                (24, -0.075, 0.071, -0.006, 0.002, -0.008, 0, 0.011, -0.011, 0.118, -0.122),
                (25, -0.078, 0.074, -0.006, 0.002, -0.009, 0, 0.012, -0.011, 0.122, -0.126),
                (26, -0.081, 0.077, -0.006, 0.002, -0.009, 0, 0.013, -0.012, 0.126, -0.13),
                (27, -0.084, 0.08, -0.006, 0.002, -0.009, 0, 0.013, -0.013, 0.13, -0.134),
                (28, -0.087, 0.083, -0.007, 0.002, -0.009, 0, 0.014, -0.014, 0.134, -0.138),
                (29, -0.09, 0.086, -0.007, 0.003, -0.009, 0.001, 0.015, -0.015, 0.138, -0.142),
                (30, -0.093, 0.089, -0.007, 0.003, -0.009, 0.001, 0.016, -0.016, 0.142, -0.146),
                (31, -0.096, 0.091, -0.007, 0.003, -0.009, 0.001, 0.017, -0.017, 0.146, -0.15),
                (32, -0.099, 0.094, -0.007, 0.003, -0.009, 0.001, 0.018, -0.018, 0.149, -0.154),
                (33, -0.102, 0.097, -0.007, 0.003, -0.009, 0.001, 0.019, -0.019, 0.153, -0.158),
                (34, -0.104, 0.1, -0.007, 0.003, -0.009, 0.001, 0.02, -0.02, 0.157, -0.162),
                (35, -0.107, 0.103, -0.007, 0.003, -0.009, 0.001, 0.021, -0.021, 0.161, -0.166),
                (36, -0.11, 0.116, -0.007, 0.003, -0.009, 0.001, 0.022, -0.022, 0.164, -0.17),
                (37, -0.113, 0.109, -0.007, 0.004, -0.009, 0.001, 0.023, -0.023, 0.168, -0.174),
                (38, -0.116, 0.111, -0.007, 0.004, -0.009, 0.001, 0.024, -0.024, 0.172, -0.178),
                (39, -0.119, 0.114, -0.007, 0.004, -0.009, 0.002, 0.025, -0.025, 0.175, -0.181),
                (40, -0.122, 0.117, -0.007, 0.004, -0.009, 0.002, 0.026, -0.025, 0.179, -0.185),
                (41, -0.125, 0.12, -0.007, 0.004, -0.009, 0.002, 0.027, -0.026, 0.183, -0.189),
                (42, -0.128, 0.123, -0.007, 0.004, -0.009, 0.002, 0.028, -0.027, 0.187, -0.193),
                (43, -0.131, 0.126, -0.007, 0.004, -0.009, 0.002, 0.029, -0.028, 0.19, -0.197),
                (44, -0.134, 0.129, -0.007, 0.005, -0.009, 0.002, 0.03, -0.029, 0.194, -0.201),
                (45, -0.136, 0.131, -0.007, 0.005, -0.009, 0.002, 0.031, -0.03, 0.198, -0.205),
                (46, -0.139, 0.134, -0.007, 0.005, -0.009, 0.002, 0.032, -0.031, 0.202, -0.209),
                (47, -0.142, 0.137, -0.007, 0.005, -0.009, 0.002, 0.033, -0.032, 0.206, -0.213),
                (48, -0.145, 0.14, -0.007, 0.005, -0.009, 0.002, 0.034, -0.033, 0.21, -0.218),
                (49, -0.148, 0.143, -0.007, 0.005, -0.009, 0.002, 0.035, -0.034, 0.214, -0.222),
                (50, -0.151, 0.146, -0.007, 0.005, -0.009, 0.002, 0.036, -0.035, 0.219, -0.227),
                (51, -0.154, 0.149, -0.007, 0.006, -0.009, 0.002, 0.037, -0.036, 0.223, -0.231),
                (52, -0.157, 0.152, -0.007, 0.007, -0.009, 0.002, 0.038, -0.037, 0.228, -0.236),
                (53, -0.16, 0.155, -0.007, 0.01, -0.009, 0.002, 0.039, -0.038, 0.234, -0.242),
                (54, -0.163, 0.157, -0.007, 0.01, -0.009, 0.002, 0.04, -0.039, 0.242, -0.249),
                (55, -0.165, 0.159, -0.01, 0.01, -0.01, 0.002, 0.046, -0.043, 0.254, -0.267),
                (56, -0.158, 0.16, -0.02, 0.01, -0.005, 0.001, 0.046, -0.055, 0.342, -0.299),






            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Tapa Saniyesi :", tapasaniyesi)

            if deltaV_hiz_farki < 0:
                for ihz_eksilme in interpolasyonlu_veri1:
                    if ihz_eksilme[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_eksilme[1], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Eksilme):", ilkhizda1mildegisim_eksilme)
            if deltaV_hiz_farki > 0:
                for ihz_artma in interpolasyonlu_veri1:
                    if ihz_artma[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_artma[2], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Artma):", ilkhizda1mildegisim_artma)
            if mesafe_ruzgari_durum == "BAŞ":
                for musrus_bas in interpolasyonlu_veri1:
                    if musrus_bas[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(musrus_bas[3], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Baş):", mesaferuzgaridegisimi_bas)
            if mesafe_ruzgari_durum == "ARKA":
                for mesruz_arka in interpolasyonlu_veri1:
                    if mesruz_arka[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(mesruz_arka[4], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Arka):", mesaferuzgaridegisimi_arka)
            if hava_sicakligi_durum == "EKSİLME":
                for hadese_eksilme in interpolasyonlu_veri1:
                    if hadese_eksilme[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_eksilme[5], 3)
                        # print("Hava Isısı Değişim Miktarı (Eksilme):", havaisisidegisim_eksilme)
            if hava_sicakligi_durum == "ARTMA":
                for hadese_artma in interpolasyonlu_veri1:
                    if hadese_artma[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_artma[6], 3)
                        # print("Hava Isısı Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == "EKSİLME":
                for hayog_eksilme in interpolasyonlu_veri1:
                    if hayog_eksilme[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_eksilme[7], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Eksilme):", havayogunlugudegisim_eksilme)
            if hava_yogunlugu_durum == "ARTMA":
                for hayog_artma in interpolasyonlu_veri1:
                    if hayog_artma[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_artma[8], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == 0:
                havayogunlugudegisim = 0
            if merkar_durum == "EKSİLME":
                for merkar_eksilme in interpolasyonlu_veri1:
                    if merkar_eksilme[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_eksilme[9], 3)
                        # print("Mermi Kare Ağırlığı Değişim Miktarı (Eksilme):", mermikareagirligidegisim_eksilme)
            if merkar_durum == "ARTMA":
                for merkar_artma in interpolasyonlu_veri1:
                    if merkar_artma[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_artma[10], 3)
            else:
                mermikareagirligidegisim = 0
                # print("Mermi Kare Ağırlığı Değişim Miktarı (Artma):", mermikareagirligidegisim_artma)
##########################################################################################################
            gac_mesafe = plan_mesafesi_1B_obüs + toplam_mesafe_duzeltmesi
            print("toplam mesafe",nisangah5bh)
# zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (100, 5.6, 0, 00, 0.3, 0, 0, 0.7, -0.6, 0, 0, 0, 0, 0, 0, -1, 1),
                (200, 11.1, 0, 0, 0.7, 0, 0.01, 1.4, -1.2, 0, 0, 0, 0, 0, 0, -2, 2),
                (300, 16.6, 0, 0, 1, 0.1, 0.01, 2.1, -1.8, 0.1, 0, 0.1, 0, 0, 0, -4, 4),
                (400, 22.2, 0, 0, 1.3, 0.2, 0.02, 2.7, -2.4, 0.1, 0, 0.1, 0, -0.1, 0.1, -5, 5),
                (500, 27.8, 0, 0, 1.7, 0.3, 0.02, 3.4, -3, 0.2, 0, 0.2, 0, -0.1, 0.1, -6, 6),
                (600, 33.4, 2, 1.02, 2, 0.4, 0.02, 4.1, -3.6, 0.2, 0, 0.3, 0, -0.1, 0.1, -7, 7),
                (700, 39.1, 2.3, 0.87, 2.4, 0.4, 0.03, 4.7, -
                4.2, 0.3, -0.1, 0.4, -0.1, -0.1, 0.1, -8, 8),
                (800, 44.9, 2.6, 0.76, 2.7, 0.5, 0.03, 5.4, -
                4.8, 0.3, -0.1, 0.4, -0.1, -0.2, 0.2, -9, 9),
                (900, 50.6, 3, 0.68, 3, 0.6, 0.04, 6, -5.3,
                0.4, -0.1, 0.5, -0.1, -0.2, 0.2, -10, 10),
                (1000, 56.4, 3.3, 0.61, 3.4, 0.7, 0.04, 6.7, -
                5.9, 0.5, -0.1, 0.6, -0.1, -0.3, 0.3, -11, 11),
                (1100, 62.3, 3.7, 0.55, 3.7, 0.8, 0.04, 7.3, -
                6.4, 0.5, -0.1, 0.7, -0.1, -0.4, 0.4, -12, 13),
                (1200, 68.2, 4, 0.5, 4.1, 0.9, 0.05, 7.9, -
                7, 0.6, -0.2, 0.8, -0.1, -0.4, 0.4, -13, 14),
                (1300, 74.1, 4.4, 0.46, 4.4, 1, 0.05, 8.6, -
                7.5, 0.7, -0.2, 0.9, -0.1, -0.5, 0.5, -14, 15),
                (1400, 80.1, 4.8, 0.43, 4.8, 1.1, 0.06, 9.2, -
                8.1, 0.8, -0.2, 1, -0.1, -0.6, 0.6, -15, 16),
                (1500, 86.1, 5.1, 0.4, 5.2, 1.2, 0.06, 9.8, -
                8.6, 0.9, -0.2, 1.2, -0.2, -0.6, 0.6, -16, 17),
                (1600, 92.2, 5.5, 0.37, 5.5, 1.3, 0.06, 10.4, -
                9.1, 1, -0.3, 1.3, -0.2, -0.7, 0.7, -17, 17),
                (1700, 98.3, 5.8, 0.35, 5.9, 1.3, 0.07, 11.1, -
                9.7, 1.1, -0.3, 1.4, -0.2, -0.8, 0.8, -18, 18),
                (1800, 104.5, 6.2, 0.33, 6.2, 1.4, 0.07, 11.7, -
                10.2, 1.1, -0.3, 1.5, -0.2, -0.9, 0.9, -19, 19),
                (1900, 110.7, 6.6, 0.31, 6.6, 1.5, 0.08, 12.3, -
                10.7, 1.2, -0.3, 1.6, -0.2, -1, 1, -19, 20),
                (2000, 117, 6.9, 0.3, 7, 1.6, 0.08, 12.9, -
                11.2, 1.3, -0.4, 1.7, -0.2, -1.1, 1.1, -20, 21),
                (2100, 123.3, 7.3, 0.28, 7.3, 1.7, 0.08, 13.5, -
                11.8, 1.4, -0.4, 1.8, -0.2, -1.2, 1.3, -21, 22),
                (2200, 129.7, 7.7, 0.27, 7.7, 1.8, 0.09, 14.1, -
                12.3, 1.5, -0.4, 1.9, -0.2, -1.4, 1.4, -22, 23),
                (2300, 136.2, 8.1, 0.26, 8.1, 1.9, 0.09, 14.7, -
                12.8, 1.6, -0.5, 2, -0.2, -1.5, 1.5, -23, 24),
                (2400, 142.7, 8.4, 0.25, 8.5, 2.1, 0.1, 15.3, -
                13.3, 1.7, -0.5, 2.1, -0.2, -1.6, 1.6, -24, 25),
                (2500, 149.2, 8.8, 0.24, 8.9, 2.2, 0.1, 15.9, -
                13.8, 1.8, -0.6, 2.2, -0.2, -1.8, 1.8, -24, 26),
                (2600, 155.9, 9.2, 0.23, 9.2, 2.3, 0.11, 16.5, -
                14.3, 1.9, -0.6, 2.3, -0.2, -1.9, 1.9, -25, 26),
                (2700, 162.6, 9.6, 0.22, 9.6, 2.4, 0.11, 17.1, -
                14.9, 2, -0.6, 2.4, -0.2, -2, 2.1, -26, 27),
                (2800, 169.4, 10, 0.21, 10, 2.5, 0.11, 17.7, -
                15.4, 2.1, -0.7, 2.5, -0.2, -2.2, 2.2, -27, 28),
                (2900, 176.2, 10.4, 0.2, 10.4, 2.6, 0.12, 18.3, -
                15.9, 2.2, -0.7, 2.5, -0.2, -2.3, 2.4, -27, 29),
                (3000, 183.1, 10.8, 0.19, 10.8, 2.7, 0.12, 18.9, -
                16.4, 2.3, -0.8, 2.6, -0.2, -2.5, 2.5, -28, 29),
                (3100, 190.1, 11.2, 0.19, 11.2, 2.8, 0.13, 19.4, -
                16.9, 2.4, -0.8, 2.7, -0.2, -2.7, 2.7, -29, 30),
                (3200, 197.2, 11.6, 0.18, 11.6, 3, 0.13, 20, -
                17.4, 2.5, -0.9, 2.8, -0.2, -2.8, 2.9, -29, 31),
                (3300, 204.4, 12, 0.17, 12, 3.1, 0.14, 20.6, -
                17.9, 2.6, -0.9, 2.8, -0.1, -3, 3.1, -30, 32),
                (3400, 211.6, 12.4, 0.17, 12.4, 3.2, 0.14, 21.2, -
                18.4, 2.7, -1, 2.9, -0.1, -3.2, 3.3, -31, 32),
                (3500, 219, 12.8, 0.16, 12.8, 3.3, 0.14, 21.7, -
                18.9, 2.8, -1, 2.9, -0.1, -3.4, 3.4, -31, 33),
                (3600, 226.4, 13.2, 0.16, 13.2, 3.5, 0.15, 22.3, -
                19.4, 2.9, -1.1, 3, -0.1, -3.6, 3.6, -32, 34),
                (3700, 233.9, 13.6, 0.15, 13.7, 3.6, 0.15, 22.9, -
                19.9, 3, -1.1, 3, -0.1, -3.8, 3.8, -33, 34),
                (3800, 241.6, 14.1, 0.15, 14.1, 3.7, 0.16, 23.4, -
                20.4, 3.1, -1.2, 3.1, -0.1, -4, 4.1, -33, 35),
                (3900, 249.3, 14.5, 0.14, 14.5, 3.9, 0.16, 24, -
                20.9, 3.2, -1.2, 3.1, 0, -4.2, 4.3, -34, 36),
                (4000, 257.2, 14.9, 0.14, 15, 4, 0.17, 24.6, -
                21.4, 3.3, -1.3, 3.2, 0, -4.4, 4.5, -34, 36),
                (4100, 265.2, 15.4, 0.14, 15.4, 4.1, 0.17, 25.1, -
                21.9, 3.4, -1.4, 3.2, 0, -4.6, 4.7, -35, 37),
                (4200, 273.3, 15.8, 0.13, 15.9, 4.3, 0.18, 25.7, -
                22.3, 3.5, -1.4, 3.2, 0, -4.8, 4.9, -36, 38),
                (4300, 281.5, 16.3, 0.13, 16.3, 4.4, 0.18, 26.2, -
                22.8, 3.6, -1.5, 3.3, 0.1, -5.1, 5.2, -36, 38),
                (4400, 289.9, 16.7, 0.13, 16.8, 4.6, 0.19, 26.8, -
                23.3, 3.7, -1.6, 3.3, 0.1, -5.3, 5.4, -37, 39),
                (4500, 298.4, 17.2, 0.12, 17.2, 4.8, 0.19, 27.3, -
                23.8, 3.8, -1.6, 3.3, 0.1, -5.6, 5.7, -37, 39),
                (4600, 307, 17.7, 0.12, 17.7, 4.9, 0.2, 27.9, -
                24.3, 3.9, -1.7, 3.3, 0.1, -5.8, 5.9, -38, 40),
                (4700, 315.9, 18.2, 0.12, 18.2, 5.1, 0.2, 28.4, -
                24.8, 4, -1.8, 3.3, 0.2, -6, 6.2, -38, 40),
                (4800, 324.9, 18.6, 0.11, 18.7, 5.3, 0.21, 29, -
                25.3, 4.1, -1.8, 3.3, 0.2, -6.3, 6.5, -39, 41),
                (4900, 334, 19.1, 0.11, 19.1, 5.4, 0.21, 29.5, -
                25.7, 4.2, -1.9, 3.4, 0.2, -6.6, 6.7, -39, 41),
                (5000, 343.4, 19.6, 0.11, 19.6, 5.6, 0.22,
                30, -26.2, 4.3, -2, 3.4, 0.3, -6.8, 7, -40, 42),
                (5100, 353, 20.1, 0.11, 20.2, 5.8, 0.23, 30.6, -
                26.7, 4.4, -2.1, 3.4, 0.3, -7.1, 7.3, -40, 42),
                (5200, 362.8, 20.7, 0.1, 20.7, 6, 0.23, 31.1, -
                27.2, 4.6, -2.1, 3.3, 0.4, -7.4, 7.6, -40, 43),
                (5300, 372.8, 21.2, 0.1, 21.2, 6.2, 0.24, 31.6, -
                27.7, 4.7, -2.2, 3.3, 0.4, -7.7, 7.9, -41, 43),
                (5400, 383.1, 21.7, 0.1, 21.7, 6.4, 0.24, 32.2, -
                28.1, 4.8, -2.3, 3.3, 0.4, -8, 8.2, -41, 44),
                (5500, 393.7, 22.3, 0.1, 22.3, 6.6, 0.25, 32.7, -
                28.6, 4.9, -2.4, 3.3, 0.5, -8.3, 8.5, -42, 44),
                (5600, 404.6, 22.8, 0.09, 22.8, 6.9, 0.26, 33.2, -
                29.1, 5, -2.5, 3.3, 0.5, -8.6, 8.8, -42, 45),
                (5700, 415.8, 23.4, 0.09, 23.4, 7.1, 0.26, 33.7, -
                29.6, 5.1, -2.6, 3.3, 0.6, -8.9, 9.1, -42, 45),
                (5800, 427.3, 24, 0.09, 24, 7.3, 0.27, 34.2, -
                30, 5.2, -2.7, 3.2, 0.6, -9.2, 9.5, -43, 45),
                (5900, 439.3, 24.6, 0.09, 24.6, 7.6, 0.28, 34.7, -
                30.5, 5.3, -2.8, 3.2, 0.6, -9.5, 9.8, -43, 46),
                (6000, 451.7, 25.2, 0.08, 25.2, 7.9, 0.28, 35.2, -
                31, 5.4, -2.9, 3.2, 0.7, -9.8, 10.1, -43, 46),
                (6100, 464.6, 25.9, 0.08, 25.9, 8.2, 0.29, 35.7, -
                31.5, 5.5, -3, 3.1, 0.7, -10.2, 10.5, -44, 46),
                (6200, 478.1, 26.5, 0.08, 26.6, 8.5, 0.3, 36.2, -
                31.9, 5.6, -3.1, 3.1, 0.8, -10.5, 10.8, -44, 47),
                (6300, 492.2, 27.2, 0.08, 27.2, 8.8, 0.31, 36.7, -
                32.4, 5.8, -3.2, 3.1, 0.8, -10.9, 11.2, -44, 47),
                (6400, 507.1, 28, 0.08, 28, 9.2, 0.31, 37.2, -
                32.9, 5.9, -3.3, 3, 0.8, -11.2, 11.6, -44, 47),
                (6500, 522.9, 28.7, 0.07, 28.7, 9.6, 0.32,
                37.7, -33.3, 6, -3.4, 3, 0.8, -11.6, 12, -44, 47),
                (6600, 539.7, 29.5, 0.07, 29.5, 10, 0.33, 38.1, -
                33.8, 6.2, -3.6, 2.9, 0.9, -11.9, 12.4, -45, 48),
                (6700, 557.8, 30.4, 0.07, 30.4, 10.5, 0.34, 38.6, -
                34.2, 6.3, -3.7, 2.9, 0.9, -12.3, 12.8, -45, 48),
                (6800, 577.7, 31.3, 0.07, 31.3, 11, 0.35, 39, -
                34.7, 6.5, -3.8, 2.8, 0.9, -12.7, 13.2, -45, 48),
                (6900, 599.7, 32.3, 0.07, 32.3, 11.6, 0.37, 39, -
                35.2, 6.8, -4, 2.8, 0.9, -13.1, 13.6, -45, 48),
                (7000, 625, 33.5, 0.06, 33.5, 12.4, 0.38, 39, -
                35.6, 6.8, -4.1, 2.8, 0.9, -13.5, 14.1, -45, 48),
                (7100, 655.6, 34.8, 0.06, 34.8, 13.3, 0.4, 39, -
                36.1, 6.8, -4.3, 2.8, 0.8, -13.9, 14.7, -45, 48),
                (7200, 697.4, 36.6, 0.06, 36.6, 14.7, 0.42, 39, -
                36.5, 6.8, -4.5, 2.7, 0.8, -14.4, 14.7, -45, 48),


            ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == gac_mesafe:
                    nisangah5bh = round(nsg[1], 1)
                    print("Nişangah:", nisangah5bh)
            for dgl in interpolasyonlu_veri1:
                if dgl[0] == gac_mesafe:
                    dogalyandz = round(dgl[5],3)
                    print("Doğal Yan Düzeltmesi:", dogalyandz)

            for tsg in interpolasyonlu_veri1:
                if tsg[0] == gac_mesafe:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            metro_yan_duzeltmesi2 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi en yakın 1 milyem99999999999999:", metro_yan_duzeltmesi2)
            print("Arzin Dnüşü:,",arzindonusuyanduzeltmesi)
            print("Rüzgaryan dzl---------------------:,",ruzgar_yan_duzeltmesi)

            
            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi5bh = round(metro_yan_duzeltmesi2 + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi33333333333335bh:", toplam_yan_duzeltmesi5bh)
            gac_yan_duzeltmesi2 = round(toplam_yan_duzeltmesi5bh - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi2)
# Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]



            # Veri kümesi
            data1 = [
                (4000,181,159,0.014,-0.012),
(4500,213,212,0.019,-0.015),
(5000,246,275,0.024,-0.02),
(5500,280,347,0.031,-0.026),
(6000,316,430,0.041,-0.034),
(6500,353,525,0.054,-0.045),
(7000,393,633,0.071,-0.059),
(7500,434,757,0.094,-0.078),
(8000,479,899,0.125,-0.103),
(8500,526,1063,0.17,-0.138),
(9000,577,1255,0.237,-0.188),
(9500,633,1486,0.346,-0.263),
(10000,697,1773,0.558,-0.388),
(10500,776,2165,1.286,-0.646),
(11000,930,3054,1.286,-2.11),


]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == gac_mesafe:
                    dusus_acisi = round(dusus_acisi[1],3)
                    # print("Düşüş Açısı:", dusus_acisi)

            for tepe_yuksekligi in interpolasyonlu_veri1:
                if tepe_yuksekligi[0] == gac_mesafe:
                    tepe_yuksekligi = round(tepe_yuksekligi[2],3)
                    # print("Tepe Yüksekliği:", tepe_yuksekligi)

            for dtac_arti1 in interpolasyonlu_veri1:
                if dtac_arti1[0] == mesafe:
                    dtac_arti1 = round(dtac_arti1[3],3)
                    print("+1 DTAÇ:", dtac_arti1)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4],3)
                    print("-1 DTAÇ:", dtac_eksi1)
    ############################################################################################################
            """METRO DÜZELTMELERİNİ HESAPLAMAK GEREKİYOR"""
            # GTAC 1 inci oonüs için
            bt_rakimi = lne_1B_obus_rakim

            milyem_sabit_dz = 1.0186

            hedef_batarya_rakim_farki = (rakim - bt_rakimi)

            dtac = (hedef_batarya_rakim_farki / (plan_mesafesi_1B_obüs / 1000)) * milyem_sabit_dz

            # Burada dtac_arti1 ve dtac_eksi1 değerlerini sadece bir kez tanımlayın
            dtac_arti1 = 0
            dtac_eksi1 = 0

            # dtac_arti1 ve dtac_eksi1 değerlerini doğru şekilde ata
            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_arti1 = round(dusus_acisi[3], 3)
                    print("dtac_arti1",dtac_arti1)


            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4], 3)
                    print("dtac_eksi1",dtac_eksi1)

            ttac = 0
            print("Mesafe",mesafe)
            if dtac < 0:
                ttac = int(dtac * dtac_arti1)
                tac = dtac + ttac
                #print("-TAÇ", tac)

            if dtac > 0:
                ttac = int(dtac * dtac_eksi1)
                tac = dtac + ttac
                print("+TAÇ", tac)
    #########################################################################################################
            """YÜKSELİŞİ BUL"""
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,18),
(100,18),
(200,18),
(300,18),
(400,18),
(500,18),
(600,18),
(700,18),
(800,18),
(900,17),
(1000,17),
(1100,17),
(1200,17),
(1300,17),
(1400,17),
(1500,17),
(1600,16),
(1700,16),
(1800,16),
(1900,16),
(2000,16),
(2100,16),
(2200,16),
(2300,15),
(2400,15),
(2500,15),
(2600,15),
(2700,15),
(2800,15),
(2900,15),
(3000,14),
(3100,14),
(3200,14),
(3300,14),
(3400,14),
(3500,14),
(3600,13),
(3700,13),
(3800,13),
(3900,13),
(4000,13),
(4100,12),
(4200,12),
(4300,12),
(4400,12),
(4500,12),
(4600,11),
(4700,11),
(4800,11),
(4900,11),
(5000,11),
(5100,10),
(5200,10),
(5300,10),
(5400,10),
(5500,9),
(5600,9),
(5700,9),
(5800,8),
(5900,8),
(6000,8),
(6100,8),
(6200,7),
(6300,7),
(6400,7),
(6500,6),
(6600,6),
(6700,5),
(6800,5),
(6900,4),
(7000,4),
(7100,3),
(7200,3),
            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            
            for nsg in interpolasyonlu_veri1:
                    if nsg[0] == gac_mesafe:
                        birmildegisim5bh = round(nsg[1], 1)
                        print("1 Milyemlik Değişim:", birmildegisim5bh)
            yukselis5bh = round(nisangah5bh+ + tac,1)
            self.ui.sonuc_yukselis_1B.setText(str(yukselis5bh))
    #########################################################################################################
            yan_1B_5bh = round(yan_1B+toplam_yan_duzeltmesi5bh)
            self.ui.sonuc_yan_1B.setText(str(yan_1B_5bh))
            self.ui.sonuc_yan_2B.setText(str(yan2))
            self.ui.lne_barut_hakki.setText(str(secilen_barut_hakki))
            istikamet_acisi = atis_istikameti_1
    
            print("SEÇİLEN BRUT", secilen_barut_hakki)
            # self.ui.sonuc_barut_hakki_1B.setText(str(secilen_barut_hakki))
            if 2200 < yan_1B_5bh < 3000:
                self.ui.sonuc_yan_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_yukselis_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_barut_hakki_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_tapa_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_istikamet_acisi_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_mesafe_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.lne_sonuc_mesafe.setText(str(mesafe))
                self.ui.lne_sonuc_barut_hakki.setText(str(secilen_barut_hakki))
                self.ui.lne_sonuc_yan.setText(str(yan_1B_5bh))
                self.ui.lne_sonuc_yukselis.setText(str(yukselis5bh))
                self.ui.lne_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
                self.ui.lne_sonuc_bir_ml_degisiklil.setText(str(birmildegisim5bh)) 
                self.ui.lne_sonuc_yuz_m.setText(str(mesafe/1000)) 
            else:
                self.ui.sonuc_yan_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_yukselis_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_barut_hakki_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_tapa_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_istikamet_acisi_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_mesafe_1B.setStyleSheet("color: grey; font-weight: normal")
            self._doldur_sonuc("1", toplam_yan_duzeltmesi5bh, yukselis5bh, birmildegisim5bh, locals())
        #################################################################################################
        if secilen_barut_hakki == 6:
            plan_mesafesi = plan_mesafesi_1B_obüs
            batarya_rakimi = lne_1B_obus_rakim
            bt_rakimi_1 = lne_1B_obus_rakim
            hedef_rakimi = lne_hedef_rakim
            ahia = lne_1B_obus_ahia
            paralanma_yuksekliği = paralanma_yuksekliği
            mermi_kare_agirligi = mermi_kare_agirligi
            ilk_hiz_farki = lne_1B_obus_ihf
            atis_istikameti = batarya_hedef_İA_1
            atis_istikameti_1 = batarya_hedef_İA_1
            mevzi_hiz_degisikligi = 0
            mevzi_yan_duzeltmesi = 0
            mevzi_ts_duzeltmesi = 0
            tapasaniyesi = 1
            barut_isisi_str = self.ui.lne_barut_isisi.text()
            barut_isisi = float(barut_isisi_str) if barut_isisi_str else 75
            # barut_isisi = self.ui.lne_barut_isisi.text() if self.ui.lne_barut_isisi.text() else 75
            print("Barutttt İSİSİ",barut_isisi)
            
            # 1. GİRİŞ MESAFESİNİ BUL / MANUEL HASSASİYETİ

            # mesafe = 1866
            # en_yakin_10_metre = round(mesafe / 10) * 10
            # print("1836'nın en yakın 10 metreye çevrilmiş hali:", en_yakin_10_metre)

            # BATARYA RAKIMINI EN YAKIN 10 M YE ÇEVİR.
            batarya_rakimi = round(batarya_rakimi/10)*10
            # PARALANMA NOKTASININ RAKIMINI BUL.
            paralanma_noktasi_rakimi = hedef_rakimi + paralanma_yuksekliği
            # HEDEF-BATARYA YÜKSEKLİK FARKINI BUL
            yukseklik_farki = paralanma_noktasi_rakimi-batarya_rakimi
            yukseklik_farki_1 = paralanma_noktasi_rakimi-bt_rakimi_1
            # yukseklik_farki = round(yukseklik_farki/100)*100
            ################# B cetvelini yüklüyoruz.################
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
                # Veri kümesi
            data1 = [
                    (2500,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2600,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2700,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2800,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2900,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3000,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3100,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3200,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3300,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3400,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3500,0,0,0,1,1,1,1,1,2,2,2,2,3,3,3),
(3600,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(3700,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(3800,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(3900,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4000,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4100,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4200,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4300,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4400,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4500,1,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(4600,1,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(4700,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(4800,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(4900,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5000,1,1,1,1,2,2,2,2,2,2,2,3,3,3,3),
(5100,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5200,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5300,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(5400,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(5500,1,1,1,2,2,2,2,2,2,3,3,3,3,3,3),
(5600,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(5700,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(5800,1,1,2,2,2,2,2,2,3,3,3,3,3,3,3),
(5900,1,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6000,1,2,2,2,2,2,2,2,3,3,3,3,3,3,4),
(6100,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(6200,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(6300,2,2,2,2,2,2,2,3,3,3,3,3,3,4,4),
(6400,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(6500,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(6600,2,2,2,2,2,3,3,3,3,3,3,3,4,4,4),
(6700,2,2,2,2,2,3,3,3,3,3,3,3,4,4,4),
(6800,2,2,2,2,3,3,3,3,3,3,3,4,4,4,4),
(6900,2,2,2,2,3,3,3,3,3,3,3,4,4,4,4),
(7000,2,2,2,3,3,3,3,3,3,3,3,4,4,4,4),
(7100,2,2,2,3,3,3,3,3,3,3,4,4,4,4,4),
(7200,2,2,3,3,3,3,3,3,3,3,4,4,4,4,4),
(7300,2,2,3,3,3,3,3,3,3,4,4,4,4,4,4),
(7400,2,3,3,3,3,3,3,3,4,4,4,4,4,4,5),
(7500,3,3,3,3,3,3,3,3,4,4,4,4,4,4,5),
(7600,3,3,3,3,3,3,3,4,4,4,4,4,4,5,5),
(7700,3,3,3,3,3,3,3,4,4,4,4,4,5,5,5),
(7800,3,3,3,3,3,3,4,4,4,4,4,4,5,5,5),
(7900,3,3,3,3,3,4,4,4,4,4,4,5,5,5,5),
(8000,3,3,3,3,3,4,4,4,4,4,5,5,5,5,5),
(8100,3,3,3,3,4,4,4,4,4,4,5,5,5,5,5),
(8200,3,3,3,4,4,4,4,4,4,5,5,5,5,5,6),
(8300,3,3,4,4,4,4,4,4,5,5,5,5,5,6,6),
(8400,3,3,4,4,4,4,4,5,5,5,5,5,6,6,6),
(8500,3,4,4,4,4,4,5,5,5,5,5,6,6,6,6),
(8600,4,4,4,4,4,4,5,5,5,5,6,6,6,6,6),
(8700,4,4,4,4,4,5,5,5,5,6,6,6,6,6,6),
(8800,4,4,4,4,5,5,5,5,6,6,6,6,6,6,6),
(8900,4,4,4,5,5,5,5,6,6,6,6,6,6,6,6),
(9000,4,4,5,5,5,6,6,6,6,6,6,6,6,6,6),


            ]

            data2 = [
                (2500,0,0,-12,-6,0,7,14,22,31,40,50,61,73,85,97),
(2600,0,-17,-12,-6,0,7,15,23,32,42,52,63,75,87,100),
(2700,0,-18,-13,-7,0,7,15,24,33,43,54,65,77,90,104),
(2800,0,-19,-13,-7,0,8,16,25,34,45,56,67,80,93,107),
(2900,0,-20,-14,-7,0,8,17,26,36,46,58,70,82,96,110),
(3000,-27,-21,-15,-8,0,8,17,27,37,48,60,72,85,99,114),
(3100,-28,-22,-15,-8,0,9,18,28,38,50,62,74,88,102,117),
(3200,-29,-23,-16,-8,0,9,18,29,40,51,64,77,91,106,121),
(3300,-31,-24,-17,-9,0,9,19,30,41,53,66,80,94,109,125),
(3400,-32,-25,-17,-9,0,10,20,31,43,55,68,82,97,113,129),
(3500,-33,-26,-18,-9,0,10,21,32,44,57,71,85,100,116,133),
(3600,-35,-27,-19,-10,0,10,21,33,46,59,73,88,104,120,138),
(3700,-36,-28,-19,-10,0,11,22,34,47,61,76,91,107,124,142),
(3800,-38,-29,-20,-10,0,11,23,36,49,63,78,94,111,128,147),
(3900,-39,-31,-21,-11,0,12,24,37,51,65,81,97,114,132,151),
(4000,-41,-32,-22,-11,0,12,25,38,52,68,83,100,118,137,156),
(4100,-43,-33,-23,-12,0,12,26,40,54,70,86,104,122,141,161),
(4200,-44,-34,-24,-12,0,13,26,41,56,72,89,107,126,145,166),
(4300,-46,-36,-24,-13,0,13,27,42,58,75,92,110,130,150,171),
(4400,-48,-37,-25,-13,0,14,28,44,60,77,95,114,134,155,177),
(4500,-50,-38,-26,-13,0,14,29,45,62,80,98,118,138,160,182),
(4600,-51,-40,-27,-14,0,15,30,47,64,82,101,122,143,165,188),
(4700,-53,-41,-28,-14,0,15,31,48,66,85,105,125,147,170,194),
(4800,-55,-43,-29,-15,0,16,32,50,68,88,108,129,152,175,200),
(4900,-57,-44,-30,-15,0,16,34,52,71,91,112,134,157,181,206),
(5000,-59,-46,-31,-16,0,17,35,53,73,94,115,138,162,186,213),
(5100,-61,-47,-32,-17,0,17,36,55,75,96,119,142,167,192,219),
(5200,-64,-49,-33,-17,0,18,37,57,78,100,122,147,172,198,226),
(5300,-66,-51,-35,-18,0,19,38,59,80,103,126,151,177,204,233),
(5400,-68,-52,-36,-18,0,19,39,60,83,106,130,156,182,210,240),
(5500,-70,-54,-37,-19,0,20,41,62,85,109,134,161,188,217,247),
(5600,-73,-56,-38,-20,0,20,42,64,88,113,138,165,194,224,255),
(5700,-75,-58,-39,-20,0,21,43,66,91,116,143,171,200,230,263),
(5800,-77,-59,-41,-21,0,22,45,68,93,120,147,176,206,237,271),
(5900,-80,-61,-42,-21,0,22,46,71,96,123,152,181,212,245,279),
(6000,-83,-63,-43,-22,0,23,47,73,99,127,156,187,219,252,287),
(6100,-85,-65,-45,-23,0,24,49,75,102,131,161,192,225,260,296),
(6200,-88,-67,-46,-23,0,25,50,77,105,135,166,198,232,268,305),
(6300,-91,-69,-47,-24,0,25,52,80,109,139,171,204,239,276,315),
(6400,-93,-72,-49,-25,0,26,53,82,112,143,176,211,247,285,325),
(6500,-96,-74,-50,-26,0,27,55,85,115,148,182,217,255,294,335),
(6600,-99,-76,-52,-27,0,28,57,87,119,152,187,224,263,303,346),
(6700,-102,-78,-53,-27,0,29,59,90,123,157,193,231,271,313,357),
(6800,-105,-81,-55,-28,0,29,60,93,126,162,199,238,280,323,369),
(6900,-109,-83,-57,-29,0,30,62,95,130,167,206,246,289,334,381),
(7000,-112,-86,-59,-30,0,31,64,98,134,172,212,254,298,345,394),
(7100,-115,-88,-60,-31,0,32,66,102,139,178,219,262,308,356,408),
(7200,-119,-91,-62,-32,0,33,68,105,143,184,226,271,319,369,422),
(7300,-123,-94,-64,-33,0,34,70,108,148,190,234,280,330,382,438),
(7400,-126,-97,-66,-34,0,35,73,112,153,196,242,290,341,396,455),
(7500,-130,-100,-68,-35,0,37,75,115,158,203,250,301,354,411,473),
(7600,-134,-103,-70,-36,0,38,77,119,163,210,259,312,368,428,493),
(7700,-139,-106,-73,-37,0,39,80,123,169,217,269,324,382,446,515),
(7800,-143,-110,-75,-38,0,40,83,128,175,226,279,337,399,466,540),
(7900,-147,-113,-77,-40,0,42,86,132,182,234,291,351,417,489,569),
(8000,-152,-117,-80,-41,0,43,89,137,189,244,303,367,437,515,604),
(8100,-157,-121,-83,-42,0,45,92,143,197,254,317,385,461,548,651),
(8200,-163,-125,-86,-44,0,46,96,149,205,266,333,407,491,591,738),
(8300,-168,-129,-89,-46,0,48,100,155,215,280,352,433,531,680,0),
(8400,-174,-134,-92,-47,0,50,104,162,226,296,375,471,0,0,0),
(8500,-180,-139,-95,-49,0,53,109,171,239,316,409,0,0,0,0),
(8600,-187,-145,-99,-51,0,55,115,181,256,346,0,0,0,0,0),
(8700,-195,-151,-104,-54,0,58,122,195,283,0,0,0,0,0,0),
(8800,-203,-157,-109,-56,0,62,132,217,0,0,0,0,0,0,0),
(8900,-212,-165,-114,-60,0,67,150,0,0,0,0,0,0,0,0),
(9000,-222,-173,-121,-64,0,0,0,0,0,0,0,0,0,0,0),
            ]
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            interpolasyonlu_veri2 = veri_kumesi_interpolasyonu(data2)
            # print(interpolasyonlu_veri2)

            hedefin_toptan_yuksekligi = yukseklik_farki
            hedefin_toptan_yuksekligi = round(hedefin_toptan_yuksekligi/100)*100
            print("Hedefin Toptan Yüksekliği: ", hedefin_toptan_yuksekligi)
            mesafe = plan_mesafesi
            print("mesafe = ", mesafe)

            if hedefin_toptan_yuksekligi == -400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[1])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[2])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[3])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[4])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -0:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[5])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[6])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[7])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[8])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[9])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 500:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[10])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 600:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[11])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 700:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 800:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 900:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[13])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 1000:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[14])
                        print("Seçilen Bölge:", bolge)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")
            # TXT dosyasının adı ve dosya yolu
            dosya_adı = METRO_DOSYA

            # İlk satırı atlayarak dosyanın geri kalanını yeni bir liste olarak alın
            metrorpr = []
            if not os.path.isfile(dosya_adı):
                QMessageBox.warning(self, "Metro", "METRAP.txt bulunamadı.")
                return None
            with open(dosya_adı, "r") as dosya:
                for satır in dosya:
                    metrorpr.append(satır.strip())

            # Satırları yan yana birleştirerek elde etmek için join() yöntemini kullanıyoruz
            birlesik_veri = "\n".join(metrorpr[1:])

            rapor_tipi = birlesik_veri[:5]
            oktant = birlesik_veri[5:6]
            metro_koor = birlesik_veri[6:9] + " " + birlesik_veri[9:12]
            tarih = birlesik_veri[13:15]
            saat = birlesik_veri[15:17]
            dakika = str(int(birlesik_veri[17:18]) * 6)
            if dakika == "0":
                dakika = str(int(birlesik_veri[17:18]) * 6) + "0"
            ara = ":"
            saat_dakika = saat + ara + str(dakika)

            gecerlilik = birlesik_veri[18:19]
            metro_istasyonu_rakimi = str(int(birlesik_veri[19:22]) * 10)
            yogunluk = int(birlesik_veri[22:25]) / 10

            print("Rapor Tipi:", rapor_tipi)
            print("Oktant:", oktant)
            print("Metro Koordinat:", metro_koor)
            print("Tarih:", tarih)
            print("Saat:", saat_dakika)
            print("Geçerlilik:", gecerlilik)
            print("Metro İstasyonu Rakımı:", metro_istasyonu_rakimi)
            print("Yoğunluk:", yogunluk)
            if bolge == 0:
                ruzgar_istikameti = int(birlesik_veri[28:30]) * 100
                ruzgar_hizi = int(birlesik_veri[30:32])
                if birlesik_veri[32] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[32:35]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[32:35]) / 10
                if birlesik_veri[35] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[35:38]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[35:38]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)  # 39-40-41 3 say

            elif bolge == 1:
                ruzgar_istikameti = int(birlesik_veri[41:43]) * 100
                ruzgar_hizi = int(birlesik_veri[43:45])
                if birlesik_veri[45] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[45:48]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[45:48]) / 10
                if birlesik_veri[48] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[48:51]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[48:51]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)
            elif bolge == 2:
                ruzgar_istikameti = int(birlesik_veri[54:56])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[56:58])  # 3
                if birlesik_veri[58] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[58:61])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[58:61])/10  # 4
                if birlesik_veri[61] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[61:64])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[61:64])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 3:
                ruzgar_istikameti = int(birlesik_veri[67:69])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[69:71])  # 3
                if birlesik_veri[71] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[71:74])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[71:74])/10  # 4
                if birlesik_veri[74] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[74:77])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[74:77])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 4:
                ruzgar_istikameti = int(birlesik_veri[80:82])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[82:84])  # 3
                if birlesik_veri[84] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[84:87])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[84:87])/10  # 4
                if birlesik_veri[87] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[87:90])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[87:90])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 5:
                ruzgar_istikameti = int(birlesik_veri[93:95])*100  # 3
                ruzgar_hizi = int(birlesik_veri[95:97])  # 3
                if birlesik_veri[97] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[97:100])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[97:100])/10  # 4
                if birlesik_veri[100] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[100:103])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[100:103])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 6:
                ruzgar_istikameti = int(birlesik_veri[106:108])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[108:110])  # 3
                if birlesik_veri[110] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[110:113])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[110:113])/10  # 4
                if birlesik_veri[113] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[113:116])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[113:116])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 7:
                ruzgar_istikameti = int(birlesik_veri[119:121])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[121:123])  # 3
                if birlesik_veri[123] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[123:126])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[123:126])/10  # 4
                if birlesik_veri[126] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[126:129])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[126:129])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 8:
                ruzgar_istikameti = int(birlesik_veri[132:134])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[134:136])  # 3
                if birlesik_veri[136] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[136:139])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[136:139])/10  # 4
                if birlesik_veri[139] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[139:142])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[139:142])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 9:
                ruzgar_istikameti = int(birlesik_veri[145:147])*100  # 3
                ruzgar_hizi = int(birlesik_veri[147:149])  # 3
                if birlesik_veri[149] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[149:152])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[149:152])/10  # 4
                if birlesik_veri[152] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[152:155])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[152:155])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 10:
                ruzgar_istikameti = int(birlesik_veri[158:160])*100  # 3
                ruzgar_hizi = int(birlesik_veri[160:162])  # 3
                if birlesik_veri[162] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[162:165])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[162:165])/10  # 4
                if birlesik_veri[165] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[165:168])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[165:168])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)
            # TAMAMLAYICI MESAFEYİ BULUYORUZ.
            if hedefin_toptan_yuksekligi == -400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[1])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[2])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[3])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[4])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -0:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[5])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[6])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[7])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[8])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[9])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 500:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[10])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 600:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[11])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 700:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 800:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 900:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[13])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 1000:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[14])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")
            bt_metro_ist_rakim_farki = batarya_rakimi - int(metro_istasyonu_rakimi)

            giris_mesafesi = round((int(mesafe) + int(tamamlayici_mesafe))/100)*100
            print("Giriş Mesafesi :", giris_mesafesi)

            ruzgarin_plan_istikameti_100 = ruzgar_istikameti - atis_istikameti

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            isi_dzl = [
                (-390,0.9),
                (-380,0.9),
                (-370,0.9),
                (-360,0.8),
                (-350,0.8),
                (-340,0.8),
                (-330,0.8),
                (-320,0.7),
                (-310,0.7),
                (-300,0.7),
                (-290,0.7),
                (-280,0.7),
                (-270,0.7),
                (-260,0.6),
                (-250,0.6),
                (-240,0.6),
                (-230,0.6),
                (-220,0.5),
                (-210,0.5),
                (-200,0.5),
                (-190,0.4),
                (-180,0.4),
                (-170,0.4),
                (-160,0.3),
                (-150,0.3),
                (-140,0.3),
                (-130,0.3),
                (-120,0.2),
                (-110,0.2),
                (-100,0.2),
                (-90,0.2),
                (-80,0.2),
                (-70,0.2),
                (-60,0.1),
                (-50,0.1),
                (-40,0.1),
                (-30,0.1),
                (-20,0),
                (-10,0),
                (0,0),
                (10,0),
                (20,0),
                (30,-0.1),
                (40,-0.1),
                (50,-0.1),
                (60,-0.1),
                (70,-0.2),
                (80,-0.2),
                (90,-0.2),
                (100,-0.2),
                (110,-0.2),
                (120,-0.2),
                (130,-0.3),
                (140,-0.3),
                (150,-0.3),
                (160,-0.3),
                (170,-0.4),
                (180,-0.4),
                (190,-0.4),
                (200,-0.5),
                (210,-0.5),
                (220,-0.5),
                (230,-0.6),
                (240,-0.6),
                (250,-0.6),
                (260,-0.6),
                (270,-0.7),
                (280,-0.7),
                (290,-0.7),
                (300,-0.7),
                (310,-0.7),
                (320,-0.7),
                (330,-0.8),
                (340,-0.8),
                (350,-0.8),
                (360,-0.8),
                (370,-0.9),
                (380,-0.9),
                (390,-0.9),
            ]


            isi_data5bh = []
            for i in range(len(isi_dzl)-1):
                start = isi_dzl[i]
                end = isi_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    isi_data5bh.append((int(x), y))

            isi_data5bh.append(isi_dzl[-1])


            yogunluk_dzl = [
                (-390,3.9),
                (-380,3.8),
                (-370,3.7),
                (-360,3.6),
                (-350,3.5),
                (-340,3.4),
                (-330,3.3),
                (-320,3.2),
                (-310,3.1),
                (-300,3),
                (-290,2.9),
                (-280,2.8),
                (-270,2.7),
                (-260,2.6),
                (-250,2.5),
                (-240,2.4),
                (-230,2.3),
                (-220,2.2),
                (-210,2.1),
                (-200,2),
                (-190,1.9),
                (-180,1.8),
                (-170,1.7),
                (-160,1.6),
                (-150,1.5),
                (-140,1.4),
                (-130,1.3),
                (-120,1.2),
                (-110,1.1),
                (-100,1),
                (-90,0.9),
                (-80,0.8),
                (-70,0.7),
                (-60,0.6),
                (-50,0.5),
                (-40,0.4),
                (-30,0.3),
                (-20,0.2),
                (-10,0.1),
                (0,0),
                (10,-0.1),
                (20,-0.2),
                (30,-0.3),
                (40,-0.4),
                (50,-0.5),
                (60,-0.6),
                (70,-0.7),
                (80,-0.8),
                (90,-0.9),
                (100,-1),
                (110,-1.1),
                (120,-1.2),
                (130,-1.3),
                (140,-1.4),
                (150,-1.5),
                (160,-1.6),
                (170,-1.7),
                (180,-1.8),
                (190,-1.9),
                (200,-2),
                (210,-2.1),
                (220,-2.2),
                (230,-2.3),
                (240,-2.4),
                (250,-2.5),
                (260,-2.6),
                (270,-2.7),
                (280,-2.8),
                (290,-2.9),
                (300,-3),
                (310,-3.1),
                (320,-3.2),
                (330,-3.3),
                (340,-3.4),
                (350,-3.5),
                (360,-3.6),
                (370,-3.7),
                (380,-3.8),
                (390,-3.9),
            ]

            yogunluk_data5bh = []

            for i in range(len(yogunluk_dzl)-1):
                start = yogunluk_dzl[i]
                end = yogunluk_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    yogunluk_data5bh.append((int(x), y))

            yogunluk_data5bh.append(yogunluk_dzl[-1])
            # print(yogunluk_data5bh)


            for i in isi_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    isi_duzeltmesi = i[1]
                    print("Isı Düzeltmesi :", isi_duzeltmesi)


            for i in yogunluk_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    yogunluk_duzeltmesi = i[1]
                    print("Yoğunluk Düzeltmesi", yogunluk_duzeltmesi)

            duzeltilmis_sıcaklik_degeri = isi_duzeltmesi + hava_sicakligi
            print("Düzeltilmiş Isı Değeri:", duzeltilmis_sıcaklik_degeri)

            duzeltilmis_yogunluk_degeri = round(yogunluk_duzeltmesi + hava_yogunlugu,1)
            print("Düzeltilmiş Yoğunluk Değeri:", duzeltilmis_yogunluk_degeri)
        # RUZGAR BİLEŞENLERİ

            atis_istikameti = round(atis_istikameti/100)*100

            if ruzgar_istikameti < atis_istikameti:
                ruzgar_istikameti = ruzgar_istikameti + 6400
            else:
                ruzgar_istikameti = ruzgar_istikameti
            print("ruzgar istikameti", ruzgar_istikameti)

            print("atış istikameti", atis_istikameti)

            ruzgarin_plan_istikameti = ruzgar_istikameti - atis_istikameti

            print("Rüzgarın Plan İstikameti:", ruzgarin_plan_istikameti)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            data1 = [
                    (0,0,1),
                (100,-0.1,0.99),
                (200,-0.2,0.98),
                (300,-0.29,0.96),
                (400,-0.38,0.92),
                (500,-0.47,0.88),
                (600,-0.56,0.83),
                (700,-0.63,0.77),
                (800,-0.71,0.71),
                (900,-0.77,0.63),
                (1000,-0.83,0.56),
                (1100,-0.88,0.47),
                (1200,-0.92,0.38),
                (1300,-0.96,0.29),
                (1400,-0.98,0.2),
                (1500,-0.99,0.1),
                (1600,-1,0),
                (1700,-0.99,-0.1),
                (1800,-0.98,-0.2),
                (1900,-0.96,-0.29),
                (2000,-0.92,-0.38),
                (2100,-0.88,-0.47),
                (2200,-0.83,-0.56),
                (2300,-0.77,-0.63),
                (2400,-0.71,-0.71),
                (2500,-0.63,-0.77),
                (2600,-0.56,-0.83),
                (2700,-0.47,-0.88),
                (2800,-0.38,-0.92),
                (2900,-0.29,-0.96),
                (3000,-0.2,-0.98),
                (3100,-0.1,-0.99),
                (3200,0,-1),
                (3300,0.1,-0.99),
                (3400,0.2,-0.98),
                (3500,0.29,-0.96),
                (3600,0.38,-0.92),
                (3700,0.47,-0.88),
                (3800,0.56,-0.83),
                (3900,0.63,-0.77),
                (4000,0.71,-0.71),
                (4100,0.77,-0.63),
                (4200,0.83,-0.56),
                (4300,0.88,-0.47),
                (4400,0.92,-0.38),
                (4500,0.96,-0.29),
                (4600,0.98,-0.2),
                (4700,0.99,-0.1),
                (4800,1,0),
                (4900,0.99,0.1),
                (5000,0.98,0.2),
                (5100,0.96,0.29),
                (5200,0.92,0.38),
                (5300,0.88,0.47),
                (5400,0.83,0.56),
                (5500,0.77,0.63),
                (5600,0.71,0.71),
                (5700,0.63,0.77),
                (5800,0.56,0.83),
                (5900,0.47,0.88),
                (6000,0.38,0.92),
                (6100,0.29,0.96),
                (6200,0.2,0.98),
                (6300,0.1,0.99),
                (6400,0,1),


            ]


            # veri_noktasi[0] = MESAFE
            # veri_noktasi[1] = YAN RÜZGARI
            # veri_noktasi[2] = MESAFE RÜZGARI

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Mesafe :", mesafe)

            for rzdlz in interpolasyonlu_veri1:
                if rzdlz[0] == ruzgarin_plan_istikameti:
                    yan_ruzgari_bileseni = rzdlz[1]
                    mesafe_ruzgari_bileseni = rzdlz[2]
                    print("Yan Rüzgarı Bileşeni:", yan_ruzgari_bileseni)
                    print("mesafe Rüzgarı Bileşeni:", mesafe_ruzgari_bileseni)


            # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (2500,112.8,7.8,0.27,7.7,1.6,0.19,7.9,-8,2.9,-3.1,7,-8.1,-2.9,2.8,-11,12),
                (2600,118,8.1,0.26,8.1,1.7,0.19,8,-8.2,3.2,-3.3,7.6,-8.6,-3,2.9,-11,12),
                (2700,123.2,8.4,0.25,8.4,1.8,0.2,8.2,-8.4,3.4,-3.5,8.1,-9.1,-3.2,3.1,-11,12),
                (2800,128.4,8.8,0.24,8.8,1.9,0.2,8.4,-8.5,3.6,-3.7,8.7,-9.6,-3.4,3.2,-11,12),
                (2900,133.7,9.1,0.23,9.1,2,0.21,8.5,-8.7,3.8,-3.9,9.3,-10.1,-3.5,3.4,-11,12),
                (3000,139,9.5,0.22,9.4,2,0.21,8.7,-8.8,4.1,-4.1,9.9,-10.6,-3.7,3.6,-11,12),
                (3100,144.4,9.8,0.21,9.8,2.1,0.22,88,-9,4.3,-4.3,10.5,-11.1,-3.9,3.8,-11,12),
                (3200,149.8,10.2,0.2,10.1,2.2,0.22,9,-9.1,4.5,-4.5,11,-11.6,-4.1,3.9,-11,12),
                (3300,155.3,10.5,0.2,10.5,2.3,0.23,9.2,-9.3,4.8,-4.7,11.6,-12.1,-4.3,4.1,-11,12),
                (3400,160.8,10.9,0.19,10.8,2.4,0.23,9.3,-9.4,5,-4.9,12.2,-12.6,-4.5,4.3,-11,12),
                (3500,166.4,11.2,0.19,11.2,2.5,0.23,9.5,-9.6,5.3,-5,12.8,-13.1,-4.6,4.5,-11,12),
                (3600,172,11.6,0.18,11.6,2.6,0.24,9.6,-9.7,5.5,-5.2,13.4,-13.6,-4.8,4.7,-11,12),
                (3700,177.7,12,0.17,11.9,2.7,0.24,9.8,-9.9,5.7,-5.4,13.9,-14.1,-5,4.9,-11,12),
                (3800,183.4,12.3,0.17,12.3,2.8,0.25,9.9,-10,6,-5.6,14.5,-14.6,-5.3,5.1,-11,12),
                (3900,189.2,12.7,0.16,12.7,2.9,0.25,10.1,-10.1,6.2,-5.8,15.1,-15,-5.5,5.3,-11,12),
                (4000,195,13.1,0.16,13,3,0.25,10.2,-10.3,6.5,-6,15.6,-15.5,-5.7,5.6,-11,12),
                (4100,200.9,13.4,0.16,13.4,3.1,0.26,10.4,-10.4,6.7,-6.2,16.2,-16,-5.9,5.8,-11,12),
                (4200,206.8,13.8,0.15,13.8,3.2,0.26,10.5,-10.5,6.9,-6.4,16.8,-16.5,-6.1,6,-11,11),
                (4300,212.8,14.2,0.15,14.1,3.3,0.27,10.7,-10.7,7.2,-6.6,17.3,-16.9,-6.4,6.3,-11,11),
                (4400,218.9,14.6,0.14,14.5,3.4,0.27,10.8,-10.8,7.4,-6.8,17.9,-17.4,-6.6,6.5,-10,11),
                (4500,225,15,0.14,14.9,3.5,0.28,11,-11,7.7,-7,18.4,-17.8,-6.8,6.7,-10,11),
                (4600,231.2,15.4,0.14,15.3,3.6,0.28,11.2,-11.1,7.9,-7.2,18.9,-18.3,-7.1,7,-10,11),
                (4700,237.4,15.7,0.13,15.7,3.8,0.28,11.3,-11.2,8.1,-7.4,19.4,-18.7,-7.3,7.3,-10,11),
                (4800,243.7,16.1,0.13,16.1,3.9,0.29,11.5,-11.4,8.4,-7.6,20,-19.2,-7.6,7.5,-10,11),
                (4900,250.1,16.5,0.13,16.5,4,0.29,11.6,-11.5,8.6,-7.8,20.5,-19.6,-7.9,7.8,-10,11),
                (5000,256.6,16.9,0.12,16.9,4.1,0.3,11.8,-11.7,8.9,-8,21,-20,-8.1,8.1,-9,10),
                (5100,263.1,17.3,0.12,17.3,4.2,0.3,12,-11.8,9.1,-8.2,21.5,-20.4,-8.4,8.4,-9,10),
                (5200,269.7,17.8,0.12,17.7,4.4,0.3,12.1,-11.9,9.3,-8.3,21.9,-20.9,-8.7,8.6,-9,10),
                (5300,276.4,18.2,0.12,18.1,4.5,0.31,12.3,-12.1,9.6,-8.5,22.4,-21.3,-9,8.9,-9,10),
                (5400,283.2,18.6,0.11,18.5,4.6,0.31,12.5,-12.2,9.8,-8.7,22.9,-21.7,-9.2,9.2,-9,10),
                (5500,290.1,19,0.11,18.9,4.7,0.32,12.6,-12.4,10,-8.9,23.3,-22.1,-9.5,9.5,-8,10),
                (5600,297.1,19.4,0.11,19.3,4.9,0.32,12.8,-12.5,10.3,-9.1,23.8,-22.5,-9.8,9.9,-8,9),
                (5700,304.1,19.9,0.11,19.7,5,0.33,13,-12.7,10.5,-9.3,24.2,-22.8,-10.1,10.2,-8,9),
                (5800,311.3,20.3,0.1,20.2,5.2,0.33,13.1,-12.8,10.7,-9.5,24.7,-23.2,-10.5,10.5,-8,9),
                (5900,318.6,20.7,0.1,20.6,5.3,0.33,13.3,-13,11,-9.7,25.1,-23.6,-10.8,10.8,-7,9),
                (6000,326,21.2,0.1,21.1,5.4,0.34,13.5,-13.1,11.2,-9.8,25.5,-24,-11.1,11.2,-7,9),
                (6100,333.5,21.6,0.1,21.5,5.6,0.34,13.7,-13.3,11.4,-10,25.9,-24.3,-11.4,11.5,-7,8),
                (6200,341.1,22.1,0.1,22,5.8,0.35,13.8,-13.4,11.7,-10.2,26.3,-24.7,-11.8,11.9,-6,8),
                (6300,348.8,22.5,0.09,22.4,5.9,0.35,14,-13.6,11.9,-10.4,26.7,-25,-12.1,12.2,-6,8),
                (6400,356.7,23,0.09,22.9,6.1,0.36,14.2,-13.7,12.1,-10.6,27.1,-25.4,-12.4,12.6,-6,8),
                (6500,364.8,23.5,0.09,23.4,6.2,0.36,14.4,-13.9,12.3,-10.7,27.5,-25.7,-12.8,12.9,-6,7),
                (6600,372.9,24,0.09,23.8,6.4,0.37,14.6,-14.1,12.6,-10.9,27.8,-26,-13.2,13.3,-5,7),
                (6700,381.3,24.5,0.09,24.3,6.6,0.37,14.8,-14.2,128,-11.1,28.2,-26.3,-13.5,13.7,-5,7),
                (6800,389.8,25,0.09,24.8,6.8,0.38,14.9,-14.4,13,-11.3,28.5,-26.6,-13.9,14.1,-5,6),
                (6900,398.5,25.5,0.08,25.3,7,0.38,15.1,-14.5,13.2,-11.4,28.9,-27,-14.3,14.5,-4,6),
                (7000,407.4,26,0.08,25.8,7.2,0.39,15.3,-14.7,13.4,-11.6,29.2,-27.3,-14.6,14.9,-4,6),
                (7100,416.5,26.5,0.08,26.4,7.4,0.39,15.5,-14.9,13.6,-11.8,29.5,-27.5,-15,15.3,-3,6),
                (7200,425.9,27.1,0.08,26.9,7.6,0.4,15.7,-15.1,13.9,-12,29.8,-27.8,-15.4,15.7,-3,5),
                (7300,435.5,27.6,0.08,27.4,7.8,0.4,15.9,-15.2,14.1,-12.1,30.1,-28.1,-15.8,16.1,-3,5),
                (7400,445.3,28.2,0.08,28,8,0.41,16.1,-15.4,14.3,-12.3,30.4,-28.4,-16.2,16.6,-2,5),
                (7500,455.5,28.7,0.08,28.6,8.3,0.42,16.3,-15.6,14.5,-12.5,30.7,-28.6,-16.6,17,-2,4),
                (7600,465.9,29.3,0.07,29.1,8.5,0.42,16.6,-15.8,14.7,-12.6,30.9,-28.9,-17.1,17.5,-2,4),
                (7700,476.8,29.9,0.07,29.7,8.8,0.43,16.8,-15.9,14.9,-12.8,31.2,-29.2,-17.5,17.9,-1,4),
                (7800,488,30.6,0.07,30.4,9.1,0.44,17,-16.1,15.1,-13,31.4,-29.4,-17.9,18.4,-1,3),
                (7900,499.7,31.2,0.07,31,9.4,0.44,17.2,-16.3,15.3,-13.1,31.6,-29.6,-18.4,18.9,0,3),
                (8000,511.8,31.9,0.07,31.7,9.7,0.45,17.5,-16.5,15.5,-13.3,31.8,-29.9,-18.8,19.4,0,2),
                (8100,524.6,32.6,0.07,32.4,10,0.46,17.7,-16.7,15.7,-13.4,32,-30.1,-19.3,19.9,1,2),
                (8200,538,33.3,0.07,33.1,10.4,0.47,17.9,-16.9,15.9,-13.6,32.2,-30.3,-19.7,20.4,1,2),
                (8300,552.2,34.1,0.06,33.8,10.7,0.47,18.2,-17.1,15.9,-13.8,32.4,-30.5,-20.2,20.9,2,1),
                (8400,567.3,34.9,0.06,34.6,11.2,0.48,18.4,-17.3,15.9,-13.9,32.5,-30.7,-20.7,21.5,2,1),
                (8500,583.7,35.7,0.06,35.5,11.6,0.49,18.7,-17.5,15.9,-14.1,32.6,-30.9,-21.2,22,3,0),
                (8600,601.4,36.6,0.06,36.4,12.2,0.5,19,-17.7,15.9,-14.2,32.7,-31,-21.7,22.6,3,0),
                (8700,621.2,37.7,0.06,37.4,12.8,0.52,19.3,-18,15.9,-14.4,32.6,-31.2,-22.2,23.3,4,0),
                (8800,643.9,38.8,0.06,38.5,13.5,0.53,19.6,-18.2,15.9,-14.5,32.6,-31.3,-22.7,24,4,-1),
                (8900,671.1,40.2,0.05,39.9,14.4,0.55,19.6,-18.4,15.9,-14.6,32.6,-31.5,-23.3,24,5,-1),
                (9000,708.2,42,0.05,41.7,15.7,0.57,19.6,-18.7,15.9,-14.8,32.6,-31.6,-23.9,24,6,-2),

            ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == giris_mesafesi:
                    nisangah6bh = round(nsg[1], 3)
                    # print("Nişangah:", nisangah)

            for tsg in interpolasyonlu_veri1:
                if tsg[0] == giris_mesafesi:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            for ts in interpolasyonlu_veri1:
                if ts[0] == giris_mesafesi:
                    deltaTS = round(ts[3], 3)
                    # print("Delta Tapa Saniyesi:", deltaTS)


            for ucs in interpolasyonlu_veri1:
                if ucs[0] == giris_mesafesi:
                    ucussuresi = round(ucs[4], 3)
                    # print("Uçuş Süresi:", ucussuresi)

            for dgl in interpolasyonlu_veri1:
                if dgl[0] == giris_mesafesi:
                    dogalyandz = round(dgl[5], 3)
                    # print("Doğal Yan Düzeltmesi:", dogalyandz)

            for yrd in interpolasyonlu_veri1:
                if yrd[0] == giris_mesafesi:
                    yan_ruzgari_duzeltme_faktoru = round(yrd[6], 1)
                    # print("Yan Rüzgarı Düzeltmesi:", yanruzgaridz)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_eksilme = round(ihz[7], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimeksilme)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_martma = round(ihz[8], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimartma)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_bas = round(mrd[9], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (BAŞ):", mesaferuzgaridzl_bas)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_arka = round(mrd[10], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (ARKA):", mesaferuzgaridzl_arka)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_eksilme = round(hvs[11], 3)
                    # print("Hava Isısı Düzeltmesi (Eksilme):", havaisisi_eksilme)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_artma = round(hvs[12], 3)
                    # print("Hava Isısı Düzeltmesi (Artma):", havaisisi_artma)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_eksilme = round(hyd[13], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", havayogunlugu_eksilme)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_artma = round(hyd[14], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", havayogunlugu_artma)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_eksilme = round(mka[15], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", mermikareagirligi_eksilme)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_artma = round(mka[16], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", mermikareagirligi_artma)

            print("Yan Rüzgarı Düzeltme Faktörü:", yan_ruzgari_duzeltme_faktoru)

            mesafe_ruzgari = ruzgar_hizi * mesafe_ruzgari_bileseni
            print("mesafe ruzgari", mesafe_ruzgari)

            ruzgar_yan_duzeltmesi = round(
                (ruzgar_hizi * yan_ruzgari_bileseni * yan_ruzgari_duzeltme_faktoru), 1)
            print("Rüzgar Yan Düzeltmesi", ruzgar_yan_duzeltmesi)

            dogalyandz = round(dogalyandz, 1)
            print("Doğal Yan Düzeltmesi:", dogalyandz)

            atis_istikameti = round(atis_istikameti/400)*400

            # dünyanın dönmesi düzeltmesi


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (2000,0.3,0.3,0.3,0.3,0.3,0.3,0.3,0.3,0.3),
                (3000,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.5,0.5),
                (4000,0.6,0.6,0.6,0.6,0.6,0.6,0.6,0.6,0.6),
                (5000,0.7,0.7,0.7,0.7,0.8,0.8,0.8,0.8,0.8),
                (6000,0.8,0.8,0.9,0.9,1,1,1,1.1,1.1),
                (7000,1,1,1,1.1,1.2,1.2,1.3,1.3,1.3),
                (8000,1.1,1.1,1.2,1.3,1.4,1.5,1.6,1.7,1.7),
                (9000,1.3,1.3,1.4,1.6,1.8,2.1,2.3,2.4,2.4),
            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            if atis_istikameti == 00 or atis_istikameti == 6400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[1], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 6000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[2], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 5600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[3], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 5200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[4], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 4800:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[5], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2000 or atis_istikameti == 4400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[6], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2400 or atis_istikameti == 4000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[7], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2800 or atis_istikameti == 3600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[8], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 3200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[9], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (2000,0,-3,-5,-8,-10,-11,-13,-13,-14,0,3,5,8,10,11,13,14),
                (3000,0,-4,-8,-11,-14,-16,-18,-19,-20,0,4,8,11,14,16,19,20),
                (4000,0,-5,-10,-14,-18,-21,-23,-25,-25,0,5,10,14,18,21,25,25),
                (5000,0,-6,-11,-17,-21,-25,-28,-29,-30,0,6,11,17,21,25,29,30),
                (6000,0,-7,-13,-19,-24,-28,-32,-33,-34,0,7,13,19,24,28,33,34),
                (7000,0,-7,-14,-21,-26,-31,-34,-37,-37,0,7,14,21,26,31,37,37),
                (8000,0,-8,-15,-22,-27,-32,-36,-38,-39,0,8,15,22,27,32,38,39),
                (9000,0,-7,-14,-20,-26,-30,-33,-35,-36,0,7,14,20,26,30,35,36),




            ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)
            if atis_istikameti == 0 or atis_istikameti == 3200 or atis_istikameti == 6400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[1], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 200 or atis_istikameti == 3000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[2], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 2800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[3], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 600 or atis_istikameti == 2600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[4], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 2400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[5], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1000 or atis_istikameti == 2200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[6], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 2000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[7], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1400 or atis_istikameti == 1800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[8], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 1600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[9], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            if atis_istikameti == 3400 or atis_istikameti == 6200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[11], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3600 or atis_istikameti == 6000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[12], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3800 or atis_istikameti == 5800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[13], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4000 or atis_istikameti == 5600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[14], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4200 or atis_istikameti == 5400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[15], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4400 or atis_istikameti == 5200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[16], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4600 or atis_istikameti == 5000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[17], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4800 or atis_istikameti == 4800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[18], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            print("Dünyanın Dönmesi Düzeltmesi:", arzindonusuyanduzeltmesi)

            metro_yan_duzeltmesi = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            metro_yan_duzeltmesi_1 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi:", metro_yan_duzeltmesi)
            metro_yan_duzeltmesi = round(metro_yan_duzeltmesi/1)*1
            print("Metro Yan Düzeltmesi en yakın 1 milyem:", metro_yan_duzeltmesi)


            # METRO MESAFE DÜZELTMESİ
            mesafe_ruzgari_fark = round(mesafe_ruzgari-0, 1)
            if mesafe_ruzgari_fark > 0:
                mesafe_ruzgari_durum = "BAŞ"
            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_durum = "ARKA"
            else:
                mesafe_ruzgari_durum = 0
            print("Mesafe Ruzgari Fark :", mesafe_ruzgari_fark)
            hava_sicakligi_fark = round(duzeltilmis_sıcaklik_degeri - 100, 1)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_durum = "EKSİLME"
            if hava_sicakligi_fark > 0:
                hava_sicakligi_durum = "ARTMA"
            else:
                hava_sicakligi_durum = 0
            print("Hava Sıcaklığı Fark :", hava_sicakligi_fark)
            hava_yogunlugu_fark = round(duzeltilmis_yogunluk_degeri - 100, 1)
            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_durum = "EKSİLME"
            if hava_yogunlugu_fark > 0:
                hava_yogunlugu_durum = "ARTMA"
            else:
                hava_yogunlugu_durum = 0
            print("Hava Yoğunluğu Fark :", hava_yogunlugu_fark)
            mermi_kare_agirligi_fark = round(mermi_kare_agirligi-2, 1)
            if mermi_kare_agirligi_fark < 0:
                merkar_durum = "EKSİLME"
            if mermi_kare_agirligi_fark > 0:
                merkar_durum = "ARTMA"
            if mermi_kare_agirligi_fark == 0:
                merkar_durum = 0

            print("Mermi Kare Ağırlığı Fark :", mermi_kare_agirligi_fark)

            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_arka
            else:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_bas
            if mesafe_ruzgari_fark == 0:
                mesafe_ruzgari_duzeltme_birimi = 0

            print("Mesafe Rüzgarı Düzeltme Birimi:", mesafe_ruzgari_duzeltme_birimi)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_duzeltme_birimi = havaisisi_eksilme
            else:
                hava_sicakligi_duzeltme_birimi = havaisisi_artma
            if hava_sicakligi_fark == 0:
                hava_sicakligi_duzeltme_birimi = 0
            print("Hava Isısı Düzeltme Birimi:", hava_sicakligi_duzeltme_birimi)

            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_eksilme
            else:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_artma
            if hava_yogunlugu_fark == 0:
                hava_yogunlugu_duzeltme_birimi = 0
            print("Hava Yoğunluğu Düzeltme Birimi:", hava_yogunlugu_duzeltme_birimi)

            if mermi_kare_agirligi_fark < 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_eksilme
            elif mermi_kare_agirligi_fark > 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_artma
            else:
                mermi_kare_agirligi_duzeltme_birimi = 0
            mesafe_ruzgari_fark = abs(mesafe_ruzgari_fark)
            hava_sicakligi_fark = abs(hava_sicakligi_fark)
            hava_yogunlugu_fark = abs(hava_yogunlugu_fark)


            print("Mermi Kare Ağırlığı Düzeltme Birimi:",
                mermi_kare_agirligi_duzeltme_birimi)

            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi = round(metro_yan_duzeltmesi + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)
            gac_yan_duzeltmesi = round(toplam_yan_duzeltmesi - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            brt_isisi_ilk_hiz_dzl = [
                                (-40,-12.5),
                                (-30,-11.6),
                                (-20,-10.6),
                                (-10,-9.6),
                                (0,-8.6),
                                (10,-7.5),
                                (20,-6.4),
                                (30,-5.2),
                                (40,-4),
                                (50,-2.7),
                                (60,-1.4),
                                (70,0),
                                (80,1.4),
                                (90,2.9),
                                (100,4.4),
                                (110,5.9),
                                (120,7.5),
                                (130,9.1),]

            barut_isisi_dzl = []
            for i in range(len(brt_isisi_ilk_hiz_dzl)-1):
                start = brt_isisi_ilk_hiz_dzl[i]
                end = brt_isisi_ilk_hiz_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    barut_isisi_dzl.append((int(x), y))

            barut_isisi_dzl.append(brt_isisi_ilk_hiz_dzl[-1])


            for i in barut_isisi_dzl:
                if i[0] == barut_isisi:
                    barut_isisi_duzeltmesi = round(i[1], 1)
            # HIZ DEĞİŞİKLİĞİNİN HESAPLANMASI
            hiz_degisikligi = round(ilk_hiz_farki + mevzi_hiz_degisikligi, 1)
            deltaV_hiz_farki = round(hiz_degisikligi + barut_isisi_duzeltmesi, 1)
            print("Hız Değişikliği:", hiz_degisikligi)
            print("Barut Isısı Düzeltmesi:", barut_isisi_duzeltmesi)
            print("Delta V Hız Farkı:", deltaV_hiz_farki)

            if deltaV_hiz_farki > 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_martma
            elif deltaV_hiz_farki < 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_eksilme
            else:
                ilk_hiz_duzeltme_birimi = 0
            print("İlk Hız Düzeltme Birimi:", ilk_hiz_duzeltme_birimi)

            deltaV_mesafe_duzeltmesi = abs(
                round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            if deltaV_hiz_farki > 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))*-1
            elif deltaV_hiz_farki < 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            else:
                deltaV_mesafe_duzeltmesi = 0
            print("Delta V Mesafe Düzeltmesi:", deltaV_mesafe_duzeltmesi)

            #### TOPLAM MESAF DÜZELTMESİ ###
            toplam_mesafe_duzeltmesi6 = round(metro_mesafe_duzeltmesi + deltaV_mesafe_duzeltmesi)
            print("Toplam Mesafe Düzeltmesi:", toplam_mesafe_duzeltmesi6)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)

            # toplam yan düzeltmede 1 milyem fazla mesafe düzeltmesinde işaret + ve 335 çıktı
            # baiek de yan 4 sl mesafe -326 çıktı

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,0,0,0,0,0,0,0,0,0,0),
            (1,0,0,0,0,0,0,0,0,0,0),
            (2,-0.006,0.006,0,0,0,0,0,0,0.011,-0.011),
            (3,-0.008,0.008,0,0.001,0,0.001,0.001,-0.001,0.016,-0.016),
            (4,-0.01,0.011,0,0.001,0,0.002,0.001,-0.001,0.019,-0.019),
            (5,-0.012,0.012,0,0.002,-0.001,0.004,0.002,-0.002,0.022,-0.022),
            (6,-0.013,0.014,-0.001,0.002,-0.002,0.006,0.003,-0.002,0.024,-0.024),
            (7,-0.015,0.016,-0.001,0.003,-0.004,0.008,0.003,-0.003,0.026,-0.026),
            (8,-0.016,0.017,-0.002,0.004,-0.006,0.01,0.004,-0.004,0.028,-0.028),
            (9,-0.017,0.018,-0.003,0.005,-0.008,0.013,0.004,-0.004,0.029,-0.03),
            (10,-0.018,0.019,-0.003,0.006,-0.011,0.015,0.005,-0.005,0.031,-0.031),
            (11,-0.019,0.02,-0.004,0.006,-0.013,0.018,0.006,-0.005,0.032,-0.032),
            (12,-0.021,0.021,-0.005,0.007,-0.016,0.02,0.006,-0.006,0.033,-0.034),
            (13,-0.022,0.022,-0.006,0.008,-0.018,0.023,0.007,-0.007,0.031,-0.035),
            (14,-0.023,0.023,-0.007,0.009,-0.021,0.025,0.008,-0.007,0.035,-0.036),
            (15,-0.024,0.024,-0.007,0.01,-0.024,0.028,0.008,-0.008,0.036,-0.037),
            (16,-0.025,0.025,-0.008,0.011,-0.027,0.031,0.009,-0.009,0.037,-0.038),
            (17,-0.026,0.026,-0.009,0.012,-0.029,0.033,0.01,-0.009,0.0373,-0.039),
            (18,-0.027,0.027,-0.01,0.012,-0.032,0.035,0.011,-0.01,0.038,-0.04),
            (19,-0.028,0.028,-0.011,0.013,-0.035,0.038,0.012,-0.011,0.039,-0.04),
            (20,-0.029,0.029,-0.011,0.014,-0.037,0.04,0.012,-0.012,0.039,-0.041),
            (21,-0.03,0.031,-0.012,0.015,-0.04,0.042,0.0133,-0.013,0.04,-0.042),
            (22,-0.031,0.032,-0.013,0.015,-0.042,0.045,0.014,-0.013,0,-0.043),
            (23,-0.032,0.033,-0.014,0.016,-0.045,0.047,0.015,-0.014,0.041,-0.043),
            (24,-0.033,0.034,-0.014,0.017,-0.047,0.049,0.016,-0.015,0.042,-0.044),
            (25,-0.035,0.035,-0.015,0.017,-0.049,0.051,0.017,-0.016,0.042,-0.045),
            (26,-0.036,0.036,-0.016,0.018,-0.052,0.053,0.018,-0.017,0.043,-0.045),
            (27,-0.037,0.037,-0.017,0.019,-0.054,0.055,0.02,-0.018,0.043,-0.046),
            (28,-0.038,0.038,-0.017,0.019,-0.056,0.057,0.021,-0.02,0.043,-0.046),
            (29,-0.039,0.039,-0.018,0.02,-0.058,0.059,0.022,-0.021,0.044,-0.047),
            (30,-0.04,0.04,-0.018,0.02,-0.06,0.06,0.023,-0.022,0.044,-0.048),
            (31,-0.042,0.042,-0.019,0.021,-0.062,0.062,0.024,-0.023,0.045,-0.048),
            (32,-0.043,0.043,-0.02,0.021,-0.064,0.064,0.026,-0.024,0.045,-0.049),
            (33,-0.044,0.044,-0.02,0.022,-0.066,0.065,0.027,-0.025,0.046,-0.049),
            (34,-0.046,0.045,-0.021,0.022,-0.068,0.067,0.028,-0.027,0.046,-0.05),
            (35,-0.047,0.047,-0.021,0.022,-0.07,0.069,0.029,-0.028,0.046,-0.051),
            (36,-0.048,0.048,-0.021,0.023,-0.071,0.07,0.031,-0.029,0.047,-0.051),
            (37,-0.05,0.049,-0.022,0.023,-0.073,0.072,0.032,-0.03,0.047,-0.052),
            (38,-0.051,0.05,-0.022,0.023,-0.075,0.073,0.033,-0.032,0.048,-0.053),
            (39,-0.052,0.052,-0.023,0.024,-0.076,0.075,0.035,-0.033,0.048,-0.053),
            (40,-0.054,0.053,-0.023,0.024,-0.078,0.076,0.036,-0.034,0.049,-0.054),
            (41,-0.055,0.054,-0.023,0.024,-0.079,0.077,0.038,-0.036,0.049,-0.055),
            (42,-0.057,0.056,-0.023,0.024,-0.081,0.079,0.039,-0.037,0.05,-0.055),
            (43,-0.058,0.057,-0.024,0.024,-0.082,0.08,0.04,-0.038,0.05,-0.056),
            (44,-0.06,0.059,-0.024,0.024,-0.084,0.081,0.042,-0.04,0.051,-0.057),
            (45,-0.061,0.06,-0.024,0.024,-0.085,0.083,0.043,-0.041,0.051,-0.058),
            (46,-0.063,0.062,-0.024,0.024,-0.087,0.084,0.045,-0.042,0.052,-0.059),
            (47,-0.064,0.063,-0.024,0.024,-0.088,0.085,0.046,-0.044,0.053,-0.06),
            (48,-0.066,0.064,-0.024,0.024,-0.089,0.087,0.048,-0.045,0.054,-0.061),
            (49,-0.067,0.066,-0.024,0.024,-0.091,0.088,0.049,-0.046,0.054,-0.062),
            (50,-0.069,0.067,-0.024,0.024,-0.092,0.089,0.05,-0.048,0.055,-0.063),
            (51,-0.07,0.069,-0.024,0.024,-0.093,0.091,0.052,-0.049,0.056,-0.064),
            (52,-0.072,0.07,-0.024,0.023,-0.095,0.092,0.053,-0.05,0.057,-0.065),
            (53,-0.074,0.072,-0.024,0.023,-0.096,0.093,0.054,-0.052,0.059,-0.067),
            (54,-0.075,0.074,-0.023,0.023,-0.097,0.094,0.056,-0.053,0.06,-0.068),
            (55,-0.077,0.075,-0.023,0.022,-0.099,0.095,0.057,-0.054,0.061,-0.07),
            (56,-0.079,0.077,-0.023,0.022,-0.1,0.097,0.058,-0.055,0.063,-0.072),
            (57,-0.08,0.078,-0.022,0.021,-0.101,0.098,0.06,-0.057,0.065,-0.074),
            (58,-0.082,0.08,-0.022,0.02,-0.102,0.099,0.061,-0.058,0.0673,-0.076),
            (59,-0.084,0.082,-0.021,0.02,-0.103,0.1,0.062,-0.059,0.069,-0.078),
            (60,-0.086,0.084,-0.02,0.02,-0.104,0.101,0.064,-0.06,0.072,-0.082),
            (61,-0.087,0.086,-0.02,0.024,-0.105,0.102,0.065,-0.062,0.078,-0.086),
            (62,-0.09,0.088,-0.02,0.024,-0.104,0.101,0.068,-0.064,0.089,-0.097),
            (63,-0.094,0.091,-0.025,0.024,-0.098,0.097,0.079,-0.071,0.125,-0.131),
            (64,-0.097,0.103,-0.036,0.024,-0.098,0.087,0.079,-0.085,0.267,-0.195),








            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Tapa Saniyesi :", tapasaniyesi)

            if deltaV_hiz_farki < 0:
                for ihz_eksilme in interpolasyonlu_veri1:
                    if ihz_eksilme[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_eksilme[1], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Eksilme):", ilkhizda1mildegisim_eksilme)
            if deltaV_hiz_farki > 0:
                for ihz_artma in interpolasyonlu_veri1:
                    if ihz_artma[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_artma[2], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Artma):", ilkhizda1mildegisim_artma)
            if mesafe_ruzgari_durum == "BAŞ":
                for musrus_bas in interpolasyonlu_veri1:
                    if musrus_bas[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(musrus_bas[3], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Baş):", mesaferuzgaridegisimi_bas)
            if mesafe_ruzgari_durum == "ARKA":
                for mesruz_arka in interpolasyonlu_veri1:
                    if mesruz_arka[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(mesruz_arka[4], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Arka):", mesaferuzgaridegisimi_arka)
            if hava_sicakligi_durum == "EKSİLME":
                for hadese_eksilme in interpolasyonlu_veri1:
                    if hadese_eksilme[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_eksilme[5], 3)
                        # print("Hava Isısı Değişim Miktarı (Eksilme):", havaisisidegisim_eksilme)
            if hava_sicakligi_durum == "ARTMA":
                for hadese_artma in interpolasyonlu_veri1:
                    if hadese_artma[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_artma[6], 3)
                        # print("Hava Isısı Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == "EKSİLME":
                for hayog_eksilme in interpolasyonlu_veri1:
                    if hayog_eksilme[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_eksilme[7], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Eksilme):", havayogunlugudegisim_eksilme)
            if hava_yogunlugu_durum == "ARTMA":
                for hayog_artma in interpolasyonlu_veri1:
                    if hayog_artma[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_artma[8], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == 0:
                havayogunlugudegisim = 0
            if merkar_durum == "EKSİLME":
                for merkar_eksilme in interpolasyonlu_veri1:
                    if merkar_eksilme[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_eksilme[9], 3)
                        # print("Mermi Kare Ağırlığı Değişim Miktarı (Eksilme):", mermikareagirligidegisim_eksilme)
            if merkar_durum == "ARTMA":
                for merkar_artma in interpolasyonlu_veri1:
                    if merkar_artma[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_artma[10], 3)
            else:
                mermikareagirligidegisim = 0
                # print("Mermi Kare Ağırlığı Değişim Miktarı (Artma):", mermikareagirligidegisim_artma)
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            
            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

            # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            # Veri kümesi
            data1 = [
                (1000,58,14,0.003,-0.003),
                (1500,90,32,0.007,-0.007),
                (2000,124,59,0.014,-0.013),
                (2500,161,96,0.023,-0.022),
                (3000,200,142,0.035,-0.033),
                (3500,242,201,0.053,-0.049),
                (4000,288,273,0.076,-0.07),
                (4500,338,362,0.109,-0.098),
                (5000,393,471,0.156,-0.138),
                (5500,455,606,0.228,-0.196),
                (6000,526,777,0.351,-0.288),
                (6500,611,1007,0.616,-0.455),
                (7000,728,1365,0.616,-0.894),

            ]
            gac_mesafe = plan_mesafesi_1B_obüs + toplam_mesafe_duzeltmesi6
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            # Veri kümesi
            data1 = [
                (2500,112.8,7.8,0.27,7.7,1.6,0.19,7.9,-8,2.9,-3.1,7,-8.1,-2.9,2.8,-11,12),
                (2600,118,8.1,0.26,8.1,1.7,0.19,8,-8.2,3.2,-3.3,7.6,-8.6,-3,2.9,-11,12),
                (2700,123.2,8.4,0.25,8.4,1.8,0.2,8.2,-8.4,3.4,-3.5,8.1,-9.1,-3.2,3.1,-11,12),
                (2800,128.4,8.8,0.24,8.8,1.9,0.2,8.4,-8.5,3.6,-3.7,8.7,-9.6,-3.4,3.2,-11,12),
                (2900,133.7,9.1,0.23,9.1,2,0.21,8.5,-8.7,3.8,-3.9,9.3,-10.1,-3.5,3.4,-11,12),
                (3000,139,9.5,0.22,9.4,2,0.21,8.7,-8.8,4.1,-4.1,9.9,-10.6,-3.7,3.6,-11,12),
                (3100,144.4,9.8,0.21,9.8,2.1,0.22,88,-9,4.3,-4.3,10.5,-11.1,-3.9,3.8,-11,12),
                (3200,149.8,10.2,0.2,10.1,2.2,0.22,9,-9.1,4.5,-4.5,11,-11.6,-4.1,3.9,-11,12),
                (3300,155.3,10.5,0.2,10.5,2.3,0.23,9.2,-9.3,4.8,-4.7,11.6,-12.1,-4.3,4.1,-11,12),
                (3400,160.8,10.9,0.19,10.8,2.4,0.23,9.3,-9.4,5,-4.9,12.2,-12.6,-4.5,4.3,-11,12),
                (3500,166.4,11.2,0.19,11.2,2.5,0.23,9.5,-9.6,5.3,-5,12.8,-13.1,-4.6,4.5,-11,12),
                (3600,172,11.6,0.18,11.6,2.6,0.24,9.6,-9.7,5.5,-5.2,13.4,-13.6,-4.8,4.7,-11,12),
                (3700,177.7,12,0.17,11.9,2.7,0.24,9.8,-9.9,5.7,-5.4,13.9,-14.1,-5,4.9,-11,12),
                (3800,183.4,12.3,0.17,12.3,2.8,0.25,9.9,-10,6,-5.6,14.5,-14.6,-5.3,5.1,-11,12),
                (3900,189.2,12.7,0.16,12.7,2.9,0.25,10.1,-10.1,6.2,-5.8,15.1,-15,-5.5,5.3,-11,12),
                (4000,195,13.1,0.16,13,3,0.25,10.2,-10.3,6.5,-6,15.6,-15.5,-5.7,5.6,-11,12),
                (4100,200.9,13.4,0.16,13.4,3.1,0.26,10.4,-10.4,6.7,-6.2,16.2,-16,-5.9,5.8,-11,12),
                (4200,206.8,13.8,0.15,13.8,3.2,0.26,10.5,-10.5,6.9,-6.4,16.8,-16.5,-6.1,6,-11,11),
                (4300,212.8,14.2,0.15,14.1,3.3,0.27,10.7,-10.7,7.2,-6.6,17.3,-16.9,-6.4,6.3,-11,11),
                (4400,218.9,14.6,0.14,14.5,3.4,0.27,10.8,-10.8,7.4,-6.8,17.9,-17.4,-6.6,6.5,-10,11),
                (4500,225,15,0.14,14.9,3.5,0.28,11,-11,7.7,-7,18.4,-17.8,-6.8,6.7,-10,11),
                (4600,231.2,15.4,0.14,15.3,3.6,0.28,11.2,-11.1,7.9,-7.2,18.9,-18.3,-7.1,7,-10,11),
                (4700,237.4,15.7,0.13,15.7,3.8,0.28,11.3,-11.2,8.1,-7.4,19.4,-18.7,-7.3,7.3,-10,11),
                (4800,243.7,16.1,0.13,16.1,3.9,0.29,11.5,-11.4,8.4,-7.6,20,-19.2,-7.6,7.5,-10,11),
                (4900,250.1,16.5,0.13,16.5,4,0.29,11.6,-11.5,8.6,-7.8,20.5,-19.6,-7.9,7.8,-10,11),
                (5000,256.6,16.9,0.12,16.9,4.1,0.3,11.8,-11.7,8.9,-8,21,-20,-8.1,8.1,-9,10),
                (5100,263.1,17.3,0.12,17.3,4.2,0.3,12,-11.8,9.1,-8.2,21.5,-20.4,-8.4,8.4,-9,10),
                (5200,269.7,17.8,0.12,17.7,4.4,0.3,12.1,-11.9,9.3,-8.3,21.9,-20.9,-8.7,8.6,-9,10),
                (5300,276.4,18.2,0.12,18.1,4.5,0.31,12.3,-12.1,9.6,-8.5,22.4,-21.3,-9,8.9,-9,10),
                (5400,283.2,18.6,0.11,18.5,4.6,0.31,12.5,-12.2,9.8,-8.7,22.9,-21.7,-9.2,9.2,-9,10),
                (5500,290.1,19,0.11,18.9,4.7,0.32,12.6,-12.4,10,-8.9,23.3,-22.1,-9.5,9.5,-8,10),
                (5600,297.1,19.4,0.11,19.3,4.9,0.32,12.8,-12.5,10.3,-9.1,23.8,-22.5,-9.8,9.9,-8,9),
                (5700,304.1,19.9,0.11,19.7,5,0.33,13,-12.7,10.5,-9.3,24.2,-22.8,-10.1,10.2,-8,9),
                (5800,311.3,20.3,0.1,20.2,5.2,0.33,13.1,-12.8,10.7,-9.5,24.7,-23.2,-10.5,10.5,-8,9),
                (5900,318.6,20.7,0.1,20.6,5.3,0.33,13.3,-13,11,-9.7,25.1,-23.6,-10.8,10.8,-7,9),
                (6000,326,21.2,0.1,21.1,5.4,0.34,13.5,-13.1,11.2,-9.8,25.5,-24,-11.1,11.2,-7,9),
                (6100,333.5,21.6,0.1,21.5,5.6,0.34,13.7,-13.3,11.4,-10,25.9,-24.3,-11.4,11.5,-7,8),
                (6200,341.1,22.1,0.1,22,5.8,0.35,13.8,-13.4,11.7,-10.2,26.3,-24.7,-11.8,11.9,-6,8),
                (6300,348.8,22.5,0.09,22.4,5.9,0.35,14,-13.6,11.9,-10.4,26.7,-25,-12.1,12.2,-6,8),
                (6400,356.7,23,0.09,22.9,6.1,0.36,14.2,-13.7,12.1,-10.6,27.1,-25.4,-12.4,12.6,-6,8),
                (6500,364.8,23.5,0.09,23.4,6.2,0.36,14.4,-13.9,12.3,-10.7,27.5,-25.7,-12.8,12.9,-6,7),
                (6600,372.9,24,0.09,23.8,6.4,0.37,14.6,-14.1,12.6,-10.9,27.8,-26,-13.2,13.3,-5,7),
                (6700,381.3,24.5,0.09,24.3,6.6,0.37,14.8,-14.2,128,-11.1,28.2,-26.3,-13.5,13.7,-5,7),
                (6800,389.8,25,0.09,24.8,6.8,0.38,14.9,-14.4,13,-11.3,28.5,-26.6,-13.9,14.1,-5,6),
                (6900,398.5,25.5,0.08,25.3,7,0.38,15.1,-14.5,13.2,-11.4,28.9,-27,-14.3,14.5,-4,6),
                (7000,407.4,26,0.08,25.8,7.2,0.39,15.3,-14.7,13.4,-11.6,29.2,-27.3,-14.6,14.9,-4,6),
                (7100,416.5,26.5,0.08,26.4,7.4,0.39,15.5,-14.9,13.6,-11.8,29.5,-27.5,-15,15.3,-3,6),
                (7200,425.9,27.1,0.08,26.9,7.6,0.4,15.7,-15.1,13.9,-12,29.8,-27.8,-15.4,15.7,-3,5),
                (7300,435.5,27.6,0.08,27.4,7.8,0.4,15.9,-15.2,14.1,-12.1,30.1,-28.1,-15.8,16.1,-3,5),
                (7400,445.3,28.2,0.08,28,8,0.41,16.1,-15.4,14.3,-12.3,30.4,-28.4,-16.2,16.6,-2,5),
                (7500,455.5,28.7,0.08,28.6,8.3,0.42,16.3,-15.6,14.5,-12.5,30.7,-28.6,-16.6,17,-2,4),
                (7600,465.9,29.3,0.07,29.1,8.5,0.42,16.6,-15.8,14.7,-12.6,30.9,-28.9,-17.1,17.5,-2,4),
                (7700,476.8,29.9,0.07,29.7,8.8,0.43,16.8,-15.9,14.9,-12.8,31.2,-29.2,-17.5,17.9,-1,4),
                (7800,488,30.6,0.07,30.4,9.1,0.44,17,-16.1,15.1,-13,31.4,-29.4,-17.9,18.4,-1,3),
                (7900,499.7,31.2,0.07,31,9.4,0.44,17.2,-16.3,15.3,-13.1,31.6,-29.6,-18.4,18.9,0,3),
                (8000,511.8,31.9,0.07,31.7,9.7,0.45,17.5,-16.5,15.5,-13.3,31.8,-29.9,-18.8,19.4,0,2),
                (8100,524.6,32.6,0.07,32.4,10,0.46,17.7,-16.7,15.7,-13.4,32,-30.1,-19.3,19.9,1,2),
                (8200,538,33.3,0.07,33.1,10.4,0.47,17.9,-16.9,15.9,-13.6,32.2,-30.3,-19.7,20.4,1,2),
                (8300,552.2,34.1,0.06,33.8,10.7,0.47,18.2,-17.1,15.9,-13.8,32.4,-30.5,-20.2,20.9,2,1),
                (8400,567.3,34.9,0.06,34.6,11.2,0.48,18.4,-17.3,15.9,-13.9,32.5,-30.7,-20.7,21.5,2,1),
                (8500,583.7,35.7,0.06,35.5,11.6,0.49,18.7,-17.5,15.9,-14.1,32.6,-30.9,-21.2,22,3,0),
                (8600,601.4,36.6,0.06,36.4,12.2,0.5,19,-17.7,15.9,-14.2,32.7,-31,-21.7,22.6,3,0),
                (8700,621.2,37.7,0.06,37.4,12.8,0.52,19.3,-18,15.9,-14.4,32.6,-31.2,-22.2,23.3,4,0),
                (8800,643.9,38.8,0.06,38.5,13.5,0.53,19.6,-18.2,15.9,-14.5,32.6,-31.3,-22.7,24,4,-1),
                (8900,671.1,40.2,0.05,39.9,14.4,0.55,19.6,-18.4,15.9,-14.6,32.6,-31.5,-23.3,24,5,-1),
                (9000,708.2,42,0.05,41.7,15.7,0.57,19.6,-18.7,15.9,-14.8,32.6,-31.6,-23.9,24,6,-2),]


            
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == gac_mesafe:
                    nisangah6bh = round(nsg[1], 1)
                    print("Nişangah:", nisangah6bh)
            for dgl in interpolasyonlu_veri1:
                if dgl[0] == gac_mesafe:
                    dogalyandz = round(dgl[5],3)
                    print("Doğal Yan Düzeltmesi:", dogalyandz)
            for tsg in interpolasyonlu_veri1:
                if tsg[0] == gac_mesafe:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            metro_yan_duzeltmesi2 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi en yakın 1 milyem99999999999999:", metro_yan_duzeltmesi2)
            print("Arzin Dnüşü:,",arzindonusuyanduzeltmesi)
            print("Rüzgaryan dzl---------------------:,",ruzgar_yan_duzeltmesi)

            
            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi6bh = round(metro_yan_duzeltmesi2 + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi3333333333333:", toplam_yan_duzeltmesi6bh)
            gac_yan_duzeltmesi2 = round(toplam_yan_duzeltmesi6bh - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi2)
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == gac_mesafe:
                    dusus_acisi = round(dusus_acisi[1],3)
                    print("Düşüş Açısı:", dusus_acisi)

            for tepe_yuksekligi in interpolasyonlu_veri1:
                if tepe_yuksekligi[0] == gac_mesafe:
                    tepe_yuksekligi = round(tepe_yuksekligi[2],3)
                    print("Tepe Yüksekliği:", tepe_yuksekligi)

            for dtac_arti1 in interpolasyonlu_veri1:
                if dtac_arti1[0] == mesafe:
                    dtac_arti1 = round(dtac_arti1[3],3)
                    print("+1 DTAÇ:", dtac_arti1)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4],3)
                    print("-1 DTAÇ:", dtac_eksi1)
    ############################################################################################################
            """METRO DÜZELTMELERİNİ HESAPLAMAK GEREKİYOR"""
            # GTAC 1 inci oonüs için
            bt_rakimi = lne_1B_obus_rakim

            milyem_sabit_dz = 1.0186

            hedef_batarya_rakim_farki = (rakim - bt_rakimi)

            dtac = (hedef_batarya_rakim_farki / (plan_mesafesi_1B_obüs / 1000)) * milyem_sabit_dz

            # Burada dtac_arti1 ve dtac_eksi1 değerlerini sadece bir kez tanımlayın
            dtac_arti1 = 0
            dtac_eksi1 = 0

            # dtac_arti1 ve dtac_eksi1 değerlerini doğru şekilde ata
            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_arti1 = round(dusus_acisi[3], 3)
                    print("dtac_arti1",dtac_arti1)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4], 3)
                    print("dtac_eksi1",dtac_eksi1)


            ttac = 0
            print("Mesafe",mesafe)
            if dtac < 0:
                ttac = int(dtac * dtac_arti1)
                tac = dtac + ttac
                #print("-TAÇ", tac)

            if dtac > 0:
                ttac = int(dtac * dtac_eksi1)
                tac = dtac + ttac
                print("+TAÇ", tac)
    #########################################################################################################
            """YÜKSELİŞİ BUL"""
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,26),
(100,26),
(200,26),
(300,25),
(400,25),
(500,24),
(600,24),
(700,24),
(800,23),
(900,23),
(1000,23),
(1100,22),
(1200,22),
(1300,22),
(1400,22),
(1500,21),
(1600,21),
(1700,21),
(1800,21),
(1900,21),
(2000,20),
(2100,20),
(2200,20),
(2300,20),
(2400,20),
(2500,19),
(2600,19),
(2700,19),
(2800,19),
(2900,19),
(3000,19),
(3100,19),
(3200,18),
(3300,18),
(3400,18),
(3500,18),
(3600,18),
(3700,18),
(3800,17),
(3900,17),
(4000,17),
(4100,17),
(4200,17),
(4300,17),
(4400,16),
(4500,16),
(4600,16),
(4700,16),
(4800,16),
(4900,16),
(5000,15),
(5100,15),
(5200,15),
(5300,15),
(5400,15),
(5500,14),
(5600,14),
(5700,14),
(5800,14),
(5900,14),
(6000,13),
(6100,13),
(6200,13),
(6300,13),
(6400,13),
(6500,12),
(6600,12),
(6700,12),
(6800,12),
(6900,11),
(7000,11),
(7100,11),
(7200,11),
(7300,10),
(7400,10),
(7500,10),
(7600,9),
(7700,9),
(7800,9),
(7900,8),
(8000,8),
(8100,8),
(8200,7),
(8300,7),
(8400,6),
(8500,6),
(8600,5),
(8700,5),
(8800,4),
(8900,3),
(9000,3),

            ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)
            
            for nsg in interpolasyonlu_veri1:
                    if nsg[0] == gac_mesafe:
                        birmildegisim6bh = round(nsg[1], 1)
                        print("1 Milyemlik Değişim:", birmildegisim6bh)


            yukselis6bh = round(nisangah6bh+ + tac,1)
            self.ui.sonuc_yukselis_1B.setText(str(yukselis6bh))
    #########################################################################################################
            yan_1B_6bh = round(yan_1B+toplam_yan_duzeltmesi6bh)
            self.ui.sonuc_yan_1B.setText(str(yan_1B_6bh))
            self.ui.sonuc_yan_2B.setText(str(yan2))
            self.ui.lne_barut_hakki.setText(str(secilen_barut_hakki))
            self.ui.sonuc_barut_hakki_1B.setText(str(secilen_barut_hakki))
            istikamet_acisi = atis_istikameti_1

            if 2200 < yan_1B_6bh < 3000:
                self.ui.sonuc_yan_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_yukselis_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_barut_hakki_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_tapa_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_istikamet_acisi_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_mesafe_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.lne_sonuc_mesafe.setText(str(mesafe))
                self.ui.lne_sonuc_barut_hakki.setText(str(secilen_barut_hakki))
                self.ui.lne_sonuc_yan.setText(str(yan_1B_6bh))
                self.ui.lne_sonuc_yukselis.setText(str(yukselis6bh))
                self.ui.lne_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
                self.ui.lne_sonuc_bir_ml_degisiklil.setText(str(birmildegisim6bh)) 
                self.ui.lne_sonuc_yuz_m.setText(str(mesafe/1000)) 
            else:
                self.ui.sonuc_yan_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_yukselis_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_barut_hakki_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_tapa_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_istikamet_acisi_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_mesafe_1B.setStyleSheet("color: grey; font-weight: normal")
            self._doldur_sonuc("1", toplam_yan_duzeltmesi6bh, yukselis6bh, birmildegisim6bh, locals())
##########################################################################################################
        if secilen_barut_hakki == 7:
            plan_mesafesi = plan_mesafesi_1B_obüs
            batarya_rakimi = lne_1B_obus_rakim
            bt_rakimi_1 = lne_1B_obus_rakim
            hedef_rakimi = lne_hedef_rakim
            ahia = lne_1B_obus_ahia
            paralanma_yuksekliği = paralanma_yuksekliği
            mermi_kare_agirligi = mermi_kare_agirligi
            ilk_hiz_farki = lne_1B_obus_ihf
            atis_istikameti = batarya_hedef_İA_1
            atis_istikameti_1 = batarya_hedef_İA_1
            mevzi_hiz_degisikligi = 0
            mevzi_yan_duzeltmesi = 0
            mevzi_ts_duzeltmesi = 0
            tapasaniyesi = 1
            barut_isisi_str = self.ui.lne_barut_isisi.text()
            barut_isisi = float(barut_isisi_str) if barut_isisi_str else 75
            # barut_isisi = self.ui.lne_barut_isisi.text() if self.ui.lne_barut_isisi.text() else 75
            print("Barutttt İSİSİ",barut_isisi)
            
            # 1. GİRİŞ MESAFESİNİ BUL / MANUEL HASSASİYETİ

            # mesafe = 1866
            # en_yakin_10_metre = round(mesafe / 10) * 10
            # print("1836'nın en yakın 10 metreye çevrilmiş hali:", en_yakin_10_metre)

            # BATARYA RAKIMINI EN YAKIN 10 M YE ÇEVİR.
            batarya_rakimi = round(batarya_rakimi/10)*10
            # PARALANMA NOKTASININ RAKIMINI BUL.
            paralanma_noktasi_rakimi = hedef_rakimi + paralanma_yuksekliği
            # HEDEF-BATARYA YÜKSEKLİK FARKINI BUL
            yukseklik_farki = paralanma_noktasi_rakimi-batarya_rakimi
            yukseklik_farki_1 = paralanma_noktasi_rakimi-bt_rakimi_1
            # yukseklik_farki = round(yukseklik_farki/100)*100
            ################# B cetvelini yüklüyoruz.################
            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4100,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4200,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4300,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4400,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4500,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4600,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4700,0,0,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4800,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4900,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(5000,0,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(5100,1,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(5200,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5300,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5400,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5500,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5600,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5700,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5800,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5900,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(6000,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(6100,1,1,1,2,2,2,2,2,2,3,3,3,3,3,3),
(6200,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(6300,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(6400,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(6500,1,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6600,1,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6700,2,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6800,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(6900,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(7000,2,2,2,2,2,2,3,3,3,3,3,3,3,3,4),
(7100,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(7200,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(7300,2,2,2,2,2,3,3,3,3,3,3,3,3,4,4),
(7400,2,2,2,2,2,3,3,3,3,3,3,3,4,4,4),
(7500,2,2,2,2,3,3,3,3,3,3,3,3,4,4,4),
(7600,2,2,2,2,3,3,3,3,3,3,3,3,4,4,4),
(7700,2,2,2,3,3,3,3,3,3,3,3,4,4,4,4),
(7800,2,2,2,3,3,3,3,3,3,3,3,4,4,4,4),
(7900,2,2,3,3,3,3,3,3,3,3,4,4,4,4,4),
(8000,2,2,3,3,3,3,3,3,3,3,4,4,4,4,4),
(8100,2,3,3,3,3,3,3,3,3,4,4,4,4,4,4),
(8200,2,3,3,3,3,3,3,3,3,4,4,4,4,4,4),
(8300,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4),
(8400,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4),
(8500,3,3,3,3,3,3,3,4,4,4,4,4,4,4,5),
(8600,3,3,3,3,3,3,3,4,4,4,4,4,4,4,5),
(8700,3,3,3,3,3,3,4,4,4,4,4,4,4,5,5),
(8800,3,3,3,3,3,3,4,4,4,4,4,4,4,5,5),
(8900,3,3,3,3,3,4,4,4,4,4,4,4,5,5,5),
(9000,3,3,3,3,4,4,4,4,4,4,4,5,5,5,5),
(9100,3,3,3,3,4,4,4,4,4,4,4,5,5,5,5),
(9200,3,3,3,4,4,4,4,4,4,4,5,5,5,5,5),
(9300,3,3,4,4,4,4,4,4,4,5,5,5,5,5,5),
(9400,3,3,4,4,4,4,4,4,4,5,5,5,5,5,5),
(9500,3,4,4,4,4,4,4,4,5,5,5,5,5,5,5),
(9600,4,4,4,4,4,4,4,5,5,5,5,5,5,5,5),
(9700,4,4,4,4,4,4,4,5,5,5,5,5,5,5,5),
(9800,4,4,4,4,4,4,5,5,5,5,5,5,5,5,6),
(9900,4,4,4,4,4,5,5,5,5,5,5,5,5,6,6),
(10000,4,4,4,4,5,5,5,5,5,5,5,5,6,6,6),
(10100,4,4,4,5,5,5,5,5,5,5,5,6,6,6,6),
(10200,4,4,4,5,5,5,5,5,5,5,6,6,6,6,6),
(10300,4,4,5,5,5,5,5,5,5,6,6,6,6,6,6),
(10400,4,5,5,5,5,5,5,5,6,6,6,6,6,6,6),
(10500,5,5,5,5,5,5,5,6,6,6,6,6,6,6,6),
(10600,5,5,5,5,5,5,6,6,6,6,6,6,6,6,6),
(10700,5,5,5,5,5,6,6,6,6,6,6,6,6,6,6),
(10800,5,5,5,5,6,6,6,6,6,6,6,6,6,6,6),
(10900,5,5,5,6,6,6,6,6,6,6,6,6,6,6,6),
(11000,5,5,6,6,6,6,6,6,6,6,6,6,6,6,6),
    ]

            data2 = [
(4000,-23,-18,-13,-7,0,7,16,25,34,45,56,67,80,93,107),
(4100,-24,-19,-13,-7,0,8,16,25,35,45,56,68,81,94,108),
(4200,-24,-19,-13,-7,0,8,16,25,35,46,57,69,82,95,109),
(4300,-25,-20,-14,-7,0,8,17,26,36,47,58,70,83,96,111),
(4400,-25,-20,-14,-7,0,8,17,26,36,47,59,71,84,98,112),
(4500,-26,-21,-14,-8,0,8,17,27,37,48,60,72,85,99,114),
(4600,-27,-21,-15,-8,0,8,17,27,38,49,61,73,87,101,115),
(4700,-28,-22,-15,-8,0,9,18,28,38,50,62,74,88,102,117),
(4800,-28,-22,-16,-8,0,9,18,28,39,51,63,76,89,104,119),
(4900,-29,-23,-16,-8,0,9,19,29,40,52,64,77,91,105,121),
(5000,-30,-23,-16,-8,0,9,19,29,41,53,65,78,92,107,123),
(5100,-31,-24,-17,-9,0,9,19,30,41,54,66,80,94,109,125),
(5200,-32,-25,-17,-9,0,10,20,31,42,55,68,81,96,111,127),
(5300,-33,-25,-18,-9,0,10,20,31,43,56,69,83,98,113,130),
(5400,-34,-26,-18,-9,0,10,21,32,44,57,71,85,100,116,132),
(5500,-35,-27,-19,-10,0,10,21,33,45,58,72,87,102,118,135),
(5600,-36,-28,-19,-10,0,11,22,34,46,60,74,89,104,121,138),
(5700,-37,-28,-20,-10,0,11,22,34,47,61,75,91,106,123,141),
(5800,-38,-29,-20,-10,0,11,23,35,48,62,77,93,109,126,144),
(5900,-39,-30,-21,-11,0,11,23,36,50,64,79,95,111,129,147),
(6000,-40,-31,-21,-11,0,12,24,37,51,66,81,97,114,132,150),
(6100,-41,-32,-22,-11,0,12,25,38,52,67,83,99,117,135,154),
(6200,-42,-33,-23,-12,0,12,25,39,54,69,85,102,120,138,158),
(6300,-44,-34,-23,-12,0,13,26,40,55,71,87,104,123,142,162),
(6400,-45,-35,-24,-12,0,13,27,41,56,72,89,107,126,145,166),
(6500,-46,-36,-25,-13,0,13,27,42,58,74,92,110,129,149,170),
(6600,-48,-37,-25,-13,0,14,28,43,60,76,94,113,132,153,174),
(6700,-49,-38,-26,-13,0,14,29,45,61,78,97,116,136,157,178),
(6800,-51,-39,-27,-14,0,15,30,46,63,81,99,119,139,161,183),
(6900,-52,-40,-28,-14,0,15,31,47,65,83,102,122,143,165,188),
(7000,-54,-42,-28,-15,0,15,32,49,66,85,105,125,147,169,193),
(7100,-56,-43,-29,-15,0,16,32,50,68,88,108,129,151,174,198),
(7200,-57,-44,-30,-15,0,16,33,51,70,90,111,132,155,179,204),
(7300,-59,-46,-31,-16,0,17,34,53,72,93,114,136,159,184,209),
(7400,-61,-47,-32,-16,0,17,35,54,74,95,117,140,164,189,215),
(7500,-63,-48,-33,-17,0,18,36,56,76,98,120,144,169,194,221),
(7600,-65,-50,-34,-17,0,18,37,58,79,101,124,148,173,200,228),
(7700,-67,-51,-35,-18,0,19,39,59,81,104,127,152,178,206,234),
(7800,-69,-53,-36,-19,0,19,40,61,83,107,131,157,184,212,241),
(7900,-71,-55,-37,-19,0,20,41,63,86,110,135,161,189,218,248),
(8000,-73,-56,-38,-20,0,21,42,65,88,113,139,166,195,224,255),
(8100,-76,-58,-40,-20,0,21,43,67,91,117,143,171,200,231,263),
(8200,-78,-60,-41,-21,0,22,45,69,94,120,148,176,206,238,271),
(8300,-80,-62,-42,-22,0,23,46,71,97,124,152,182,213,245,279),
(8400,-83,-64,-43,-22,0,23,48,73,100,128,157,187,219,253,288),
(8500,-86,-66,-45,-23,0,24,49,75,103,131,162,193,226,261,297),
(8600,-88,-68,-46,-24,0,25,51,78,106,136,167,199,233,269,307),
(8700,-91,-70,-48,-24,0,25,52,80,109,140,172,206,241,278,317),
(8800,-94,-72,-49,-25,0,26,54,83,113,144,177,212,249,287,328),
(8900,-97,-74,-51,-26,0,27,55,85,116,149,183,219,257,297,339),
(9000,-100,-77,-52,-27,0,28,57,88,120,154,189,227,266,307,351),
(9100,-103,-79,-54,-28,0,29,59,91,124,159,196,234,275,318,363),
(9200,-106,-82,-56,-28,0,30,61,94,128,164,202,2343,285,330,377),
(9300,-110,-84,-57,-29,0,31,63,97,133,170,210,251,295,342,391),
(9400,-113,-87,-59,-30,0,32,65,100,137,176,217,261,306,355,407),
(9500,-117,-90,-61,-31,0,33,68,104,142,183,225,270,318,369,424),
(9600,-121,-93,-63,-32,0,34,70,108,147,189,234,281,331,385,443),
(9700,-125,-96,-66,-34,0,35,72,112,153,197,243,293,346,402,464),
(9800,-129,-99,-68,-35,0,37,75,116,159,205,253,305,361,422,487),
(9900,-134,-103,-70,-36,0,38,78,120,165,213,264,319,379,443,515),
(10000,-139,-107,-73,-37,0,39,81,125,172,223,277,335,399,469,549),
(10100,-144,-111,-76,-39,0,41,85,131,180,233,290,353,422,501,574),
(10200,-149,-115,-79,-40,0,43,88,137,189,245,306,375,452,544,675),
(10300,-155,-119,-82,-42,0,45,92,143,199,259,326,402,493,644,675),
(10400,-161,-124,-85,-44,0,47,97,151,210,276,351,442,493,644,675),
(10500,-167,-129,-89,-46,0,49,102,160,224,298,389,442,493,644,675),
(10600,-175,-135,-93,-48,0,52,108,171,244,336,389,442,493,644,675),
(10700,-183,-141,-97,-51,0,55,116,187,287,336,389,442,493,644,675),
(10800,-191,-148,-103,-54,0,59,129,187,287,336,389,442,493,644,675),
(10900,-201,-157,-109,-57,0,67,129,187,287,336,389,442,493,644,675),
(11000,-213,-167,-117,-63,0,67,129,187,287,336,389,442,493,644,675),



]


            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            interpolasyonlu_veri2 = veri_kumesi_interpolasyonu(data2)
            # print(interpolasyonlu_veri2)

            hedefin_toptan_yuksekligi = yukseklik_farki
            hedefin_toptan_yuksekligi = round(hedefin_toptan_yuksekligi/100)*100
            print("Hedefin Toptan Yüksekliği: ", hedefin_toptan_yuksekligi)
            mesafe = plan_mesafesi
            print("mesafe = ", mesafe)

            if hedefin_toptan_yuksekligi == -400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[1])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[2])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[3])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[4])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -0:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[5])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[6])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[7])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[8])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[9])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 500:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[10])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 600:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[11])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 700:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 800:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 900:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[13])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 1000:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[14])
                        print("Seçilen Bölge:", bolge)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")


            # TXT dosyasının adı ve dosya yolu
            dosya_adı = METRO_DOSYA

            # İlk satırı atlayarak dosyanın geri kalanını yeni bir liste olarak alın
            metrorpr = []
            if not os.path.isfile(dosya_adı):
                QMessageBox.warning(self, "Metro", "METRAP.txt bulunamadı.")
                return None
            with open(dosya_adı, "r") as dosya:
                for satır in dosya:
                    metrorpr.append(satır.strip())

            # Satırları yan yana birleştirerek elde etmek için join() yöntemini kullanıyoruz
            birlesik_veri = "\n".join(metrorpr[1:])

            rapor_tipi = birlesik_veri[:5]
            oktant = birlesik_veri[5:6]
            metro_koor = birlesik_veri[6:9] + " " + birlesik_veri[9:12]
            tarih = birlesik_veri[13:15]
            saat = birlesik_veri[15:17]
            dakika = str(int(birlesik_veri[17:18]) * 6)
            if dakika == "0":
                dakika = str(int(birlesik_veri[17:18]) * 6) + "0"
            ara = ":"
            saat_dakika = saat + ara + str(dakika)

            gecerlilik = birlesik_veri[18:19]
            metro_istasyonu_rakimi = str(int(birlesik_veri[19:22]) * 10)
            yogunluk = int(birlesik_veri[22:25]) / 10

            print("Rapor Tipi:", rapor_tipi)
            print("Oktant:", oktant)
            print("Metro Koordinat:", metro_koor)
            print("Tarih:", tarih)
            print("Saat:", saat_dakika)
            print("Geçerlilik:", gecerlilik)
            print("Metro İstasyonu Rakımı:", metro_istasyonu_rakimi)
            print("Yoğunluk:", yogunluk)


            if bolge == 0:
                ruzgar_istikameti = int(birlesik_veri[28:30]) * 100
                ruzgar_hizi = int(birlesik_veri[30:32])
                if birlesik_veri[32] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[32:35]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[32:35]) / 10
                if birlesik_veri[35] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[35:38]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[35:38]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)  # 39-40-41 3 say

            elif bolge == 1:
                ruzgar_istikameti = int(birlesik_veri[41:43]) * 100
                ruzgar_hizi = int(birlesik_veri[43:45])
                if birlesik_veri[45] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[45:48]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[45:48]) / 10
                if birlesik_veri[48] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[48:51]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[48:51]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)

            elif bolge == 2:
                ruzgar_istikameti = int(birlesik_veri[54:56])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[56:58])  # 3
                if birlesik_veri[58] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[58:61])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[58:61])/10  # 4
                if birlesik_veri[61] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[61:64])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[61:64])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 3:
                ruzgar_istikameti = int(birlesik_veri[67:69])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[69:71])  # 3
                if birlesik_veri[71] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[71:74])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[71:74])/10  # 4
                if birlesik_veri[74] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[74:77])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[74:77])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 4:
                ruzgar_istikameti = int(birlesik_veri[80:82])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[82:84])  # 3
                if birlesik_veri[84] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[84:87])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[84:87])/10  # 4
                if birlesik_veri[87] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[87:90])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[87:90])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 5:
                ruzgar_istikameti = int(birlesik_veri[93:95])*100  # 3
                ruzgar_hizi = int(birlesik_veri[95:97])  # 3
                if birlesik_veri[97] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[97:100])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[97:100])/10  # 4
                if birlesik_veri[100] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[100:103])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[100:103])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 6:
                ruzgar_istikameti = int(birlesik_veri[106:108])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[108:110])  # 3
                if birlesik_veri[110] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[110:113])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[110:113])/10  # 4
                if birlesik_veri[113] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[113:116])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[113:116])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 7:
                ruzgar_istikameti = int(birlesik_veri[119:121])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[121:123])  # 3
                if birlesik_veri[123] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[123:126])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[123:126])/10  # 4
                if birlesik_veri[126] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[126:129])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[126:129])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 8:
                ruzgar_istikameti = int(birlesik_veri[132:134])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[134:136])  # 3
                if birlesik_veri[136] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[136:139])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[136:139])/10  # 4
                if birlesik_veri[139] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[139:142])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[139:142])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 9:
                ruzgar_istikameti = int(birlesik_veri[145:147])*100  # 3
                ruzgar_hizi = int(birlesik_veri[147:149])  # 3
                if birlesik_veri[149] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[149:152])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[149:152])/10  # 4
                if birlesik_veri[152] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[152:155])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[152:155])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 10:
                ruzgar_istikameti = int(birlesik_veri[158:160])*100  # 3
                ruzgar_hizi = int(birlesik_veri[160:162])  # 3
                if birlesik_veri[162] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[162:165])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[162:165])/10  # 4
                if birlesik_veri[165] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[165:168])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[165:168])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)


            # TAMAMLAYICI MESAFEYİ BULUYORUZ.
            if hedefin_toptan_yuksekligi == -400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[1])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[2])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[3])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[4])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -0:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[5])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[6])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[7])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[8])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[9])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 500:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[10])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 600:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[11])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 700:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 800:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 900:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[13])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 1000:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[14])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")

            bt_metro_ist_rakim_farki = batarya_rakimi - int(metro_istasyonu_rakimi)

            giris_mesafesi = round((int(mesafe) + int(tamamlayici_mesafe))/100)*100
            print("Giriş Mesafesi :", giris_mesafesi)

            ruzgarin_plan_istikameti_100 = ruzgar_istikameti - atis_istikameti

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            isi_dzl = [
                (-390,0.9),
            (-380,0.9),
            (-370,0.9),
            (-360,0.8),
            (-350,0.8),
            (-340,0.8),
            (-330,0.8),
            (-320,0.7),
            (-310,0.7),
            (-300,0.7),
            (-290,0.7),
            (-280,0.7),
            (-270,0.7),
            (-260,0.6),
            (-250,0.6),
            (-240,0.6),
            (-230,0.6),
            (-220,0.5),
            (-210,0.5),
            (-200,0.5),
            (-190,0.4),
            (-180,0.4),
            (-170,0.4),
            (-160,0.3),
            (-150,0.3),
            (-140,0.3),
            (-130,0.3),
            (-120,0.2),
            (-110,0.2),
            (-100,0.2),
            (-90,0.2),
            (-80,0.2),
            (-70,0.2),
            (-60,0.1),
            (-50,0.1),
            (-40,0.1),
            (-30,0.1),
            (-20,0),
            (-10,0),
            (0,0),
            (10,0),
            (20,0),
            (30,-0.1),
            (40,-0.1),
            (50,-0.1),
            (60,-0.1),
            (70,-0.2),
            (80,-0.2),
            (90,-0.2),
            (100,-0.2),
            (110,-0.2),
            (120,-0.2),
            (130,-0.3),
            (140,-0.3),
            (150,-0.3),
            (160,-0.3),
            (170,-0.4),
            (180,-0.4),
            (190,-0.4),
            (200,-0.5),
            (210,-0.5),
            (220,-0.5),
            (230,-0.6),
            (240,-0.6),
            (250,-0.6),
            (260,-0.6),
            (270,-0.7),
            (280,-0.7),
            (290,-0.7),
            (300,-0.7),
            (310,-0.7),
            (320,-0.7),
            (330,-0.8),
            (340,-0.8),
            (350,-0.8),
            (360,-0.8),
            (370,-0.9),
            (380,-0.9),
            (390,-0.9),
        ]


            isi_data5bh = []
            for i in range(len(isi_dzl)-1):
                start = isi_dzl[i]
                end = isi_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    isi_data5bh.append((int(x), y))

            isi_data5bh.append(isi_dzl[-1])


            yogunluk_dzl = [
                (-390,3.9),
            (-380,3.8),
            (-370,3.7),
            (-360,3.6),
            (-350,3.5),
            (-340,3.4),
            (-330,3.3),
            (-320,3.2),
            (-310,3.1),
            (-300,3),
            (-290,2.9),
            (-280,2.8),
            (-270,2.7),
            (-260,2.6),
            (-250,2.5),
            (-240,2.4),
            (-230,2.3),
            (-220,2.2),
            (-210,2.1),
            (-200,2),
            (-190,1.9),
            (-180,1.8),
            (-170,1.7),
            (-160,1.6),
            (-150,1.5),
            (-140,1.4),
            (-130,1.3),
            (-120,1.2),
            (-110,1.1),
            (-100,1),
            (-90,0.9),
            (-80,0.8),
            (-70,0.7),
            (-60,0.6),
            (-50,0.5),
            (-40,0.4),
            (-30,0.3),
            (-20,0.2),
            (-10,0.1),
            (0,0),
            (10,-0.1),
            (20,-0.2),
            (30,-0.3),
            (40,-0.4),
            (50,-0.5),
            (60,-0.6),
            (70,-0.7),
            (80,-0.8),
            (90,-0.9),
            (100,-1),
            (110,-1.1),
            (120,-1.2),
            (130,-1.3),
            (140,-1.4),
            (150,-1.5),
            (160,-1.6),
            (170,-1.7),
            (180,-1.8),
            (190,-1.9),
            (200,-2),
            (210,-2.1),
            (220,-2.2),
            (230,-2.3),
            (240,-2.4),
            (250,-2.5),
            (260,-2.6),
            (270,-2.7),
            (280,-2.8),
            (290,-2.9),
            (300,-3),
            (310,-3.1),
            (320,-3.2),
            (330,-3.3),
            (340,-3.4),
            (350,-3.5),
            (360,-3.6),
            (370,-3.7),
            (380,-3.8),
            (390,-3.9),
    ]

            yogunluk_data5bh = []

            for i in range(len(yogunluk_dzl)-1):
                start = yogunluk_dzl[i]
                end = yogunluk_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    yogunluk_data5bh.append((int(x), y))

            yogunluk_data5bh.append(yogunluk_dzl[-1])
            # print(yogunluk_data5bh)


            for i in isi_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    isi_duzeltmesi = i[1]
                    print("Isı Düzeltmesi :", isi_duzeltmesi)


            for i in yogunluk_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    yogunluk_duzeltmesi = i[1]
                    print("Yoğunluk Düzeltmesi", yogunluk_duzeltmesi)

            duzeltilmis_sıcaklik_degeri = isi_duzeltmesi + hava_sicakligi
            print("Düzeltilmiş Isı Değeri:", duzeltilmis_sıcaklik_degeri)

            duzeltilmis_yogunluk_degeri = yogunluk_duzeltmesi + hava_yogunlugu
            print("Düzeltilmiş Yoğunluk Değeri:", duzeltilmis_yogunluk_degeri)

            # RUZGAR BİLEŞENLERİ

            atis_istikameti = round(atis_istikameti/100)*100

            if ruzgar_istikameti < atis_istikameti:
                ruzgar_istikameti = ruzgar_istikameti + 6400
            else:
                ruzgar_istikameti = ruzgar_istikameti
            print("ruzgar istikameti", ruzgar_istikameti)

            print("atış istikameti", atis_istikameti)

            ruzgarin_plan_istikameti = ruzgar_istikameti - atis_istikameti

            print("Rüzgarın Plan İstikameti:", ruzgarin_plan_istikameti)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            data1 = [
                (0,0,1),
(100,-0.1,0.99),
(200,-0.2,0.98),
(300,-0.29,0.96),
(400,-0.38,0.92),
(500,-0.47,0.88),
(600,-0.56,0.83),
(700,-0.63,0.77),
(800,-0.71,0.71),
(900,-0.77,0.63),
(1000,-0.83,0.56),
(1100,-0.88,0.47),
(1200,-0.92,0.38),
(1300,-0.96,0.29),
(1400,-0.98,0.2),
(1500,-0.99,0.1),
(1600,1,0),
(1700,-0.99,-0.1),
(1800,-0.98,-0.2),
(1900,-0.96,-0.29),
(2000,-0.92,-0.38),
(2100,-0.88,-0.47),
(2200,-0.83,-0.56),
(2300,-0.77,-0.63),
(2400,-0.71,-0.71),
(2500,-0.63,-0.77),
(2600,-0.56,-0.83),
(2700,-0.47,-0.88),
(2800,-0.38,-0.92),
(2900,-0.29,-0.96),
(3000,-0.2,-0.98),
(3100,-0.1,-0.99),
(3200,0,-1),
(3300,0.1,-0.99),
(3400,0.2,-0.98),
(3500,0.29,-0.96),
(3600,0.38,-0.92),
(3700,0.47,-0.88),
(3800,0.56,-0.83),
(3900,0.63,-0.77),
(4000,0.71,-0.71),
(4100,0.77,-0.63),
(4200,0.83,-0.56),
(4300,0.88,-0.47),
(4400,0.92,-0.38),
(4500,0.96,-0.29),
(4600,0.98,-0.2),
(4700,0.99,-0.1),
(4800,1,0),
(4900,0.99,0.1),
(5000,0.98,0.2),
(5100,0.96,0.29),
(5200,0.92,0.38),
(5300,0.88,0.47),
(5400,0.83,0.56),
(5500,0.77,0.63),
(5600,0.71,0.71),
(5700,0.63,0.77),
(5800,0.56,0.83),
(5900,0.47,0.88),
(6000,0.38,0.92),
(6100,0.29,0.96),
(6200,0.2,0.98),
(6300,0.1,0.99),
(6400,0,1),


    ]


            # veri_noktasi[0] = MESAFE
            # veri_noktasi[1] = YAN RÜZGARI
            # veri_noktasi[2] = MESAFE RÜZGARI

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Mesafe :", mesafe)

            for rzdlz in interpolasyonlu_veri1:
                if rzdlz[0] == ruzgarin_plan_istikameti:
                    yan_ruzgari_bileseni = rzdlz[1]
                    mesafe_ruzgari_bileseni = rzdlz[2]
                    print("Yan Rüzgarı Bileşeni:", yan_ruzgari_bileseni)
                    print("mesafe Rüzgarı Bileşeni:", mesafe_ruzgari_bileseni)


            # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,136.4,11.4,0.18,11.2,2.4,0.34,12,-11.9,3.1,-2.9,1.4,-3.4,-10.2,9.6,-12,13),
(4100,141.1,11.7,0.18,11.5,2.5,0.34,12.2,-12,3.3,-3.1,1.8,-3.8,-10.5,9.9,-12,13),
(4200,145.7,12.1,0.17,11.9,2.6,0.35,12.3,-12.2,3.5,-3.3,2.2,-4.2,-10.8,10.2,-12,13),
(4300,150.5,12.4,0.17,12.2,2.7,0.36,12.4,-12.3,3.7,-3.4,2.6,-4.6,-11,10.4,-12,13),
(4400,155.3,12.8,0.16,12.6,2.8,0.36,12.6,-12.4,3.9,-3.6,3,-5,-11.3,10.7,-12,13),
(4500,160.1,13.1,0.16,12.9,2.9,0.37,12.7,-12,4.2,-3.8,3.5,-5.4,-11.6,11,-11,13),
(4600,165,13.5,0.16,13.2,3,0.37,12.8,-12.7,4.4,-4,3.9,-5.9,-11.9,11.2,-11,13),
(4700,169.9,13.8,0.15,13.6,3.1,0.38,12.9,-12.8,4.6,-4.2,4.4,-6.3,-12.2,11.5,-11,12),
(4800,174.9,14.2,0.15,13.9,3.2,0.39,13,-12.9,4.8,-4.3,4.9,-6.7,-12.5,11.8,-11,12),
(4900,179.9,14.5,0.14,14.3,3.3,0.39,13.1,-13,5,-4.5,5.3,-7.2,-12.8,12.1,-11,12),
(5000,185,14.9,0.14,14.6,3.4,0.4,13.3,-13.1,5.3,-4.7,5.8,-7.6,-13.1,12.3,-10,12),
(5100,190.2,15.3,0.14,15,3.5,0.4,13.4,-13.3,5.5,-4.9,6.3,-8.1,-13.3,12.6,-10,12),
(5200,195.3,15.6,0.13,15.4,3.6,0.41,13.5,-13.4,5.7,-5.1,6.8,-8.5,-13.6,12.9,-10,11),
(5300,200.6,16,0.13,15.7,3.7,0.41,13.6,-13.5,5.9,-5.3,7.3,-9,-13.9,13.2,-10,11),
(5400,205.9,16.4,0.13,16.1,3.8,0.42,13.7,-13.6,6.2,-5.5,7.8,-9.5,-14.2,13.5,-9,11),
(5500,211.2,16.7,0.13,16.4,3.9,0.42,13.8,-13.7,6.4,-5.7,8.3,-9.9,-14.5,13.8,-9,11),
(5600,216.6,17.1,0.12,16.8,4,0.43,13.9,-13.8,6.6,-5.9,8.9,-10.4,-14.8,14.1,-9,10),
(5700,222,17.5,0.12,17.2,4.1,0.43,14,-13.9,6.9,-6.1,9.4,-10.8,-15.1,14.4,-8,10),
(5800,227.5,17.9,0.12,17.5,4.2,0.44,14.1,-14,7.1,-6.3,9.9,-11.3,-15.4,14.7,-8,10),
(5900,233,18.2,0.12,17.9,4.4,0.44,14.2,-14.1,7.4,-6.5,10.4,-11.7,-15.7,15,-8,10),
(6000,238.6,18,0.11,18.3,4.5,0.45,14.2,-14.2,7.6,-6.7,10.9,-12.2,-16,15.3,-8,9),
(6100,244.3,19,0.11,18.7,4.6,0.45,14.3,-14.3,7.9,-6.9,11.5,-12.7,-16.3,15.6,-7,9),
(6200,250,19.4,0.11,19.1,4.7,0.46,14.4,-14.3,8.1,-7.1,12,-13.1,-16.6,15.9,-7,9),
(6300,255.8,19.8,0.11,19.4,4.8,0.46,14.5,-14.4,8.4,-7.3,12.5,-13.6,-16.9,16.2,-7,8),
(6400,261.6,20.2,0.1,19.8,4.9,0.47,14.6,-14.5,8.6,-7.5,13,-14,-17.3,16.6,-6,8),
(6500,267.5,20.6,0.1,20.2,5.1,0.47,14.7,-14.6,8.9,-7.7,13.5,-14.5,-17.6,16.9,-6,8),
(6600,273.5,21,0.1,20.6,5.2,0.48,14.8,-14.7,9.1,-7.9,14,-14.9,-17.9,17.3,-5,7),
(6700,279.5,21.4,0.1,21,5.3,0.48,14.9,-14.8,9.4,-8.1,14.5,-15.3,-18.2,17.6,-5,7),
(6800,285.6,21.8,0.1,21.4,5.5,0.49,15,-14.9,9.6,-8.3,15.1,-15.8,-18.6,18,-5,7),
(6900,291.8,22.2,0.1,21.8,5.6,0.49,15.1,-15,9.9,-8.5,15.5,-16.2,-18.9,18.3,-4,6),
(7000,298,22.6,0.09,22.2,5.7,0.5,15.2,-15.1,10.1,-8.7,16,-16.6,-19.3,18.7,-4,6),
(7100,304.3,23,0.09,22.6,5.9,0.5,15.3,-15.1,10.4,-8.9,16.5,-17.1,-19.6,19.1,-3,5),
(7200,310.7,23.5,0.09,23,6,0.5,15.3,-15.2,10.6,-9.1,17,-17.5,-20,19.4,-3,5),
(7300,317.2,23.9,0.09,23.5,6.2,0.51,15.4,-15.3,10.9,-9.3,17.5,-17.9,-20.3,19.8,-2,5),
(7400,323.7,24.3,0.09,23.9,6.3,0.51,15.5,-15.4,11.1,-9.5,17.9,-18.3,-20.7,20.2,-2,4),
(7500,330.4,24.8,0.09,24.3,6.5,0.52,15.6,-15.5,11.4,-9.7,18.4,-18.7,-21.1,20.6,-1,4),
(7600,337.1,25.2,0.08,24.7,6.6,0.52,15.7,-15.6,11.7,-9.9,18.9,-19.1,-21.4,21,-1,3),
(7700,343.9,25.6,0.08,25.2,6.8,0.53,15.8,-15.7,11.9,-10.2,19.3,-19.5,-21.8,21.4,0,3),
(7800,350.8,26.1,0.08,25.6,6.9,0.53,15.9,-15.8,12.2,-10.4,19.8,-19.9,-22.2,21.9,0,2),
(7900,357.9,26.6,0.08,26.1,7.1,0.54,16,-15.8,12.4,-10.6,20.2,-20.3,-22.6,22.3,1,2),
(8000,365,27,0.08,26.5,7.3,0.54,16.1,-15.9,12.7,-10.8,20.6,-20.7,-23,22.7,1,1),
(8100,372.2,27.5,0.08,27,7.4,0.55,16.2,-16,13,-11,21,-21.1,-23.4,23.2,2,1),
(8200,379.6,28,0.08,27.4,7.6,0.55,16.3,-16.1,13.2,-11.2,21.5,-21.5,-23.8,23.6,2,0),
(8300,387.1,28.4,0.08,27.9,7.8,0.56,16.4,-16.2,13.5,-11.4,21.9,-21.8,-24.3,24.1,3,0),
(8400,394.7,28.9,0.07,28.4,8,0.56,16.5,-16.3,13.7,-11.6,22.3,-22.2,-24.7,24.5,3,-1),
(8500,402.4,29.4,0.07,28.9,8.2,0.57,16.6,-16.4,14,-11.8,22.7,-22.5,-25.1,25,4,-1),
(8600,410.3,29.9,0.07,29.4,8.4,0.57,16.7,-16.5,14.3,-12,23,-22.9,-25.6,25.5,5,-2),
(8700,418.4,30.4,0.07,29.9,8.6,0.58,16.8,-16.6,14.5,-12.2,23.4,-23.2,-26,26,5,-2),
(8800,426.6,30.9,0.07,30.4,8.8,0.59,16.9,-16.7,14.8,-12.4,23.8,-23.6,-26.5,26.5,6,-3),
(8900,435,31.5,0.07,30.9,9,0.59,17,-16.8,15.1,-12.6,24.1,-23.9,-26.9,27,6,-3),
(9000,443.6,32,0.07,31.4,9.2,0.6,17.1,-16.9,15.3,-12.8,24.5,-24.3,-27.4,27.5,7,-4),
(9100,452.4,32.6,0.07,31.9,9.5,0.6,17.2,-16.9,15.6,-13,24.8,-24.6,-27.9,28,8,-5),
(9200,461.4,33.1,0.07,32.5,9.7,0.61,17.3,-17,15.9,-13.3,25.2,-24.9,-28.4,28.6,8,-5),
(9300,470.6,33.7,0.06,33.1,10,0.61,17.4,-17.1,16.2,-13.5,25.5,-25.2,-28.9,29.1,9,-6),
(9400,480.1,34.3,0.06,33.6,10.2,0.62,17.5,-17.2,16.4,-13.7,25.8,-25.5,-29.4,29.7,10,-7),
(9500,489.9,34.9,0.06,34.2,10.5,0.63,17.6,-17.4,16.7,-13.9,26.1,-25.8,-29.9,30.3,11,-7),
(9600,500,35.5,0.06,34.8,10.8,0.63,17.8,-17.5,17,-14.1,26.4,-26.1,-30.4,30.9,11,-8),
(9700,510.5,36.1,0.06,35.4,11.1,0.64,17.9,-17.6,17.3,-14.3,26.7,-26.4,-30.9,31.5,12,-9),
(9800,521.4,36.8,0.06,36.1,11.4,0.65,18,-17.7,17,-14.5,26.9,-26.6,-31.5,32.1,13,-9),
(9900,532.7,37.5,0.06,36.7,11.7,0.65,18.1,-17.8,17.9,-14.7,27.2,-26.9,-32,32.7,14,-10),
(10000,544.5,38.2,0.06,37.4,12.1,0.66,18.3,-17.9,18.3,-14.9,27.4,-27.2,-32.6,33.4,14,-11),
(10100,556.8,38.9,0.06,38.2,12.5,0.67,18.4,-18,18.3,-15.1,27.6,-27.4,-33.1,34.1,15,-11),
(10200,569.9,39.7,0.06,38.9,12.9,0.68,18.5,-18.1,18.3,-15.3,27.9,-27.7,-33.7,34.8,16,-12),
(10300,583.8,40.5,0.05,39.7,13.3,0.69,18.7,-18.2,18.3,-15.5,28.1,-27.9,-34.3,35.5,17,-13),
(10400,598.6,41.3,0.05,40.5,13.8,0.7,18.8,-18.4,18.3,-15.7,28.2,-28.1,-34.9,36.3,18,-14),
(10500,614.7,42.3,0.05,41.4,14.4,0.71,18.9,-18.5,18.3,-15.9,28.4,-28.4,-35.5,37.2,19,-15),
(10600,632.5,43.3,0.05,42.4,15,0.72,19.1,-18.6,18.3,-16.1,28.5,-28.6,-36.2,38.3,20,-16),
(10700,652.5,44.4,0.05,43.5,15.7,0.74,19.3,-18.7,18.3,-16.3,28.5,-28.8,-36.8,38.3,21,-16),
(10800,676.2,45.7,0.05,44.8,16.6,0.75,19.5,-18.9,18.3,-16.5,28.5,-29,-37.5,38.3,22,-18),
(10900,706.6,47.4,0.05,46.4,17.8,0.77,19.5,-19,18.3,-16.7,28.5,-29.1,-38.2,38.3,24,-19),
(11000,764.2,50.4,0.04,49.4,20.3,0.77,19.5,-19.2,18.3,-16.9,28.5,-29.3,-38.9,38.3,24,-20),



    ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == giris_mesafesi:
                    nisangah5bh = round(nsg[1], 3)
                    print("Nişangah:", nisangah5bh)

            for tsg in interpolasyonlu_veri1:
                if tsg[0] == giris_mesafesi:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            for ts in interpolasyonlu_veri1:
                if ts[0] == giris_mesafesi:
                    deltaTS = round(ts[3], 3)
                    # print("Delta Tapa Saniyesi:", deltaTS)


            for ucs in interpolasyonlu_veri1:
                if ucs[0] == giris_mesafesi:
                    ucussuresi = round(ucs[4], 3)
                    # print("Uçuş Süresi:", ucussuresi)

            for dgl in interpolasyonlu_veri1:
                if dgl[0] == giris_mesafesi:
                    dogalyandz = round(dgl[5], 3)
                    # print("Doğal Yan Düzeltmesi:", dogalyandz)

            for yrd in interpolasyonlu_veri1:
                if yrd[0] == giris_mesafesi:
                    yan_ruzgari_duzeltme_faktoru = round(yrd[6], 1)
                    # print("Yan Rüzgarı Düzeltmesi:", yanruzgaridz)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_eksilme = round(ihz[7], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimeksilme)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_martma = round(ihz[8], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimartma)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_bas = round(mrd[9], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (BAŞ):", mesaferuzgaridzl_bas)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_arka = round(mrd[10], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (ARKA):", mesaferuzgaridzl_arka)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_eksilme = round(hvs[11], 3)
                    # print("Hava Isısı Düzeltmesi (Eksilme):", havaisisi_eksilme)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_artma = round(hvs[12], 3)
                    # print("Hava Isısı Düzeltmesi (Artma):", havaisisi_artma)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_eksilme = round(hyd[13], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", havayogunlugu_eksilme)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_artma = round(hyd[14], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", havayogunlugu_artma)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_eksilme = round(mka[15], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", mermikareagirligi_eksilme)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_artma = round(mka[16], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", mermikareagirligi_artma)

            print("Yan Rüzgarı Düzeltme Faktörü:", yan_ruzgari_duzeltme_faktoru)

            mesafe_ruzgari = ruzgar_hizi * mesafe_ruzgari_bileseni
            print("mesafe ruzgari", mesafe_ruzgari)

            ruzgar_yan_duzeltmesi = round(
                (ruzgar_hizi * yan_ruzgari_bileseni * yan_ruzgari_duzeltme_faktoru), 1)
            print("Rüzgar Yan Düzeltmesi", ruzgar_yan_duzeltmesi)

            dogalyandz = round(dogalyandz, 1)
            print("Doğal Yan Düzeltmesi:", dogalyandz)

            atis_istikameti = round(atis_istikameti/400)*400

            # dünyanın dönmesi düzeltmesi


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,0.5,0.5,0.5,0.5,0.5,0.5,0.5,0.5,0.5),
(5000,0.6,0.6,0.6,0.6,0.6,0.7,0.7,0.7,0.7),
(6000,0.7,0.7,0.8,0.8,0.8,0.8,0.8,0.9,0.9),
(7000,0.9,0.9,0.9,0.9,1,1,1,1.1,1.1),
(8000,1,1,1,1.1,1.2,1.2,1.3,1.3,1.3),
(9000,1.1,1.2,1.2,1.3,1.4,1.5,1.5,1.6,1.6),
(10000,1.3,1.3,1.4,1.5,1.6,1.8,1.9,2,2),
(11000,1.3,1.4,1.5,1.6,1.8,2,2.1,2.2,2.3),






    ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            if atis_istikameti == 00 or atis_istikameti == 6400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[1], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 6000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[2], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 5600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[3], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 5200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[4], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 4800:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[5], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2000 or atis_istikameti == 4400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[6], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2400 or atis_istikameti == 4000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[7], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2800 or atis_istikameti == 3600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[8], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 3200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[9], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,0,-5,-9,-14,-18,-21,-23,-24,-25,0,5,9,14,18,21,24,25),
(5000,0,-6,-11,-17,-21,-25,-28,-29,-30,0,6,11,17,21,25,29,30),
(6000,0,-7,-13,-19,-24,-29,-32,-34,-35,0,7,13,19,24,29,34,35),
(7000,0,-8,-15,-22,-27,-32,-36,-38,-39,0,8,15,22,27,32,38,39),
(8000,0,-8,-16,-24,-30,-35,-39,-42,-42,0,8,16,24,30,35,42,42),
(9000,0,-9,-17,-25,-32,-37,-41,-44,-45,0,9,17,25,32,37,44,45),
(10000,0,-9,-18,-25,-32,-38,-42,-45,-46,0,9,18,25,32,38,45,46),
(11000,0,-8,-16,-23,-29,-34,-38,-40,-41,0,8,16,23,29,34,40,41),
    ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)
            if atis_istikameti == 0 or atis_istikameti == 3200 or atis_istikameti == 6400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[1], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 200 or atis_istikameti == 3000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[2], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 2800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[3], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 600 or atis_istikameti == 2600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[4], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 2400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[5], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1000 or atis_istikameti == 2200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[6], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 2000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[7], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1400 or atis_istikameti == 1800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[8], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 1600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[9], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            if atis_istikameti == 3400 or atis_istikameti == 6200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[11], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3600 or atis_istikameti == 6000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[12], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3800 or atis_istikameti == 5800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[13], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4000 or atis_istikameti == 5600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[14], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4200 or atis_istikameti == 5400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[15], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4400 or atis_istikameti == 5200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[16], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4600 or atis_istikameti == 5000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[17], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4800 or atis_istikameti == 4800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[18], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            print("Dünyanın Dönmesi Düzeltmesi:", arzindonusuyanduzeltmesi)

            metro_yan_duzeltmesi = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            metro_yan_duzeltmesi_1 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi:", metro_yan_duzeltmesi)
            metro_yan_duzeltmesi = round(metro_yan_duzeltmesi/1)*1
            print("Metro Yan Düzeltmesi en yakın 1 milyem:", metro_yan_duzeltmesi)


            # METRO MESAFE DÜZELTMESİ
            mesafe_ruzgari_fark = round(mesafe_ruzgari-0, 1)
            if mesafe_ruzgari_fark > 0:
                mesafe_ruzgari_durum = "BAŞ"
            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_durum = "ARKA"
            else:
                mesafe_ruzgari_durum = 0
            print("Mesafe Ruzgari Fark :", mesafe_ruzgari_fark)
            hava_sicakligi_fark = round(duzeltilmis_sıcaklik_degeri - 100, 1)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_durum = "EKSİLME"
            if hava_sicakligi_fark > 0:
                hava_sicakligi_durum = "ARTMA"
            else:
                hava_sicakligi_durum = 0
            print("Hava Sıcaklığı Fark :", hava_sicakligi_fark)
            hava_yogunlugu_fark = round(duzeltilmis_yogunluk_degeri - 100, 1)
            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_durum = "EKSİLME"
            if hava_yogunlugu_fark > 0:
                hava_yogunlugu_durum = "ARTMA"
            else:
                hava_yogunlugu_durum = 0
            print("Hava Yoğunluğu Fark :", hava_yogunlugu_fark)
            mermi_kare_agirligi_fark = round(mermi_kare_agirligi-2, 1)
            if mermi_kare_agirligi_fark < 0:
                merkar_durum = "EKSİLME"
            if mermi_kare_agirligi_fark > 0:
                merkar_durum = "ARTMA"
            if mermi_kare_agirligi_fark == 0:
                merkar_durum = 0

            print("Mermi Kare Ağırlığı Fark :", mermi_kare_agirligi_fark)

            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_arka
            else:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_bas
            if mesafe_ruzgari_fark == 0:
                mesafe_ruzgari_duzeltme_birimi = 0

            print("Mesafe Rüzgarı Düzeltme Birimi:", mesafe_ruzgari_duzeltme_birimi)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_duzeltme_birimi = havaisisi_eksilme
            else:
                hava_sicakligi_duzeltme_birimi = havaisisi_artma
            if hava_sicakligi_fark == 0:
                hava_sicakligi_duzeltme_birimi = 0
            print("Hava Isısı Düzeltme Birimi:", hava_sicakligi_duzeltme_birimi)

            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_eksilme
            else:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_artma
            if hava_yogunlugu_fark == 0:
                hava_yogunlugu_duzeltme_birimi = 0
            print("Hava Yoğunluğu Düzeltme Birimi:", hava_yogunlugu_duzeltme_birimi)

            if mermi_kare_agirligi_fark < 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_eksilme
            elif mermi_kare_agirligi_fark > 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_artma
            else:
                mermi_kare_agirligi_duzeltme_birimi = 0
            mesafe_ruzgari_fark = abs(mesafe_ruzgari_fark)
            hava_sicakligi_fark = abs(hava_sicakligi_fark)
            hava_yogunlugu_fark = abs(hava_yogunlugu_fark)


            print("Mermi Kare Ağırlığı Düzeltme Birimi:",
                mermi_kare_agirligi_duzeltme_birimi)

            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi = round(metro_yan_duzeltmesi + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)
            gac_yan_duzeltmesi = round(toplam_yan_duzeltmesi - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi)

            # Lineer interpolasyon fonksiyonu

            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            brt_isisi_ilk_hiz_dzl = [
                (-40,-14),
(-30,-12.9),
(-20,-11.8),
(-10,-10.7),
(0,-9.5),
(10,-8.2),
(20,-6.9),
(30,-5.6),
(40,-4.3),
(50,-2.9),
(60,-1.5),
(70,0),
(80,1.5),
(90,3.1),
(100,4.7),
(110,6.3),
(120,7.9),
(130,9.6),



    ]
            barut__isisi_dzl2 = 3
            barut_isisi_dzl = []
            for i in range(len(brt_isisi_ilk_hiz_dzl)-1):
                start = brt_isisi_ilk_hiz_dzl[i]
                end = brt_isisi_ilk_hiz_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    barut_isisi_dzl.append((int(x), y))

            barut_isisi_dzl.append(brt_isisi_ilk_hiz_dzl[-1])
            print("Barut Isısı",barut_isisi)

            for i in barut_isisi_dzl:
                if i[0] == barut_isisi:
                    print("mesafe",i[0])
                    barut__isisi_dzl2 = round(i[1], 1)
                    print("İ1:",i[1])
                    print("Barurrrrrrrr",barut__isisi_dzl2)
            # HIZ DEĞİŞİKLİĞİNİN HESAPLANMASI
            hiz_degisikligi = round(ilk_hiz_farki + mevzi_hiz_degisikligi, 1)
            deltaV_hiz_farki = round(hiz_degisikligi + barut__isisi_dzl2, 1)
            print("Hız Değişikliği:", hiz_degisikligi)
            print("Barut Isısı Düzeltmesi2:", barut__isisi_dzl2)
            print("Delta V Hız Farkı:", deltaV_hiz_farki)

            if deltaV_hiz_farki > 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_martma
            elif deltaV_hiz_farki < 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_eksilme
            else:
                ilk_hiz_duzeltme_birimi = 0
            print("İlk Hız Düzeltme Birimi:", ilk_hiz_duzeltme_birimi)

            deltaV_mesafe_duzeltmesi = abs(
                round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            if deltaV_hiz_farki > 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))*-1
            elif deltaV_hiz_farki < 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            else:
                deltaV_mesafe_duzeltmesi = 0
            print("Delta V Mesafe Düzeltmesi:", deltaV_mesafe_duzeltmesi)

            #### TOPLAM MESAF DÜZELTMESİ ###
            toplam_mesafe_duzeltmesi = round(metro_mesafe_duzeltmesi + deltaV_mesafe_duzeltmesi)
            print("Toplam Mesafe Düzeltmesi:", toplam_mesafe_duzeltmesi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)

            # toplam yan düzeltmede 1 milyem fazla mesafe düzeltmesinde işaret + ve 335 çıktı
            # baiek de yan 4 sl mesafe -326 çıktı

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,0,0,0,0,0,0,0,0,0,0),
(1,0,0,0,0,0,0,0,0,0,0),
(2,-0.005,0.005,0,0,0,0,0,0,0.011,-0.011),
(3,-0.007,0.007,0,0,0,0,0.001,-0.001,0.016,-0.016),
(4,-0.009,0.009,0,0,0,-0.001,0.002,-0.002,0.02,-0.02),
(5,-0.011,0.011,0,0,0.001,-0.001,0.002,-0.002,0.024,-0.024),
(6,-0.013,0.013,0,0,0.001,-0.002,0.004,-0.004,0.027,-0.027),
(7,-0.015,0.015,0,0,0.003,-0.003,0.005,-0.005,0.029,-0.03),
(8,-0.017,0.017,0,0,0.004,-0.004,0.007,-0.007,0.031,-0.032),
(9,-0.019,0.019,0,0,0.006,-0.004,0.009,-0.009,0.032,-0.033),
(10,-0.021,0.021,0,0,0.006,-0.004,0.011,-0.011,0.033,-0.035),
(11,-0.022,0.022,0,0.001,0.007,-0.004,0.013,-0.012,0.034,-0.036),
(12,-0.024,0.024,0,0.001,0.007,-0.003,0.015,-0.013,0.035,-0.037),
(13,-0.025,0.025,-0.001,0.001,0.006,-0.001,0.016,-0.015,0.036,-0.038),
(14,-0.026,0.026,-0.001,0.002,0.005,0,0.018,-0.016,0.036,-0.038),
(15,-0.027,0.027,-0.002,0.003,0.004,0.002,0.019,-0.017,0.037,-0.039),
(16,-0.028,0.028,-0.002,0.003,0.002,0.003,0.02,-0.018,0.037,-0.04),
(17,-0.029,0.029,-0.003,0.004,0.001,0.005,0.022,-0.019,0.038,-0.041),
(18,-0.03,0.03,-0.003,0.004,-0.001,0.007,0.023,-0.021,0.038,-0.041),
(19,-0.031,0.031,-0.004,0.005,-0.003,0.01,0.024,-0.022,0.039,-0.042),
(20,-0.032,0.032,-0.004,0.006,-0.005,0.012,0.025,-0.023,0.039,-0.042),
(21,-0.033,0.033,-0.005,0.006,-0.007,0.014,0.027,-0.024,0.04,-0.043),
(22,-0.034,0.034,-0.005,0.007,-0.009,0.016,0.028,-0.025,0.04,-0.043),
(23,-0.035,0.035,-0.006,0.008,-0.012,0.018,0.029,-0.026,0.04,-0.043),
(24,-0.036,0.036,-0.007,0.008,-0.014,0.021,0.03,-0.027,0.04,-0.044),
(25,-0.036,0.037,-0.007,0.009,-0.016,0.023,0.031,-0.028,0.04,-0.044),
(26,-0.037,0.037,-0.008,0.01,-0.019,0.025,0.033,-0.029,0.04,-0.044),
(27,-0.038,0.038,-0.008,0.01,-0.021,0.027,0.034,-0.03,0.04,-0.044),
(28,-0.039,0.039,-0.009,0.011,-0.023,0.029,0.035,-0.032,0.04,-0.044),
(29,-0.04,0.04,-0.01,0.012,-0.025,0.031,0.036,-0.033,0.04,-0.044),
(30,-0.041,0.041,-0.01,0.012,-0.027,0.033,0.038,-0.034,0.04,-0.044),
(31,-0.041,0.042,-0.011,0.013,-0.03,0.035,0.039,-0.036,0.04,-0.044),
(32,-0.042,0.043,-0.012,0.014,-0.032,0.037,0.041,-0.037,0.04,-0.044),
(33,-0.043,0.043,-0.012,0.014,-0.034,0.039,0.042,-0.038,0.039,-0.044),
(34,-0.044,0.044,-0.013,0.015,-0.036,0.041,0.044,-0.04,0.039,-0.044),
(35,-0.045,0.045,-0.013,0.015,-0.038,0.043,0.045,-0.041,0.039,-0.043),
(36,-0.046,0.046,-0.014,0.016,-0.04,0.045,0.047,-0.043,0.038,-0.043),
(37,-0.047,0.047,-0.014,0.016,-0.042,0.047,0.048,-0.044,0.038,-0.043),
(38,-0.048,0.048,-0.015,0.017,-0.043,0.048,0.05,-0.046,0.037,-0.043),
(39,-0.048,0.049,-0.015,0.017,-0.045,0.05,0.052,-0.047,0.037,-0.042),
(40,-0.049,0.05,-0.016,0.018,-0.047,0.052,0.053,-0.049,0.036,-0.042),
(41,-0.05,0.051,-0.016,0.018,-0.049,0.053,0.055,-0.05,0.036,-0.042),
(42,-0.051,0.051,-0.017,0.019,-0.05,0.055,0.057,-0.052,0.035,-0.041),
(43,-0.052,0.052,-0.017,0.019,-0.052,0.056,0.059,-0.054,0.035,-0.041),
(44,-0.053,0.053,-0.018,0.019,-0.053,0.058,0.061,-0.055,0.034,-0.041),
(45,-0.054,0.054,-0.018,0.02,-0.055,0.059,0.062,-0.057,0.034,-0.04),
(46,-0.055,0.055,-0.018,0.02,-0.056,0.061,0.064,-0.059,0.033,-0.04),
(47,-0.056,0.056,-0.019,0.02,-0.058,0.062,0.066,-0.061,0.033,-0.04),
(48,-0.057,0.057,-0.019,0.02,-0.059,0.063,0.068,-0.062,0.032,-0.039),
(49,-0.058,0.059,-0.019,0.021,-0.061,0.065,0.07,-0.064,0.032,-0.039),
(50,-0.059,0.06,-0.02,0.021,-0.062,0.066,0.072,-0.066,0.031,-0.039),
(51,-0.061,0.061,-0.02,0.021,-0.063,0.067,0.074,-0.068,0.031,-0.038),
(52,-0.062,0.062,-0.02,0.021,-0.064,0.068,0.076,-0.07,0.031,-0.038),
(53,-0.063,0.063,-0.02,0.021,-0.065,0.07,0.078,-0.071,0.03,-0.038),
(54,-0.064,0.064,-0.02,0.021,-0.067,0.071,0.08,-0.073,0.03,-0.038),
(55,-0.065,0.065,-0.02,0.021,-0.068,0.072,0.082,-0.075,0.029,-0.038),
(56,-0.066,0.066,-0.02,0.021,-0.069,0.073,0.084,-0.077,0.029,-0.038),
(57,-0.067,0.067,-0.021,0.021,-0.07,0.074,0.086,-0.079,0.029,-0.038),
(58,-0.069,0.069,-0.021,0.021,-0.071,0.075,0.088,-0.081,0.029,-0.038),
(59,-0.07,0.07,-0.02,0.021,-0.072,0.076,0.09,-0.082,0.029,-0.038),
(60,-0.071,0.071,-0.02,0.021,-0.073,0.077,0.092,-0.084,0.029,-0.038),
(61,-0.072,0.072,-0.02,0.02,-0.074,0.078,0.094,-0.086,0.029,-0.039),
(62,-0.074,0.074,-0.02,0.02,-0.075,0.079,0.096,-0.088,0.03,-0.039),
(63,-0.075,0.075,-0.02,0.02,-0.076,0.08,0.098,-0.09,0.03,-0.04),
(64,-0.076,0.076,-0.02,0.02,-0.076,0.081,0.1,-0.092,0.031,-0.04),
(65,-0.077,0.078,-0.019,0.019,-0.077,0.082,0.102,-0.093,0.032,-0.042),
(66,-0.079,0.079,-0.019,0.019,-0.078,0.083,0.104,-0.095,0.033,-0.043),
(67,-0.08,0.08,-0.019,0.019,-0.079,0.084,0.106,-0.097,0.035,-0.045),
(68,-0.082,0.082,-0.018,0.019,-0.079,0.085,0.108,-0.099,0.038,-0.048),
(69,-0.083,0.083,-0.018,0.022,-0.079,0.085,0.11,-0.101,0.044,-0.053),
(70,-0.085,0.086,-0.019,0.043,-0.078,0.084,0.114,-0.104,0.058,-0.065),
(71,-0.09,0.09,-0.023,0.043,-0.073,0.08,0.123,-0.11,0.103,-0.1),
(72,-0.093,0.09,-0.033,0.043,-0.073,0.072,0.123,-0.121,0.103,-0.162),
(73,-0.105,0.09,-0.047,0.043,-0.073,0.054,0.123,-0.142,0.103,-0.284),


    ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Tapa Saniyesi :", tapasaniyesi)

            if deltaV_hiz_farki < 0:
                for ihz_eksilme in interpolasyonlu_veri1:
                    if ihz_eksilme[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_eksilme[1], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Eksilme):", ilkhizda1mildegisim_eksilme)
            if deltaV_hiz_farki > 0:
                for ihz_artma in interpolasyonlu_veri1:
                    if ihz_artma[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_artma[2], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Artma):", ilkhizda1mildegisim_artma)
            if mesafe_ruzgari_durum == "BAŞ":
                for musrus_bas in interpolasyonlu_veri1:
                    if musrus_bas[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(musrus_bas[3], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Baş):", mesaferuzgaridegisimi_bas)
            if mesafe_ruzgari_durum == "ARKA":
                for mesruz_arka in interpolasyonlu_veri1:
                    if mesruz_arka[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(mesruz_arka[4], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Arka):", mesaferuzgaridegisimi_arka)
            if hava_sicakligi_durum == "EKSİLME":
                for hadese_eksilme in interpolasyonlu_veri1:
                    if hadese_eksilme[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_eksilme[5], 3)
                        # print("Hava Isısı Değişim Miktarı (Eksilme):", havaisisidegisim_eksilme)
            if hava_sicakligi_durum == "ARTMA":
                for hadese_artma in interpolasyonlu_veri1:
                    if hadese_artma[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_artma[6], 3)
                        # print("Hava Isısı Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == "EKSİLME":
                for hayog_eksilme in interpolasyonlu_veri1:
                    if hayog_eksilme[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_eksilme[7], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Eksilme):", havayogunlugudegisim_eksilme)
            if hava_yogunlugu_durum == "ARTMA":
                for hayog_artma in interpolasyonlu_veri1:
                    if hayog_artma[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_artma[8], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == 0:
                havayogunlugudegisim = 0
            if merkar_durum == "EKSİLME":
                for merkar_eksilme in interpolasyonlu_veri1:
                    if merkar_eksilme[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_eksilme[9], 3)
                        # print("Mermi Kare Ağırlığı Değişim Miktarı (Eksilme):", mermikareagirligidegisim_eksilme)
            if merkar_durum == "ARTMA":
                for merkar_artma in interpolasyonlu_veri1:
                    if merkar_artma[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_artma[10], 3)
            else:
                mermikareagirligidegisim = 0
                # print("Mermi Kare Ağırlığı Değişim Miktarı (Artma):", mermikareagirligidegisim_artma)
##########################################################################################################
            gac_mesafe = plan_mesafesi_1B_obüs + toplam_mesafe_duzeltmesi
            print("toplam mesafe",nisangah5bh)
# zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
               (4000,136.4,11.4,0.18,11.2,2.4,0.34,12,-11.9,3.1,-2.9,1.4,-3.4,-10.2,9.6,-12,13),
(4100,141.1,11.7,0.18,11.5,2.5,0.34,12.2,-12,3.3,-3.1,1.8,-3.8,-10.5,9.9,-12,13),
(4200,145.7,12.1,0.17,11.9,2.6,0.35,12.3,-12.2,3.5,-3.3,2.2,-4.2,-10.8,10.2,-12,13),
(4300,150.5,12.4,0.17,12.2,2.7,0.36,12.4,-12.3,3.7,-3.4,2.6,-4.6,-11,10.4,-12,13),
(4400,155.3,12.8,0.16,12.6,2.8,0.36,12.6,-12.4,3.9,-3.6,3,-5,-11.3,10.7,-12,13),
(4500,160.1,13.1,0.16,12.9,2.9,0.37,12.7,-12,4.2,-3.8,3.5,-5.4,-11.6,11,-11,13),
(4600,165,13.5,0.16,13.2,3,0.37,12.8,-12.7,4.4,-4,3.9,-5.9,-11.9,11.2,-11,13),
(4700,169.9,13.8,0.15,13.6,3.1,0.38,12.9,-12.8,4.6,-4.2,4.4,-6.3,-12.2,11.5,-11,12),
(4800,174.9,14.2,0.15,13.9,3.2,0.39,13,-12.9,4.8,-4.3,4.9,-6.7,-12.5,11.8,-11,12),
(4900,179.9,14.5,0.14,14.3,3.3,0.39,13.1,-13,5,-4.5,5.3,-7.2,-12.8,12.1,-11,12),
(5000,185,14.9,0.14,14.6,3.4,0.4,13.3,-13.1,5.3,-4.7,5.8,-7.6,-13.1,12.3,-10,12),
(5100,190.2,15.3,0.14,15,3.5,0.4,13.4,-13.3,5.5,-4.9,6.3,-8.1,-13.3,12.6,-10,12),
(5200,195.3,15.6,0.13,15.4,3.6,0.41,13.5,-13.4,5.7,-5.1,6.8,-8.5,-13.6,12.9,-10,11),
(5300,200.6,16,0.13,15.7,3.7,0.41,13.6,-13.5,5.9,-5.3,7.3,-9,-13.9,13.2,-10,11),
(5400,205.9,16.4,0.13,16.1,3.8,0.42,13.7,-13.6,6.2,-5.5,7.8,-9.5,-14.2,13.5,-9,11),
(5500,211.2,16.7,0.13,16.4,3.9,0.42,13.8,-13.7,6.4,-5.7,8.3,-9.9,-14.5,13.8,-9,11),
(5600,216.6,17.1,0.12,16.8,4,0.43,13.9,-13.8,6.6,-5.9,8.9,-10.4,-14.8,14.1,-9,10),
(5700,222,17.5,0.12,17.2,4.1,0.43,14,-13.9,6.9,-6.1,9.4,-10.8,-15.1,14.4,-8,10),
(5800,227.5,17.9,0.12,17.5,4.2,0.44,14.1,-14,7.1,-6.3,9.9,-11.3,-15.4,14.7,-8,10),
(5900,233,18.2,0.12,17.9,4.4,0.44,14.2,-14.1,7.4,-6.5,10.4,-11.7,-15.7,15,-8,10),
(6000,238.6,18,0.11,18.3,4.5,0.45,14.2,-14.2,7.6,-6.7,10.9,-12.2,-16,15.3,-8,9),
(6100,244.3,19,0.11,18.7,4.6,0.45,14.3,-14.3,7.9,-6.9,11.5,-12.7,-16.3,15.6,-7,9),
(6200,250,19.4,0.11,19.1,4.7,0.46,14.4,-14.3,8.1,-7.1,12,-13.1,-16.6,15.9,-7,9),
(6300,255.8,19.8,0.11,19.4,4.8,0.46,14.5,-14.4,8.4,-7.3,12.5,-13.6,-16.9,16.2,-7,8),
(6400,261.6,20.2,0.1,19.8,4.9,0.47,14.6,-14.5,8.6,-7.5,13,-14,-17.3,16.6,-6,8),
(6500,267.5,20.6,0.1,20.2,5.1,0.47,14.7,-14.6,8.9,-7.7,13.5,-14.5,-17.6,16.9,-6,8),
(6600,273.5,21,0.1,20.6,5.2,0.48,14.8,-14.7,9.1,-7.9,14,-14.9,-17.9,17.3,-5,7),
(6700,279.5,21.4,0.1,21,5.3,0.48,14.9,-14.8,9.4,-8.1,14.5,-15.3,-18.2,17.6,-5,7),
(6800,285.6,21.8,0.1,21.4,5.5,0.49,15,-14.9,9.6,-8.3,15.1,-15.8,-18.6,18,-5,7),
(6900,291.8,22.2,0.1,21.8,5.6,0.49,15.1,-15,9.9,-8.5,15.5,-16.2,-18.9,18.3,-4,6),
(7000,298,22.6,0.09,22.2,5.7,0.5,15.2,-15.1,10.1,-8.7,16,-16.6,-19.3,18.7,-4,6),
(7100,304.3,23,0.09,22.6,5.9,0.5,15.3,-15.1,10.4,-8.9,16.5,-17.1,-19.6,19.1,-3,5),
(7200,310.7,23.5,0.09,23,6,0.5,15.3,-15.2,10.6,-9.1,17,-17.5,-20,19.4,-3,5),
(7300,317.2,23.9,0.09,23.5,6.2,0.51,15.4,-15.3,10.9,-9.3,17.5,-17.9,-20.3,19.8,-2,5),
(7400,323.7,24.3,0.09,23.9,6.3,0.51,15.5,-15.4,11.1,-9.5,17.9,-18.3,-20.7,20.2,-2,4),
(7500,330.4,24.8,0.09,24.3,6.5,0.52,15.6,-15.5,11.4,-9.7,18.4,-18.7,-21.1,20.6,-1,4),
(7600,337.1,25.2,0.08,24.7,6.6,0.52,15.7,-15.6,11.7,-9.9,18.9,-19.1,-21.4,21,-1,3),
(7700,343.9,25.6,0.08,25.2,6.8,0.53,15.8,-15.7,11.9,-10.2,19.3,-19.5,-21.8,21.4,0,3),
(7800,350.8,26.1,0.08,25.6,6.9,0.53,15.9,-15.8,12.2,-10.4,19.8,-19.9,-22.2,21.9,0,2),
(7900,357.9,26.6,0.08,26.1,7.1,0.54,16,-15.8,12.4,-10.6,20.2,-20.3,-22.6,22.3,1,2),
(8000,365,27,0.08,26.5,7.3,0.54,16.1,-15.9,12.7,-10.8,20.6,-20.7,-23,22.7,1,1),
(8100,372.2,27.5,0.08,27,7.4,0.55,16.2,-16,13,-11,21,-21.1,-23.4,23.2,2,1),
(8200,379.6,28,0.08,27.4,7.6,0.55,16.3,-16.1,13.2,-11.2,21.5,-21.5,-23.8,23.6,2,0),
(8300,387.1,28.4,0.08,27.9,7.8,0.56,16.4,-16.2,13.5,-11.4,21.9,-21.8,-24.3,24.1,3,0),
(8400,394.7,28.9,0.07,28.4,8,0.56,16.5,-16.3,13.7,-11.6,22.3,-22.2,-24.7,24.5,3,-1),
(8500,402.4,29.4,0.07,28.9,8.2,0.57,16.6,-16.4,14,-11.8,22.7,-22.5,-25.1,25,4,-1),
(8600,410.3,29.9,0.07,29.4,8.4,0.57,16.7,-16.5,14.3,-12,23,-22.9,-25.6,25.5,5,-2),
(8700,418.4,30.4,0.07,29.9,8.6,0.58,16.8,-16.6,14.5,-12.2,23.4,-23.2,-26,26,5,-2),
(8800,426.6,30.9,0.07,30.4,8.8,0.59,16.9,-16.7,14.8,-12.4,23.8,-23.6,-26.5,26.5,6,-3),
(8900,435,31.5,0.07,30.9,9,0.59,17,-16.8,15.1,-12.6,24.1,-23.9,-26.9,27,6,-3),
(9000,443.6,32,0.07,31.4,9.2,0.6,17.1,-16.9,15.3,-12.8,24.5,-24.3,-27.4,27.5,7,-4),
(9100,452.4,32.6,0.07,31.9,9.5,0.6,17.2,-16.9,15.6,-13,24.8,-24.6,-27.9,28,8,-5),
(9200,461.4,33.1,0.07,32.5,9.7,0.61,17.3,-17,15.9,-13.3,25.2,-24.9,-28.4,28.6,8,-5),
(9300,470.6,33.7,0.06,33.1,10,0.61,17.4,-17.1,16.2,-13.5,25.5,-25.2,-28.9,29.1,9,-6),
(9400,480.1,34.3,0.06,33.6,10.2,0.62,17.5,-17.2,16.4,-13.7,25.8,-25.5,-29.4,29.7,10,-7),
(9500,489.9,34.9,0.06,34.2,10.5,0.63,17.6,-17.4,16.7,-13.9,26.1,-25.8,-29.9,30.3,11,-7),
(9600,500,35.5,0.06,34.8,10.8,0.63,17.8,-17.5,17,-14.1,26.4,-26.1,-30.4,30.9,11,-8),
(9700,510.5,36.1,0.06,35.4,11.1,0.64,17.9,-17.6,17.3,-14.3,26.7,-26.4,-30.9,31.5,12,-9),
(9800,521.4,36.8,0.06,36.1,11.4,0.65,18,-17.7,17,-14.5,26.9,-26.6,-31.5,32.1,13,-9),
(9900,532.7,37.5,0.06,36.7,11.7,0.65,18.1,-17.8,17.9,-14.7,27.2,-26.9,-32,32.7,14,-10),
(10000,544.5,38.2,0.06,37.4,12.1,0.66,18.3,-17.9,18.3,-14.9,27.4,-27.2,-32.6,33.4,14,-11),
(10100,556.8,38.9,0.06,38.2,12.5,0.67,18.4,-18,18.3,-15.1,27.6,-27.4,-33.1,34.1,15,-11),
(10200,569.9,39.7,0.06,38.9,12.9,0.68,18.5,-18.1,18.3,-15.3,27.9,-27.7,-33.7,34.8,16,-12),
(10300,583.8,40.5,0.05,39.7,13.3,0.69,18.7,-18.2,18.3,-15.5,28.1,-27.9,-34.3,35.5,17,-13),
(10400,598.6,41.3,0.05,40.5,13.8,0.7,18.8,-18.4,18.3,-15.7,28.2,-28.1,-34.9,36.3,18,-14),
(10500,614.7,42.3,0.05,41.4,14.4,0.71,18.9,-18.5,18.3,-15.9,28.4,-28.4,-35.5,37.2,19,-15),
(10600,632.5,43.3,0.05,42.4,15,0.72,19.1,-18.6,18.3,-16.1,28.5,-28.6,-36.2,38.3,20,-16),
(10700,652.5,44.4,0.05,43.5,15.7,0.74,19.3,-18.7,18.3,-16.3,28.5,-28.8,-36.8,38.3,21,-16),
(10800,676.2,45.7,0.05,44.8,16.6,0.75,19.5,-18.9,18.3,-16.5,28.5,-29,-37.5,38.3,22,-18),
(10900,706.6,47.4,0.05,46.4,17.8,0.77,19.5,-19,18.3,-16.7,28.5,-29.1,-38.2,38.3,24,-19),
(11000,764.2,50.4,0.04,49.4,20.3,0.77,19.5,-19.2,18.3,-16.9,28.5,-29.3,-38.9,38.3,24,-20),



    ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == gac_mesafe:
                    nisangah5bh = round(nsg[1], 1)
                    print("Nişangah:", nisangah5bh)
            for dgl in interpolasyonlu_veri1:
                if dgl[0] == gac_mesafe:
                    dogalyandz = round(dgl[5],3)
                    print("Doğal Yan Düzeltmesi:", dogalyandz)
            for tsg in interpolasyonlu_veri1:
                if tsg[0] == gac_mesafe:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            metro_yan_duzeltmesi2 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi en yakın 1 milyem99999999999999:", metro_yan_duzeltmesi2)
            print("Arzin Dnüşü:,",arzindonusuyanduzeltmesi)
            print("Rüzgaryan dzl---------------------:,",ruzgar_yan_duzeltmesi)

            
            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi5bh = round(metro_yan_duzeltmesi2 + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi33333333333335bh:", toplam_yan_duzeltmesi5bh)
            gac_yan_duzeltmesi2 = round(toplam_yan_duzeltmesi5bh - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi2)
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]



            # Veri kümesi
            data1 = [
                (4000,181,159,0.014,-0.012),
(4500,213,212,0.019,-0.015),
(5000,246,275,0.024,-0.02),
(5500,280,347,0.031,-0.026),
(6000,316,430,0.041,-0.034),
(6500,353,525,0.054,-0.045),
(7000,393,633,0.071,-0.059),
(7500,434,757,0.094,-0.078),
(8000,479,899,0.125,-0.103),
(8500,526,1063,0.17,-0.138),
(9000,577,1255,0.237,-0.188),
(9500,633,1486,0.346,-0.263),
(10000,697,1773,0.558,-0.388),
(10500,776,2165,1.286,-0.646),
(11000,930,3054,1.286,-2.11),


]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == gac_mesafe:
                    dusus_acisi = round(dusus_acisi[1],3)
                    # print("Düşüş Açısı:", dusus_acisi)

            for tepe_yuksekligi in interpolasyonlu_veri1:
                if tepe_yuksekligi[0] == gac_mesafe:
                    tepe_yuksekligi = round(tepe_yuksekligi[2],3)
                    # print("Tepe Yüksekliği:", tepe_yuksekligi)

            for dtac_arti1 in interpolasyonlu_veri1:
                if dtac_arti1[0] == mesafe:
                    dtac_arti1 = round(dtac_arti1[3],3)
                    print("+1 DTAÇ:", dtac_arti1)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4],3)
                    print("-1 DTAÇ:", dtac_eksi1)
    ############################################################################################################
            """METRO DÜZELTMELERİNİ HESAPLAMAK GEREKİYOR"""
            # GTAC 1 inci oonüs için
            bt_rakimi = lne_1B_obus_rakim

            milyem_sabit_dz = 1.0186

            hedef_batarya_rakim_farki = (rakim - bt_rakimi)

            dtac = (hedef_batarya_rakim_farki / (plan_mesafesi_1B_obüs / 1000)) * milyem_sabit_dz

            # Burada dtac_arti1 ve dtac_eksi1 değerlerini sadece bir kez tanımlayın
            dtac_arti1 = 0
            dtac_eksi1 = 0

            # dtac_arti1 ve dtac_eksi1 değerlerini doğru şekilde ata
            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_arti1 = round(dusus_acisi[3], 3)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_eksi1 = round(dusus_acisi[4], 3)

            ttac = 0
            print("Mesafe",mesafe)
            print("dtacccccccccc",dtac)
            print("dtacccccccccc++++",dtac_arti1)
            print("dtacccccccccc",dtac_eksi1)
            if dtac < 0:
                ttac = int(dtac * dtac_arti1)
                tac = dtac + ttac
                #print("-TAÇ", tac)

            if dtac > 0:

                ttac = int(dtac * dtac_eksi1)
                tac = dtac + ttac
                print("+TAÇ", tac)
    #########################################################################################################
            """YÜKSELİŞİ BUL"""
            # tac = 0
            yukselis5bh = round(nisangah5bh+ + tac,1)
            self.ui.sonuc_yukselis_1B.setText(str(yukselis5bh))
    #########################################################################################################
            yan_1B_5bh = round(yan_1B+toplam_yan_duzeltmesi5bh)####
            self.ui.sonuc_yan_1B.setText(str(yan_1B_5bh))####
            self.ui.lne_barut_hakki.setText(str(secilen_barut_hakki))
            self.ui.sonuc_barut_hakki_1B.setText(str(secilen_barut_hakki))

            if 2200 < yan_1B_5bh < 3000:
                self.ui.sonuc_yan_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_yukselis_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_barut_hakki_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_tapa_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_istikamet_acisi_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_mesafe_1B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.lne_sonuc_mesafe.setText(str(mesafe))
                self.ui.lne_sonuc_barut_hakki.setText(str(secilen_barut_hakki))
                self.ui.lne_sonuc_yan.setText(str(yan_1B_5bh))
                self.ui.lne_sonuc_yukselis.setText(str(yukselis5bh))
                self.ui.lne_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
                # self.ui.lne_sonuc_bir_ml_degisiklil.setText(str(birmildegisim5bh)) 
                self.ui.lne_sonuc_yuz_m.setText(str(mesafe/1000)) 
                """SONRA YAPACAĞIM"""
                # self.ui.lbl_sonuc_barut_hakki_2.setText(str(secilen_barut_hakki))
                # self.ui.lbl_sonuc_yan_2.setText(str(yan_2B_5bh))
                # self.ui.lbl_sonuc_yukselis_2.setText(str(yukselis5bh))
                # self.ui.lbl_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
            else:
                self.ui.sonuc_yan_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_yukselis_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_barut_hakki_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_tapa_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_istikamet_acisi_1B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_mesafe_1B.setStyleSheet("color: grey; font-weight: normal")
            self._doldur_sonuc("1", toplam_yan_duzeltmesi5bh, yukselis5bh, locals().get("birmildegisim5bh", 1), locals())
        
      
##########################################################################################################

 
        if secilen_barut_hakki == 5:
            plan_mesafesi = plan_mesafesi_2B_obüs
            batarya_rakimi = lne_2B_obus_rakim
            bt_rakimi_1 = lne_2B_obus_rakim
            hedef_rakimi = lne_hedef_rakim
            ahia = lne_2B_obus_ahia
            paralanma_yuksekliği = paralanma_yuksekliği
            mermi_kare_agirligi = mermi_kare_agirligi
            ilk_hiz_farki = lne_2B_obus_ihf
            atis_istikameti = batarya_hedef_İA_2
            atis_istikameti_1 = batarya_hedef_İA_2
            mevzi_hiz_degisikligi = 0
            mevzi_yan_duzeltmesi = 0
            mevzi_ts_duzeltmesi = 0
            tapasaniyesi = 1
            barut_isisi_str = self.ui.lne_barut_isisi.text()
            barut_isisi = float(barut_isisi_str) if barut_isisi_str else 75
            # barut_isisi = self.ui.lne_barut_isisi.text() if self.ui.lne_barut_isisi.text() else 75
            print("Barutttt İSİSİ",barut_isisi)
                # 1. GİRİŞ MESAFESİNİ BUL / MANUEL HASSASİYETİ

            # mesafe = 1866
            # en_yakin_10_metre = round(mesafe / 10) * 10
            # print("1836'nın en yakın 10 metreye çevrilmiş hali:", en_yakin_10_metre)

            # BATARYA RAKIMINI EN YAKIN 10 M YE ÇEVİR.
            batarya_rakimi = round(batarya_rakimi/10)*10
            # PARALANMA NOKTASININ RAKIMINI BUL.
            paralanma_noktasi_rakimi = hedef_rakimi + paralanma_yuksekliği
            # HEDEF-BATARYA YÜKSEKLİK FARKINI BUL
            yukseklik_farki = paralanma_noktasi_rakimi-batarya_rakimi
            yukseklik_farki_1 = paralanma_noktasi_rakimi-bt_rakimi_1
            # yukseklik_farki = round(yukseklik_farki/100)*100
            ################# B cetvelini yüklüyoruz.################
            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (100, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (200, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (300, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (400, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (500, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (600, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (700, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (800, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (900, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1000, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1100, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1200, 0, 0, 0, 0, 0, 0, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1300, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1400, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1500, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1600, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1700, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1800, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (1900, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2000, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2100, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2200, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2300, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2400, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2500, 0, 0, 0, 0, 0, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2600, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2700, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2800, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (2900, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (3000, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (3100, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3),
                (3200, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3300, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3400, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3500, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3600, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3700, 0, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3),
                (3800, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3),
                (3900, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4000, 0, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4100, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4200, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4300, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3),
                (4400, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4500, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4600, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4700, 1, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3),
                (4800, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3),
                (4900, 1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3),
                (5000, 1, 1, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3),
                (5100, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3),
                (5200, 1, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4),
                (5300, 1, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4),
                (5400, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4),
                (5500, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4),
                (5600, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 4, 4),
                (5700, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4),
                (5800, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4),
                (5900, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4),
                (6000, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4),
                (6100, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4),
                (6200, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 5),
                (6300, 2, 2, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5),
                (6400, 2, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5),
                (6500, 2, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 5, 5, 5),
                (6600, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5),
                (6700, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5),
                (6800, 3, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5),
                (6900, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5),
                (7000, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5),
                (7100, 3, 3, 3, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, 5),
                (7200, 3, 3, 4, 4, 4, 5, 5, 5, 5, 5, 5, 5, 5, 5, 5),
            ]

            data2 = [
                (100, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0),
                (200, 0, 0, 0, 0, 0, 1, 2, 3, 5, 7, 8, 10, 12, 14, 0),
                (300, 0, 0, 0, 0, 0, 1, 3, 5, 8, 10, 13, 16, 19, 22, 25),
                (400, 0, 0, 0, 0, 0, 2, 4, 7, 10, 14, 17, 21, 25, 28, 32),
                (500, 0, 0, 0, 0, 0, 3, 6, 9, 13, 17, 21, 26, 30, 35, 40),
                (600, 0, 0, 0, 0, 0, 3, 7, 11, 15, 20, 25, 30, 36, 41, 47),
                (700, 0, 0, 0, 0, 0, 4, 8, 13, 18, 23, 29, 35, 41, 47, 54),
                (800, 0, 0, 0, 0, 0, 4, 9, 14, 20, 26, 32, 39, 46, 53, 61),
                (900, 0, 0, 0, 0, 0, 5, 10, 16, 22, 29, 36, 43, 51, 59, 68),
                (1000, 0, 0, 0, 0, 0, 5, 11, 18, 24, 32, 39, 47, 56, 65, 74),
                (1100, 0, 0, 0, 0, 0, 6, 12, 19, 27, 35, 43, 52, 61, 71, 81),
                (1200, 0, 0, 0, 0, 0, 7, 14, 21, 29, 38, 46, 56, 66, 76, 87),
                (1300, 0, 0, 0, 0, 0, 7, 15, 23, 31, 40, 50, 60, 71, 82, 94),
                (1400, 0, 0, 0, -7, 0, 8, 16, 24, 34, 43, 54, 64, 76, 88, 100),
                (1500, 0, 0, 0, -8, 0, 8, 17, 26, 36, 46, 57, 69, 81, 93, 107),
                (1600, 0, 0, 0, -8, 0, 9, 18, 28, 38, 49, 61, 73, 86, 99, 113),
                (1700, 0, 0, 0, -9, 0, 9, 19, 30, 41, 52, 64, 77, 91, 105, 120),
                (1800, 0, 0, 0, -9, 0, 10, 20, 31, 43, 55, 68, 82, 96, 111, 126),
                (1900, 0, 0, -19, -10, 0, 10, 22, 33, 45, 58, 72, 86, 101, 116, 133),
                (2000, 0, 0, -20, -10, 0, 11, 23, 35, 48, 61, 75, 90, 106, 122, 139),
                (2100, 0, 0, -22, -11, 0, 12, 24, 37, 50, 64, 79, 95, 111, 128, 146),
                (2200, 0, 0, -23, -12, 0, 12, 25, 38, 53, 67, 83, 99, 116, 134, 153),
                (2300, 0, -35, -24, -12, 0, 13, 26, 40, 55, 71, 87, 104, 122, 140, 160),
                (2400, 0, -36, -25, -13, 0, 13, 27, 42, 58, 74, 91, 108, 127, 146, 167),
                (2500, 0, -38, -26, -13, 0, 14, 29, 44, 60, 77, 95, 113, 132, 152, 174),
                (2600, 0, -40, -27, -14, 0, 15, 30, 46, 63, 80, 98, 118, 138, 159, 181),
                (2700, -54, -42, -28, -15, 0, 15, 31, 48, 65, 83, 102, 122, 143, 165, 188),
                (2800, -57, -43, -30, -15, 0, 16, 32, 50, 68, 87, 106, 127, 149, 171, 195),
                (2900, -59, -45, -31, -16, 0, 16, 34, 52, 70, 90, 111, 132, 154, 178, 202),
                (3000, -61, -47, -32, -16, 0, 17, 35, 54, 73, 93, 115, 137, 160, 184, 210),
                (3100, -64, -49, -33, -17, 0, 18, 36, 56, 76, 97, 119, 142, 166, 191, 218),
                (3200, -69, -53, -36, -18, 0, 19, 39, 60, 79, 100, 123, 147, 172, 198, 225),
                (3300, -66, -51, -34, -18, 0, 18, 38, 58, 81, 104, 127, 152, 178, 205, 233),
                (3400, -71, -54, -37, -19, 0, 20, 40, 62, 84, 107, 132, 157, 184, 212, 241),
                (3500, -74, -56, -38, -20, 0, 20, 42, 64, 87, 111, 136, 163, 190, 219, 249),
                (3600, -76, -58, -40, -20, 0, 21, 43, 66, 90, 115, 141, 168, 197, 226, 258),
                (3700, -79, -60, -41, -21, 0, 22, 44, 68, 93, 119, 145, 174, 203, 234, 266),
                (3800, -81, -62, -42, -22, 0, 22, 46, 70, 96, 122, 150, 179, 210, 241, 275),
                (3900, -84, -64, -44, -22, 0, 23, 47, 73, 99, 126, 155, 185, 216, 249, 284),
                (4000, -87, -66, -45, -23, 0, 24, 49, 75, 102, 130, 160, 191, 223, 257, 293),
                (4100, -89, -68, -46, -24, 0, 25, 50, 77, 105, 134, 165, 197, 230, 265, 302),
                (4200, -92, -70, -48, -24, 0, 25, 52, 80, 108, 139, 170, 203, 238, 274, 312),
                (4300, -95, -73, -49, -25, 0, 26, 54, 82, 112, 143, 175, 209, 245, 282, 322),
                (4400, -98, -75, -51, -26, 0, 27, 55, 85, 115, 147, 181, 216, 253, 291, 332),
                (4500, -101, -77, -52, -27, 0, 28, 57, 87, 119, 152, 186, 222, 260, 300, 342),
                (4600, -104, -79, -54, -27, 0, 29, 59, 90, 122, 156, 192, 229, 269, 310, 353),
                (4700, -107, -82, -55, -28, 0, 29, 60, 92, 126, 161, 198, 236, 277, 320, 365),
                (4800, -110, -84, -57, -29, 0, 30, 62, 95, 130, 166, 204, 244, 286, 330, 376),
                (4900, -113, -86, -59, -30, 0, 31, 64, 98, 134, 171, 210, 251, 295, 340, 389),
                (5000, -116, -89, -60, -31, 0, 32, 66, 101, 138, 176, 217, 259, 304, 351, 401),
                (5100, -120, -91, -62, -32, 0, 33, 68, 104, 142, 182, 223, 267, 314, 363, 415),
                (5200, -123, -94, -64, -33, 0, 34, 70, 107, 146, 187, 230, 276, 324, 375, 429),
                (5300, -126, -97, -66, -34, 0, 35, 72, 110, 151, 193, 238, 285, 335, 387, 444),
                (5400, -130, -100, -68, -35, 0, 36, 74,
                114, 155, 199, 245, 294, 346, 401, 460),
                (5500, -134, -102, -70, -36, 0, 37, 76,
                117, 160, 205, 253, 304, 358, 415, 477),
                (5600, -138, -105, -72, -37, 0, 38, 79,
                121, 165, 212, 262, 314, 370, 430, 495),
                (5700, -142, -108, -74, -38, 0, 40, 81,
                125, 171, 219, 271, 325, 384, 447, 515),
                (5800, -146, -112, -76, -39, 0, 41, 84,
                129, 176, 227, 280, 337, 399, 465, 537),
                (5900, -150, -115, -78, -40, 0, 42, 86,
                133, 182, 235, 290, 350, 414, 485, 562),
                (6000, -154, -118, -81, -41, 0, 43, 89,
                137, 189, 243, 301, 364, 432, 507, 590),
                (6100, -159, -122, -83, -43, 0, 45, 92,
                142, 195, 252, 313, 379, 452, 532, 625),
                (6200, -164, -126, -86, -44, 0, 46, 95,
                147, 203, 262, 327, 397, 475, 563, 669),
                (6300, -169, -130, -89, -45, 0, 48, 99,
                153, 211, 273, 342, 417, 502, 603, 739),
                (6400, -174, -134, -91, -47, 0, 50, 103, 159, 220, 286, 359, 441, 538, 666, 0),
                (6500, -180, -138, -95, -49, 0, 52, 107, 166, 230, 301, 380, 472, 594, 0, 0),
                (6600, -185, -143, -98, -50, 0, 54, 111, 174, 242, 318, 407, 523, 0, 0, 0),
                (6700, -192, -148, -101, -52, 0, 56, 116, 183, 256, 341, 452, 0, 0, 0, 0),
                (6800, -199, -153, -105, -54, 0, 59, 123, 193, 275, 381, 0, 0, 0, 0, 0),
                (6900, -206, -159, -110, -57, 0, 62, 130, 208, 310, 0, 0, 0, 0, 0, 0),
                (7000, -214, -166, -115, -60, 0, 66, 140, 238, 0, 0, 0, 0, 0, 0, 0),
                (7100, -223, -173, -120, -63, 0, 71, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (7200, -233, -182, -127, -67, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),

            ]


            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            interpolasyonlu_veri2 = veri_kumesi_interpolasyonu(data2)
            # print(interpolasyonlu_veri2)

            hedefin_toptan_yuksekligi = yukseklik_farki
            hedefin_toptan_yuksekligi = round(hedefin_toptan_yuksekligi/100)*100
            print("Hedefin Toptan Yüksekliği: ", hedefin_toptan_yuksekligi)
            mesafe = plan_mesafesi
            print("mesafe = ", mesafe)

            if hedefin_toptan_yuksekligi == -400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[1])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[2])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[3])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[4])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -0:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[5])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[6])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[7])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[8])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[9])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 500:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[10])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 600:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[11])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 700:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 800:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 900:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[13])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 1000:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[14])
                        print("Seçilen Bölge:", bolge)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")


            # TXT dosyasının adı ve dosya yolu
            dosya_adı = METRO_DOSYA

            # İlk satırı atlayarak dosyanın geri kalanını yeni bir liste olarak alın
            metrorpr = []
            if not os.path.isfile(dosya_adı):
                QMessageBox.warning(self, "Metro", "METRAP.txt bulunamadı.")
                return None
            with open(dosya_adı, "r") as dosya:
                for satır in dosya:
                    metrorpr.append(satır.strip())

            # Satırları yan yana birleştirerek elde etmek için join() yöntemini kullanıyoruz
            birlesik_veri = "\n".join(metrorpr[1:])

            rapor_tipi = birlesik_veri[:5]
            oktant = birlesik_veri[5:6]
            metro_koor = birlesik_veri[6:9] + " " + birlesik_veri[9:12]
            tarih = birlesik_veri[13:15]
            saat = birlesik_veri[15:17]
            dakika = str(int(birlesik_veri[17:18]) * 6)
            if dakika == "0":
                dakika = str(int(birlesik_veri[17:18]) * 6) + "0"
            ara = ":"
            saat_dakika = saat + ara + str(dakika)

            gecerlilik = birlesik_veri[18:19]
            metro_istasyonu_rakimi = str(int(birlesik_veri[19:22]) * 10)
            yogunluk = int(birlesik_veri[22:25]) / 10

            print("Rapor Tipi:", rapor_tipi)
            print("Oktant:", oktant)
            print("Metro Koordinat:", metro_koor)
            print("Tarih:", tarih)
            print("Saat:", saat_dakika)
            print("Geçerlilik:", gecerlilik)
            print("Metro İstasyonu Rakımı:", metro_istasyonu_rakimi)
            print("Yoğunluk:", yogunluk)


            if bolge == 0:
                ruzgar_istikameti = int(birlesik_veri[28:30]) * 100
                ruzgar_hizi = int(birlesik_veri[30:32])
                if birlesik_veri[32] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[32:35]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[32:35]) / 10
                if birlesik_veri[35] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[35:38]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[35:38]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)  # 39-40-41 3 say

            elif bolge == 1:
                ruzgar_istikameti = int(birlesik_veri[41:43]) * 100
                ruzgar_hizi = int(birlesik_veri[43:45])
                if birlesik_veri[45] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[45:48]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[45:48]) / 10
                if birlesik_veri[48] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[48:51]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[48:51]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)

            elif bolge == 2:
                ruzgar_istikameti = int(birlesik_veri[54:56])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[56:58])  # 3
                if birlesik_veri[58] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[58:61])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[58:61])/10  # 4
                if birlesik_veri[61] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[61:64])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[61:64])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 3:
                ruzgar_istikameti = int(birlesik_veri[67:69])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[69:71])  # 3
                if birlesik_veri[71] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[71:74])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[71:74])/10  # 4
                if birlesik_veri[74] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[74:77])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[74:77])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 4:
                ruzgar_istikameti = int(birlesik_veri[80:82])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[82:84])  # 3
                if birlesik_veri[84] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[84:87])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[84:87])/10  # 4
                if birlesik_veri[87] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[87:90])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[87:90])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 5:
                ruzgar_istikameti = int(birlesik_veri[93:95])*100  # 3
                ruzgar_hizi = int(birlesik_veri[95:97])  # 3
                if birlesik_veri[97] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[97:100])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[97:100])/10  # 4
                if birlesik_veri[100] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[100:103])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[100:103])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 6:
                ruzgar_istikameti = int(birlesik_veri[106:108])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[108:110])  # 3
                if birlesik_veri[110] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[110:113])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[110:113])/10  # 4
                if birlesik_veri[113] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[113:116])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[113:116])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 7:
                ruzgar_istikameti = int(birlesik_veri[119:121])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[121:123])  # 3
                if birlesik_veri[123] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[123:126])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[123:126])/10  # 4
                if birlesik_veri[126] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[126:129])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[126:129])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 8:
                ruzgar_istikameti = int(birlesik_veri[132:134])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[134:136])  # 3
                if birlesik_veri[136] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[136:139])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[136:139])/10  # 4
                if birlesik_veri[139] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[139:142])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[139:142])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 9:
                ruzgar_istikameti = int(birlesik_veri[145:147])*100  # 3
                ruzgar_hizi = int(birlesik_veri[147:149])  # 3
                if birlesik_veri[149] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[149:152])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[149:152])/10  # 4
                if birlesik_veri[152] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[152:155])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[152:155])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 10:
                ruzgar_istikameti = int(birlesik_veri[158:160])*100  # 3
                ruzgar_hizi = int(birlesik_veri[160:162])  # 3
                if birlesik_veri[162] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[162:165])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[162:165])/10  # 4
                if birlesik_veri[165] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[165:168])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[165:168])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)


            # TAMAMLAYICI MESAFEYİ BULUYORUZ.
            if hedefin_toptan_yuksekligi == -400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[1])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[2])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[3])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[4])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -0:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[5])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[6])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[7])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[8])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[9])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 500:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[10])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 600:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[11])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 700:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 800:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 900:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[13])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 1000:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[14])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")

            bt_metro_ist_rakim_farki = batarya_rakimi - int(metro_istasyonu_rakimi)

            giris_mesafesi = round((int(mesafe) + int(tamamlayici_mesafe))/100)*100
            print("Giriş Mesafesi :", giris_mesafesi)

            ruzgarin_plan_istikameti_100 = ruzgar_istikameti - atis_istikameti

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            isi_dzl = [
                (-390, 0.9),
                (-380, 0.9),
                (-370, 0.9),
                (-360, 0.8),
                (-350, 0.8),
                (-340, 0.8),
                (-330, 0.8),
                (-320, 0.7),
                (-310, 0.7),
                (-300, 0.7),
                (-290, 0.7),
                (-280, 0.7),
                (-270, 0.7),
                (-260, 0.6),
                (-250, 0.6),
                (-240, 0.6),
                (-230, 0.6),
                (-220, 0.5),
                (-210, 0.5),
                (-200, 0.5),
                (-190, 0.4),
                (-180, 0.4),
                (-170, 0.4),
                (-160, 0.3),
                (-150, 0.3),
                (-140, 0.3),
                (-130, 0.3),
                (-120, 0.2),
                (-110, 0.2),
                (-100, 0.2),
                (-90, 0.2),
                (-80, 0.2),
                (-70, 0.2),
                (-60, 0.1),
                (-50, 0.1),
                (-40, 0.1),
                (-30, 0.1),
                (-20, 0),
                (-10, 0),
                (0, 0),
                (10, 0),
                (20, 0),
                (30, -0.1),
                (40, -0.1),
                (50, -0.1),
                (60, -0.1),
                (70, -0.2),
                (80, -0.2),
                (90, -0.2),
                (100, -0.2),
                (110, -0.2),
                (120, -0.2),
                (130, -0.3),
                (140, -0.3),
                (150, -0.3),
                (160, -0.3),
                (170, -0.4),
                (180, -0.4),
                (190, -0.4),
                (200, -0.5),
                (210, -0.5),
                (220, -0.5),
                (230, -0.6),
                (240, -0.6),
                (250, -0.6),
                (260, -0.6),
                (270, -0.7),
                (280, -0.7),
                (290, -0.7),
                (300, -0.7),
                (310, -0.7),
                (320, -0.7),
                (330, -0.8),
                (340, -0.8),
                (350, -0.8),
                (360, -0.8),
                (370, -0.9),
                (380, -0.9),
                (390, -0.9),

            ]


            isi_data5bh = []
            for i in range(len(isi_dzl)-1):
                start = isi_dzl[i]
                end = isi_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    isi_data5bh.append((int(x), y))

            isi_data5bh.append(isi_dzl[-1])


            yogunluk_dzl = [
                (-390, 3.9),
                (-380, 3.8),
                (-370, 3.7),
                (-360, 3.6),
                (-350, 3.5),
                (-340, 3.4),
                (-330, 3.3),
                (-320, 3.2),
                (-310, 3.1),
                (-300, 3),
                (-290, 2.9),
                (-280, 2.8),
                (-270, 2.7),
                (-260, 2.6),
                (-250, 2.5),
                (-240, 2.4),
                (-230, 2.3),
                (-220, 2.2),
                (-210, 2.1),
                (-200, 2),
                (-190, 1.0),
                (-180, 1.8),
                (-170, 1.7),
                (-160, 1.6),
                (-150, 1.5),
                (-140, 1.4),
                (-130, 1.3),
                (-120, 1.2),
                (-110, 1.1),
                (-100, 1),
                (-90, 0.9),
                (-80, 0.8),
                (-70, 0.7),
                (-60, 0.6),
                (-50, 0.5),
                (-40, 0.4),
                (-30, 0.3),
                (-20, 0.2),
                (-10, 0.1),
                (0, 0),
                (10, -0.1),
                (20, -0.2),
                (30, -0.3),
                (40, -0.4),
                (50, -0.5),
                (60, -0.6),
                (70, -0.7),
                (80, -0.8),
                (90, -0.9),
                (100, -1),
                (110, -1.1),
                (120, -1.2),
                (130, -1.3),
                (140, -1.4),
                (150, -1.5),
                (160, -1.6),
                (170, -1.7),
                (180, -1.8),
                (190, -1.9),
                (200, -2),
                (210, -2.1),
                (220, -2.2),
                (230, -2.3),
                (240, -2.4),
                (250, -2.5),
                (260, -2.6),
                (270, -2.7),
                (280, -2.8),
                (290, -2.9),
                (300, -3),
                (310, -3.1),
                (320, -3.2),
                (330, -3.3),
                (340, -3.4),
                (350, -3.5),
                (360, -3.6),
                (370, -3.7),
                (380, -3.8),
                (390, -3.9),
            ]

            yogunluk_data5bh = []

            for i in range(len(yogunluk_dzl)-1):
                start = yogunluk_dzl[i]
                end = yogunluk_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    yogunluk_data5bh.append((int(x), y))

            yogunluk_data5bh.append(yogunluk_dzl[-1])
            # print(yogunluk_data5bh)


            for i in isi_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    isi_duzeltmesi = i[1]
                    print("Isı Düzeltmesi :", isi_duzeltmesi)


            for i in yogunluk_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    yogunluk_duzeltmesi = i[1]
                    print("Yoğunluk Düzeltmesi", yogunluk_duzeltmesi)

            duzeltilmis_sıcaklik_degeri = isi_duzeltmesi + hava_sicakligi
            print("Düzeltilmiş Isı Değeri:", duzeltilmis_sıcaklik_degeri)

            duzeltilmis_yogunluk_degeri = yogunluk_duzeltmesi + hava_yogunlugu
            print("Düzeltilmiş Yoğunluk Değeri:", duzeltilmis_yogunluk_degeri)

            # RUZGAR BİLEŞENLERİ

            atis_istikameti = round(atis_istikameti/100)*100

            if ruzgar_istikameti < atis_istikameti:
                ruzgar_istikameti = ruzgar_istikameti + 6400
            else:
                ruzgar_istikameti = ruzgar_istikameti
            print("ruzgar istikameti", ruzgar_istikameti)

            print("atış istikameti", atis_istikameti)

            ruzgarin_plan_istikameti = ruzgar_istikameti - atis_istikameti

            print("Rüzgarın Plan İstikameti:", ruzgarin_plan_istikameti)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            data1 = [
                (0, 0, 1),
                (100, -0.1, 0.99),
                (200, -0.2, 0.98),
                (300, -0.29, 0.96),
                (400, -0.38, 0.92),
                (500, -0.47, 0.88),
                (600, -0.56, 0.83),
                (700, -0.63, 0.77),
                (800, -0.71, 0.71),
                (900, -0.77, 0.63),
                (1000, -0.83, 0.56),
                (1100, -0.88, 0.47),
                (1200, -0.92, 0.38),
                (1300, -0.96, 0.29),
                (1400, -0.98, 0.2),
                (1500, -0.99, 0.1),
                (1600, -1, 0),
                (1700, -0.99, -0.1),
                (1800, -0.98, -0.2),
                (1900, -0.96, -0.29),
                (2000, -0.92, -0.38),
                (2100, -0.88, -0.47),
                (2200, -0.83, -0.56),
                (2300, -0.77, -0.63),
                (2400, -0.71, -0.71),
                (2500, -0.63, -0.77),
                (2600, -0.56, -0.83),
                (2700, -0.47, -0.88),
                (2800, -0.38, -0.92),
                (2900, -0.29, -0.96),
                (3000, -0.2, -0.98),
                (3100, -0.1, -0.99),
                (3200, 0, -1),
                (3300, 0.1, -0.99),
                (3400, 0.2, -0.98),
                (3500, 0.29, -0.96),
                (3600, 0.38, -0.92),
                (3700, 0.47, -0.88),
                (3800, 0.56, -0.83),
                (3900, 0.63, -0.77),
                (4000, 0.71, -0.71),
                (4100, 0.77, -0.63),
                (4200, 0.83, -0.56),
                (4300, 0.88, -0.47),
                (4400, 0.92, -0.38),
                (4500, 0.96, -0.29),
                (4600, 0.98, -0.2),
                (4700, 0.99, -0.1),
                (4800, 1, 0),
                (4900, 0.99, 0.1),
                (5000, 0.98, 0.2),
                (5100, 0.96, 0.29),
                (5200, 0.92, 0.38),
                (5300, 0.88, 0.47),
                (5400, 0.83, 0.56),
                (5500, 0.77, 0.63),
                (5600, 0.71, 0.71),
                (5700, 0.63, 0.77),
                (5800, 0.56, 0.83),
                (5900, 0.47, 0.88),
                (6000, 0.38, 0.92),
                (6100, 0.29, 0.96),
                (6200, 0.2, 0.98),
                (6300, 0.1, 0.99),
                (6400, 0, 1),
            ]


            # veri_noktasi[0] = MESAFE
            # veri_noktasi[1] = YAN RÜZGARI
            # veri_noktasi[2] = MESAFE RÜZGARI

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Mesafe :", mesafe)

            for rzdlz in interpolasyonlu_veri1:
                if rzdlz[0] == ruzgarin_plan_istikameti:
                    yan_ruzgari_bileseni = rzdlz[1]
                    mesafe_ruzgari_bileseni = rzdlz[2]
                    print("Yan Rüzgarı Bileşeni:", yan_ruzgari_bileseni)
                    print("mesafe Rüzgarı Bileşeni:", mesafe_ruzgari_bileseni)


            # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (100, 5.6, 0, 00, 0.3, 0, 0, 0.7, -0.6, 0, 0, 0, 0, 0, 0, -1, 1),
                (200, 11.1, 0, 0, 0.7, 0, 0.01, 1.4, -1.2, 0, 0, 0, 0, 0, 0, -2, 2),
                (300, 16.6, 0, 0, 1, 0.1, 0.01, 2.1, -1.8, 0.1, 0, 0.1, 0, 0, 0, -4, 4),
                (400, 22.2, 0, 0, 1.3, 0.2, 0.02, 2.7, -2.4, 0.1, 0, 0.1, 0, -0.1, 0.1, -5, 5),
                (500, 27.8, 0, 0, 1.7, 0.3, 0.02, 3.4, -3, 0.2, 0, 0.2, 0, -0.1, 0.1, -6, 6),
                (600, 33.4, 2, 1.02, 2, 0.4, 0.02, 4.1, -3.6, 0.2, 0, 0.3, 0, -0.1, 0.1, -7, 7),
                (700, 39.1, 2.3, 0.87, 2.4, 0.4, 0.03, 4.7, -
                4.2, 0.3, -0.1, 0.4, -0.1, -0.1, 0.1, -8, 8),
                (800, 44.9, 2.6, 0.76, 2.7, 0.5, 0.03, 5.4, -
                4.8, 0.3, -0.1, 0.4, -0.1, -0.2, 0.2, -9, 9),
                (900, 50.6, 3, 0.68, 3, 0.6, 0.04, 6, -5.3,
                0.4, -0.1, 0.5, -0.1, -0.2, 0.2, -10, 10),
                (1000, 56.4, 3.3, 0.61, 3.4, 0.7, 0.04, 6.7, -
                5.9, 0.5, -0.1, 0.6, -0.1, -0.3, 0.3, -11, 11),
                (1100, 62.3, 3.7, 0.55, 3.7, 0.8, 0.04, 7.3, -
                6.4, 0.5, -0.1, 0.7, -0.1, -0.4, 0.4, -12, 13),
                (1200, 68.2, 4, 0.5, 4.1, 0.9, 0.05, 7.9, -
                7, 0.6, -0.2, 0.8, -0.1, -0.4, 0.4, -13, 14),
                (1300, 74.1, 4.4, 0.46, 4.4, 1, 0.05, 8.6, -
                7.5, 0.7, -0.2, 0.9, -0.1, -0.5, 0.5, -14, 15),
                (1400, 80.1, 4.8, 0.43, 4.8, 1.1, 0.06, 9.2, -
                8.1, 0.8, -0.2, 1, -0.1, -0.6, 0.6, -15, 16),
                (1500, 86.1, 5.1, 0.4, 5.2, 1.2, 0.06, 9.8, -
                8.6, 0.9, -0.2, 1.2, -0.2, -0.6, 0.6, -16, 17),
                (1600, 92.2, 5.5, 0.37, 5.5, 1.3, 0.06, 10.4, -
                9.1, 1, -0.3, 1.3, -0.2, -0.7, 0.7, -17, 17),
                (1700, 98.3, 5.8, 0.35, 5.9, 1.3, 0.07, 11.1, -
                9.7, 1.1, -0.3, 1.4, -0.2, -0.8, 0.8, -18, 18),
                (1800, 104.5, 6.2, 0.33, 6.2, 1.4, 0.07, 11.7, -
                10.2, 1.1, -0.3, 1.5, -0.2, -0.9, 0.9, -19, 19),
                (1900, 110.7, 6.6, 0.31, 6.6, 1.5, 0.08, 12.3, -
                10.7, 1.2, -0.3, 1.6, -0.2, -1, 1, -19, 20),
                (2000, 117, 6.9, 0.3, 7, 1.6, 0.08, 12.9, -
                11.2, 1.3, -0.4, 1.7, -0.2, -1.1, 1.1, -20, 21),
                (2100, 123.3, 7.3, 0.28, 7.3, 1.7, 0.08, 13.5, -
                11.8, 1.4, -0.4, 1.8, -0.2, -1.2, 1.3, -21, 22),
                (2200, 129.7, 7.7, 0.27, 7.7, 1.8, 0.09, 14.1, -
                12.3, 1.5, -0.4, 1.9, -0.2, -1.4, 1.4, -22, 23),
                (2300, 136.2, 8.1, 0.26, 8.1, 1.9, 0.09, 14.7, -
                12.8, 1.6, -0.5, 2, -0.2, -1.5, 1.5, -23, 24),
                (2400, 142.7, 8.4, 0.25, 8.5, 2.1, 0.1, 15.3, -
                13.3, 1.7, -0.5, 2.1, -0.2, -1.6, 1.6, -24, 25),
                (2500, 149.2, 8.8, 0.24, 8.9, 2.2, 0.1, 15.9, -
                13.8, 1.8, -0.6, 2.2, -0.2, -1.8, 1.8, -24, 26),
                (2600, 155.9, 9.2, 0.23, 9.2, 2.3, 0.11, 16.5, -
                14.3, 1.9, -0.6, 2.3, -0.2, -1.9, 1.9, -25, 26),
                (2700, 162.6, 9.6, 0.22, 9.6, 2.4, 0.11, 17.1, -
                14.9, 2, -0.6, 2.4, -0.2, -2, 2.1, -26, 27),
                (2800, 169.4, 10, 0.21, 10, 2.5, 0.11, 17.7, -
                15.4, 2.1, -0.7, 2.5, -0.2, -2.2, 2.2, -27, 28),
                (2900, 176.2, 10.4, 0.2, 10.4, 2.6, 0.12, 18.3, -
                15.9, 2.2, -0.7, 2.5, -0.2, -2.3, 2.4, -27, 29),
                (3000, 183.1, 10.8, 0.19, 10.8, 2.7, 0.12, 18.9, -
                16.4, 2.3, -0.8, 2.6, -0.2, -2.5, 2.5, -28, 29),
                (3100, 190.1, 11.2, 0.19, 11.2, 2.8, 0.13, 19.4, -
                16.9, 2.4, -0.8, 2.7, -0.2, -2.7, 2.7, -29, 30),
                (3200, 197.2, 11.6, 0.18, 11.6, 3, 0.13, 20, -
                17.4, 2.5, -0.9, 2.8, -0.2, -2.8, 2.9, -29, 31),
                (3300, 204.4, 12, 0.17, 12, 3.1, 0.14, 20.6, -
                17.9, 2.6, -0.9, 2.8, -0.1, -3, 3.1, -30, 32),
                (3400, 211.6, 12.4, 0.17, 12.4, 3.2, 0.14, 21.2, -
                18.4, 2.7, -1, 2.9, -0.1, -3.2, 3.3, -31, 32),
                (3500, 219, 12.8, 0.16, 12.8, 3.3, 0.14, 21.7, -
                18.9, 2.8, -1, 2.9, -0.1, -3.4, 3.4, -31, 33),
                (3600, 226.4, 13.2, 0.16, 13.2, 3.5, 0.15, 22.3, -
                19.4, 2.9, -1.1, 3, -0.1, -3.6, 3.6, -32, 34),
                (3700, 233.9, 13.6, 0.15, 13.7, 3.6, 0.15, 22.9, -
                19.9, 3, -1.1, 3, -0.1, -3.8, 3.8, -33, 34),
                (3800, 241.6, 14.1, 0.15, 14.1, 3.7, 0.16, 23.4, -
                20.4, 3.1, -1.2, 3.1, -0.1, -4, 4.1, -33, 35),
                (3900, 249.3, 14.5, 0.14, 14.5, 3.9, 0.16, 24, -
                20.9, 3.2, -1.2, 3.1, 0, -4.2, 4.3, -34, 36),
                (4000, 257.2, 14.9, 0.14, 15, 4, 0.17, 24.6, -
                21.4, 3.3, -1.3, 3.2, 0, -4.4, 4.5, -34, 36),
                (4100, 265.2, 15.4, 0.14, 15.4, 4.1, 0.17, 25.1, -
                21.9, 3.4, -1.4, 3.2, 0, -4.6, 4.7, -35, 37),
                (4200, 273.3, 15.8, 0.13, 15.9, 4.3, 0.18, 25.7, -
                22.3, 3.5, -1.4, 3.2, 0, -4.8, 4.9, -36, 38),
                (4300, 281.5, 16.3, 0.13, 16.3, 4.4, 0.18, 26.2, -
                22.8, 3.6, -1.5, 3.3, 0.1, -5.1, 5.2, -36, 38),
                (4400, 289.9, 16.7, 0.13, 16.8, 4.6, 0.19, 26.8, -
                23.3, 3.7, -1.6, 3.3, 0.1, -5.3, 5.4, -37, 39),
                (4500, 298.4, 17.2, 0.12, 17.2, 4.8, 0.19, 27.3, -
                23.8, 3.8, -1.6, 3.3, 0.1, -5.6, 5.7, -37, 39),
                (4600, 307, 17.7, 0.12, 17.7, 4.9, 0.2, 27.9, -
                24.3, 3.9, -1.7, 3.3, 0.1, -5.8, 5.9, -38, 40),
                (4700, 315.9, 18.2, 0.12, 18.2, 5.1, 0.2, 28.4, -
                24.8, 4, -1.8, 3.3, 0.2, -6, 6.2, -38, 40),
                (4800, 324.9, 18.6, 0.11, 18.7, 5.3, 0.21, 29, -
                25.3, 4.1, -1.8, 3.3, 0.2, -6.3, 6.5, -39, 41),
                (4900, 334, 19.1, 0.11, 19.1, 5.4, 0.21, 29.5, -
                25.7, 4.2, -1.9, 3.4, 0.2, -6.6, 6.7, -39, 41),
                (5000, 343.4, 19.6, 0.11, 19.6, 5.6, 0.22,
                30, -26.2, 4.3, -2, 3.4, 0.3, -6.8, 7, -40, 42),
                (5100, 353, 20.1, 0.11, 20.2, 5.8, 0.23, 30.6, -
                26.7, 4.4, -2.1, 3.4, 0.3, -7.1, 7.3, -40, 42),
                (5200, 362.8, 20.7, 0.1, 20.7, 6, 0.23, 31.1, -
                27.2, 4.6, -2.1, 3.3, 0.4, -7.4, 7.6, -40, 43),
                (5300, 372.8, 21.2, 0.1, 21.2, 6.2, 0.24, 31.6, -
                27.7, 4.7, -2.2, 3.3, 0.4, -7.7, 7.9, -41, 43),
                (5400, 383.1, 21.7, 0.1, 21.7, 6.4, 0.24, 32.2, -
                28.1, 4.8, -2.3, 3.3, 0.4, -8, 8.2, -41, 44),
                (5500, 393.7, 22.3, 0.1, 22.3, 6.6, 0.25, 32.7, -
                28.6, 4.9, -2.4, 3.3, 0.5, -8.3, 8.5, -42, 44),
                (5600, 404.6, 22.8, 0.09, 22.8, 6.9, 0.26, 33.2, -
                29.1, 5, -2.5, 3.3, 0.5, -8.6, 8.8, -42, 45),
                (5700, 415.8, 23.4, 0.09, 23.4, 7.1, 0.26, 33.7, -
                29.6, 5.1, -2.6, 3.3, 0.6, -8.9, 9.1, -42, 45),
                (5800, 427.3, 24, 0.09, 24, 7.3, 0.27, 34.2, -
                30, 5.2, -2.7, 3.2, 0.6, -9.2, 9.5, -43, 45),
                (5900, 439.3, 24.6, 0.09, 24.6, 7.6, 0.28, 34.7, -
                30.5, 5.3, -2.8, 3.2, 0.6, -9.5, 9.8, -43, 46),
                (6000, 451.7, 25.2, 0.08, 25.2, 7.9, 0.28, 35.2, -
                31, 5.4, -2.9, 3.2, 0.7, -9.8, 10.1, -43, 46),
                (6100, 464.6, 25.9, 0.08, 25.9, 8.2, 0.29, 35.7, -
                31.5, 5.5, -3, 3.1, 0.7, -10.2, 10.5, -44, 46),
                (6200, 478.1, 26.5, 0.08, 26.6, 8.5, 0.3, 36.2, -
                31.9, 5.6, -3.1, 3.1, 0.8, -10.5, 10.8, -44, 47),
                (6300, 492.2, 27.2, 0.08, 27.2, 8.8, 0.31, 36.7, -
                32.4, 5.8, -3.2, 3.1, 0.8, -10.9, 11.2, -44, 47),
                (6400, 507.1, 28, 0.08, 28, 9.2, 0.31, 37.2, -
                32.9, 5.9, -3.3, 3, 0.8, -11.2, 11.6, -44, 47),
                (6500, 522.9, 28.7, 0.07, 28.7, 9.6, 0.32,
                37.7, -33.3, 6, -3.4, 3, 0.8, -11.6, 12, -44, 47),
                (6600, 539.7, 29.5, 0.07, 29.5, 10, 0.33, 38.1, -
                33.8, 6.2, -3.6, 2.9, 0.9, -11.9, 12.4, -45, 48),
                (6700, 557.8, 30.4, 0.07, 30.4, 10.5, 0.34, 38.6, -
                34.2, 6.3, -3.7, 2.9, 0.9, -12.3, 12.8, -45, 48),
                (6800, 577.7, 31.3, 0.07, 31.3, 11, 0.35, 39, -
                34.7, 6.5, -3.8, 2.8, 0.9, -12.7, 13.2, -45, 48),
                (6900, 599.7, 32.3, 0.07, 32.3, 11.6, 0.37, 39, -
                35.2, 6.8, -4, 2.8, 0.9, -13.1, 13.6, -45, 48),
                (7000, 625, 33.5, 0.06, 33.5, 12.4, 0.38, 39, -
                35.6, 6.8, -4.1, 2.8, 0.9, -13.5, 14.1, -45, 48),
                (7100, 655.6, 34.8, 0.06, 34.8, 13.3, 0.4, 39, -
                36.1, 6.8, -4.3, 2.8, 0.8, -13.9, 14.7, -45, 48),
                (7200, 697.4, 36.6, 0.06, 36.6, 14.7, 0.42, 39, -
                36.5, 6.8, -4.5, 2.7, 0.8, -14.4, 14.7, -45, 48),


            ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == giris_mesafesi:
                    nisangah5bh = round(nsg[1], 3)
                    print("Nişangah:", nisangah5bh)

            for tsg in interpolasyonlu_veri1:
                if tsg[0] == giris_mesafesi:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            for ts in interpolasyonlu_veri1:
                if ts[0] == giris_mesafesi:
                    deltaTS = round(ts[3], 3)
                    # print("Delta Tapa Saniyesi:", deltaTS)


            for ucs in interpolasyonlu_veri1:
                if ucs[0] == giris_mesafesi:
                    ucussuresi = round(ucs[4], 3)
                    # print("Uçuş Süresi:", ucussuresi)

            for dgl in interpolasyonlu_veri1:
                if dgl[0] == giris_mesafesi:
                    dogalyandz = round(dgl[5], 3)
                    # print("Doğal Yan Düzeltmesi:", dogalyandz)

            for yrd in interpolasyonlu_veri1:
                if yrd[0] == giris_mesafesi:
                    yan_ruzgari_duzeltme_faktoru = round(yrd[6], 1)
                    # print("Yan Rüzgarı Düzeltmesi:", yanruzgaridz)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_eksilme = round(ihz[7], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimeksilme)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_martma = round(ihz[8], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimartma)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_bas = round(mrd[9], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (BAŞ):", mesaferuzgaridzl_bas)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_arka = round(mrd[10], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (ARKA):", mesaferuzgaridzl_arka)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_eksilme = round(hvs[11], 3)
                    # print("Hava Isısı Düzeltmesi (Eksilme):", havaisisi_eksilme)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_artma = round(hvs[12], 3)
                    # print("Hava Isısı Düzeltmesi (Artma):", havaisisi_artma)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_eksilme = round(hyd[13], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", havayogunlugu_eksilme)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_artma = round(hyd[14], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", havayogunlugu_artma)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_eksilme = round(mka[15], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", mermikareagirligi_eksilme)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_artma = round(mka[16], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", mermikareagirligi_artma)

            print("Yan Rüzgarı Düzeltme Faktörü:", yan_ruzgari_duzeltme_faktoru)

            mesafe_ruzgari = ruzgar_hizi * mesafe_ruzgari_bileseni
            print("mesafe ruzgari", mesafe_ruzgari)

            ruzgar_yan_duzeltmesi = round(
                (ruzgar_hizi * yan_ruzgari_bileseni * yan_ruzgari_duzeltme_faktoru), 1)
            print("Rüzgar Yan Düzeltmesi", ruzgar_yan_duzeltmesi)

            dogalyandz = round(dogalyandz, 1)
            print("Doğal Yan Düzeltmesi:", dogalyandz)

            atis_istikameti = round(atis_istikameti/400)*400

            # dünyanın dönmesi düzeltmesi


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (1000, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2),
                (1500, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2, 0.2),
                (2000, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3),
                (2500, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4),
                (3000, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5),
                (3500, 0.5, 0.5, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6),
                (4000, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.7, 0.7, 0.7),
                (4500, 0.7, 0.7, 0.7, 0.8, 0.8, 0.8, 0.8, 0.9, 0.9),
                (5000, 0.8, 0.8, 0.8, 0.8, 0.9, 0.9, 1, 1, 1),
                (5500, 0.9, 0.9, 0.9, 1, 1, 1.1, 1.1, 1.1, 1.2),
                (6000, 0.9, 1, 1, 1.1, 1.1, 1.2, 1.3, 1.3, 1.3),
                (6500, 1, 1, 1.1, 1.2, 1.3, 1.4, 1.5, 1.5, 1.6),
                (7000, 1.1, 1.1, 1.2, 1.3, 1.5, 1.6, 1.8, 1.9, 1.9),




            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            if atis_istikameti == 00 or atis_istikameti == 6400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[1], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 6000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[2], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 5600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[3], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 5200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[4], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 4800:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[5], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2000 or atis_istikameti == 4400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[6], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2400 or atis_istikameti == 4000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[7], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2800 or atis_istikameti == 3600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[8], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 3200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[9], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (1000, 0, -1, -3, -4, -5, -6, -6, -7, -7, 0, 1, 3, 4, 5, 6, 7, 7),
                (1500, 0, -2, -4, -6, -7, -8, -9, -10, -10, 0, 2, 4, 6, 7, 8, 10, 10),
                (2000, 0, -3, -5, -7, -9, -11, -12, -13, -13, 0, 3, 5, 7, 9, 11, 13, 13),
                (2500, 0, -3, -6, -9, -11, -13, -15, -16, -16, 0, 3, 6, 9, 11, 13, 16, 16),
                (3000, 0, -4, -7, -10, -13, -16, -17, -18, -19, 0, 4, 7, 10, 13, 16, 18, 19),
                (3500, 0, -4, -8, -12, -15, -18, -20, -21, -21, 0, 4, 8, 12, 15, 18, 21, 21),
                (4000, 0, -5, -9, -13, -17, -20, -22, -23, -24, 0, 5, 9, 13, 17, 20, 23, 24),
                (4500, 0, -5, -10, -14, -18, -21, -24, -25, -26, 0, 5, 10, 14, 18, 21, 25, 26),
                (5000, 0, -5, -11, -15, -19, -23, -25, -27, -28, 0, 5, 11, 15, 19, 23, 27, 28),
                (5500, 0, -6, -11, -16, -21, -24, -27, -29, -29, 0, 6, 11, 16, 21, 24, 29, 29),
                (6000, 0, -6, -12, -17, -21, -25, -28, -30, -30, 0, 6, 12, 17, 21, 25, 30, 30),
                (6500, 0, -6, -12, -17, -22, -26, -29, -30, -31, 0, 6, 12, 17, 22, 26, 30, 31),
                (7000, 0, -6, -12, -17, -22, -25, -28, -30, -31, 0, 6, 12, 17, 22, 25, 30, 31),



            ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)
            if atis_istikameti == 0 or atis_istikameti == 3200 or atis_istikameti == 6400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[1], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 200 or atis_istikameti == 3000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[2], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 2800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[3], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 600 or atis_istikameti == 2600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[4], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 2400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[5], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1000 or atis_istikameti == 2200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[6], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 2000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[7], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1400 or atis_istikameti == 1800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[8], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 1600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[9], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            if atis_istikameti == 3400 or atis_istikameti == 6200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[11], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3600 or atis_istikameti == 6000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[12], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3800 or atis_istikameti == 5800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[13], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4000 or atis_istikameti == 5600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[14], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4200 or atis_istikameti == 5400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[15], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4400 or atis_istikameti == 5200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[16], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4600 or atis_istikameti == 5000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[17], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4800 or atis_istikameti == 4800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[18], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            print("Dünyanın Dönmesi Düzeltmesi:", arzindonusuyanduzeltmesi)

            metro_yan_duzeltmesi = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            metro_yan_duzeltmesi_1 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi:", metro_yan_duzeltmesi)
            metro_yan_duzeltmesi = round(metro_yan_duzeltmesi/1)*1
            print("Metro Yan Düzeltmesi en yakın 1 milyem:", metro_yan_duzeltmesi)


            # METRO MESAFE DÜZELTMESİ
            mesafe_ruzgari_fark = round(mesafe_ruzgari-0, 1)
            if mesafe_ruzgari_fark > 0:
                mesafe_ruzgari_durum = "BAŞ"
            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_durum = "ARKA"
            else:
                mesafe_ruzgari_durum = 0
            print("Mesafe Ruzgari Fark :", mesafe_ruzgari_fark)
            hava_sicakligi_fark = round(duzeltilmis_sıcaklik_degeri - 100, 1)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_durum = "EKSİLME"
            if hava_sicakligi_fark > 0:
                hava_sicakligi_durum = "ARTMA"
            else:
                hava_sicakligi_durum = 0
            print("Hava Sıcaklığı Fark :", hava_sicakligi_fark)
            hava_yogunlugu_fark = round(duzeltilmis_yogunluk_degeri - 100, 1)
            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_durum = "EKSİLME"
            if hava_yogunlugu_fark > 0:
                hava_yogunlugu_durum = "ARTMA"
            else:
                hava_yogunlugu_durum = 0
            print("Hava Yoğunluğu Fark :", hava_yogunlugu_fark)
            mermi_kare_agirligi_fark = round(mermi_kare_agirligi-2, 1)
            if mermi_kare_agirligi_fark < 0:
                merkar_durum = "EKSİLME"
            if mermi_kare_agirligi_fark > 0:
                merkar_durum = "ARTMA"
            if mermi_kare_agirligi_fark == 0:
                merkar_durum = 0

            print("Mermi Kare Ağırlığı Fark :", mermi_kare_agirligi_fark)

            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_arka
            else:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_bas
            if mesafe_ruzgari_fark == 0:
                mesafe_ruzgari_duzeltme_birimi = 0

            print("Mesafe Rüzgarı Düzeltme Birimi:", mesafe_ruzgari_duzeltme_birimi)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_duzeltme_birimi = havaisisi_eksilme
            else:
                hava_sicakligi_duzeltme_birimi = havaisisi_artma
            if hava_sicakligi_fark == 0:
                hava_sicakligi_duzeltme_birimi = 0
            print("Hava Isısı Düzeltme Birimi:", hava_sicakligi_duzeltme_birimi)

            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_eksilme
            else:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_artma
            if hava_yogunlugu_fark == 0:
                hava_yogunlugu_duzeltme_birimi = 0
            print("Hava Yoğunluğu Düzeltme Birimi:", hava_yogunlugu_duzeltme_birimi)

            if mermi_kare_agirligi_fark < 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_eksilme
            elif mermi_kare_agirligi_fark > 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_artma
            else:
                mermi_kare_agirligi_duzeltme_birimi = 0
            mesafe_ruzgari_fark = abs(mesafe_ruzgari_fark)
            hava_sicakligi_fark = abs(hava_sicakligi_fark)
            hava_yogunlugu_fark = abs(hava_yogunlugu_fark)


            print("Mermi Kare Ağırlığı Düzeltme Birimi:",
                mermi_kare_agirligi_duzeltme_birimi)

            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi = round(metro_yan_duzeltmesi + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)
            gac_yan_duzeltmesi = round(toplam_yan_duzeltmesi - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi)

            # Lineer interpolasyon fonksiyonu

            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            brt_isisi_ilk_hiz_dzl = [
                (-40, -8.5),
                (-30, -8),
                (-20, -7.4),
                (-10, -6.7),
                (0, -6),
                (10, -5.3),
                (20, -4.5),
                (30, -3.7),
                (40, -2.8),
                (50, -1.9),
                (60, -1),
                (70, 0),
                (80, 1),
                (90, 2.1),
                (100, 3.2),
                (110, 4.4),
                (120, 5.6),
                (130, 6.9),

            ]
            barut__isisi_dzl2 = 3
            barut_isisi_dzl = []
            for i in range(len(brt_isisi_ilk_hiz_dzl)-1):
                start = brt_isisi_ilk_hiz_dzl[i]
                end = brt_isisi_ilk_hiz_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    barut_isisi_dzl.append((int(x), y))

            barut_isisi_dzl.append(brt_isisi_ilk_hiz_dzl[-1])
            print("Barut Isısı",barut_isisi)

            for i in barut_isisi_dzl:
                if i[0] == barut_isisi:
                    print("mesafe",i[0])
                    barut__isisi_dzl2 = round(i[1], 1)
                    print("İ1:",i[1])
                    print("Barurrrrrrrr",barut__isisi_dzl2)
            # HIZ DEĞİŞİKLİĞİNİN HESAPLANMASI
            hiz_degisikligi = round(ilk_hiz_farki + mevzi_hiz_degisikligi, 1)
            deltaV_hiz_farki = round(hiz_degisikligi + barut__isisi_dzl2, 1)
            print("Hız Değişikliği:", hiz_degisikligi)
            print("Barut Isısı Düzeltmesi2:", barut__isisi_dzl2)
            print("Delta V Hız Farkı:", deltaV_hiz_farki)

            if deltaV_hiz_farki > 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_martma
            elif deltaV_hiz_farki < 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_eksilme
            else:
                ilk_hiz_duzeltme_birimi = 0
            print("İlk Hız Düzeltme Birimi:", ilk_hiz_duzeltme_birimi)

            deltaV_mesafe_duzeltmesi = abs(
                round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            if deltaV_hiz_farki > 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))*-1
            elif deltaV_hiz_farki < 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            else:
                deltaV_mesafe_duzeltmesi = 0
            print("Delta V Mesafe Düzeltmesi:", deltaV_mesafe_duzeltmesi)

            #### TOPLAM MESAF DÜZELTMESİ ###
            toplam_mesafe_duzeltmesi = round(metro_mesafe_duzeltmesi + deltaV_mesafe_duzeltmesi)
            print("Toplam Mesafe Düzeltmesi:", toplam_mesafe_duzeltmesi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)

            # toplam yan düzeltmede 1 milyem fazla mesafe düzeltmesinde işaret + ve 335 çıktı
            # baiek de yan 4 sl mesafe -326 çıktı

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (2, -0.007, 0.007, 0, 0, 0, 0, 0, 0, 0.013, -0.013),
                (3, -0.01, 0.01, 0, 0, -0.001, 0, 0, 0, 0.019, -0.019),
                (4, -0.013, 0.013, -0.001, 0, -0.001, 0, 0, 0, 0.024, -0.025),
                (5, -0.017, 0.016, -0.001, 0, -0.001, 0, 0.001, -0.001, 0.03, -0.03),
                (6, -0.02, 0.019, -0.001, 0, -0.002, 0, 0.001, -0.001, 0.035, -0.036),
                (7, -0.023, 0.022, -0.002, 0, -0.002, 0, 0.001, -0.001, 0.041, -0.041),
                (8, -0.026, 0.025, -0.002, 0, -0.003, 0, 0.001, -0.001, 0.046, -0.047),
                (9, -0.029, 0.028, -0.002, 0.001, -0.004, 0.001, 0.002, -0.002, 0.051, -0.052),
                (10, -0.032, 0.031, -0.003, 0.001, -0.004, 0.001, 0.002, -0.002, 0.056, -0.057),
                (11, -0.036, 0.034, -0.003, 0.001, -0.005, 0.001, 0.003, -0.003, 0.061, -0.062),
                (12, -0.039, 0.037, -0.003, 0.001, -0.005, 0.001, 0.003, -0.003, 0.066, -0.067),
                (13, -0.042, 0.04, -0.003, 0.001, -0.005, 0.001, 0.003, -0.003, 0.07, -0.072),
                (14, -0.045, 0.043, -0.004, 0.001, -0.006, 0.001, 0.004, -0.004, 0.075, -0.077),
                (15, -0.048, 0.046, -0.004, 0.001, -0.006, 0.001, 0.005, -0.004, 0.08, -0.082),
                (16, -0.051, 0.048, -0.004, 0.001, -0.007, 0.001, 0.005, -0.005, 0.084, -0.086),
                (17, -0.054, 0.051, -0.005, 0.001, -0.007, 0.001, 0.006, -0.006, 0.089, -0.091),
                (18, -0.057, 0.054, -0.005, 0.001, -0.007, 0.001, 0.006, -0.006, 0.093, -0.096),
                (19, -0.06, 0.057, -0.005, 0.001, -0.007, 0.001, 0.007, -0.007, 0.097, -0.1),
                (20, -0.063, 0.06, -0.005, 0.002, -0.008, 0, 0.008, -0.008, 0.102, -0.105),
                (21, -0.066, 0.063, -0.005, 0.002, -0.008, 0, 0.008, -0.008, 0.106, -0.109),
                (22, -0.069, 0.066, -0.006, 0.002, -0.008, 0, 0.009, -0.009, 0.11, -0.113),
                (23, -0.072, 0.068, -0.006, 0.002, -0.008, 0, 0.01, -0.01, 0.114, -0.118),
                (24, -0.075, 0.071, -0.006, 0.002, -0.008, 0, 0.011, -0.011, 0.118, -0.122),
                (25, -0.078, 0.074, -0.006, 0.002, -0.009, 0, 0.012, -0.011, 0.122, -0.126),
                (26, -0.081, 0.077, -0.006, 0.002, -0.009, 0, 0.013, -0.012, 0.126, -0.13),
                (27, -0.084, 0.08, -0.006, 0.002, -0.009, 0, 0.013, -0.013, 0.13, -0.134),
                (28, -0.087, 0.083, -0.007, 0.002, -0.009, 0, 0.014, -0.014, 0.134, -0.138),
                (29, -0.09, 0.086, -0.007, 0.003, -0.009, 0.001, 0.015, -0.015, 0.138, -0.142),
                (30, -0.093, 0.089, -0.007, 0.003, -0.009, 0.001, 0.016, -0.016, 0.142, -0.146),
                (31, -0.096, 0.091, -0.007, 0.003, -0.009, 0.001, 0.017, -0.017, 0.146, -0.15),
                (32, -0.099, 0.094, -0.007, 0.003, -0.009, 0.001, 0.018, -0.018, 0.149, -0.154),
                (33, -0.102, 0.097, -0.007, 0.003, -0.009, 0.001, 0.019, -0.019, 0.153, -0.158),
                (34, -0.104, 0.1, -0.007, 0.003, -0.009, 0.001, 0.02, -0.02, 0.157, -0.162),
                (35, -0.107, 0.103, -0.007, 0.003, -0.009, 0.001, 0.021, -0.021, 0.161, -0.166),
                (36, -0.11, 0.116, -0.007, 0.003, -0.009, 0.001, 0.022, -0.022, 0.164, -0.17),
                (37, -0.113, 0.109, -0.007, 0.004, -0.009, 0.001, 0.023, -0.023, 0.168, -0.174),
                (38, -0.116, 0.111, -0.007, 0.004, -0.009, 0.001, 0.024, -0.024, 0.172, -0.178),
                (39, -0.119, 0.114, -0.007, 0.004, -0.009, 0.002, 0.025, -0.025, 0.175, -0.181),
                (40, -0.122, 0.117, -0.007, 0.004, -0.009, 0.002, 0.026, -0.025, 0.179, -0.185),
                (41, -0.125, 0.12, -0.007, 0.004, -0.009, 0.002, 0.027, -0.026, 0.183, -0.189),
                (42, -0.128, 0.123, -0.007, 0.004, -0.009, 0.002, 0.028, -0.027, 0.187, -0.193),
                (43, -0.131, 0.126, -0.007, 0.004, -0.009, 0.002, 0.029, -0.028, 0.19, -0.197),
                (44, -0.134, 0.129, -0.007, 0.005, -0.009, 0.002, 0.03, -0.029, 0.194, -0.201),
                (45, -0.136, 0.131, -0.007, 0.005, -0.009, 0.002, 0.031, -0.03, 0.198, -0.205),
                (46, -0.139, 0.134, -0.007, 0.005, -0.009, 0.002, 0.032, -0.031, 0.202, -0.209),
                (47, -0.142, 0.137, -0.007, 0.005, -0.009, 0.002, 0.033, -0.032, 0.206, -0.213),
                (48, -0.145, 0.14, -0.007, 0.005, -0.009, 0.002, 0.034, -0.033, 0.21, -0.218),
                (49, -0.148, 0.143, -0.007, 0.005, -0.009, 0.002, 0.035, -0.034, 0.214, -0.222),
                (50, -0.151, 0.146, -0.007, 0.005, -0.009, 0.002, 0.036, -0.035, 0.219, -0.227),
                (51, -0.154, 0.149, -0.007, 0.006, -0.009, 0.002, 0.037, -0.036, 0.223, -0.231),
                (52, -0.157, 0.152, -0.007, 0.007, -0.009, 0.002, 0.038, -0.037, 0.228, -0.236),
                (53, -0.16, 0.155, -0.007, 0.01, -0.009, 0.002, 0.039, -0.038, 0.234, -0.242),
                (54, -0.163, 0.157, -0.007, 0.01, -0.009, 0.002, 0.04, -0.039, 0.242, -0.249),
                (55, -0.165, 0.159, -0.01, 0.01, -0.01, 0.002, 0.046, -0.043, 0.254, -0.267),
                (56, -0.158, 0.16, -0.02, 0.01, -0.005, 0.001, 0.046, -0.055, 0.342, -0.299),






            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Tapa Saniyesi :", tapasaniyesi)

            if deltaV_hiz_farki < 0:
                for ihz_eksilme in interpolasyonlu_veri1:
                    if ihz_eksilme[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_eksilme[1], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Eksilme):", ilkhizda1mildegisim_eksilme)
            if deltaV_hiz_farki > 0:
                for ihz_artma in interpolasyonlu_veri1:
                    if ihz_artma[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_artma[2], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Artma):", ilkhizda1mildegisim_artma)
            if mesafe_ruzgari_durum == "BAŞ":
                for musrus_bas in interpolasyonlu_veri1:
                    if musrus_bas[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(musrus_bas[3], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Baş):", mesaferuzgaridegisimi_bas)
            if mesafe_ruzgari_durum == "ARKA":
                for mesruz_arka in interpolasyonlu_veri1:
                    if mesruz_arka[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(mesruz_arka[4], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Arka):", mesaferuzgaridegisimi_arka)
            if hava_sicakligi_durum == "EKSİLME":
                for hadese_eksilme in interpolasyonlu_veri1:
                    if hadese_eksilme[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_eksilme[5], 3)
                        # print("Hava Isısı Değişim Miktarı (Eksilme):", havaisisidegisim_eksilme)
            if hava_sicakligi_durum == "ARTMA":
                for hadese_artma in interpolasyonlu_veri1:
                    if hadese_artma[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_artma[6], 3)
                        # print("Hava Isısı Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == "EKSİLME":
                for hayog_eksilme in interpolasyonlu_veri1:
                    if hayog_eksilme[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_eksilme[7], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Eksilme):", havayogunlugudegisim_eksilme)
            if hava_yogunlugu_durum == "ARTMA":
                for hayog_artma in interpolasyonlu_veri1:
                    if hayog_artma[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_artma[8], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == 0:
                havayogunlugudegisim = 0
            if merkar_durum == "EKSİLME":
                for merkar_eksilme in interpolasyonlu_veri1:
                    if merkar_eksilme[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_eksilme[9], 3)
                        # print("Mermi Kare Ağırlığı Değişim Miktarı (Eksilme):", mermikareagirligidegisim_eksilme)
            if merkar_durum == "ARTMA":
                for merkar_artma in interpolasyonlu_veri1:
                    if merkar_artma[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_artma[10], 3)
            else:
                mermikareagirligidegisim = 0
                # print("Mermi Kare Ağırlığı Değişim Miktarı (Artma):", mermikareagirligidegisim_artma)
##########################################################################################################
            gac_mesafe = plan_mesafesi_2B_obüs + toplam_mesafe_duzeltmesi
            print("toplam mesafe",nisangah5bh)
# zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
                (100, 5.6, 0, 00, 0.3, 0, 0, 0.7, -0.6, 0, 0, 0, 0, 0, 0, -1, 1),
                (200, 11.1, 0, 0, 0.7, 0, 0.01, 1.4, -1.2, 0, 0, 0, 0, 0, 0, -2, 2),
                (300, 16.6, 0, 0, 1, 0.1, 0.01, 2.1, -1.8, 0.1, 0, 0.1, 0, 0, 0, -4, 4),
                (400, 22.2, 0, 0, 1.3, 0.2, 0.02, 2.7, -2.4, 0.1, 0, 0.1, 0, -0.1, 0.1, -5, 5),
                (500, 27.8, 0, 0, 1.7, 0.3, 0.02, 3.4, -3, 0.2, 0, 0.2, 0, -0.1, 0.1, -6, 6),
                (600, 33.4, 2, 1.02, 2, 0.4, 0.02, 4.1, -3.6, 0.2, 0, 0.3, 0, -0.1, 0.1, -7, 7),
                (700, 39.1, 2.3, 0.87, 2.4, 0.4, 0.03, 4.7, -
                4.2, 0.3, -0.1, 0.4, -0.1, -0.1, 0.1, -8, 8),
                (800, 44.9, 2.6, 0.76, 2.7, 0.5, 0.03, 5.4, -
                4.8, 0.3, -0.1, 0.4, -0.1, -0.2, 0.2, -9, 9),
                (900, 50.6, 3, 0.68, 3, 0.6, 0.04, 6, -5.3,
                0.4, -0.1, 0.5, -0.1, -0.2, 0.2, -10, 10),
                (1000, 56.4, 3.3, 0.61, 3.4, 0.7, 0.04, 6.7, -
                5.9, 0.5, -0.1, 0.6, -0.1, -0.3, 0.3, -11, 11),
                (1100, 62.3, 3.7, 0.55, 3.7, 0.8, 0.04, 7.3, -
                6.4, 0.5, -0.1, 0.7, -0.1, -0.4, 0.4, -12, 13),
                (1200, 68.2, 4, 0.5, 4.1, 0.9, 0.05, 7.9, -
                7, 0.6, -0.2, 0.8, -0.1, -0.4, 0.4, -13, 14),
                (1300, 74.1, 4.4, 0.46, 4.4, 1, 0.05, 8.6, -
                7.5, 0.7, -0.2, 0.9, -0.1, -0.5, 0.5, -14, 15),
                (1400, 80.1, 4.8, 0.43, 4.8, 1.1, 0.06, 9.2, -
                8.1, 0.8, -0.2, 1, -0.1, -0.6, 0.6, -15, 16),
                (1500, 86.1, 5.1, 0.4, 5.2, 1.2, 0.06, 9.8, -
                8.6, 0.9, -0.2, 1.2, -0.2, -0.6, 0.6, -16, 17),
                (1600, 92.2, 5.5, 0.37, 5.5, 1.3, 0.06, 10.4, -
                9.1, 1, -0.3, 1.3, -0.2, -0.7, 0.7, -17, 17),
                (1700, 98.3, 5.8, 0.35, 5.9, 1.3, 0.07, 11.1, -
                9.7, 1.1, -0.3, 1.4, -0.2, -0.8, 0.8, -18, 18),
                (1800, 104.5, 6.2, 0.33, 6.2, 1.4, 0.07, 11.7, -
                10.2, 1.1, -0.3, 1.5, -0.2, -0.9, 0.9, -19, 19),
                (1900, 110.7, 6.6, 0.31, 6.6, 1.5, 0.08, 12.3, -
                10.7, 1.2, -0.3, 1.6, -0.2, -1, 1, -19, 20),
                (2000, 117, 6.9, 0.3, 7, 1.6, 0.08, 12.9, -
                11.2, 1.3, -0.4, 1.7, -0.2, -1.1, 1.1, -20, 21),
                (2100, 123.3, 7.3, 0.28, 7.3, 1.7, 0.08, 13.5, -
                11.8, 1.4, -0.4, 1.8, -0.2, -1.2, 1.3, -21, 22),
                (2200, 129.7, 7.7, 0.27, 7.7, 1.8, 0.09, 14.1, -
                12.3, 1.5, -0.4, 1.9, -0.2, -1.4, 1.4, -22, 23),
                (2300, 136.2, 8.1, 0.26, 8.1, 1.9, 0.09, 14.7, -
                12.8, 1.6, -0.5, 2, -0.2, -1.5, 1.5, -23, 24),
                (2400, 142.7, 8.4, 0.25, 8.5, 2.1, 0.1, 15.3, -
                13.3, 1.7, -0.5, 2.1, -0.2, -1.6, 1.6, -24, 25),
                (2500, 149.2, 8.8, 0.24, 8.9, 2.2, 0.1, 15.9, -
                13.8, 1.8, -0.6, 2.2, -0.2, -1.8, 1.8, -24, 26),
                (2600, 155.9, 9.2, 0.23, 9.2, 2.3, 0.11, 16.5, -
                14.3, 1.9, -0.6, 2.3, -0.2, -1.9, 1.9, -25, 26),
                (2700, 162.6, 9.6, 0.22, 9.6, 2.4, 0.11, 17.1, -
                14.9, 2, -0.6, 2.4, -0.2, -2, 2.1, -26, 27),
                (2800, 169.4, 10, 0.21, 10, 2.5, 0.11, 17.7, -
                15.4, 2.1, -0.7, 2.5, -0.2, -2.2, 2.2, -27, 28),
                (2900, 176.2, 10.4, 0.2, 10.4, 2.6, 0.12, 18.3, -
                15.9, 2.2, -0.7, 2.5, -0.2, -2.3, 2.4, -27, 29),
                (3000, 183.1, 10.8, 0.19, 10.8, 2.7, 0.12, 18.9, -
                16.4, 2.3, -0.8, 2.6, -0.2, -2.5, 2.5, -28, 29),
                (3100, 190.1, 11.2, 0.19, 11.2, 2.8, 0.13, 19.4, -
                16.9, 2.4, -0.8, 2.7, -0.2, -2.7, 2.7, -29, 30),
                (3200, 197.2, 11.6, 0.18, 11.6, 3, 0.13, 20, -
                17.4, 2.5, -0.9, 2.8, -0.2, -2.8, 2.9, -29, 31),
                (3300, 204.4, 12, 0.17, 12, 3.1, 0.14, 20.6, -
                17.9, 2.6, -0.9, 2.8, -0.1, -3, 3.1, -30, 32),
                (3400, 211.6, 12.4, 0.17, 12.4, 3.2, 0.14, 21.2, -
                18.4, 2.7, -1, 2.9, -0.1, -3.2, 3.3, -31, 32),
                (3500, 219, 12.8, 0.16, 12.8, 3.3, 0.14, 21.7, -
                18.9, 2.8, -1, 2.9, -0.1, -3.4, 3.4, -31, 33),
                (3600, 226.4, 13.2, 0.16, 13.2, 3.5, 0.15, 22.3, -
                19.4, 2.9, -1.1, 3, -0.1, -3.6, 3.6, -32, 34),
                (3700, 233.9, 13.6, 0.15, 13.7, 3.6, 0.15, 22.9, -
                19.9, 3, -1.1, 3, -0.1, -3.8, 3.8, -33, 34),
                (3800, 241.6, 14.1, 0.15, 14.1, 3.7, 0.16, 23.4, -
                20.4, 3.1, -1.2, 3.1, -0.1, -4, 4.1, -33, 35),
                (3900, 249.3, 14.5, 0.14, 14.5, 3.9, 0.16, 24, -
                20.9, 3.2, -1.2, 3.1, 0, -4.2, 4.3, -34, 36),
                (4000, 257.2, 14.9, 0.14, 15, 4, 0.17, 24.6, -
                21.4, 3.3, -1.3, 3.2, 0, -4.4, 4.5, -34, 36),
                (4100, 265.2, 15.4, 0.14, 15.4, 4.1, 0.17, 25.1, -
                21.9, 3.4, -1.4, 3.2, 0, -4.6, 4.7, -35, 37),
                (4200, 273.3, 15.8, 0.13, 15.9, 4.3, 0.18, 25.7, -
                22.3, 3.5, -1.4, 3.2, 0, -4.8, 4.9, -36, 38),
                (4300, 281.5, 16.3, 0.13, 16.3, 4.4, 0.18, 26.2, -
                22.8, 3.6, -1.5, 3.3, 0.1, -5.1, 5.2, -36, 38),
                (4400, 289.9, 16.7, 0.13, 16.8, 4.6, 0.19, 26.8, -
                23.3, 3.7, -1.6, 3.3, 0.1, -5.3, 5.4, -37, 39),
                (4500, 298.4, 17.2, 0.12, 17.2, 4.8, 0.19, 27.3, -
                23.8, 3.8, -1.6, 3.3, 0.1, -5.6, 5.7, -37, 39),
                (4600, 307, 17.7, 0.12, 17.7, 4.9, 0.2, 27.9, -
                24.3, 3.9, -1.7, 3.3, 0.1, -5.8, 5.9, -38, 40),
                (4700, 315.9, 18.2, 0.12, 18.2, 5.1, 0.2, 28.4, -
                24.8, 4, -1.8, 3.3, 0.2, -6, 6.2, -38, 40),
                (4800, 324.9, 18.6, 0.11, 18.7, 5.3, 0.21, 29, -
                25.3, 4.1, -1.8, 3.3, 0.2, -6.3, 6.5, -39, 41),
                (4900, 334, 19.1, 0.11, 19.1, 5.4, 0.21, 29.5, -
                25.7, 4.2, -1.9, 3.4, 0.2, -6.6, 6.7, -39, 41),
                (5000, 343.4, 19.6, 0.11, 19.6, 5.6, 0.22,
                30, -26.2, 4.3, -2, 3.4, 0.3, -6.8, 7, -40, 42),
                (5100, 353, 20.1, 0.11, 20.2, 5.8, 0.23, 30.6, -
                26.7, 4.4, -2.1, 3.4, 0.3, -7.1, 7.3, -40, 42),
                (5200, 362.8, 20.7, 0.1, 20.7, 6, 0.23, 31.1, -
                27.2, 4.6, -2.1, 3.3, 0.4, -7.4, 7.6, -40, 43),
                (5300, 372.8, 21.2, 0.1, 21.2, 6.2, 0.24, 31.6, -
                27.7, 4.7, -2.2, 3.3, 0.4, -7.7, 7.9, -41, 43),
                (5400, 383.1, 21.7, 0.1, 21.7, 6.4, 0.24, 32.2, -
                28.1, 4.8, -2.3, 3.3, 0.4, -8, 8.2, -41, 44),
                (5500, 393.7, 22.3, 0.1, 22.3, 6.6, 0.25, 32.7, -
                28.6, 4.9, -2.4, 3.3, 0.5, -8.3, 8.5, -42, 44),
                (5600, 404.6, 22.8, 0.09, 22.8, 6.9, 0.26, 33.2, -
                29.1, 5, -2.5, 3.3, 0.5, -8.6, 8.8, -42, 45),
                (5700, 415.8, 23.4, 0.09, 23.4, 7.1, 0.26, 33.7, -
                29.6, 5.1, -2.6, 3.3, 0.6, -8.9, 9.1, -42, 45),
                (5800, 427.3, 24, 0.09, 24, 7.3, 0.27, 34.2, -
                30, 5.2, -2.7, 3.2, 0.6, -9.2, 9.5, -43, 45),
                (5900, 439.3, 24.6, 0.09, 24.6, 7.6, 0.28, 34.7, -
                30.5, 5.3, -2.8, 3.2, 0.6, -9.5, 9.8, -43, 46),
                (6000, 451.7, 25.2, 0.08, 25.2, 7.9, 0.28, 35.2, -
                31, 5.4, -2.9, 3.2, 0.7, -9.8, 10.1, -43, 46),
                (6100, 464.6, 25.9, 0.08, 25.9, 8.2, 0.29, 35.7, -
                31.5, 5.5, -3, 3.1, 0.7, -10.2, 10.5, -44, 46),
                (6200, 478.1, 26.5, 0.08, 26.6, 8.5, 0.3, 36.2, -
                31.9, 5.6, -3.1, 3.1, 0.8, -10.5, 10.8, -44, 47),
                (6300, 492.2, 27.2, 0.08, 27.2, 8.8, 0.31, 36.7, -
                32.4, 5.8, -3.2, 3.1, 0.8, -10.9, 11.2, -44, 47),
                (6400, 507.1, 28, 0.08, 28, 9.2, 0.31, 37.2, -
                32.9, 5.9, -3.3, 3, 0.8, -11.2, 11.6, -44, 47),
                (6500, 522.9, 28.7, 0.07, 28.7, 9.6, 0.32,
                37.7, -33.3, 6, -3.4, 3, 0.8, -11.6, 12, -44, 47),
                (6600, 539.7, 29.5, 0.07, 29.5, 10, 0.33, 38.1, -
                33.8, 6.2, -3.6, 2.9, 0.9, -11.9, 12.4, -45, 48),
                (6700, 557.8, 30.4, 0.07, 30.4, 10.5, 0.34, 38.6, -
                34.2, 6.3, -3.7, 2.9, 0.9, -12.3, 12.8, -45, 48),
                (6800, 577.7, 31.3, 0.07, 31.3, 11, 0.35, 39, -
                34.7, 6.5, -3.8, 2.8, 0.9, -12.7, 13.2, -45, 48),
                (6900, 599.7, 32.3, 0.07, 32.3, 11.6, 0.37, 39, -
                35.2, 6.8, -4, 2.8, 0.9, -13.1, 13.6, -45, 48),
                (7000, 625, 33.5, 0.06, 33.5, 12.4, 0.38, 39, -
                35.6, 6.8, -4.1, 2.8, 0.9, -13.5, 14.1, -45, 48),
                (7100, 655.6, 34.8, 0.06, 34.8, 13.3, 0.4, 39, -
                36.1, 6.8, -4.3, 2.8, 0.8, -13.9, 14.7, -45, 48),
                (7200, 697.4, 36.6, 0.06, 36.6, 14.7, 0.42, 39, -
                36.5, 6.8, -4.5, 2.7, 0.8, -14.4, 14.7, -45, 48),


            ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == gac_mesafe:
                    nisangah5bh = round(nsg[1], 1)
                    print("Nişangah:", nisangah5bh)
            for dgl in interpolasyonlu_veri1:
                if dgl[0] == gac_mesafe:
                    dogalyandz = round(dgl[5],3)
                    print("Doğal Yan Düzeltmesi:", dogalyandz)
            for tsg in interpolasyonlu_veri1:
                if tsg[0] == gac_mesafe:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            metro_yan_duzeltmesi2 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi en yakın 1 milyem99999999999999:", metro_yan_duzeltmesi2)
            print("Arzin Dnüşü:,",arzindonusuyanduzeltmesi)
            print("Rüzgaryan dzl---------------------:,",ruzgar_yan_duzeltmesi)

            
            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi5bh = round(metro_yan_duzeltmesi2 + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi33333333333335bh:", toplam_yan_duzeltmesi5bh)
            gac_yan_duzeltmesi2 = round(toplam_yan_duzeltmesi5bh - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi2)
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]



            # Veri kümesi
            data1 = [
                (1000,58,14,0.003,-0.003),
                (1500,90,32,0.007,-0.007),
                (2000,124,59,0.014,-0.013),
                (2500,161,96,0.023,-0.022),
                (3000,200,142,0.035,-0.033),
                (3500,242,201,0.053,-0.049),
                (4000,288,273,0.076,-0.07),
                (4500,338,362,0.109,-0.098),
                (5000,393,471,0.156,-0.138),
                (5500,455,606,0.228,-0.196),
                (6000,526,777,0.351,-0.288),
                (6500,611,1007,0.616,-0.455),
                (7000,728,1365,0.616,-0.894),

            ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == gac_mesafe:
                    dusus_acisi = round(dusus_acisi[1],3)
                    # print("Düşüş Açısı:", dusus_acisi)

            for tepe_yuksekligi in interpolasyonlu_veri1:
                if tepe_yuksekligi[0] == gac_mesafe:
                    tepe_yuksekligi = round(tepe_yuksekligi[2],3)
                    # print("Tepe Yüksekliği:", tepe_yuksekligi)

            for dtac_arti1 in interpolasyonlu_veri1:
                if dtac_arti1[0] == mesafe:
                    dtac_arti1 = round(dtac_arti1[3],3)
                    # print("+1 DTAÇ:", dtac_arti1)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4],3)
                    # print("-1 DTAÇ:", dtac_eksi1)
    ############################################################################################################
            """METRO DÜZELTMELERİNİ HESAPLAMAK GEREKİYOR"""
            # GTAC 1 inci oonüs için
            bt_rakimi = lne_1B_obus_rakim

            milyem_sabit_dz = 1.0186

            hedef_batarya_rakim_farki = (rakim - bt_rakimi)

            dtac = (hedef_batarya_rakim_farki / (plan_mesafesi_1B_obüs / 1000)) * milyem_sabit_dz

            # Burada dtac_arti1 ve dtac_eksi1 değerlerini sadece bir kez tanımlayın
            dtac_arti1 = 0
            dtac_eksi1 = 0

            # dtac_arti1 ve dtac_eksi1 değerlerini doğru şekilde ata
            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_arti1 = round(dusus_acisi[3], 3)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4], 3)

            ttac = 0
            print("Mesafe",mesafe)
            if dtac < 0:
                ttac = int(dtac * dtac_arti1)
                tac = dtac + ttac
                #print("-TAÇ", tac)

            if dtac > 0:
                ttac = int(dtac * dtac_eksi1)
                tac = dtac + ttac
                print("+TAÇ", tac)
    #########################################################################################################
            """YÜKSELİŞİ BUL"""
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,18),
(100,18),
(200,18),
(300,18),
(400,18),
(500,18),
(600,18),
(700,18),
(800,18),
(900,17),
(1000,17),
(1100,17),
(1200,17),
(1300,17),
(1400,17),
(1500,17),
(1600,16),
(1700,16),
(1800,16),
(1900,16),
(2000,16),
(2100,16),
(2200,16),
(2300,15),
(2400,15),
(2500,15),
(2600,15),
(2700,15),
(2800,15),
(2900,15),
(3000,14),
(3100,14),
(3200,14),
(3300,14),
(3400,14),
(3500,14),
(3600,13),
(3700,13),
(3800,13),
(3900,13),
(4000,13),
(4100,12),
(4200,12),
(4300,12),
(4400,12),
(4500,12),
(4600,11),
(4700,11),
(4800,11),
(4900,11),
(5000,11),
(5100,10),
(5200,10),
(5300,10),
(5400,10),
(5500,9),
(5600,9),
(5700,9),
(5800,8),
(5900,8),
(6000,8),
(6100,8),
(6200,7),
(6300,7),
(6400,7),
(6500,6),
(6600,6),
(6700,5),
(6800,5),
(6900,4),
(7000,4),
(7100,3),
(7200,3),
            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            
            for nsg in interpolasyonlu_veri1:
                    if nsg[0] == gac_mesafe:
                        birmildegisim5bh = round(nsg[1], 1)
                        print("1 Milyemlik Değişim:", birmildegisim5bh)

            yukselis5bh = round(nisangah5bh+ + tac,1)
            self.ui.sonuc_yukselis_2B.setText(str(yukselis5bh))
    #########################################################################################################
            yan_2B_5bh = round(yan_2B+toplam_yan_duzeltmesi5bh)
            self.ui.sonuc_yan_2B.setText(str(yan_2B_5bh))
            self.ui.lne_barut_hakki.setText(str(secilen_barut_hakki))
            self.ui.sonuc_barut_hakki_2B.setText(str(secilen_barut_hakki))

            if 2200 < yan_2B_5bh < 3000:
                self.ui.sonuc_yan_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_yukselis_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_barut_hakki_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_tapa_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_istikamet_acisi_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_mesafe_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.lne_sonuc_mesafe.setText(str(mesafe))
                self.ui.lne_sonuc_barut_hakki.setText(str(secilen_barut_hakki))
                self.ui.lne_sonuc_yan.setText(str(yan_2B_5bh))
                self.ui.lne_sonuc_yukselis.setText(str(yukselis5bh))
                self.ui.lne_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
                self.ui.lne_sonuc_bir_ml_degisiklil.setText(str(birmildegisim5bh)) 
                self.ui.lne_sonuc_yuz_m.setText(str(mesafe/1000)) 
                """SONRA YAPACAĞIM"""
                # self.ui.lbl_sonuc_barut_hakki_2.setText(str(secilen_barut_hakki))
                # self.ui.lbl_sonuc_yan_2.setText(str(yan_2B_5bh))
                # self.ui.lbl_sonuc_yukselis_2.setText(str(yukselis5bh))
                # self.ui.lbl_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
            else:
                self.ui.sonuc_yan_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_yukselis_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_barut_hakki_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_tapa_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_istikamet_acisi_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_mesafe_2B.setStyleSheet("color: grey; font-weight: normal")
            self._doldur_sonuc("2", toplam_yan_duzeltmesi5bh, yukselis5bh, birmildegisim5bh, locals())
        #################################################################################################
        if secilen_barut_hakki == 6:
            plan_mesafesi = plan_mesafesi_2B_obüs
            batarya_rakimi = lne_2B_obus_rakim
            bt_rakimi_1 = lne_2B_obus_rakim
            hedef_rakimi = lne_hedef_rakim
            ahia = lne_2B_obus_ahia
            paralanma_yuksekliği = paralanma_yuksekliği
            mermi_kare_agirligi = mermi_kare_agirligi
            ilk_hiz_farki = lne_2B_obus_ihf
            atis_istikameti = batarya_hedef_İA_1
            atis_istikameti_1 = batarya_hedef_İA_1
            mevzi_hiz_degisikligi = 0
            mevzi_yan_duzeltmesi = 0
            mevzi_ts_duzeltmesi = 0
            tapasaniyesi = 1
            barut_isisi_str = self.ui.lne_barut_isisi.text()
            barut_isisi = float(barut_isisi_str) if barut_isisi_str else 75
            # barut_isisi = self.ui.lne_barut_isisi.text() if self.ui.lne_barut_isisi.text() else 75
            print("Barutttt İSİSİ",barut_isisi)
            
            # 1. GİRİŞ MESAFESİNİ BUL / MANUEL HASSASİYETİ

            # mesafe = 1866
            # en_yakin_10_metre = round(mesafe / 10) * 10
            # print("1836'nın en yakın 10 metreye çevrilmiş hali:", en_yakin_10_metre)

            # BATARYA RAKIMINI EN YAKIN 10 M YE ÇEVİR.
            batarya_rakimi = round(batarya_rakimi/10)*10
            # PARALANMA NOKTASININ RAKIMINI BUL.
            paralanma_noktasi_rakimi = hedef_rakimi + paralanma_yuksekliği
            # HEDEF-BATARYA YÜKSEKLİK FARKINI BUL
            yukseklik_farki = paralanma_noktasi_rakimi-batarya_rakimi
            yukseklik_farki_1 = paralanma_noktasi_rakimi-bt_rakimi_1
            # yukseklik_farki = round(yukseklik_farki/100)*100
            ################# B cetvelini yüklüyoruz.################
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
                # Veri kümesi
            data1 = [
                    (2500,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2600,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2700,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2800,0,0,0,0,0,1,1,1,2,2,2,2,3,3,3),
(2900,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3000,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3100,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3200,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3300,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3400,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3),
(3500,0,0,0,1,1,1,1,1,2,2,2,2,3,3,3),
(3600,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(3700,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(3800,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(3900,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4000,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4100,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4200,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4300,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4400,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4500,1,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(4600,1,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(4700,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(4800,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(4900,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5000,1,1,1,1,2,2,2,2,2,2,2,3,3,3,3),
(5100,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5200,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5300,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(5400,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(5500,1,1,1,2,2,2,2,2,2,3,3,3,3,3,3),
(5600,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(5700,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(5800,1,1,2,2,2,2,2,2,3,3,3,3,3,3,3),
(5900,1,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6000,1,2,2,2,2,2,2,2,3,3,3,3,3,3,4),
(6100,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(6200,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(6300,2,2,2,2,2,2,2,3,3,3,3,3,3,4,4),
(6400,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(6500,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(6600,2,2,2,2,2,3,3,3,3,3,3,3,4,4,4),
(6700,2,2,2,2,2,3,3,3,3,3,3,3,4,4,4),
(6800,2,2,2,2,3,3,3,3,3,3,3,4,4,4,4),
(6900,2,2,2,2,3,3,3,3,3,3,3,4,4,4,4),
(7000,2,2,2,3,3,3,3,3,3,3,3,4,4,4,4),
(7100,2,2,2,3,3,3,3,3,3,3,4,4,4,4,4),
(7200,2,2,3,3,3,3,3,3,3,3,4,4,4,4,4),
(7300,2,2,3,3,3,3,3,3,3,4,4,4,4,4,4),
(7400,2,3,3,3,3,3,3,3,4,4,4,4,4,4,5),
(7500,3,3,3,3,3,3,3,3,4,4,4,4,4,4,5),
(7600,3,3,3,3,3,3,3,4,4,4,4,4,4,5,5),
(7700,3,3,3,3,3,3,3,4,4,4,4,4,5,5,5),
(7800,3,3,3,3,3,3,4,4,4,4,4,4,5,5,5),
(7900,3,3,3,3,3,4,4,4,4,4,4,5,5,5,5),
(8000,3,3,3,3,3,4,4,4,4,4,5,5,5,5,5),
(8100,3,3,3,3,4,4,4,4,4,4,5,5,5,5,5),
(8200,3,3,3,4,4,4,4,4,4,5,5,5,5,5,6),
(8300,3,3,4,4,4,4,4,4,5,5,5,5,5,6,6),
(8400,3,3,4,4,4,4,4,5,5,5,5,5,6,6,6),
(8500,3,4,4,4,4,4,5,5,5,5,5,6,6,6,6),
(8600,4,4,4,4,4,4,5,5,5,5,6,6,6,6,6),
(8700,4,4,4,4,4,5,5,5,5,6,6,6,6,6,6),
(8800,4,4,4,4,5,5,5,5,6,6,6,6,6,6,6),
(8900,4,4,4,5,5,5,5,6,6,6,6,6,6,6,6),
(9000,4,4,5,5,5,6,6,6,6,6,6,6,6,6,6),


            ]

            data2 = [
                (2500,0,0,-12,-6,0,7,14,22,31,40,50,61,73,85,97),
(2600,0,-17,-12,-6,0,7,15,23,32,42,52,63,75,87,100),
(2700,0,-18,-13,-7,0,7,15,24,33,43,54,65,77,90,104),
(2800,0,-19,-13,-7,0,8,16,25,34,45,56,67,80,93,107),
(2900,0,-20,-14,-7,0,8,17,26,36,46,58,70,82,96,110),
(3000,-27,-21,-15,-8,0,8,17,27,37,48,60,72,85,99,114),
(3100,-28,-22,-15,-8,0,9,18,28,38,50,62,74,88,102,117),
(3200,-29,-23,-16,-8,0,9,18,29,40,51,64,77,91,106,121),
(3300,-31,-24,-17,-9,0,9,19,30,41,53,66,80,94,109,125),
(3400,-32,-25,-17,-9,0,10,20,31,43,55,68,82,97,113,129),
(3500,-33,-26,-18,-9,0,10,21,32,44,57,71,85,100,116,133),
(3600,-35,-27,-19,-10,0,10,21,33,46,59,73,88,104,120,138),
(3700,-36,-28,-19,-10,0,11,22,34,47,61,76,91,107,124,142),
(3800,-38,-29,-20,-10,0,11,23,36,49,63,78,94,111,128,147),
(3900,-39,-31,-21,-11,0,12,24,37,51,65,81,97,114,132,151),
(4000,-41,-32,-22,-11,0,12,25,38,52,68,83,100,118,137,156),
(4100,-43,-33,-23,-12,0,12,26,40,54,70,86,104,122,141,161),
(4200,-44,-34,-24,-12,0,13,26,41,56,72,89,107,126,145,166),
(4300,-46,-36,-24,-13,0,13,27,42,58,75,92,110,130,150,171),
(4400,-48,-37,-25,-13,0,14,28,44,60,77,95,114,134,155,177),
(4500,-50,-38,-26,-13,0,14,29,45,62,80,98,118,138,160,182),
(4600,-51,-40,-27,-14,0,15,30,47,64,82,101,122,143,165,188),
(4700,-53,-41,-28,-14,0,15,31,48,66,85,105,125,147,170,194),
(4800,-55,-43,-29,-15,0,16,32,50,68,88,108,129,152,175,200),
(4900,-57,-44,-30,-15,0,16,34,52,71,91,112,134,157,181,206),
(5000,-59,-46,-31,-16,0,17,35,53,73,94,115,138,162,186,213),
(5100,-61,-47,-32,-17,0,17,36,55,75,96,119,142,167,192,219),
(5200,-64,-49,-33,-17,0,18,37,57,78,100,122,147,172,198,226),
(5300,-66,-51,-35,-18,0,19,38,59,80,103,126,151,177,204,233),
(5400,-68,-52,-36,-18,0,19,39,60,83,106,130,156,182,210,240),
(5500,-70,-54,-37,-19,0,20,41,62,85,109,134,161,188,217,247),
(5600,-73,-56,-38,-20,0,20,42,64,88,113,138,165,194,224,255),
(5700,-75,-58,-39,-20,0,21,43,66,91,116,143,171,200,230,263),
(5800,-77,-59,-41,-21,0,22,45,68,93,120,147,176,206,237,271),
(5900,-80,-61,-42,-21,0,22,46,71,96,123,152,181,212,245,279),
(6000,-83,-63,-43,-22,0,23,47,73,99,127,156,187,219,252,287),
(6100,-85,-65,-45,-23,0,24,49,75,102,131,161,192,225,260,296),
(6200,-88,-67,-46,-23,0,25,50,77,105,135,166,198,232,268,305),
(6300,-91,-69,-47,-24,0,25,52,80,109,139,171,204,239,276,315),
(6400,-93,-72,-49,-25,0,26,53,82,112,143,176,211,247,285,325),
(6500,-96,-74,-50,-26,0,27,55,85,115,148,182,217,255,294,335),
(6600,-99,-76,-52,-27,0,28,57,87,119,152,187,224,263,303,346),
(6700,-102,-78,-53,-27,0,29,59,90,123,157,193,231,271,313,357),
(6800,-105,-81,-55,-28,0,29,60,93,126,162,199,238,280,323,369),
(6900,-109,-83,-57,-29,0,30,62,95,130,167,206,246,289,334,381),
(7000,-112,-86,-59,-30,0,31,64,98,134,172,212,254,298,345,394),
(7100,-115,-88,-60,-31,0,32,66,102,139,178,219,262,308,356,408),
(7200,-119,-91,-62,-32,0,33,68,105,143,184,226,271,319,369,422),
(7300,-123,-94,-64,-33,0,34,70,108,148,190,234,280,330,382,438),
(7400,-126,-97,-66,-34,0,35,73,112,153,196,242,290,341,396,455),
(7500,-130,-100,-68,-35,0,37,75,115,158,203,250,301,354,411,473),
(7600,-134,-103,-70,-36,0,38,77,119,163,210,259,312,368,428,493),
(7700,-139,-106,-73,-37,0,39,80,123,169,217,269,324,382,446,515),
(7800,-143,-110,-75,-38,0,40,83,128,175,226,279,337,399,466,540),
(7900,-147,-113,-77,-40,0,42,86,132,182,234,291,351,417,489,569),
(8000,-152,-117,-80,-41,0,43,89,137,189,244,303,367,437,515,604),
(8100,-157,-121,-83,-42,0,45,92,143,197,254,317,385,461,548,651),
(8200,-163,-125,-86,-44,0,46,96,149,205,266,333,407,491,591,738),
(8300,-168,-129,-89,-46,0,48,100,155,215,280,352,433,531,680,0),
(8400,-174,-134,-92,-47,0,50,104,162,226,296,375,471,0,0,0),
(8500,-180,-139,-95,-49,0,53,109,171,239,316,409,0,0,0,0),
(8600,-187,-145,-99,-51,0,55,115,181,256,346,0,0,0,0,0),
(8700,-195,-151,-104,-54,0,58,122,195,283,0,0,0,0,0,0),
(8800,-203,-157,-109,-56,0,62,132,217,0,0,0,0,0,0,0),
(8900,-212,-165,-114,-60,0,67,150,0,0,0,0,0,0,0,0),
(9000,-222,-173,-121,-64,0,0,0,0,0,0,0,0,0,0,0),
            ]
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            interpolasyonlu_veri2 = veri_kumesi_interpolasyonu(data2)
            # print(interpolasyonlu_veri2)

            hedefin_toptan_yuksekligi = yukseklik_farki
            hedefin_toptan_yuksekligi = round(hedefin_toptan_yuksekligi/100)*100
            print("Hedefin Toptan Yüksekliği: ", hedefin_toptan_yuksekligi)
            mesafe = plan_mesafesi
            print("mesafe = ", mesafe)

            if hedefin_toptan_yuksekligi == -400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[1])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[2])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[3])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[4])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -0:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[5])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[6])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[7])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[8])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[9])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 500:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[10])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 600:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[11])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 700:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 800:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 900:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[13])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 1000:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[14])
                        print("Seçilen Bölge:", bolge)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")
            # TXT dosyasının adı ve dosya yolu
            dosya_adı = METRO_DOSYA

            # İlk satırı atlayarak dosyanın geri kalanını yeni bir liste olarak alın
            metrorpr = []
            if not os.path.isfile(dosya_adı):
                QMessageBox.warning(self, "Metro", "METRAP.txt bulunamadı.")
                return None
            with open(dosya_adı, "r") as dosya:
                for satır in dosya:
                    metrorpr.append(satır.strip())

            # Satırları yan yana birleştirerek elde etmek için join() yöntemini kullanıyoruz
            birlesik_veri = "\n".join(metrorpr[1:])

            rapor_tipi = birlesik_veri[:5]
            oktant = birlesik_veri[5:6]
            metro_koor = birlesik_veri[6:9] + " " + birlesik_veri[9:12]
            tarih = birlesik_veri[13:15]
            saat = birlesik_veri[15:17]
            dakika = str(int(birlesik_veri[17:18]) * 6)
            if dakika == "0":
                dakika = str(int(birlesik_veri[17:18]) * 6) + "0"
            ara = ":"
            saat_dakika = saat + ara + str(dakika)

            gecerlilik = birlesik_veri[18:19]
            metro_istasyonu_rakimi = str(int(birlesik_veri[19:22]) * 10)
            yogunluk = int(birlesik_veri[22:25]) / 10

            print("Rapor Tipi:", rapor_tipi)
            print("Oktant:", oktant)
            print("Metro Koordinat:", metro_koor)
            print("Tarih:", tarih)
            print("Saat:", saat_dakika)
            print("Geçerlilik:", gecerlilik)
            print("Metro İstasyonu Rakımı:", metro_istasyonu_rakimi)
            print("Yoğunluk:", yogunluk)
            if bolge == 0:
                ruzgar_istikameti = int(birlesik_veri[28:30]) * 100
                ruzgar_hizi = int(birlesik_veri[30:32])
                if birlesik_veri[32] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[32:35]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[32:35]) / 10
                if birlesik_veri[35] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[35:38]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[35:38]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)  # 39-40-41 3 say

            elif bolge == 1:
                ruzgar_istikameti = int(birlesik_veri[41:43]) * 100
                ruzgar_hizi = int(birlesik_veri[43:45])
                if birlesik_veri[45] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[45:48]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[45:48]) / 10
                if birlesik_veri[48] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[48:51]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[48:51]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)
            elif bolge == 2:
                ruzgar_istikameti = int(birlesik_veri[54:56])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[56:58])  # 3
                if birlesik_veri[58] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[58:61])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[58:61])/10  # 4
                if birlesik_veri[61] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[61:64])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[61:64])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 3:
                ruzgar_istikameti = int(birlesik_veri[67:69])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[69:71])  # 3
                if birlesik_veri[71] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[71:74])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[71:74])/10  # 4
                if birlesik_veri[74] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[74:77])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[74:77])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 4:
                ruzgar_istikameti = int(birlesik_veri[80:82])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[82:84])  # 3
                if birlesik_veri[84] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[84:87])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[84:87])/10  # 4
                if birlesik_veri[87] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[87:90])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[87:90])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 5:
                ruzgar_istikameti = int(birlesik_veri[93:95])*100  # 3
                ruzgar_hizi = int(birlesik_veri[95:97])  # 3
                if birlesik_veri[97] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[97:100])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[97:100])/10  # 4
                if birlesik_veri[100] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[100:103])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[100:103])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 6:
                ruzgar_istikameti = int(birlesik_veri[106:108])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[108:110])  # 3
                if birlesik_veri[110] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[110:113])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[110:113])/10  # 4
                if birlesik_veri[113] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[113:116])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[113:116])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 7:
                ruzgar_istikameti = int(birlesik_veri[119:121])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[121:123])  # 3
                if birlesik_veri[123] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[123:126])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[123:126])/10  # 4
                if birlesik_veri[126] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[126:129])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[126:129])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 8:
                ruzgar_istikameti = int(birlesik_veri[132:134])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[134:136])  # 3
                if birlesik_veri[136] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[136:139])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[136:139])/10  # 4
                if birlesik_veri[139] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[139:142])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[139:142])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 9:
                ruzgar_istikameti = int(birlesik_veri[145:147])*100  # 3
                ruzgar_hizi = int(birlesik_veri[147:149])  # 3
                if birlesik_veri[149] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[149:152])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[149:152])/10  # 4
                if birlesik_veri[152] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[152:155])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[152:155])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 10:
                ruzgar_istikameti = int(birlesik_veri[158:160])*100  # 3
                ruzgar_hizi = int(birlesik_veri[160:162])  # 3
                if birlesik_veri[162] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[162:165])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[162:165])/10  # 4
                if birlesik_veri[165] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[165:168])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[165:168])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)
            # TAMAMLAYICI MESAFEYİ BULUYORUZ.
            if hedefin_toptan_yuksekligi == -400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[1])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[2])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[3])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[4])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -0:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[5])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[6])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[7])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[8])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[9])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 500:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[10])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 600:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[11])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 700:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 800:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 900:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[13])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 1000:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[14])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")
            bt_metro_ist_rakim_farki = batarya_rakimi - int(metro_istasyonu_rakimi)

            giris_mesafesi = round((int(mesafe) + int(tamamlayici_mesafe))/100)*100
            print("Giriş Mesafesi :", giris_mesafesi)

            ruzgarin_plan_istikameti_100 = ruzgar_istikameti - atis_istikameti

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            isi_dzl = [
                (-390,0.9),
                (-380,0.9),
                (-370,0.9),
                (-360,0.8),
                (-350,0.8),
                (-340,0.8),
                (-330,0.8),
                (-320,0.7),
                (-310,0.7),
                (-300,0.7),
                (-290,0.7),
                (-280,0.7),
                (-270,0.7),
                (-260,0.6),
                (-250,0.6),
                (-240,0.6),
                (-230,0.6),
                (-220,0.5),
                (-210,0.5),
                (-200,0.5),
                (-190,0.4),
                (-180,0.4),
                (-170,0.4),
                (-160,0.3),
                (-150,0.3),
                (-140,0.3),
                (-130,0.3),
                (-120,0.2),
                (-110,0.2),
                (-100,0.2),
                (-90,0.2),
                (-80,0.2),
                (-70,0.2),
                (-60,0.1),
                (-50,0.1),
                (-40,0.1),
                (-30,0.1),
                (-20,0),
                (-10,0),
                (0,0),
                (10,0),
                (20,0),
                (30,-0.1),
                (40,-0.1),
                (50,-0.1),
                (60,-0.1),
                (70,-0.2),
                (80,-0.2),
                (90,-0.2),
                (100,-0.2),
                (110,-0.2),
                (120,-0.2),
                (130,-0.3),
                (140,-0.3),
                (150,-0.3),
                (160,-0.3),
                (170,-0.4),
                (180,-0.4),
                (190,-0.4),
                (200,-0.5),
                (210,-0.5),
                (220,-0.5),
                (230,-0.6),
                (240,-0.6),
                (250,-0.6),
                (260,-0.6),
                (270,-0.7),
                (280,-0.7),
                (290,-0.7),
                (300,-0.7),
                (310,-0.7),
                (320,-0.7),
                (330,-0.8),
                (340,-0.8),
                (350,-0.8),
                (360,-0.8),
                (370,-0.9),
                (380,-0.9),
                (390,-0.9),
            ]


            isi_data5bh = []
            for i in range(len(isi_dzl)-1):
                start = isi_dzl[i]
                end = isi_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    isi_data5bh.append((int(x), y))

            isi_data5bh.append(isi_dzl[-1])


            yogunluk_dzl = [
                (-390,3.9),
                (-380,3.8),
                (-370,3.7),
                (-360,3.6),
                (-350,3.5),
                (-340,3.4),
                (-330,3.3),
                (-320,3.2),
                (-310,3.1),
                (-300,3),
                (-290,2.9),
                (-280,2.8),
                (-270,2.7),
                (-260,2.6),
                (-250,2.5),
                (-240,2.4),
                (-230,2.3),
                (-220,2.2),
                (-210,2.1),
                (-200,2),
                (-190,1.9),
                (-180,1.8),
                (-170,1.7),
                (-160,1.6),
                (-150,1.5),
                (-140,1.4),
                (-130,1.3),
                (-120,1.2),
                (-110,1.1),
                (-100,1),
                (-90,0.9),
                (-80,0.8),
                (-70,0.7),
                (-60,0.6),
                (-50,0.5),
                (-40,0.4),
                (-30,0.3),
                (-20,0.2),
                (-10,0.1),
                (0,0),
                (10,-0.1),
                (20,-0.2),
                (30,-0.3),
                (40,-0.4),
                (50,-0.5),
                (60,-0.6),
                (70,-0.7),
                (80,-0.8),
                (90,-0.9),
                (100,-1),
                (110,-1.1),
                (120,-1.2),
                (130,-1.3),
                (140,-1.4),
                (150,-1.5),
                (160,-1.6),
                (170,-1.7),
                (180,-1.8),
                (190,-1.9),
                (200,-2),
                (210,-2.1),
                (220,-2.2),
                (230,-2.3),
                (240,-2.4),
                (250,-2.5),
                (260,-2.6),
                (270,-2.7),
                (280,-2.8),
                (290,-2.9),
                (300,-3),
                (310,-3.1),
                (320,-3.2),
                (330,-3.3),
                (340,-3.4),
                (350,-3.5),
                (360,-3.6),
                (370,-3.7),
                (380,-3.8),
                (390,-3.9),
            ]

            yogunluk_data5bh = []

            for i in range(len(yogunluk_dzl)-1):
                start = yogunluk_dzl[i]
                end = yogunluk_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    yogunluk_data5bh.append((int(x), y))

            yogunluk_data5bh.append(yogunluk_dzl[-1])
            # print(yogunluk_data5bh)


            for i in isi_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    isi_duzeltmesi = i[1]
                    print("Isı Düzeltmesi :", isi_duzeltmesi)


            for i in yogunluk_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    yogunluk_duzeltmesi = i[1]
                    print("Yoğunluk Düzeltmesi", yogunluk_duzeltmesi)

            duzeltilmis_sıcaklik_degeri = isi_duzeltmesi + hava_sicakligi
            print("Düzeltilmiş Isı Değeri:", duzeltilmis_sıcaklik_degeri)

            duzeltilmis_yogunluk_degeri = round(yogunluk_duzeltmesi + hava_yogunlugu,1)
            print("Düzeltilmiş Yoğunluk Değeri:", duzeltilmis_yogunluk_degeri)
        # RUZGAR BİLEŞENLERİ

            atis_istikameti = round(atis_istikameti/100)*100

            if ruzgar_istikameti < atis_istikameti:
                ruzgar_istikameti = ruzgar_istikameti + 6400
            else:
                ruzgar_istikameti = ruzgar_istikameti
            print("ruzgar istikameti", ruzgar_istikameti)

            print("atış istikameti", atis_istikameti)

            ruzgarin_plan_istikameti = ruzgar_istikameti - atis_istikameti

            print("Rüzgarın Plan İstikameti:", ruzgarin_plan_istikameti)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            data1 = [
                    (0,0,1),
                (100,-0.1,0.99),
                (200,-0.2,0.98),
                (300,-0.29,0.96),
                (400,-0.38,0.92),
                (500,-0.47,0.88),
                (600,-0.56,0.83),
                (700,-0.63,0.77),
                (800,-0.71,0.71),
                (900,-0.77,0.63),
                (1000,-0.83,0.56),
                (1100,-0.88,0.47),
                (1200,-0.92,0.38),
                (1300,-0.96,0.29),
                (1400,-0.98,0.2),
                (1500,-0.99,0.1),
                (1600,-1,0),
                (1700,-0.99,-0.1),
                (1800,-0.98,-0.2),
                (1900,-0.96,-0.29),
                (2000,-0.92,-0.38),
                (2100,-0.88,-0.47),
                (2200,-0.83,-0.56),
                (2300,-0.77,-0.63),
                (2400,-0.71,-0.71),
                (2500,-0.63,-0.77),
                (2600,-0.56,-0.83),
                (2700,-0.47,-0.88),
                (2800,-0.38,-0.92),
                (2900,-0.29,-0.96),
                (3000,-0.2,-0.98),
                (3100,-0.1,-0.99),
                (3200,0,-1),
                (3300,0.1,-0.99),
                (3400,0.2,-0.98),
                (3500,0.29,-0.96),
                (3600,0.38,-0.92),
                (3700,0.47,-0.88),
                (3800,0.56,-0.83),
                (3900,0.63,-0.77),
                (4000,0.71,-0.71),
                (4100,0.77,-0.63),
                (4200,0.83,-0.56),
                (4300,0.88,-0.47),
                (4400,0.92,-0.38),
                (4500,0.96,-0.29),
                (4600,0.98,-0.2),
                (4700,0.99,-0.1),
                (4800,1,0),
                (4900,0.99,0.1),
                (5000,0.98,0.2),
                (5100,0.96,0.29),
                (5200,0.92,0.38),
                (5300,0.88,0.47),
                (5400,0.83,0.56),
                (5500,0.77,0.63),
                (5600,0.71,0.71),
                (5700,0.63,0.77),
                (5800,0.56,0.83),
                (5900,0.47,0.88),
                (6000,0.38,0.92),
                (6100,0.29,0.96),
                (6200,0.2,0.98),
                (6300,0.1,0.99),
                (6400,0,1),


            ]


            # veri_noktasi[0] = MESAFE
            # veri_noktasi[1] = YAN RÜZGARI
            # veri_noktasi[2] = MESAFE RÜZGARI

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Mesafe :", mesafe)

            for rzdlz in interpolasyonlu_veri1:
                if rzdlz[0] == ruzgarin_plan_istikameti:
                    yan_ruzgari_bileseni = rzdlz[1]
                    mesafe_ruzgari_bileseni = rzdlz[2]
                    print("Yan Rüzgarı Bileşeni:", yan_ruzgari_bileseni)
                    print("mesafe Rüzgarı Bileşeni:", mesafe_ruzgari_bileseni)


            # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (2500,112.8,7.8,0.27,7.7,1.6,0.19,7.9,-8,2.9,-3.1,7,-8.1,-2.9,2.8,-11,12),
                (2600,118,8.1,0.26,8.1,1.7,0.19,8,-8.2,3.2,-3.3,7.6,-8.6,-3,2.9,-11,12),
                (2700,123.2,8.4,0.25,8.4,1.8,0.2,8.2,-8.4,3.4,-3.5,8.1,-9.1,-3.2,3.1,-11,12),
                (2800,128.4,8.8,0.24,8.8,1.9,0.2,8.4,-8.5,3.6,-3.7,8.7,-9.6,-3.4,3.2,-11,12),
                (2900,133.7,9.1,0.23,9.1,2,0.21,8.5,-8.7,3.8,-3.9,9.3,-10.1,-3.5,3.4,-11,12),
                (3000,139,9.5,0.22,9.4,2,0.21,8.7,-8.8,4.1,-4.1,9.9,-10.6,-3.7,3.6,-11,12),
                (3100,144.4,9.8,0.21,9.8,2.1,0.22,88,-9,4.3,-4.3,10.5,-11.1,-3.9,3.8,-11,12),
                (3200,149.8,10.2,0.2,10.1,2.2,0.22,9,-9.1,4.5,-4.5,11,-11.6,-4.1,3.9,-11,12),
                (3300,155.3,10.5,0.2,10.5,2.3,0.23,9.2,-9.3,4.8,-4.7,11.6,-12.1,-4.3,4.1,-11,12),
                (3400,160.8,10.9,0.19,10.8,2.4,0.23,9.3,-9.4,5,-4.9,12.2,-12.6,-4.5,4.3,-11,12),
                (3500,166.4,11.2,0.19,11.2,2.5,0.23,9.5,-9.6,5.3,-5,12.8,-13.1,-4.6,4.5,-11,12),
                (3600,172,11.6,0.18,11.6,2.6,0.24,9.6,-9.7,5.5,-5.2,13.4,-13.6,-4.8,4.7,-11,12),
                (3700,177.7,12,0.17,11.9,2.7,0.24,9.8,-9.9,5.7,-5.4,13.9,-14.1,-5,4.9,-11,12),
                (3800,183.4,12.3,0.17,12.3,2.8,0.25,9.9,-10,6,-5.6,14.5,-14.6,-5.3,5.1,-11,12),
                (3900,189.2,12.7,0.16,12.7,2.9,0.25,10.1,-10.1,6.2,-5.8,15.1,-15,-5.5,5.3,-11,12),
                (4000,195,13.1,0.16,13,3,0.25,10.2,-10.3,6.5,-6,15.6,-15.5,-5.7,5.6,-11,12),
                (4100,200.9,13.4,0.16,13.4,3.1,0.26,10.4,-10.4,6.7,-6.2,16.2,-16,-5.9,5.8,-11,12),
                (4200,206.8,13.8,0.15,13.8,3.2,0.26,10.5,-10.5,6.9,-6.4,16.8,-16.5,-6.1,6,-11,11),
                (4300,212.8,14.2,0.15,14.1,3.3,0.27,10.7,-10.7,7.2,-6.6,17.3,-16.9,-6.4,6.3,-11,11),
                (4400,218.9,14.6,0.14,14.5,3.4,0.27,10.8,-10.8,7.4,-6.8,17.9,-17.4,-6.6,6.5,-10,11),
                (4500,225,15,0.14,14.9,3.5,0.28,11,-11,7.7,-7,18.4,-17.8,-6.8,6.7,-10,11),
                (4600,231.2,15.4,0.14,15.3,3.6,0.28,11.2,-11.1,7.9,-7.2,18.9,-18.3,-7.1,7,-10,11),
                (4700,237.4,15.7,0.13,15.7,3.8,0.28,11.3,-11.2,8.1,-7.4,19.4,-18.7,-7.3,7.3,-10,11),
                (4800,243.7,16.1,0.13,16.1,3.9,0.29,11.5,-11.4,8.4,-7.6,20,-19.2,-7.6,7.5,-10,11),
                (4900,250.1,16.5,0.13,16.5,4,0.29,11.6,-11.5,8.6,-7.8,20.5,-19.6,-7.9,7.8,-10,11),
                (5000,256.6,16.9,0.12,16.9,4.1,0.3,11.8,-11.7,8.9,-8,21,-20,-8.1,8.1,-9,10),
                (5100,263.1,17.3,0.12,17.3,4.2,0.3,12,-11.8,9.1,-8.2,21.5,-20.4,-8.4,8.4,-9,10),
                (5200,269.7,17.8,0.12,17.7,4.4,0.3,12.1,-11.9,9.3,-8.3,21.9,-20.9,-8.7,8.6,-9,10),
                (5300,276.4,18.2,0.12,18.1,4.5,0.31,12.3,-12.1,9.6,-8.5,22.4,-21.3,-9,8.9,-9,10),
                (5400,283.2,18.6,0.11,18.5,4.6,0.31,12.5,-12.2,9.8,-8.7,22.9,-21.7,-9.2,9.2,-9,10),
                (5500,290.1,19,0.11,18.9,4.7,0.32,12.6,-12.4,10,-8.9,23.3,-22.1,-9.5,9.5,-8,10),
                (5600,297.1,19.4,0.11,19.3,4.9,0.32,12.8,-12.5,10.3,-9.1,23.8,-22.5,-9.8,9.9,-8,9),
                (5700,304.1,19.9,0.11,19.7,5,0.33,13,-12.7,10.5,-9.3,24.2,-22.8,-10.1,10.2,-8,9),
                (5800,311.3,20.3,0.1,20.2,5.2,0.33,13.1,-12.8,10.7,-9.5,24.7,-23.2,-10.5,10.5,-8,9),
                (5900,318.6,20.7,0.1,20.6,5.3,0.33,13.3,-13,11,-9.7,25.1,-23.6,-10.8,10.8,-7,9),
                (6000,326,21.2,0.1,21.1,5.4,0.34,13.5,-13.1,11.2,-9.8,25.5,-24,-11.1,11.2,-7,9),
                (6100,333.5,21.6,0.1,21.5,5.6,0.34,13.7,-13.3,11.4,-10,25.9,-24.3,-11.4,11.5,-7,8),
                (6200,341.1,22.1,0.1,22,5.8,0.35,13.8,-13.4,11.7,-10.2,26.3,-24.7,-11.8,11.9,-6,8),
                (6300,348.8,22.5,0.09,22.4,5.9,0.35,14,-13.6,11.9,-10.4,26.7,-25,-12.1,12.2,-6,8),
                (6400,356.7,23,0.09,22.9,6.1,0.36,14.2,-13.7,12.1,-10.6,27.1,-25.4,-12.4,12.6,-6,8),
                (6500,364.8,23.5,0.09,23.4,6.2,0.36,14.4,-13.9,12.3,-10.7,27.5,-25.7,-12.8,12.9,-6,7),
                (6600,372.9,24,0.09,23.8,6.4,0.37,14.6,-14.1,12.6,-10.9,27.8,-26,-13.2,13.3,-5,7),
                (6700,381.3,24.5,0.09,24.3,6.6,0.37,14.8,-14.2,128,-11.1,28.2,-26.3,-13.5,13.7,-5,7),
                (6800,389.8,25,0.09,24.8,6.8,0.38,14.9,-14.4,13,-11.3,28.5,-26.6,-13.9,14.1,-5,6),
                (6900,398.5,25.5,0.08,25.3,7,0.38,15.1,-14.5,13.2,-11.4,28.9,-27,-14.3,14.5,-4,6),
                (7000,407.4,26,0.08,25.8,7.2,0.39,15.3,-14.7,13.4,-11.6,29.2,-27.3,-14.6,14.9,-4,6),
                (7100,416.5,26.5,0.08,26.4,7.4,0.39,15.5,-14.9,13.6,-11.8,29.5,-27.5,-15,15.3,-3,6),
                (7200,425.9,27.1,0.08,26.9,7.6,0.4,15.7,-15.1,13.9,-12,29.8,-27.8,-15.4,15.7,-3,5),
                (7300,435.5,27.6,0.08,27.4,7.8,0.4,15.9,-15.2,14.1,-12.1,30.1,-28.1,-15.8,16.1,-3,5),
                (7400,445.3,28.2,0.08,28,8,0.41,16.1,-15.4,14.3,-12.3,30.4,-28.4,-16.2,16.6,-2,5),
                (7500,455.5,28.7,0.08,28.6,8.3,0.42,16.3,-15.6,14.5,-12.5,30.7,-28.6,-16.6,17,-2,4),
                (7600,465.9,29.3,0.07,29.1,8.5,0.42,16.6,-15.8,14.7,-12.6,30.9,-28.9,-17.1,17.5,-2,4),
                (7700,476.8,29.9,0.07,29.7,8.8,0.43,16.8,-15.9,14.9,-12.8,31.2,-29.2,-17.5,17.9,-1,4),
                (7800,488,30.6,0.07,30.4,9.1,0.44,17,-16.1,15.1,-13,31.4,-29.4,-17.9,18.4,-1,3),
                (7900,499.7,31.2,0.07,31,9.4,0.44,17.2,-16.3,15.3,-13.1,31.6,-29.6,-18.4,18.9,0,3),
                (8000,511.8,31.9,0.07,31.7,9.7,0.45,17.5,-16.5,15.5,-13.3,31.8,-29.9,-18.8,19.4,0,2),
                (8100,524.6,32.6,0.07,32.4,10,0.46,17.7,-16.7,15.7,-13.4,32,-30.1,-19.3,19.9,1,2),
                (8200,538,33.3,0.07,33.1,10.4,0.47,17.9,-16.9,15.9,-13.6,32.2,-30.3,-19.7,20.4,1,2),
                (8300,552.2,34.1,0.06,33.8,10.7,0.47,18.2,-17.1,15.9,-13.8,32.4,-30.5,-20.2,20.9,2,1),
                (8400,567.3,34.9,0.06,34.6,11.2,0.48,18.4,-17.3,15.9,-13.9,32.5,-30.7,-20.7,21.5,2,1),
                (8500,583.7,35.7,0.06,35.5,11.6,0.49,18.7,-17.5,15.9,-14.1,32.6,-30.9,-21.2,22,3,0),
                (8600,601.4,36.6,0.06,36.4,12.2,0.5,19,-17.7,15.9,-14.2,32.7,-31,-21.7,22.6,3,0),
                (8700,621.2,37.7,0.06,37.4,12.8,0.52,19.3,-18,15.9,-14.4,32.6,-31.2,-22.2,23.3,4,0),
                (8800,643.9,38.8,0.06,38.5,13.5,0.53,19.6,-18.2,15.9,-14.5,32.6,-31.3,-22.7,24,4,-1),
                (8900,671.1,40.2,0.05,39.9,14.4,0.55,19.6,-18.4,15.9,-14.6,32.6,-31.5,-23.3,24,5,-1),
                (9000,708.2,42,0.05,41.7,15.7,0.57,19.6,-18.7,15.9,-14.8,32.6,-31.6,-23.9,24,6,-2),

            ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == giris_mesafesi:
                    nisangah6bh = round(nsg[1], 3)
                    # print("Nişangah:", nisangah)

            for tsg in interpolasyonlu_veri1:
                if tsg[0] == giris_mesafesi:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            for ts in interpolasyonlu_veri1:
                if ts[0] == giris_mesafesi:
                    deltaTS = round(ts[3], 3)
                    # print("Delta Tapa Saniyesi:", deltaTS)


            for ucs in interpolasyonlu_veri1:
                if ucs[0] == giris_mesafesi:
                    ucussuresi = round(ucs[4], 3)
                    # print("Uçuş Süresi:", ucussuresi)

            for dgl in interpolasyonlu_veri1:
                if dgl[0] == giris_mesafesi:
                    dogalyandz = round(dgl[5], 3)
                    # print("Doğal Yan Düzeltmesi:", dogalyandz)

            for yrd in interpolasyonlu_veri1:
                if yrd[0] == giris_mesafesi:
                    yan_ruzgari_duzeltme_faktoru = round(yrd[6], 1)
                    # print("Yan Rüzgarı Düzeltmesi:", yanruzgaridz)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_eksilme = round(ihz[7], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimeksilme)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_martma = round(ihz[8], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimartma)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_bas = round(mrd[9], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (BAŞ):", mesaferuzgaridzl_bas)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_arka = round(mrd[10], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (ARKA):", mesaferuzgaridzl_arka)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_eksilme = round(hvs[11], 3)
                    # print("Hava Isısı Düzeltmesi (Eksilme):", havaisisi_eksilme)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_artma = round(hvs[12], 3)
                    # print("Hava Isısı Düzeltmesi (Artma):", havaisisi_artma)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_eksilme = round(hyd[13], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", havayogunlugu_eksilme)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_artma = round(hyd[14], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", havayogunlugu_artma)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_eksilme = round(mka[15], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", mermikareagirligi_eksilme)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_artma = round(mka[16], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", mermikareagirligi_artma)

            print("Yan Rüzgarı Düzeltme Faktörü:", yan_ruzgari_duzeltme_faktoru)

            mesafe_ruzgari = ruzgar_hizi * mesafe_ruzgari_bileseni
            print("mesafe ruzgari", mesafe_ruzgari)

            ruzgar_yan_duzeltmesi = round(
                (ruzgar_hizi * yan_ruzgari_bileseni * yan_ruzgari_duzeltme_faktoru), 1)
            print("Rüzgar Yan Düzeltmesi", ruzgar_yan_duzeltmesi)

            dogalyandz = round(dogalyandz, 1)
            print("Doğal Yan Düzeltmesi:", dogalyandz)

            atis_istikameti = round(atis_istikameti/400)*400

            # dünyanın dönmesi düzeltmesi


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (2000,0.3,0.3,0.3,0.3,0.3,0.3,0.3,0.3,0.3),
                (3000,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.5,0.5),
                (4000,0.6,0.6,0.6,0.6,0.6,0.6,0.6,0.6,0.6),
                (5000,0.7,0.7,0.7,0.7,0.8,0.8,0.8,0.8,0.8),
                (6000,0.8,0.8,0.9,0.9,1,1,1,1.1,1.1),
                (7000,1,1,1,1.1,1.2,1.2,1.3,1.3,1.3),
                (8000,1.1,1.1,1.2,1.3,1.4,1.5,1.6,1.7,1.7),
                (9000,1.3,1.3,1.4,1.6,1.8,2.1,2.3,2.4,2.4),
            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            if atis_istikameti == 00 or atis_istikameti == 6400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[1], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 6000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[2], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 5600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[3], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 5200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[4], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 4800:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[5], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2000 or atis_istikameti == 4400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[6], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2400 or atis_istikameti == 4000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[7], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2800 or atis_istikameti == 3600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[8], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 3200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[9], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (2000,0,-3,-5,-8,-10,-11,-13,-13,-14,0,3,5,8,10,11,13,14),
                (3000,0,-4,-8,-11,-14,-16,-18,-19,-20,0,4,8,11,14,16,19,20),
                (4000,0,-5,-10,-14,-18,-21,-23,-25,-25,0,5,10,14,18,21,25,25),
                (5000,0,-6,-11,-17,-21,-25,-28,-29,-30,0,6,11,17,21,25,29,30),
                (6000,0,-7,-13,-19,-24,-28,-32,-33,-34,0,7,13,19,24,28,33,34),
                (7000,0,-7,-14,-21,-26,-31,-34,-37,-37,0,7,14,21,26,31,37,37),
                (8000,0,-8,-15,-22,-27,-32,-36,-38,-39,0,8,15,22,27,32,38,39),
                (9000,0,-7,-14,-20,-26,-30,-33,-35,-36,0,7,14,20,26,30,35,36),




            ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)
            if atis_istikameti == 0 or atis_istikameti == 3200 or atis_istikameti == 6400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[1], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 200 or atis_istikameti == 3000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[2], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 2800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[3], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 600 or atis_istikameti == 2600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[4], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 2400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[5], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1000 or atis_istikameti == 2200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[6], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 2000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[7], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1400 or atis_istikameti == 1800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[8], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 1600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[9], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            if atis_istikameti == 3400 or atis_istikameti == 6200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[11], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3600 or atis_istikameti == 6000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[12], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3800 or atis_istikameti == 5800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[13], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4000 or atis_istikameti == 5600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[14], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4200 or atis_istikameti == 5400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[15], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4400 or atis_istikameti == 5200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[16], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4600 or atis_istikameti == 5000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[17], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4800 or atis_istikameti == 4800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[18], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            print("Dünyanın Dönmesi Düzeltmesi:", arzindonusuyanduzeltmesi)

            metro_yan_duzeltmesi = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            metro_yan_duzeltmesi_1 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi:", metro_yan_duzeltmesi)
            metro_yan_duzeltmesi = round(metro_yan_duzeltmesi/1)*1
            print("Metro Yan Düzeltmesi en yakın 1 milyem:", metro_yan_duzeltmesi)


            # METRO MESAFE DÜZELTMESİ
            mesafe_ruzgari_fark = round(mesafe_ruzgari-0, 1)
            if mesafe_ruzgari_fark > 0:
                mesafe_ruzgari_durum = "BAŞ"
            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_durum = "ARKA"
            else:
                mesafe_ruzgari_durum = 0
            print("Mesafe Ruzgari Fark :", mesafe_ruzgari_fark)
            hava_sicakligi_fark = round(duzeltilmis_sıcaklik_degeri - 100, 1)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_durum = "EKSİLME"
            if hava_sicakligi_fark > 0:
                hava_sicakligi_durum = "ARTMA"
            else:
                hava_sicakligi_durum = 0
            print("Hava Sıcaklığı Fark :", hava_sicakligi_fark)
            hava_yogunlugu_fark = round(duzeltilmis_yogunluk_degeri - 100, 1)
            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_durum = "EKSİLME"
            if hava_yogunlugu_fark > 0:
                hava_yogunlugu_durum = "ARTMA"
            else:
                hava_yogunlugu_durum = 0
            print("Hava Yoğunluğu Fark :", hava_yogunlugu_fark)
            mermi_kare_agirligi_fark = round(mermi_kare_agirligi-2, 1)
            if mermi_kare_agirligi_fark < 0:
                merkar_durum = "EKSİLME"
            if mermi_kare_agirligi_fark > 0:
                merkar_durum = "ARTMA"
            if mermi_kare_agirligi_fark == 0:
                merkar_durum = 0

            print("Mermi Kare Ağırlığı Fark :", mermi_kare_agirligi_fark)

            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_arka
            else:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_bas
            if mesafe_ruzgari_fark == 0:
                mesafe_ruzgari_duzeltme_birimi = 0

            print("Mesafe Rüzgarı Düzeltme Birimi:", mesafe_ruzgari_duzeltme_birimi)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_duzeltme_birimi = havaisisi_eksilme
            else:
                hava_sicakligi_duzeltme_birimi = havaisisi_artma
            if hava_sicakligi_fark == 0:
                hava_sicakligi_duzeltme_birimi = 0
            print("Hava Isısı Düzeltme Birimi:", hava_sicakligi_duzeltme_birimi)

            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_eksilme
            else:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_artma
            if hava_yogunlugu_fark == 0:
                hava_yogunlugu_duzeltme_birimi = 0
            print("Hava Yoğunluğu Düzeltme Birimi:", hava_yogunlugu_duzeltme_birimi)

            if mermi_kare_agirligi_fark < 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_eksilme
            elif mermi_kare_agirligi_fark > 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_artma
            else:
                mermi_kare_agirligi_duzeltme_birimi = 0
            mesafe_ruzgari_fark = abs(mesafe_ruzgari_fark)
            hava_sicakligi_fark = abs(hava_sicakligi_fark)
            hava_yogunlugu_fark = abs(hava_yogunlugu_fark)


            print("Mermi Kare Ağırlığı Düzeltme Birimi:",
                mermi_kare_agirligi_duzeltme_birimi)

            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi = round(metro_yan_duzeltmesi + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)
            gac_yan_duzeltmesi = round(toplam_yan_duzeltmesi - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            brt_isisi_ilk_hiz_dzl = [
                                (-40,-12.5),
                                (-30,-11.6),
                                (-20,-10.6),
                                (-10,-9.6),
                                (0,-8.6),
                                (10,-7.5),
                                (20,-6.4),
                                (30,-5.2),
                                (40,-4),
                                (50,-2.7),
                                (60,-1.4),
                                (70,0),
                                (80,1.4),
                                (90,2.9),
                                (100,4.4),
                                (110,5.9),
                                (120,7.5),
                                (130,9.1),]

            barut_isisi_dzl = []
            for i in range(len(brt_isisi_ilk_hiz_dzl)-1):
                start = brt_isisi_ilk_hiz_dzl[i]
                end = brt_isisi_ilk_hiz_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    barut_isisi_dzl.append((int(x), y))

            barut_isisi_dzl.append(brt_isisi_ilk_hiz_dzl[-1])


            for i in barut_isisi_dzl:
                if i[0] == barut_isisi:
                    barut_isisi_duzeltmesi = round(i[1], 1)
            # HIZ DEĞİŞİKLİĞİNİN HESAPLANMASI
            hiz_degisikligi = round(ilk_hiz_farki + mevzi_hiz_degisikligi, 1)
            deltaV_hiz_farki = round(hiz_degisikligi + barut_isisi_duzeltmesi, 1)
            print("Hız Değişikliği:", hiz_degisikligi)
            print("Barut Isısı Düzeltmesi:", barut_isisi_duzeltmesi)
            print("Delta V Hız Farkı:", deltaV_hiz_farki)

            if deltaV_hiz_farki > 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_martma
            elif deltaV_hiz_farki < 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_eksilme
            else:
                ilk_hiz_duzeltme_birimi = 0
            print("İlk Hız Düzeltme Birimi:", ilk_hiz_duzeltme_birimi)

            deltaV_mesafe_duzeltmesi = abs(
                round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            if deltaV_hiz_farki > 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))*-1
            elif deltaV_hiz_farki < 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            else:
                deltaV_mesafe_duzeltmesi = 0
            print("Delta V Mesafe Düzeltmesi:", deltaV_mesafe_duzeltmesi)

            #### TOPLAM MESAF DÜZELTMESİ ###
            toplam_mesafe_duzeltmesi = round(metro_mesafe_duzeltmesi + deltaV_mesafe_duzeltmesi)
            print("Toplam Mesafe Düzeltmesi:", toplam_mesafe_duzeltmesi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)

            # toplam yan düzeltmede 1 milyem fazla mesafe düzeltmesinde işaret + ve 335 çıktı
            # baiek de yan 4 sl mesafe -326 çıktı

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,0,0,0,0,0,0,0,0,0,0),
            (1,0,0,0,0,0,0,0,0,0,0),
            (2,-0.006,0.006,0,0,0,0,0,0,0.011,-0.011),
            (3,-0.008,0.008,0,0.001,0,0.001,0.001,-0.001,0.016,-0.016),
            (4,-0.01,0.011,0,0.001,0,0.002,0.001,-0.001,0.019,-0.019),
            (5,-0.012,0.012,0,0.002,-0.001,0.004,0.002,-0.002,0.022,-0.022),
            (6,-0.013,0.014,-0.001,0.002,-0.002,0.006,0.003,-0.002,0.024,-0.024),
            (7,-0.015,0.016,-0.001,0.003,-0.004,0.008,0.003,-0.003,0.026,-0.026),
            (8,-0.016,0.017,-0.002,0.004,-0.006,0.01,0.004,-0.004,0.028,-0.028),
            (9,-0.017,0.018,-0.003,0.005,-0.008,0.013,0.004,-0.004,0.029,-0.03),
            (10,-0.018,0.019,-0.003,0.006,-0.011,0.015,0.005,-0.005,0.031,-0.031),
            (11,-0.019,0.02,-0.004,0.006,-0.013,0.018,0.006,-0.005,0.032,-0.032),
            (12,-0.021,0.021,-0.005,0.007,-0.016,0.02,0.006,-0.006,0.033,-0.034),
            (13,-0.022,0.022,-0.006,0.008,-0.018,0.023,0.007,-0.007,0.031,-0.035),
            (14,-0.023,0.023,-0.007,0.009,-0.021,0.025,0.008,-0.007,0.035,-0.036),
            (15,-0.024,0.024,-0.007,0.01,-0.024,0.028,0.008,-0.008,0.036,-0.037),
            (16,-0.025,0.025,-0.008,0.011,-0.027,0.031,0.009,-0.009,0.037,-0.038),
            (17,-0.026,0.026,-0.009,0.012,-0.029,0.033,0.01,-0.009,0.0373,-0.039),
            (18,-0.027,0.027,-0.01,0.012,-0.032,0.035,0.011,-0.01,0.038,-0.04),
            (19,-0.028,0.028,-0.011,0.013,-0.035,0.038,0.012,-0.011,0.039,-0.04),
            (20,-0.029,0.029,-0.011,0.014,-0.037,0.04,0.012,-0.012,0.039,-0.041),
            (21,-0.03,0.031,-0.012,0.015,-0.04,0.042,0.0133,-0.013,0.04,-0.042),
            (22,-0.031,0.032,-0.013,0.015,-0.042,0.045,0.014,-0.013,0,-0.043),
            (23,-0.032,0.033,-0.014,0.016,-0.045,0.047,0.015,-0.014,0.041,-0.043),
            (24,-0.033,0.034,-0.014,0.017,-0.047,0.049,0.016,-0.015,0.042,-0.044),
            (25,-0.035,0.035,-0.015,0.017,-0.049,0.051,0.017,-0.016,0.042,-0.045),
            (26,-0.036,0.036,-0.016,0.018,-0.052,0.053,0.018,-0.017,0.043,-0.045),
            (27,-0.037,0.037,-0.017,0.019,-0.054,0.055,0.02,-0.018,0.043,-0.046),
            (28,-0.038,0.038,-0.017,0.019,-0.056,0.057,0.021,-0.02,0.043,-0.046),
            (29,-0.039,0.039,-0.018,0.02,-0.058,0.059,0.022,-0.021,0.044,-0.047),
            (30,-0.04,0.04,-0.018,0.02,-0.06,0.06,0.023,-0.022,0.044,-0.048),
            (31,-0.042,0.042,-0.019,0.021,-0.062,0.062,0.024,-0.023,0.045,-0.048),
            (32,-0.043,0.043,-0.02,0.021,-0.064,0.064,0.026,-0.024,0.045,-0.049),
            (33,-0.044,0.044,-0.02,0.022,-0.066,0.065,0.027,-0.025,0.046,-0.049),
            (34,-0.046,0.045,-0.021,0.022,-0.068,0.067,0.028,-0.027,0.046,-0.05),
            (35,-0.047,0.047,-0.021,0.022,-0.07,0.069,0.029,-0.028,0.046,-0.051),
            (36,-0.048,0.048,-0.021,0.023,-0.071,0.07,0.031,-0.029,0.047,-0.051),
            (37,-0.05,0.049,-0.022,0.023,-0.073,0.072,0.032,-0.03,0.047,-0.052),
            (38,-0.051,0.05,-0.022,0.023,-0.075,0.073,0.033,-0.032,0.048,-0.053),
            (39,-0.052,0.052,-0.023,0.024,-0.076,0.075,0.035,-0.033,0.048,-0.053),
            (40,-0.054,0.053,-0.023,0.024,-0.078,0.076,0.036,-0.034,0.049,-0.054),
            (41,-0.055,0.054,-0.023,0.024,-0.079,0.077,0.038,-0.036,0.049,-0.055),
            (42,-0.057,0.056,-0.023,0.024,-0.081,0.079,0.039,-0.037,0.05,-0.055),
            (43,-0.058,0.057,-0.024,0.024,-0.082,0.08,0.04,-0.038,0.05,-0.056),
            (44,-0.06,0.059,-0.024,0.024,-0.084,0.081,0.042,-0.04,0.051,-0.057),
            (45,-0.061,0.06,-0.024,0.024,-0.085,0.083,0.043,-0.041,0.051,-0.058),
            (46,-0.063,0.062,-0.024,0.024,-0.087,0.084,0.045,-0.042,0.052,-0.059),
            (47,-0.064,0.063,-0.024,0.024,-0.088,0.085,0.046,-0.044,0.053,-0.06),
            (48,-0.066,0.064,-0.024,0.024,-0.089,0.087,0.048,-0.045,0.054,-0.061),
            (49,-0.067,0.066,-0.024,0.024,-0.091,0.088,0.049,-0.046,0.054,-0.062),
            (50,-0.069,0.067,-0.024,0.024,-0.092,0.089,0.05,-0.048,0.055,-0.063),
            (51,-0.07,0.069,-0.024,0.024,-0.093,0.091,0.052,-0.049,0.056,-0.064),
            (52,-0.072,0.07,-0.024,0.023,-0.095,0.092,0.053,-0.05,0.057,-0.065),
            (53,-0.074,0.072,-0.024,0.023,-0.096,0.093,0.054,-0.052,0.059,-0.067),
            (54,-0.075,0.074,-0.023,0.023,-0.097,0.094,0.056,-0.053,0.06,-0.068),
            (55,-0.077,0.075,-0.023,0.022,-0.099,0.095,0.057,-0.054,0.061,-0.07),
            (56,-0.079,0.077,-0.023,0.022,-0.1,0.097,0.058,-0.055,0.063,-0.072),
            (57,-0.08,0.078,-0.022,0.021,-0.101,0.098,0.06,-0.057,0.065,-0.074),
            (58,-0.082,0.08,-0.022,0.02,-0.102,0.099,0.061,-0.058,0.0673,-0.076),
            (59,-0.084,0.082,-0.021,0.02,-0.103,0.1,0.062,-0.059,0.069,-0.078),
            (60,-0.086,0.084,-0.02,0.02,-0.104,0.101,0.064,-0.06,0.072,-0.082),
            (61,-0.087,0.086,-0.02,0.024,-0.105,0.102,0.065,-0.062,0.078,-0.086),
            (62,-0.09,0.088,-0.02,0.024,-0.104,0.101,0.068,-0.064,0.089,-0.097),
            (63,-0.094,0.091,-0.025,0.024,-0.098,0.097,0.079,-0.071,0.125,-0.131),
            (64,-0.097,0.103,-0.036,0.024,-0.098,0.087,0.079,-0.085,0.267,-0.195),








            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Tapa Saniyesi :", tapasaniyesi)

            if deltaV_hiz_farki < 0:
                for ihz_eksilme in interpolasyonlu_veri1:
                    if ihz_eksilme[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_eksilme[1], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Eksilme):", ilkhizda1mildegisim_eksilme)
            if deltaV_hiz_farki > 0:
                for ihz_artma in interpolasyonlu_veri1:
                    if ihz_artma[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_artma[2], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Artma):", ilkhizda1mildegisim_artma)
            if mesafe_ruzgari_durum == "BAŞ":
                for musrus_bas in interpolasyonlu_veri1:
                    if musrus_bas[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(musrus_bas[3], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Baş):", mesaferuzgaridegisimi_bas)
            if mesafe_ruzgari_durum == "ARKA":
                for mesruz_arka in interpolasyonlu_veri1:
                    if mesruz_arka[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(mesruz_arka[4], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Arka):", mesaferuzgaridegisimi_arka)
            if hava_sicakligi_durum == "EKSİLME":
                for hadese_eksilme in interpolasyonlu_veri1:
                    if hadese_eksilme[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_eksilme[5], 3)
                        # print("Hava Isısı Değişim Miktarı (Eksilme):", havaisisidegisim_eksilme)
            if hava_sicakligi_durum == "ARTMA":
                for hadese_artma in interpolasyonlu_veri1:
                    if hadese_artma[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_artma[6], 3)
                        # print("Hava Isısı Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == "EKSİLME":
                for hayog_eksilme in interpolasyonlu_veri1:
                    if hayog_eksilme[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_eksilme[7], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Eksilme):", havayogunlugudegisim_eksilme)
            if hava_yogunlugu_durum == "ARTMA":
                for hayog_artma in interpolasyonlu_veri1:
                    if hayog_artma[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_artma[8], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == 0:
                havayogunlugudegisim = 0
            if merkar_durum == "EKSİLME":
                for merkar_eksilme in interpolasyonlu_veri1:
                    if merkar_eksilme[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_eksilme[9], 3)
                        # print("Mermi Kare Ağırlığı Değişim Miktarı (Eksilme):", mermikareagirligidegisim_eksilme)
            if merkar_durum == "ARTMA":
                for merkar_artma in interpolasyonlu_veri1:
                    if merkar_artma[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_artma[10], 3)
            else:
                mermikareagirligidegisim = 0
                # print("Mermi Kare Ağırlığı Değişim Miktarı (Artma):", mermikareagirligidegisim_artma)
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            
            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

            # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            # Veri kümesi
            data1 = [
                (1000,58,14,0.003,-0.003),
                (1500,90,32,0.007,-0.007),
                (2000,124,59,0.014,-0.013),
                (2500,161,96,0.023,-0.022),
                (3000,200,142,0.035,-0.033),
                (3500,242,201,0.053,-0.049),
                (4000,288,273,0.076,-0.07),
                (4500,338,362,0.109,-0.098),
                (5000,393,471,0.156,-0.138),
                (5500,455,606,0.228,-0.196),
                (6000,526,777,0.351,-0.288),
                (6500,611,1007,0.616,-0.455),
                (7000,728,1365,0.616,-0.894),

            ]
            gac_mesafe = plan_mesafesi_2B_obüs + toplam_mesafe_duzeltmesi
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            # Veri kümesi
            data1 = [
                (2500,112.8,7.8,0.27,7.7,1.6,0.19,7.9,-8,2.9,-3.1,7,-8.1,-2.9,2.8,-11,12),
                (2600,118,8.1,0.26,8.1,1.7,0.19,8,-8.2,3.2,-3.3,7.6,-8.6,-3,2.9,-11,12),
                (2700,123.2,8.4,0.25,8.4,1.8,0.2,8.2,-8.4,3.4,-3.5,8.1,-9.1,-3.2,3.1,-11,12),
                (2800,128.4,8.8,0.24,8.8,1.9,0.2,8.4,-8.5,3.6,-3.7,8.7,-9.6,-3.4,3.2,-11,12),
                (2900,133.7,9.1,0.23,9.1,2,0.21,8.5,-8.7,3.8,-3.9,9.3,-10.1,-3.5,3.4,-11,12),
                (3000,139,9.5,0.22,9.4,2,0.21,8.7,-8.8,4.1,-4.1,9.9,-10.6,-3.7,3.6,-11,12),
                (3100,144.4,9.8,0.21,9.8,2.1,0.22,88,-9,4.3,-4.3,10.5,-11.1,-3.9,3.8,-11,12),
                (3200,149.8,10.2,0.2,10.1,2.2,0.22,9,-9.1,4.5,-4.5,11,-11.6,-4.1,3.9,-11,12),
                (3300,155.3,10.5,0.2,10.5,2.3,0.23,9.2,-9.3,4.8,-4.7,11.6,-12.1,-4.3,4.1,-11,12),
                (3400,160.8,10.9,0.19,10.8,2.4,0.23,9.3,-9.4,5,-4.9,12.2,-12.6,-4.5,4.3,-11,12),
                (3500,166.4,11.2,0.19,11.2,2.5,0.23,9.5,-9.6,5.3,-5,12.8,-13.1,-4.6,4.5,-11,12),
                (3600,172,11.6,0.18,11.6,2.6,0.24,9.6,-9.7,5.5,-5.2,13.4,-13.6,-4.8,4.7,-11,12),
                (3700,177.7,12,0.17,11.9,2.7,0.24,9.8,-9.9,5.7,-5.4,13.9,-14.1,-5,4.9,-11,12),
                (3800,183.4,12.3,0.17,12.3,2.8,0.25,9.9,-10,6,-5.6,14.5,-14.6,-5.3,5.1,-11,12),
                (3900,189.2,12.7,0.16,12.7,2.9,0.25,10.1,-10.1,6.2,-5.8,15.1,-15,-5.5,5.3,-11,12),
                (4000,195,13.1,0.16,13,3,0.25,10.2,-10.3,6.5,-6,15.6,-15.5,-5.7,5.6,-11,12),
                (4100,200.9,13.4,0.16,13.4,3.1,0.26,10.4,-10.4,6.7,-6.2,16.2,-16,-5.9,5.8,-11,12),
                (4200,206.8,13.8,0.15,13.8,3.2,0.26,10.5,-10.5,6.9,-6.4,16.8,-16.5,-6.1,6,-11,11),
                (4300,212.8,14.2,0.15,14.1,3.3,0.27,10.7,-10.7,7.2,-6.6,17.3,-16.9,-6.4,6.3,-11,11),
                (4400,218.9,14.6,0.14,14.5,3.4,0.27,10.8,-10.8,7.4,-6.8,17.9,-17.4,-6.6,6.5,-10,11),
                (4500,225,15,0.14,14.9,3.5,0.28,11,-11,7.7,-7,18.4,-17.8,-6.8,6.7,-10,11),
                (4600,231.2,15.4,0.14,15.3,3.6,0.28,11.2,-11.1,7.9,-7.2,18.9,-18.3,-7.1,7,-10,11),
                (4700,237.4,15.7,0.13,15.7,3.8,0.28,11.3,-11.2,8.1,-7.4,19.4,-18.7,-7.3,7.3,-10,11),
                (4800,243.7,16.1,0.13,16.1,3.9,0.29,11.5,-11.4,8.4,-7.6,20,-19.2,-7.6,7.5,-10,11),
                (4900,250.1,16.5,0.13,16.5,4,0.29,11.6,-11.5,8.6,-7.8,20.5,-19.6,-7.9,7.8,-10,11),
                (5000,256.6,16.9,0.12,16.9,4.1,0.3,11.8,-11.7,8.9,-8,21,-20,-8.1,8.1,-9,10),
                (5100,263.1,17.3,0.12,17.3,4.2,0.3,12,-11.8,9.1,-8.2,21.5,-20.4,-8.4,8.4,-9,10),
                (5200,269.7,17.8,0.12,17.7,4.4,0.3,12.1,-11.9,9.3,-8.3,21.9,-20.9,-8.7,8.6,-9,10),
                (5300,276.4,18.2,0.12,18.1,4.5,0.31,12.3,-12.1,9.6,-8.5,22.4,-21.3,-9,8.9,-9,10),
                (5400,283.2,18.6,0.11,18.5,4.6,0.31,12.5,-12.2,9.8,-8.7,22.9,-21.7,-9.2,9.2,-9,10),
                (5500,290.1,19,0.11,18.9,4.7,0.32,12.6,-12.4,10,-8.9,23.3,-22.1,-9.5,9.5,-8,10),
                (5600,297.1,19.4,0.11,19.3,4.9,0.32,12.8,-12.5,10.3,-9.1,23.8,-22.5,-9.8,9.9,-8,9),
                (5700,304.1,19.9,0.11,19.7,5,0.33,13,-12.7,10.5,-9.3,24.2,-22.8,-10.1,10.2,-8,9),
                (5800,311.3,20.3,0.1,20.2,5.2,0.33,13.1,-12.8,10.7,-9.5,24.7,-23.2,-10.5,10.5,-8,9),
                (5900,318.6,20.7,0.1,20.6,5.3,0.33,13.3,-13,11,-9.7,25.1,-23.6,-10.8,10.8,-7,9),
                (6000,326,21.2,0.1,21.1,5.4,0.34,13.5,-13.1,11.2,-9.8,25.5,-24,-11.1,11.2,-7,9),
                (6100,333.5,21.6,0.1,21.5,5.6,0.34,13.7,-13.3,11.4,-10,25.9,-24.3,-11.4,11.5,-7,8),
                (6200,341.1,22.1,0.1,22,5.8,0.35,13.8,-13.4,11.7,-10.2,26.3,-24.7,-11.8,11.9,-6,8),
                (6300,348.8,22.5,0.09,22.4,5.9,0.35,14,-13.6,11.9,-10.4,26.7,-25,-12.1,12.2,-6,8),
                (6400,356.7,23,0.09,22.9,6.1,0.36,14.2,-13.7,12.1,-10.6,27.1,-25.4,-12.4,12.6,-6,8),
                (6500,364.8,23.5,0.09,23.4,6.2,0.36,14.4,-13.9,12.3,-10.7,27.5,-25.7,-12.8,12.9,-6,7),
                (6600,372.9,24,0.09,23.8,6.4,0.37,14.6,-14.1,12.6,-10.9,27.8,-26,-13.2,13.3,-5,7),
                (6700,381.3,24.5,0.09,24.3,6.6,0.37,14.8,-14.2,128,-11.1,28.2,-26.3,-13.5,13.7,-5,7),
                (6800,389.8,25,0.09,24.8,6.8,0.38,14.9,-14.4,13,-11.3,28.5,-26.6,-13.9,14.1,-5,6),
                (6900,398.5,25.5,0.08,25.3,7,0.38,15.1,-14.5,13.2,-11.4,28.9,-27,-14.3,14.5,-4,6),
                (7000,407.4,26,0.08,25.8,7.2,0.39,15.3,-14.7,13.4,-11.6,29.2,-27.3,-14.6,14.9,-4,6),
                (7100,416.5,26.5,0.08,26.4,7.4,0.39,15.5,-14.9,13.6,-11.8,29.5,-27.5,-15,15.3,-3,6),
                (7200,425.9,27.1,0.08,26.9,7.6,0.4,15.7,-15.1,13.9,-12,29.8,-27.8,-15.4,15.7,-3,5),
                (7300,435.5,27.6,0.08,27.4,7.8,0.4,15.9,-15.2,14.1,-12.1,30.1,-28.1,-15.8,16.1,-3,5),
                (7400,445.3,28.2,0.08,28,8,0.41,16.1,-15.4,14.3,-12.3,30.4,-28.4,-16.2,16.6,-2,5),
                (7500,455.5,28.7,0.08,28.6,8.3,0.42,16.3,-15.6,14.5,-12.5,30.7,-28.6,-16.6,17,-2,4),
                (7600,465.9,29.3,0.07,29.1,8.5,0.42,16.6,-15.8,14.7,-12.6,30.9,-28.9,-17.1,17.5,-2,4),
                (7700,476.8,29.9,0.07,29.7,8.8,0.43,16.8,-15.9,14.9,-12.8,31.2,-29.2,-17.5,17.9,-1,4),
                (7800,488,30.6,0.07,30.4,9.1,0.44,17,-16.1,15.1,-13,31.4,-29.4,-17.9,18.4,-1,3),
                (7900,499.7,31.2,0.07,31,9.4,0.44,17.2,-16.3,15.3,-13.1,31.6,-29.6,-18.4,18.9,0,3),
                (8000,511.8,31.9,0.07,31.7,9.7,0.45,17.5,-16.5,15.5,-13.3,31.8,-29.9,-18.8,19.4,0,2),
                (8100,524.6,32.6,0.07,32.4,10,0.46,17.7,-16.7,15.7,-13.4,32,-30.1,-19.3,19.9,1,2),
                (8200,538,33.3,0.07,33.1,10.4,0.47,17.9,-16.9,15.9,-13.6,32.2,-30.3,-19.7,20.4,1,2),
                (8300,552.2,34.1,0.06,33.8,10.7,0.47,18.2,-17.1,15.9,-13.8,32.4,-30.5,-20.2,20.9,2,1),
                (8400,567.3,34.9,0.06,34.6,11.2,0.48,18.4,-17.3,15.9,-13.9,32.5,-30.7,-20.7,21.5,2,1),
                (8500,583.7,35.7,0.06,35.5,11.6,0.49,18.7,-17.5,15.9,-14.1,32.6,-30.9,-21.2,22,3,0),
                (8600,601.4,36.6,0.06,36.4,12.2,0.5,19,-17.7,15.9,-14.2,32.7,-31,-21.7,22.6,3,0),
                (8700,621.2,37.7,0.06,37.4,12.8,0.52,19.3,-18,15.9,-14.4,32.6,-31.2,-22.2,23.3,4,0),
                (8800,643.9,38.8,0.06,38.5,13.5,0.53,19.6,-18.2,15.9,-14.5,32.6,-31.3,-22.7,24,4,-1),
                (8900,671.1,40.2,0.05,39.9,14.4,0.55,19.6,-18.4,15.9,-14.6,32.6,-31.5,-23.3,24,5,-1),
                (9000,708.2,42,0.05,41.7,15.7,0.57,19.6,-18.7,15.9,-14.8,32.6,-31.6,-23.9,24,6,-2),]


            
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == gac_mesafe:
                    nisangah6bh = round(nsg[1], 1)
                    print("Nişangah:", nisangah6bh)
            for dgl in interpolasyonlu_veri1:
                if dgl[0] == gac_mesafe:
                    dogalyandz = round(dgl[5],3)
                    print("Doğal Yan Düzeltmesi:", dogalyandz)
            for tsg in interpolasyonlu_veri1:
                if tsg[0] == gac_mesafe:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)       

            metro_yan_duzeltmesi2 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi en yakın 1 milyem99999999999999:", metro_yan_duzeltmesi2)
            print("Arzin Dnüşü:,",arzindonusuyanduzeltmesi)
            print("Rüzgaryan dzl---------------------:,",ruzgar_yan_duzeltmesi)

            
            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi6bh = round(metro_yan_duzeltmesi2 + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi3333333333333:", toplam_yan_duzeltmesi6bh)
            gac_yan_duzeltmesi2 = round(toplam_yan_duzeltmesi6bh - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi2)
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == gac_mesafe:
                    dusus_acisi = round(dusus_acisi[1],3)
                    # print("Düşüş Açısı:", dusus_acisi)

            for tepe_yuksekligi in interpolasyonlu_veri1:
                if tepe_yuksekligi[0] == gac_mesafe:
                    tepe_yuksekligi = round(tepe_yuksekligi[2],3)
                    # print("Tepe Yüksekliği:", tepe_yuksekligi)

            for dtac_arti1 in interpolasyonlu_veri1:
                if dtac_arti1[0] == mesafe:
                    dtac_arti1 = round(dtac_arti1[3],3)
                    # print("+1 DTAÇ:", dtac_arti1)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4],3)
                    # print("-1 DTAÇ:", dtac_eksi1)
    ############################################################################################################
            """METRO DÜZELTMELERİNİ HESAPLAMAK GEREKİYOR"""
            # GTAC 1 inci oonüs için
            bt_rakimi = lne_1B_obus_rakim

            milyem_sabit_dz = 1.0186

            hedef_batarya_rakim_farki = (rakim - bt_rakimi)

            dtac = (hedef_batarya_rakim_farki / (plan_mesafesi_1B_obüs / 1000)) * milyem_sabit_dz

            # Burada dtac_arti1 ve dtac_eksi1 değerlerini sadece bir kez tanımlayın
            dtac_arti1 = 0
            dtac_eksi1 = 0

            # dtac_arti1 ve dtac_eksi1 değerlerini doğru şekilde ata
            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_arti1 = round(dusus_acisi[3], 3)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4], 3)

            ttac = 0
            print("Mesafe",mesafe)
            if dtac < 0:
                ttac = int(dtac * dtac_arti1)
                tac = dtac + ttac
                #print("-TAÇ", tac)

            if dtac > 0:
                ttac = int(dtac * dtac_eksi1)
                tac = dtac + ttac
                print("+TAÇ", tac)
    #########################################################################################################
            """YÜKSELİŞİ BUL"""
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,26),
(100,26),
(200,26),
(300,25),
(400,25),
(500,24),
(600,24),
(700,24),
(800,23),
(900,23),
(1000,23),
(1100,22),
(1200,22),
(1300,22),
(1400,22),
(1500,21),
(1600,21),
(1700,21),
(1800,21),
(1900,21),
(2000,20),
(2100,20),
(2200,20),
(2300,20),
(2400,20),
(2500,19),
(2600,19),
(2700,19),
(2800,19),
(2900,19),
(3000,19),
(3100,19),
(3200,18),
(3300,18),
(3400,18),
(3500,18),
(3600,18),
(3700,18),
(3800,17),
(3900,17),
(4000,17),
(4100,17),
(4200,17),
(4300,17),
(4400,16),
(4500,16),
(4600,16),
(4700,16),
(4800,16),
(4900,16),
(5000,15),
(5100,15),
(5200,15),
(5300,15),
(5400,15),
(5500,14),
(5600,14),
(5700,14),
(5800,14),
(5900,14),
(6000,13),
(6100,13),
(6200,13),
(6300,13),
(6400,13),
(6500,12),
(6600,12),
(6700,12),
(6800,12),
(6900,11),
(7000,11),
(7100,11),
(7200,11),
(7300,10),
(7400,10),
(7500,10),
(7600,9),
(7700,9),
(7800,9),
(7900,8),
(8000,8),
(8100,8),
(8200,7),
(8300,7),
(8400,6),
(8500,6),
(8600,5),
(8700,5),
(8800,4),
(8900,3),
(9000,3),

            ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            
            for nsg in interpolasyonlu_veri1:
                    if nsg[0] == gac_mesafe:
                        birmildegisim6bh = round(nsg[1], 1)
                        print("1 Milyemlik Değişim:", birmildegisim6bh)

            yukselis6bh = round(nisangah6bh+ + tac,1)
            self.ui.sonuc_yukselis_2B.setText(str(yukselis6bh))
    #########################################################################################################
            yan_2B_6bh = round(yan_2B+toplam_yan_duzeltmesi6bh)
            self.ui.sonuc_yan_2B.setText(str(yan_2B_6bh))
            self.ui.lne_barut_hakki.setText(str(secilen_barut_hakki))
            self.ui.sonuc_barut_hakki_2B.setText(str(secilen_barut_hakki))
            if 2200 < yan_2B_6bh < 3000:
                self.ui.sonuc_yan_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_yukselis_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_barut_hakki_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_tapa_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_istikamet_acisi_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_mesafe_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.lne_sonuc_mesafe.setText(str(mesafe))
                self.ui.lne_sonuc_barut_hakki.setText(str(secilen_barut_hakki))
                self.ui.lne_sonuc_yan.setText(str(yan_2B_6bh))
                self.ui.lne_sonuc_yukselis.setText(str(yukselis6bh))
                self.ui.lne_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
                self.ui.lne_sonuc_bir_ml_degisiklil.setText(str(birmildegisim6bh)) 
                self.ui.lne_sonuc_yuz_m.setText(str(mesafe/1000)) 
            else:
                self.ui.sonuc_yan_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_yukselis_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_barut_hakki_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_tapa_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_istikamet_acisi_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_mesafe_2B.setStyleSheet("color: grey; font-weight: normal")
            self._doldur_sonuc("2", toplam_yan_duzeltmesi6bh, yukselis6bh, birmildegisim6bh, locals())

##########################################################################################################
        if secilen_barut_hakki == 7:
            plan_mesafesi = plan_mesafesi_2B_obüs
            batarya_rakimi = lne_2B_obus_rakim
            bt_rakimi_1 = lne_2B_obus_rakim
            hedef_rakimi = lne_hedef_rakim
            ahia = lne_2B_obus_ahia
            paralanma_yuksekliği = paralanma_yuksekliği
            mermi_kare_agirligi = mermi_kare_agirligi
            ilk_hiz_farki = lne_2B_obus_ihf
            atis_istikameti = batarya_hedef_İA_2
            atis_istikameti_1 = batarya_hedef_İA_2
            mevzi_hiz_degisikligi = 0
            mevzi_yan_duzeltmesi = 0
            mevzi_ts_duzeltmesi = 0
            tapasaniyesi = 1
            barut_isisi_str = self.ui.lne_barut_isisi.text()
            barut_isisi = float(barut_isisi_str) if barut_isisi_str else 75
            # barut_isisi = self.ui.lne_barut_isisi.text() if self.ui.lne_barut_isisi.text() else 75
            print("Barutttt İSİSİ",barut_isisi)
                # 1. GİRİŞ MESAFESİNİ BUL / MANUEL HASSASİYETİ

            # mesafe = 1866
            # en_yakin_10_metre = round(mesafe / 10) * 10
            # print("1836'nın en yakın 10 metreye çevrilmiş hali:", en_yakin_10_metre)

            # BATARYA RAKIMINI EN YAKIN 10 M YE ÇEVİR.
            batarya_rakimi = round(batarya_rakimi/10)*10
            # PARALANMA NOKTASININ RAKIMINI BUL.
            paralanma_noktasi_rakimi = hedef_rakimi + paralanma_yuksekliği
            # HEDEF-BATARYA YÜKSEKLİK FARKINI BUL
            yukseklik_farki = paralanma_noktasi_rakimi-batarya_rakimi
            yukseklik_farki_1 = paralanma_noktasi_rakimi-bt_rakimi_1
            # yukseklik_farki = round(yukseklik_farki/100)*100
            ################# B cetvelini yüklüyoruz.################
            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4100,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4200,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4300,0,0,0,1,1,1,1,2,2,2,2,2,3,3,3),
(4400,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4500,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4600,0,0,1,1,1,1,1,2,2,2,2,2,3,3,3),
(4700,0,0,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4800,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(4900,0,1,1,1,1,1,2,2,2,2,2,2,3,3,3),
(5000,0,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(5100,1,1,1,1,1,1,2,2,2,2,2,3,3,3,3),
(5200,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5300,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5400,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5500,1,1,1,1,1,2,2,2,2,2,2,3,3,3,3),
(5600,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5700,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5800,1,1,1,1,2,2,2,2,2,2,3,3,3,3,3),
(5900,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(6000,1,1,1,2,2,2,2,2,2,2,3,3,3,3,3),
(6100,1,1,1,2,2,2,2,2,2,3,3,3,3,3,3),
(6200,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(6300,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(6400,1,1,2,2,2,2,2,2,2,3,3,3,3,3,3),
(6500,1,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6600,1,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6700,2,2,2,2,2,2,2,2,3,3,3,3,3,3,3),
(6800,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(6900,2,2,2,2,2,2,2,3,3,3,3,3,3,3,4),
(7000,2,2,2,2,2,2,3,3,3,3,3,3,3,3,4),
(7100,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(7200,2,2,2,2,2,2,3,3,3,3,3,3,3,4,4),
(7300,2,2,2,2,2,3,3,3,3,3,3,3,3,4,4),
(7400,2,2,2,2,2,3,3,3,3,3,3,3,4,4,4),
(7500,2,2,2,2,3,3,3,3,3,3,3,3,4,4,4),
(7600,2,2,2,2,3,3,3,3,3,3,3,3,4,4,4),
(7700,2,2,2,3,3,3,3,3,3,3,3,4,4,4,4),
(7800,2,2,2,3,3,3,3,3,3,3,3,4,4,4,4),
(7900,2,2,3,3,3,3,3,3,3,3,4,4,4,4,4),
(8000,2,2,3,3,3,3,3,3,3,3,4,4,4,4,4),
(8100,2,3,3,3,3,3,3,3,3,4,4,4,4,4,4),
(8200,2,3,3,3,3,3,3,3,3,4,4,4,4,4,4),
(8300,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4),
(8400,3,3,3,3,3,3,3,3,4,4,4,4,4,4,4),
(8500,3,3,3,3,3,3,3,4,4,4,4,4,4,4,5),
(8600,3,3,3,3,3,3,3,4,4,4,4,4,4,4,5),
(8700,3,3,3,3,3,3,4,4,4,4,4,4,4,5,5),
(8800,3,3,3,3,3,3,4,4,4,4,4,4,4,5,5),
(8900,3,3,3,3,3,4,4,4,4,4,4,4,5,5,5),
(9000,3,3,3,3,4,4,4,4,4,4,4,5,5,5,5),
(9100,3,3,3,3,4,4,4,4,4,4,4,5,5,5,5),
(9200,3,3,3,4,4,4,4,4,4,4,5,5,5,5,5),
(9300,3,3,4,4,4,4,4,4,4,5,5,5,5,5,5),
(9400,3,3,4,4,4,4,4,4,4,5,5,5,5,5,5),
(9500,3,4,4,4,4,4,4,4,5,5,5,5,5,5,5),
(9600,4,4,4,4,4,4,4,5,5,5,5,5,5,5,5),
(9700,4,4,4,4,4,4,4,5,5,5,5,5,5,5,5),
(9800,4,4,4,4,4,4,5,5,5,5,5,5,5,5,6),
(9900,4,4,4,4,4,5,5,5,5,5,5,5,5,6,6),
(10000,4,4,4,4,5,5,5,5,5,5,5,5,6,6,6),
(10100,4,4,4,5,5,5,5,5,5,5,5,6,6,6,6),
(10200,4,4,4,5,5,5,5,5,5,5,6,6,6,6,6),
(10300,4,4,5,5,5,5,5,5,5,6,6,6,6,6,6),
(10400,4,5,5,5,5,5,5,5,6,6,6,6,6,6,6),
(10500,5,5,5,5,5,5,5,6,6,6,6,6,6,6,6),
(10600,5,5,5,5,5,5,6,6,6,6,6,6,6,6,6),
(10700,5,5,5,5,5,6,6,6,6,6,6,6,6,6,6),
(10800,5,5,5,5,6,6,6,6,6,6,6,6,6,6,6),
(10900,5,5,5,6,6,6,6,6,6,6,6,6,6,6,6),
(11000,5,5,6,6,6,6,6,6,6,6,6,6,6,6,6),
    ]

            data2 = [
(4000,-23,-18,-13,-7,0,7,16,25,34,45,56,67,80,93,107),
(4100,-24,-19,-13,-7,0,8,16,25,35,45,56,68,81,94,108),
(4200,-24,-19,-13,-7,0,8,16,25,35,46,57,69,82,95,109),
(4300,-25,-20,-14,-7,0,8,17,26,36,47,58,70,83,96,111),
(4400,-25,-20,-14,-7,0,8,17,26,36,47,59,71,84,98,112),
(4500,-26,-21,-14,-8,0,8,17,27,37,48,60,72,85,99,114),
(4600,-27,-21,-15,-8,0,8,17,27,38,49,61,73,87,101,115),
(4700,-28,-22,-15,-8,0,9,18,28,38,50,62,74,88,102,117),
(4800,-28,-22,-16,-8,0,9,18,28,39,51,63,76,89,104,119),
(4900,-29,-23,-16,-8,0,9,19,29,40,52,64,77,91,105,121),
(5000,-30,-23,-16,-8,0,9,19,29,41,53,65,78,92,107,123),
(5100,-31,-24,-17,-9,0,9,19,30,41,54,66,80,94,109,125),
(5200,-32,-25,-17,-9,0,10,20,31,42,55,68,81,96,111,127),
(5300,-33,-25,-18,-9,0,10,20,31,43,56,69,83,98,113,130),
(5400,-34,-26,-18,-9,0,10,21,32,44,57,71,85,100,116,132),
(5500,-35,-27,-19,-10,0,10,21,33,45,58,72,87,102,118,135),
(5600,-36,-28,-19,-10,0,11,22,34,46,60,74,89,104,121,138),
(5700,-37,-28,-20,-10,0,11,22,34,47,61,75,91,106,123,141),
(5800,-38,-29,-20,-10,0,11,23,35,48,62,77,93,109,126,144),
(5900,-39,-30,-21,-11,0,11,23,36,50,64,79,95,111,129,147),
(6000,-40,-31,-21,-11,0,12,24,37,51,66,81,97,114,132,150),
(6100,-41,-32,-22,-11,0,12,25,38,52,67,83,99,117,135,154),
(6200,-42,-33,-23,-12,0,12,25,39,54,69,85,102,120,138,158),
(6300,-44,-34,-23,-12,0,13,26,40,55,71,87,104,123,142,162),
(6400,-45,-35,-24,-12,0,13,27,41,56,72,89,107,126,145,166),
(6500,-46,-36,-25,-13,0,13,27,42,58,74,92,110,129,149,170),
(6600,-48,-37,-25,-13,0,14,28,43,60,76,94,113,132,153,174),
(6700,-49,-38,-26,-13,0,14,29,45,61,78,97,116,136,157,178),
(6800,-51,-39,-27,-14,0,15,30,46,63,81,99,119,139,161,183),
(6900,-52,-40,-28,-14,0,15,31,47,65,83,102,122,143,165,188),
(7000,-54,-42,-28,-15,0,15,32,49,66,85,105,125,147,169,193),
(7100,-56,-43,-29,-15,0,16,32,50,68,88,108,129,151,174,198),
(7200,-57,-44,-30,-15,0,16,33,51,70,90,111,132,155,179,204),
(7300,-59,-46,-31,-16,0,17,34,53,72,93,114,136,159,184,209),
(7400,-61,-47,-32,-16,0,17,35,54,74,95,117,140,164,189,215),
(7500,-63,-48,-33,-17,0,18,36,56,76,98,120,144,169,194,221),
(7600,-65,-50,-34,-17,0,18,37,58,79,101,124,148,173,200,228),
(7700,-67,-51,-35,-18,0,19,39,59,81,104,127,152,178,206,234),
(7800,-69,-53,-36,-19,0,19,40,61,83,107,131,157,184,212,241),
(7900,-71,-55,-37,-19,0,20,41,63,86,110,135,161,189,218,248),
(8000,-73,-56,-38,-20,0,21,42,65,88,113,139,166,195,224,255),
(8100,-76,-58,-40,-20,0,21,43,67,91,117,143,171,200,231,263),
(8200,-78,-60,-41,-21,0,22,45,69,94,120,148,176,206,238,271),
(8300,-80,-62,-42,-22,0,23,46,71,97,124,152,182,213,245,279),
(8400,-83,-64,-43,-22,0,23,48,73,100,128,157,187,219,253,288),
(8500,-86,-66,-45,-23,0,24,49,75,103,131,162,193,226,261,297),
(8600,-88,-68,-46,-24,0,25,51,78,106,136,167,199,233,269,307),
(8700,-91,-70,-48,-24,0,25,52,80,109,140,172,206,241,278,317),
(8800,-94,-72,-49,-25,0,26,54,83,113,144,177,212,249,287,328),
(8900,-97,-74,-51,-26,0,27,55,85,116,149,183,219,257,297,339),
(9000,-100,-77,-52,-27,0,28,57,88,120,154,189,227,266,307,351),
(9100,-103,-79,-54,-28,0,29,59,91,124,159,196,234,275,318,363),
(9200,-106,-82,-56,-28,0,30,61,94,128,164,202,2343,285,330,377),
(9300,-110,-84,-57,-29,0,31,63,97,133,170,210,251,295,342,391),
(9400,-113,-87,-59,-30,0,32,65,100,137,176,217,261,306,355,407),
(9500,-117,-90,-61,-31,0,33,68,104,142,183,225,270,318,369,424),
(9600,-121,-93,-63,-32,0,34,70,108,147,189,234,281,331,385,443),
(9700,-125,-96,-66,-34,0,35,72,112,153,197,243,293,346,402,464),
(9800,-129,-99,-68,-35,0,37,75,116,159,205,253,305,361,422,487),
(9900,-134,-103,-70,-36,0,38,78,120,165,213,264,319,379,443,515),
(10000,-139,-107,-73,-37,0,39,81,125,172,223,277,335,399,469,549),
(10100,-144,-111,-76,-39,0,41,85,131,180,233,290,353,422,501,574),
(10200,-149,-115,-79,-40,0,43,88,137,189,245,306,375,452,544,675),
(10300,-155,-119,-82,-42,0,45,92,143,199,259,326,402,493,644,675),
(10400,-161,-124,-85,-44,0,47,97,151,210,276,351,442,493,644,675),
(10500,-167,-129,-89,-46,0,49,102,160,224,298,389,442,493,644,675),
(10600,-175,-135,-93,-48,0,52,108,171,244,336,389,442,493,644,675),
(10700,-183,-141,-97,-51,0,55,116,187,287,336,389,442,493,644,675),
(10800,-191,-148,-103,-54,0,59,129,187,287,336,389,442,493,644,675),
(10900,-201,-157,-109,-57,0,67,129,187,287,336,389,442,493,644,675),
(11000,-213,-167,-117,-63,0,67,129,187,287,336,389,442,493,644,675),



]


            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            interpolasyonlu_veri2 = veri_kumesi_interpolasyonu(data2)
            # print(interpolasyonlu_veri2)

            hedefin_toptan_yuksekligi = yukseklik_farki
            hedefin_toptan_yuksekligi = round(hedefin_toptan_yuksekligi/100)*100
            print("Hedefin Toptan Yüksekliği: ", hedefin_toptan_yuksekligi)
            mesafe = plan_mesafesi
            print("mesafe = ", mesafe)

            if hedefin_toptan_yuksekligi == -400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[1])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[2])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[3])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[4])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == -0:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[5])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 100:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[6])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 200:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[7])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 300:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[8])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 400:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[9])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 500:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[10])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 600:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[11])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 700:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 800:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[12])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 900:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[13])
                        print("Seçilen Bölge:", bolge)

            elif hedefin_toptan_yuksekligi == 1000:
                for veri_noktasi in interpolasyonlu_veri1:
                    if veri_noktasi[0] == mesafe:
                        bolge = int(veri_noktasi[14])
                        print("Seçilen Bölge:", bolge)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")


            # TXT dosyasının adı ve dosya yolu
            dosya_adı = METRO_DOSYA

            # İlk satırı atlayarak dosyanın geri kalanını yeni bir liste olarak alın
            metrorpr = []
            if not os.path.isfile(dosya_adı):
                QMessageBox.warning(self, "Metro", "METRAP.txt bulunamadı.")
                return None
            with open(dosya_adı, "r") as dosya:
                for satır in dosya:
                    metrorpr.append(satır.strip())

            # Satırları yan yana birleştirerek elde etmek için join() yöntemini kullanıyoruz
            birlesik_veri = "\n".join(metrorpr[1:])

            rapor_tipi = birlesik_veri[:5]
            oktant = birlesik_veri[5:6]
            metro_koor = birlesik_veri[6:9] + " " + birlesik_veri[9:12]
            tarih = birlesik_veri[13:15]
            saat = birlesik_veri[15:17]
            dakika = str(int(birlesik_veri[17:18]) * 6)
            if dakika == "0":
                dakika = str(int(birlesik_veri[17:18]) * 6) + "0"
            ara = ":"
            saat_dakika = saat + ara + str(dakika)

            gecerlilik = birlesik_veri[18:19]
            metro_istasyonu_rakimi = str(int(birlesik_veri[19:22]) * 10)
            yogunluk = int(birlesik_veri[22:25]) / 10

            print("Rapor Tipi:", rapor_tipi)
            print("Oktant:", oktant)
            print("Metro Koordinat:", metro_koor)
            print("Tarih:", tarih)
            print("Saat:", saat_dakika)
            print("Geçerlilik:", gecerlilik)
            print("Metro İstasyonu Rakımı:", metro_istasyonu_rakimi)
            print("Yoğunluk:", yogunluk)


            if bolge == 0:
                ruzgar_istikameti = int(birlesik_veri[28:30]) * 100
                ruzgar_hizi = int(birlesik_veri[30:32])
                if birlesik_veri[32] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[32:35]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[32:35]) / 10
                if birlesik_veri[35] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[35:38]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[35:38]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)  # 39-40-41 3 say

            elif bolge == 1:
                ruzgar_istikameti = int(birlesik_veri[41:43]) * 100
                ruzgar_hizi = int(birlesik_veri[43:45])
                if birlesik_veri[45] == "0":
                    hava_sicakligi = int("1" + birlesik_veri[45:48]) / 10
                else:
                    hava_sicakligi = int(birlesik_veri[45:48]) / 10
                if birlesik_veri[48] == "0":
                    hava_yogunlugu = int("1" + birlesik_veri[48:51]) / 10
                else:
                    hava_yogunlugu = int(birlesik_veri[48:51]) / 10
                print("ruzgar istikameti:", ruzgar_istikameti)
                print("ruzgar hızı:", ruzgar_hizi)
                print("hava sıcaklığı:", hava_sicakligi)
                print("hava yoğunluğu:", hava_yogunlugu)

            elif bolge == 2:
                ruzgar_istikameti = int(birlesik_veri[54:56])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[56:58])  # 3
                if birlesik_veri[58] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[58:61])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[58:61])/10  # 4
                if birlesik_veri[61] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[61:64])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[61:64])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 3:
                ruzgar_istikameti = int(birlesik_veri[67:69])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[69:71])  # 3
                if birlesik_veri[71] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[71:74])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[71:74])/10  # 4
                if birlesik_veri[74] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[74:77])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[74:77])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 4:
                ruzgar_istikameti = int(birlesik_veri[80:82])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[82:84])  # 3
                if birlesik_veri[84] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[84:87])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[84:87])/10  # 4
                if birlesik_veri[87] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[87:90])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[87:90])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 5:
                ruzgar_istikameti = int(birlesik_veri[93:95])*100  # 3
                ruzgar_hizi = int(birlesik_veri[95:97])  # 3
                if birlesik_veri[97] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[97:100])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[97:100])/10  # 4
                if birlesik_veri[100] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[100:103])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[100:103])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 6:
                ruzgar_istikameti = int(birlesik_veri[106:108])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[108:110])  # 3
                if birlesik_veri[110] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[110:113])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[110:113])/10  # 4
                if birlesik_veri[113] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[113:116])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[113:116])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 7:
                ruzgar_istikameti = int(birlesik_veri[119:121])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[121:123])  # 3
                if birlesik_veri[123] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[123:126])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[123:126])/10  # 4
                if birlesik_veri[126] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[126:129])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[126:129])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 8:
                ruzgar_istikameti = int(birlesik_veri[132:134])*100  # 3    54-64 arası
                ruzgar_hizi = int(birlesik_veri[134:136])  # 3
                if birlesik_veri[136] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[136:139])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[136:139])/10  # 4
                if birlesik_veri[139] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[139:142])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[139:142])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 9:
                ruzgar_istikameti = int(birlesik_veri[145:147])*100  # 3
                ruzgar_hizi = int(birlesik_veri[147:149])  # 3
                if birlesik_veri[149] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[149:152])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[149:152])/10  # 4
                if birlesik_veri[152] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[152:155])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[152:155])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)

            elif bolge == 10:
                ruzgar_istikameti = int(birlesik_veri[158:160])*100  # 3
                ruzgar_hizi = int(birlesik_veri[160:162])  # 3
                if birlesik_veri[162] == "0":
                    hava_sicakligi = int("1"+birlesik_veri[162:165])/10  # 4
                else:
                    hava_sicakligi = int(birlesik_veri[162:165])/10  # 4
                if birlesik_veri[165] == "0":
                    hava_yogunlugu = int("1"+birlesik_veri[165:168])/10  # 4
                else:
                    hava_yogunlugu = int(birlesik_veri[165:168])/10  # 4
                print("ruzgar istikameti", ruzgar_istikameti)
                print("ruzgar hızı", ruzgar_hizi)
                print("hava sıcaklığı", hava_sicakligi)
                print("hava yoğunluğu", hava_yogunlugu)


            # TAMAMLAYICI MESAFEYİ BULUYORUZ.
            if hedefin_toptan_yuksekligi == -400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[1])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[2])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[3])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[4])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == -0:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[5])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 100:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[6])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 200:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[7])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 300:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[8])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 400:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[9])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 500:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[10])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 600:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[11])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 700:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 800:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[12])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 900:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[13])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)

            elif hedefin_toptan_yuksekligi == 1000:
                for t_mesafe in interpolasyonlu_veri2:
                    if t_mesafe[0] == mesafe:
                        tamamlayici_mesafe = int(t_mesafe[14])
                        tamamlayici_mesafe = int(tamamlayici_mesafe)
                        print("Tamamlamlayıcı Mesafe:", tamamlayici_mesafe)
            else:
                print("Rakım Farkını en yakın 100 e yuvarlamayı unutma")

            bt_metro_ist_rakim_farki = batarya_rakimi - int(metro_istasyonu_rakimi)

            giris_mesafesi = round((int(mesafe) + int(tamamlayici_mesafe))/100)*100
            print("Giriş Mesafesi :", giris_mesafesi)

            ruzgarin_plan_istikameti_100 = ruzgar_istikameti - atis_istikameti

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            isi_dzl = [
                (-390,0.9),
            (-380,0.9),
            (-370,0.9),
            (-360,0.8),
            (-350,0.8),
            (-340,0.8),
            (-330,0.8),
            (-320,0.7),
            (-310,0.7),
            (-300,0.7),
            (-290,0.7),
            (-280,0.7),
            (-270,0.7),
            (-260,0.6),
            (-250,0.6),
            (-240,0.6),
            (-230,0.6),
            (-220,0.5),
            (-210,0.5),
            (-200,0.5),
            (-190,0.4),
            (-180,0.4),
            (-170,0.4),
            (-160,0.3),
            (-150,0.3),
            (-140,0.3),
            (-130,0.3),
            (-120,0.2),
            (-110,0.2),
            (-100,0.2),
            (-90,0.2),
            (-80,0.2),
            (-70,0.2),
            (-60,0.1),
            (-50,0.1),
            (-40,0.1),
            (-30,0.1),
            (-20,0),
            (-10,0),
            (0,0),
            (10,0),
            (20,0),
            (30,-0.1),
            (40,-0.1),
            (50,-0.1),
            (60,-0.1),
            (70,-0.2),
            (80,-0.2),
            (90,-0.2),
            (100,-0.2),
            (110,-0.2),
            (120,-0.2),
            (130,-0.3),
            (140,-0.3),
            (150,-0.3),
            (160,-0.3),
            (170,-0.4),
            (180,-0.4),
            (190,-0.4),
            (200,-0.5),
            (210,-0.5),
            (220,-0.5),
            (230,-0.6),
            (240,-0.6),
            (250,-0.6),
            (260,-0.6),
            (270,-0.7),
            (280,-0.7),
            (290,-0.7),
            (300,-0.7),
            (310,-0.7),
            (320,-0.7),
            (330,-0.8),
            (340,-0.8),
            (350,-0.8),
            (360,-0.8),
            (370,-0.9),
            (380,-0.9),
            (390,-0.9),
        ]


            isi_data5bh = []
            for i in range(len(isi_dzl)-1):
                start = isi_dzl[i]
                end = isi_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    isi_data5bh.append((int(x), y))

            isi_data5bh.append(isi_dzl[-1])


            yogunluk_dzl = [
                (-390,3.9),
            (-380,3.8),
            (-370,3.7),
            (-360,3.6),
            (-350,3.5),
            (-340,3.4),
            (-330,3.3),
            (-320,3.2),
            (-310,3.1),
            (-300,3),
            (-290,2.9),
            (-280,2.8),
            (-270,2.7),
            (-260,2.6),
            (-250,2.5),
            (-240,2.4),
            (-230,2.3),
            (-220,2.2),
            (-210,2.1),
            (-200,2),
            (-190,1.9),
            (-180,1.8),
            (-170,1.7),
            (-160,1.6),
            (-150,1.5),
            (-140,1.4),
            (-130,1.3),
            (-120,1.2),
            (-110,1.1),
            (-100,1),
            (-90,0.9),
            (-80,0.8),
            (-70,0.7),
            (-60,0.6),
            (-50,0.5),
            (-40,0.4),
            (-30,0.3),
            (-20,0.2),
            (-10,0.1),
            (0,0),
            (10,-0.1),
            (20,-0.2),
            (30,-0.3),
            (40,-0.4),
            (50,-0.5),
            (60,-0.6),
            (70,-0.7),
            (80,-0.8),
            (90,-0.9),
            (100,-1),
            (110,-1.1),
            (120,-1.2),
            (130,-1.3),
            (140,-1.4),
            (150,-1.5),
            (160,-1.6),
            (170,-1.7),
            (180,-1.8),
            (190,-1.9),
            (200,-2),
            (210,-2.1),
            (220,-2.2),
            (230,-2.3),
            (240,-2.4),
            (250,-2.5),
            (260,-2.6),
            (270,-2.7),
            (280,-2.8),
            (290,-2.9),
            (300,-3),
            (310,-3.1),
            (320,-3.2),
            (330,-3.3),
            (340,-3.4),
            (350,-3.5),
            (360,-3.6),
            (370,-3.7),
            (380,-3.8),
            (390,-3.9),
    ]

            yogunluk_data5bh = []

            for i in range(len(yogunluk_dzl)-1):
                start = yogunluk_dzl[i]
                end = yogunluk_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    yogunluk_data5bh.append((int(x), y))

            yogunluk_data5bh.append(yogunluk_dzl[-1])
            # print(yogunluk_data5bh)


            for i in isi_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    isi_duzeltmesi = i[1]
                    print("Isı Düzeltmesi :", isi_duzeltmesi)


            for i in yogunluk_data5bh:
                if i[0] == bt_metro_ist_rakim_farki:
                    yogunluk_duzeltmesi = i[1]
                    print("Yoğunluk Düzeltmesi", yogunluk_duzeltmesi)

            duzeltilmis_sıcaklik_degeri = isi_duzeltmesi + hava_sicakligi
            print("Düzeltilmiş Isı Değeri:", duzeltilmis_sıcaklik_degeri)

            duzeltilmis_yogunluk_degeri = yogunluk_duzeltmesi + hava_yogunlugu
            print("Düzeltilmiş Yoğunluk Değeri:", duzeltilmis_yogunluk_degeri)

            # RUZGAR BİLEŞENLERİ

            atis_istikameti = round(atis_istikameti/100)*100

            if ruzgar_istikameti < atis_istikameti:
                ruzgar_istikameti = ruzgar_istikameti + 6400
            else:
                ruzgar_istikameti = ruzgar_istikameti
            print("ruzgar istikameti", ruzgar_istikameti)

            print("atış istikameti", atis_istikameti)

            ruzgarin_plan_istikameti = ruzgar_istikameti - atis_istikameti

            print("Rüzgarın Plan İstikameti:", ruzgarin_plan_istikameti)

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            data1 = [
                (0,0,1),
(100,-0.1,0.99),
(200,-0.2,0.98),
(300,-0.29,0.96),
(400,-0.38,0.92),
(500,-0.47,0.88),
(600,-0.56,0.83),
(700,-0.63,0.77),
(800,-0.71,0.71),
(900,-0.77,0.63),
(1000,-0.83,0.56),
(1100,-0.88,0.47),
(1200,-0.92,0.38),
(1300,-0.96,0.29),
(1400,-0.98,0.2),
(1500,-0.99,0.1),
(1600,1,0),
(1700,-0.99,-0.1),
(1800,-0.98,-0.2),
(1900,-0.96,-0.29),
(2000,-0.92,-0.38),
(2100,-0.88,-0.47),
(2200,-0.83,-0.56),
(2300,-0.77,-0.63),
(2400,-0.71,-0.71),
(2500,-0.63,-0.77),
(2600,-0.56,-0.83),
(2700,-0.47,-0.88),
(2800,-0.38,-0.92),
(2900,-0.29,-0.96),
(3000,-0.2,-0.98),
(3100,-0.1,-0.99),
(3200,0,-1),
(3300,0.1,-0.99),
(3400,0.2,-0.98),
(3500,0.29,-0.96),
(3600,0.38,-0.92),
(3700,0.47,-0.88),
(3800,0.56,-0.83),
(3900,0.63,-0.77),
(4000,0.71,-0.71),
(4100,0.77,-0.63),
(4200,0.83,-0.56),
(4300,0.88,-0.47),
(4400,0.92,-0.38),
(4500,0.96,-0.29),
(4600,0.98,-0.2),
(4700,0.99,-0.1),
(4800,1,0),
(4900,0.99,0.1),
(5000,0.98,0.2),
(5100,0.96,0.29),
(5200,0.92,0.38),
(5300,0.88,0.47),
(5400,0.83,0.56),
(5500,0.77,0.63),
(5600,0.71,0.71),
(5700,0.63,0.77),
(5800,0.56,0.83),
(5900,0.47,0.88),
(6000,0.38,0.92),
(6100,0.29,0.96),
(6200,0.2,0.98),
(6300,0.1,0.99),
(6400,0,1),


    ]


            # veri_noktasi[0] = MESAFE
            # veri_noktasi[1] = YAN RÜZGARI
            # veri_noktasi[2] = MESAFE RÜZGARI

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Mesafe :", mesafe)

            for rzdlz in interpolasyonlu_veri1:
                if rzdlz[0] == ruzgarin_plan_istikameti:
                    yan_ruzgari_bileseni = rzdlz[1]
                    mesafe_ruzgari_bileseni = rzdlz[2]
                    print("Yan Rüzgarı Bileşeni:", yan_ruzgari_bileseni)
                    print("mesafe Rüzgarı Bileşeni:", mesafe_ruzgari_bileseni)


            # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,136.4,11.4,0.18,11.2,2.4,0.34,12,-11.9,3.1,-2.9,1.4,-3.4,-10.2,9.6,-12,13),
(4100,141.1,11.7,0.18,11.5,2.5,0.34,12.2,-12,3.3,-3.1,1.8,-3.8,-10.5,9.9,-12,13),
(4200,145.7,12.1,0.17,11.9,2.6,0.35,12.3,-12.2,3.5,-3.3,2.2,-4.2,-10.8,10.2,-12,13),
(4300,150.5,12.4,0.17,12.2,2.7,0.36,12.4,-12.3,3.7,-3.4,2.6,-4.6,-11,10.4,-12,13),
(4400,155.3,12.8,0.16,12.6,2.8,0.36,12.6,-12.4,3.9,-3.6,3,-5,-11.3,10.7,-12,13),
(4500,160.1,13.1,0.16,12.9,2.9,0.37,12.7,-12,4.2,-3.8,3.5,-5.4,-11.6,11,-11,13),
(4600,165,13.5,0.16,13.2,3,0.37,12.8,-12.7,4.4,-4,3.9,-5.9,-11.9,11.2,-11,13),
(4700,169.9,13.8,0.15,13.6,3.1,0.38,12.9,-12.8,4.6,-4.2,4.4,-6.3,-12.2,11.5,-11,12),
(4800,174.9,14.2,0.15,13.9,3.2,0.39,13,-12.9,4.8,-4.3,4.9,-6.7,-12.5,11.8,-11,12),
(4900,179.9,14.5,0.14,14.3,3.3,0.39,13.1,-13,5,-4.5,5.3,-7.2,-12.8,12.1,-11,12),
(5000,185,14.9,0.14,14.6,3.4,0.4,13.3,-13.1,5.3,-4.7,5.8,-7.6,-13.1,12.3,-10,12),
(5100,190.2,15.3,0.14,15,3.5,0.4,13.4,-13.3,5.5,-4.9,6.3,-8.1,-13.3,12.6,-10,12),
(5200,195.3,15.6,0.13,15.4,3.6,0.41,13.5,-13.4,5.7,-5.1,6.8,-8.5,-13.6,12.9,-10,11),
(5300,200.6,16,0.13,15.7,3.7,0.41,13.6,-13.5,5.9,-5.3,7.3,-9,-13.9,13.2,-10,11),
(5400,205.9,16.4,0.13,16.1,3.8,0.42,13.7,-13.6,6.2,-5.5,7.8,-9.5,-14.2,13.5,-9,11),
(5500,211.2,16.7,0.13,16.4,3.9,0.42,13.8,-13.7,6.4,-5.7,8.3,-9.9,-14.5,13.8,-9,11),
(5600,216.6,17.1,0.12,16.8,4,0.43,13.9,-13.8,6.6,-5.9,8.9,-10.4,-14.8,14.1,-9,10),
(5700,222,17.5,0.12,17.2,4.1,0.43,14,-13.9,6.9,-6.1,9.4,-10.8,-15.1,14.4,-8,10),
(5800,227.5,17.9,0.12,17.5,4.2,0.44,14.1,-14,7.1,-6.3,9.9,-11.3,-15.4,14.7,-8,10),
(5900,233,18.2,0.12,17.9,4.4,0.44,14.2,-14.1,7.4,-6.5,10.4,-11.7,-15.7,15,-8,10),
(6000,238.6,18,0.11,18.3,4.5,0.45,14.2,-14.2,7.6,-6.7,10.9,-12.2,-16,15.3,-8,9),
(6100,244.3,19,0.11,18.7,4.6,0.45,14.3,-14.3,7.9,-6.9,11.5,-12.7,-16.3,15.6,-7,9),
(6200,250,19.4,0.11,19.1,4.7,0.46,14.4,-14.3,8.1,-7.1,12,-13.1,-16.6,15.9,-7,9),
(6300,255.8,19.8,0.11,19.4,4.8,0.46,14.5,-14.4,8.4,-7.3,12.5,-13.6,-16.9,16.2,-7,8),
(6400,261.6,20.2,0.1,19.8,4.9,0.47,14.6,-14.5,8.6,-7.5,13,-14,-17.3,16.6,-6,8),
(6500,267.5,20.6,0.1,20.2,5.1,0.47,14.7,-14.6,8.9,-7.7,13.5,-14.5,-17.6,16.9,-6,8),
(6600,273.5,21,0.1,20.6,5.2,0.48,14.8,-14.7,9.1,-7.9,14,-14.9,-17.9,17.3,-5,7),
(6700,279.5,21.4,0.1,21,5.3,0.48,14.9,-14.8,9.4,-8.1,14.5,-15.3,-18.2,17.6,-5,7),
(6800,285.6,21.8,0.1,21.4,5.5,0.49,15,-14.9,9.6,-8.3,15.1,-15.8,-18.6,18,-5,7),
(6900,291.8,22.2,0.1,21.8,5.6,0.49,15.1,-15,9.9,-8.5,15.5,-16.2,-18.9,18.3,-4,6),
(7000,298,22.6,0.09,22.2,5.7,0.5,15.2,-15.1,10.1,-8.7,16,-16.6,-19.3,18.7,-4,6),
(7100,304.3,23,0.09,22.6,5.9,0.5,15.3,-15.1,10.4,-8.9,16.5,-17.1,-19.6,19.1,-3,5),
(7200,310.7,23.5,0.09,23,6,0.5,15.3,-15.2,10.6,-9.1,17,-17.5,-20,19.4,-3,5),
(7300,317.2,23.9,0.09,23.5,6.2,0.51,15.4,-15.3,10.9,-9.3,17.5,-17.9,-20.3,19.8,-2,5),
(7400,323.7,24.3,0.09,23.9,6.3,0.51,15.5,-15.4,11.1,-9.5,17.9,-18.3,-20.7,20.2,-2,4),
(7500,330.4,24.8,0.09,24.3,6.5,0.52,15.6,-15.5,11.4,-9.7,18.4,-18.7,-21.1,20.6,-1,4),
(7600,337.1,25.2,0.08,24.7,6.6,0.52,15.7,-15.6,11.7,-9.9,18.9,-19.1,-21.4,21,-1,3),
(7700,343.9,25.6,0.08,25.2,6.8,0.53,15.8,-15.7,11.9,-10.2,19.3,-19.5,-21.8,21.4,0,3),
(7800,350.8,26.1,0.08,25.6,6.9,0.53,15.9,-15.8,12.2,-10.4,19.8,-19.9,-22.2,21.9,0,2),
(7900,357.9,26.6,0.08,26.1,7.1,0.54,16,-15.8,12.4,-10.6,20.2,-20.3,-22.6,22.3,1,2),
(8000,365,27,0.08,26.5,7.3,0.54,16.1,-15.9,12.7,-10.8,20.6,-20.7,-23,22.7,1,1),
(8100,372.2,27.5,0.08,27,7.4,0.55,16.2,-16,13,-11,21,-21.1,-23.4,23.2,2,1),
(8200,379.6,28,0.08,27.4,7.6,0.55,16.3,-16.1,13.2,-11.2,21.5,-21.5,-23.8,23.6,2,0),
(8300,387.1,28.4,0.08,27.9,7.8,0.56,16.4,-16.2,13.5,-11.4,21.9,-21.8,-24.3,24.1,3,0),
(8400,394.7,28.9,0.07,28.4,8,0.56,16.5,-16.3,13.7,-11.6,22.3,-22.2,-24.7,24.5,3,-1),
(8500,402.4,29.4,0.07,28.9,8.2,0.57,16.6,-16.4,14,-11.8,22.7,-22.5,-25.1,25,4,-1),
(8600,410.3,29.9,0.07,29.4,8.4,0.57,16.7,-16.5,14.3,-12,23,-22.9,-25.6,25.5,5,-2),
(8700,418.4,30.4,0.07,29.9,8.6,0.58,16.8,-16.6,14.5,-12.2,23.4,-23.2,-26,26,5,-2),
(8800,426.6,30.9,0.07,30.4,8.8,0.59,16.9,-16.7,14.8,-12.4,23.8,-23.6,-26.5,26.5,6,-3),
(8900,435,31.5,0.07,30.9,9,0.59,17,-16.8,15.1,-12.6,24.1,-23.9,-26.9,27,6,-3),
(9000,443.6,32,0.07,31.4,9.2,0.6,17.1,-16.9,15.3,-12.8,24.5,-24.3,-27.4,27.5,7,-4),
(9100,452.4,32.6,0.07,31.9,9.5,0.6,17.2,-16.9,15.6,-13,24.8,-24.6,-27.9,28,8,-5),
(9200,461.4,33.1,0.07,32.5,9.7,0.61,17.3,-17,15.9,-13.3,25.2,-24.9,-28.4,28.6,8,-5),
(9300,470.6,33.7,0.06,33.1,10,0.61,17.4,-17.1,16.2,-13.5,25.5,-25.2,-28.9,29.1,9,-6),
(9400,480.1,34.3,0.06,33.6,10.2,0.62,17.5,-17.2,16.4,-13.7,25.8,-25.5,-29.4,29.7,10,-7),
(9500,489.9,34.9,0.06,34.2,10.5,0.63,17.6,-17.4,16.7,-13.9,26.1,-25.8,-29.9,30.3,11,-7),
(9600,500,35.5,0.06,34.8,10.8,0.63,17.8,-17.5,17,-14.1,26.4,-26.1,-30.4,30.9,11,-8),
(9700,510.5,36.1,0.06,35.4,11.1,0.64,17.9,-17.6,17.3,-14.3,26.7,-26.4,-30.9,31.5,12,-9),
(9800,521.4,36.8,0.06,36.1,11.4,0.65,18,-17.7,17,-14.5,26.9,-26.6,-31.5,32.1,13,-9),
(9900,532.7,37.5,0.06,36.7,11.7,0.65,18.1,-17.8,17.9,-14.7,27.2,-26.9,-32,32.7,14,-10),
(10000,544.5,38.2,0.06,37.4,12.1,0.66,18.3,-17.9,18.3,-14.9,27.4,-27.2,-32.6,33.4,14,-11),
(10100,556.8,38.9,0.06,38.2,12.5,0.67,18.4,-18,18.3,-15.1,27.6,-27.4,-33.1,34.1,15,-11),
(10200,569.9,39.7,0.06,38.9,12.9,0.68,18.5,-18.1,18.3,-15.3,27.9,-27.7,-33.7,34.8,16,-12),
(10300,583.8,40.5,0.05,39.7,13.3,0.69,18.7,-18.2,18.3,-15.5,28.1,-27.9,-34.3,35.5,17,-13),
(10400,598.6,41.3,0.05,40.5,13.8,0.7,18.8,-18.4,18.3,-15.7,28.2,-28.1,-34.9,36.3,18,-14),
(10500,614.7,42.3,0.05,41.4,14.4,0.71,18.9,-18.5,18.3,-15.9,28.4,-28.4,-35.5,37.2,19,-15),
(10600,632.5,43.3,0.05,42.4,15,0.72,19.1,-18.6,18.3,-16.1,28.5,-28.6,-36.2,38.3,20,-16),
(10700,652.5,44.4,0.05,43.5,15.7,0.74,19.3,-18.7,18.3,-16.3,28.5,-28.8,-36.8,38.3,21,-16),
(10800,676.2,45.7,0.05,44.8,16.6,0.75,19.5,-18.9,18.3,-16.5,28.5,-29,-37.5,38.3,22,-18),
(10900,706.6,47.4,0.05,46.4,17.8,0.77,19.5,-19,18.3,-16.7,28.5,-29.1,-38.2,38.3,24,-19),
(11000,764.2,50.4,0.04,49.4,20.3,0.77,19.5,-19.2,18.3,-16.9,28.5,-29.3,-38.9,38.3,24,-20),



    ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == giris_mesafesi:
                    nisangah5bh = round(nsg[1], 3)
                    print("Nişangah:", nisangah5bh)

            for tsg in interpolasyonlu_veri1:
                if tsg[0] == giris_mesafesi:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            for ts in interpolasyonlu_veri1:
                if ts[0] == giris_mesafesi:
                    deltaTS = round(ts[3], 3)
                    # print("Delta Tapa Saniyesi:", deltaTS)


            for ucs in interpolasyonlu_veri1:
                if ucs[0] == giris_mesafesi:
                    ucussuresi = round(ucs[4], 3)
                    # print("Uçuş Süresi:", ucussuresi)

            for dgl in interpolasyonlu_veri1:
                if dgl[0] == giris_mesafesi:
                    dogalyandz = round(dgl[5], 3)
                    # print("Doğal Yan Düzeltmesi:", dogalyandz)

            for yrd in interpolasyonlu_veri1:
                if yrd[0] == giris_mesafesi:
                    yan_ruzgari_duzeltme_faktoru = round(yrd[6], 1)
                    # print("Yan Rüzgarı Düzeltmesi:", yanruzgaridz)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_eksilme = round(ihz[7], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimeksilme)

            for ihz in interpolasyonlu_veri1:
                if ihz[0] == giris_mesafesi:
                    ilkhizdzl_martma = round(ihz[8], 3)
                    # print("İlk Hızda 1 m/sn değişim dzl:", ilkhizdegisimartma)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_bas = round(mrd[9], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (BAŞ):", mesaferuzgaridzl_bas)

            for mrd in interpolasyonlu_veri1:
                if mrd[0] == giris_mesafesi:
                    mesaferuzgaridzl_arka = round(mrd[10], 3)
                    # print("Mesafe Rüzgarı Düzeltmesi (ARKA):", mesaferuzgaridzl_arka)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_eksilme = round(hvs[11], 3)
                    # print("Hava Isısı Düzeltmesi (Eksilme):", havaisisi_eksilme)

            for hvs in interpolasyonlu_veri1:
                if hvs[0] == giris_mesafesi:
                    havaisisi_artma = round(hvs[12], 3)
                    # print("Hava Isısı Düzeltmesi (Artma):", havaisisi_artma)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_eksilme = round(hyd[13], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", havayogunlugu_eksilme)

            for hyd in interpolasyonlu_veri1:
                if hyd[0] == giris_mesafesi:
                    havayogunlugu_artma = round(hyd[14], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", havayogunlugu_artma)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_eksilme = round(mka[15], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Eksilme):", mermikareagirligi_eksilme)

            for mka in interpolasyonlu_veri1:
                if mka[0] == giris_mesafesi:
                    mermikareagirligi_artma = round(mka[16], 3)
                    # print("Hava Yoğunluğu Düzeltmesi (Artma):", mermikareagirligi_artma)

            print("Yan Rüzgarı Düzeltme Faktörü:", yan_ruzgari_duzeltme_faktoru)

            mesafe_ruzgari = ruzgar_hizi * mesafe_ruzgari_bileseni
            print("mesafe ruzgari", mesafe_ruzgari)

            ruzgar_yan_duzeltmesi = round(
                (ruzgar_hizi * yan_ruzgari_bileseni * yan_ruzgari_duzeltme_faktoru), 1)
            print("Rüzgar Yan Düzeltmesi", ruzgar_yan_duzeltmesi)

            dogalyandz = round(dogalyandz, 1)
            print("Doğal Yan Düzeltmesi:", dogalyandz)

            atis_istikameti = round(atis_istikameti/400)*400

            # dünyanın dönmesi düzeltmesi


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,0.5,0.5,0.5,0.5,0.5,0.5,0.5,0.5,0.5),
(5000,0.6,0.6,0.6,0.6,0.6,0.7,0.7,0.7,0.7),
(6000,0.7,0.7,0.8,0.8,0.8,0.8,0.8,0.9,0.9),
(7000,0.9,0.9,0.9,0.9,1,1,1,1.1,1.1),
(8000,1,1,1,1.1,1.2,1.2,1.3,1.3,1.3),
(9000,1.1,1.2,1.2,1.3,1.4,1.5,1.5,1.6,1.6),
(10000,1.3,1.3,1.4,1.5,1.6,1.8,1.9,2,2),
(11000,1.3,1.4,1.5,1.6,1.8,2,2.1,2.2,2.3),






    ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            if atis_istikameti == 00 or atis_istikameti == 6400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[1], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 6000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[2], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 5600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[3], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 5200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[4], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 4800:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[5], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2000 or atis_istikameti == 4400:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[6], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2400 or atis_istikameti == 4000:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[7], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 2800 or atis_istikameti == 3600:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[8], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)

            if atis_istikameti == 3200:
                for arz in interpolasyonlu_veri1:
                    if arz[0] == giris_mesafesi:
                        arzindonusuyanduzeltmesi = round(arz[9], 1)
                        # print("Arzın Dönüşü Yan Düzeltmesi:", arzindonusuyanduzeltmesi)


            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (4000,0,-5,-9,-14,-18,-21,-23,-24,-25,0,5,9,14,18,21,24,25),
(5000,0,-6,-11,-17,-21,-25,-28,-29,-30,0,6,11,17,21,25,29,30),
(6000,0,-7,-13,-19,-24,-29,-32,-34,-35,0,7,13,19,24,29,34,35),
(7000,0,-8,-15,-22,-27,-32,-36,-38,-39,0,8,15,22,27,32,38,39),
(8000,0,-8,-16,-24,-30,-35,-39,-42,-42,0,8,16,24,30,35,42,42),
(9000,0,-9,-17,-25,-32,-37,-41,-44,-45,0,9,17,25,32,37,44,45),
(10000,0,-9,-18,-25,-32,-38,-42,-45,-46,0,9,18,25,32,38,45,46),
(11000,0,-8,-16,-23,-29,-34,-38,-40,-41,0,8,16,23,29,34,40,41),
    ]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)
            if atis_istikameti == 0 or atis_istikameti == 3200 or atis_istikameti == 6400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[1], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 200 or atis_istikameti == 3000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[2], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 400 or atis_istikameti == 2800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[3], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 600 or atis_istikameti == 2600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[4], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 800 or atis_istikameti == 2400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[5], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1000 or atis_istikameti == 2200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[6], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1200 or atis_istikameti == 2000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[7], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1400 or atis_istikameti == 1800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[8], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 1600 or atis_istikameti == 1600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[9], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            if atis_istikameti == 3400 or atis_istikameti == 6200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[11], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3600 or atis_istikameti == 6000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[12], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 3800 or atis_istikameti == 5800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[13], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4000 or atis_istikameti == 5600:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[14], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4200 or atis_istikameti == 5400:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[15], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4400 or atis_istikameti == 5200:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[16], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4600 or atis_istikameti == 5000:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[17], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)

            if atis_istikameti == 4800 or atis_istikameti == 4800:
                for arzmes in interpolasyonlu_veri1:
                    if arzmes[0] == giris_mesafesi:
                        arzindonusumesafeduzeltmesi = round(arzmes[18], 1)
                        # print("Arzın Dönüşü Mesafe Düzeltmesi:", arzindonusumesafeduzeltmesi)


            print("Dünyanın Dönmesi Düzeltmesi:", arzindonusuyanduzeltmesi)

            metro_yan_duzeltmesi = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            metro_yan_duzeltmesi_1 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi:", metro_yan_duzeltmesi)
            metro_yan_duzeltmesi = round(metro_yan_duzeltmesi/1)*1
            print("Metro Yan Düzeltmesi en yakın 1 milyem:", metro_yan_duzeltmesi)


            # METRO MESAFE DÜZELTMESİ
            mesafe_ruzgari_fark = round(mesafe_ruzgari-0, 1)
            if mesafe_ruzgari_fark > 0:
                mesafe_ruzgari_durum = "BAŞ"
            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_durum = "ARKA"
            else:
                mesafe_ruzgari_durum = 0
            print("Mesafe Ruzgari Fark :", mesafe_ruzgari_fark)
            hava_sicakligi_fark = round(duzeltilmis_sıcaklik_degeri - 100, 1)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_durum = "EKSİLME"
            if hava_sicakligi_fark > 0:
                hava_sicakligi_durum = "ARTMA"
            else:
                hava_sicakligi_durum = 0
            print("Hava Sıcaklığı Fark :", hava_sicakligi_fark)
            hava_yogunlugu_fark = round(duzeltilmis_yogunluk_degeri - 100, 1)
            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_durum = "EKSİLME"
            if hava_yogunlugu_fark > 0:
                hava_yogunlugu_durum = "ARTMA"
            else:
                hava_yogunlugu_durum = 0
            print("Hava Yoğunluğu Fark :", hava_yogunlugu_fark)
            mermi_kare_agirligi_fark = round(mermi_kare_agirligi-2, 1)
            if mermi_kare_agirligi_fark < 0:
                merkar_durum = "EKSİLME"
            if mermi_kare_agirligi_fark > 0:
                merkar_durum = "ARTMA"
            if mermi_kare_agirligi_fark == 0:
                merkar_durum = 0

            print("Mermi Kare Ağırlığı Fark :", mermi_kare_agirligi_fark)

            if mesafe_ruzgari_fark < 0:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_arka
            else:
                mesafe_ruzgari_duzeltme_birimi = mesaferuzgaridzl_bas
            if mesafe_ruzgari_fark == 0:
                mesafe_ruzgari_duzeltme_birimi = 0

            print("Mesafe Rüzgarı Düzeltme Birimi:", mesafe_ruzgari_duzeltme_birimi)
            if hava_sicakligi_fark < 0:
                hava_sicakligi_duzeltme_birimi = havaisisi_eksilme
            else:
                hava_sicakligi_duzeltme_birimi = havaisisi_artma
            if hava_sicakligi_fark == 0:
                hava_sicakligi_duzeltme_birimi = 0
            print("Hava Isısı Düzeltme Birimi:", hava_sicakligi_duzeltme_birimi)

            if hava_yogunlugu_fark < 0:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_eksilme
            else:
                hava_yogunlugu_duzeltme_birimi = havayogunlugu_artma
            if hava_yogunlugu_fark == 0:
                hava_yogunlugu_duzeltme_birimi = 0
            print("Hava Yoğunluğu Düzeltme Birimi:", hava_yogunlugu_duzeltme_birimi)

            if mermi_kare_agirligi_fark < 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_eksilme
            elif mermi_kare_agirligi_fark > 0:
                mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_artma
            else:
                mermi_kare_agirligi_duzeltme_birimi = 0
            mesafe_ruzgari_fark = abs(mesafe_ruzgari_fark)
            hava_sicakligi_fark = abs(hava_sicakligi_fark)
            hava_yogunlugu_fark = abs(hava_yogunlugu_fark)


            print("Mermi Kare Ağırlığı Düzeltme Birimi:",
                mermi_kare_agirligi_duzeltme_birimi)

            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi = round(metro_yan_duzeltmesi + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)
            gac_yan_duzeltmesi = round(toplam_yan_duzeltmesi - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi)

            # Lineer interpolasyon fonksiyonu

            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]
            brt_isisi_ilk_hiz_dzl = [
                (-40,-14),
(-30,-12.9),
(-20,-11.8),
(-10,-10.7),
(0,-9.5),
(10,-8.2),
(20,-6.9),
(30,-5.6),
(40,-4.3),
(50,-2.9),
(60,-1.5),
(70,0),
(80,1.5),
(90,3.1),
(100,4.7),
(110,6.3),
(120,7.9),
(130,9.6),



    ]
            barut__isisi_dzl2 = 3
            barut_isisi_dzl = []
            for i in range(len(brt_isisi_ilk_hiz_dzl)-1):
                start = brt_isisi_ilk_hiz_dzl[i]
                end = brt_isisi_ilk_hiz_dzl[i+1]
                diff_x = end[0] - start[0]
                diff_y = end[1] - start[1]

                for j in range(diff_x):
                    x = start[0] + j
                    y = start[1] + (j * diff_y) / diff_x
                    barut_isisi_dzl.append((int(x), y))

            barut_isisi_dzl.append(brt_isisi_ilk_hiz_dzl[-1])
            print("Barut Isısı",barut_isisi)

            for i in barut_isisi_dzl:
                if i[0] == barut_isisi:
                    print("mesafe",i[0])
                    barut__isisi_dzl2 = round(i[1], 1)
                    print("İ1:",i[1])
                    print("Barurrrrrrrr",barut__isisi_dzl2)
            # HIZ DEĞİŞİKLİĞİNİN HESAPLANMASI
            hiz_degisikligi = round(ilk_hiz_farki + mevzi_hiz_degisikligi, 1)
            deltaV_hiz_farki = round(hiz_degisikligi + barut__isisi_dzl2, 1)
            print("Hız Değişikliği:", hiz_degisikligi)
            print("Barut Isısı Düzeltmesi2:", barut__isisi_dzl2)
            print("Delta V Hız Farkı:", deltaV_hiz_farki)

            if deltaV_hiz_farki > 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_martma
            elif deltaV_hiz_farki < 0:
                ilk_hiz_duzeltme_birimi = ilkhizdzl_eksilme
            else:
                ilk_hiz_duzeltme_birimi = 0
            print("İlk Hız Düzeltme Birimi:", ilk_hiz_duzeltme_birimi)

            deltaV_mesafe_duzeltmesi = abs(
                round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            if deltaV_hiz_farki > 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))*-1
            elif deltaV_hiz_farki < 0:
                deltaV_mesafe_duzeltmesi = abs(
                    round(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi, 1))
            else:
                deltaV_mesafe_duzeltmesi = 0
            print("Delta V Mesafe Düzeltmesi:", deltaV_mesafe_duzeltmesi)

            #### TOPLAM MESAF DÜZELTMESİ ###
            toplam_mesafe_duzeltmesi = round(metro_mesafe_duzeltmesi + deltaV_mesafe_duzeltmesi)
            print("Toplam Mesafe Düzeltmesi:", toplam_mesafe_duzeltmesi)
            print("Toplam Yan Düzeltmesi:", toplam_yan_duzeltmesi)

            # toplam yan düzeltmede 1 milyem fazla mesafe düzeltmesinde işaret + ve 335 çıktı
            # baiek de yan 4 sl mesafe -326 çıktı

            # Lineer interpolasyon fonksiyonu


            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,0,0,0,0,0,0,0,0,0,0),
(1,0,0,0,0,0,0,0,0,0,0),
(2,-0.005,0.005,0,0,0,0,0,0,0.011,-0.011),
(3,-0.007,0.007,0,0,0,0,0.001,-0.001,0.016,-0.016),
(4,-0.009,0.009,0,0,0,-0.001,0.002,-0.002,0.02,-0.02),
(5,-0.011,0.011,0,0,0.001,-0.001,0.002,-0.002,0.024,-0.024),
(6,-0.013,0.013,0,0,0.001,-0.002,0.004,-0.004,0.027,-0.027),
(7,-0.015,0.015,0,0,0.003,-0.003,0.005,-0.005,0.029,-0.03),
(8,-0.017,0.017,0,0,0.004,-0.004,0.007,-0.007,0.031,-0.032),
(9,-0.019,0.019,0,0,0.006,-0.004,0.009,-0.009,0.032,-0.033),
(10,-0.021,0.021,0,0,0.006,-0.004,0.011,-0.011,0.033,-0.035),
(11,-0.022,0.022,0,0.001,0.007,-0.004,0.013,-0.012,0.034,-0.036),
(12,-0.024,0.024,0,0.001,0.007,-0.003,0.015,-0.013,0.035,-0.037),
(13,-0.025,0.025,-0.001,0.001,0.006,-0.001,0.016,-0.015,0.036,-0.038),
(14,-0.026,0.026,-0.001,0.002,0.005,0,0.018,-0.016,0.036,-0.038),
(15,-0.027,0.027,-0.002,0.003,0.004,0.002,0.019,-0.017,0.037,-0.039),
(16,-0.028,0.028,-0.002,0.003,0.002,0.003,0.02,-0.018,0.037,-0.04),
(17,-0.029,0.029,-0.003,0.004,0.001,0.005,0.022,-0.019,0.038,-0.041),
(18,-0.03,0.03,-0.003,0.004,-0.001,0.007,0.023,-0.021,0.038,-0.041),
(19,-0.031,0.031,-0.004,0.005,-0.003,0.01,0.024,-0.022,0.039,-0.042),
(20,-0.032,0.032,-0.004,0.006,-0.005,0.012,0.025,-0.023,0.039,-0.042),
(21,-0.033,0.033,-0.005,0.006,-0.007,0.014,0.027,-0.024,0.04,-0.043),
(22,-0.034,0.034,-0.005,0.007,-0.009,0.016,0.028,-0.025,0.04,-0.043),
(23,-0.035,0.035,-0.006,0.008,-0.012,0.018,0.029,-0.026,0.04,-0.043),
(24,-0.036,0.036,-0.007,0.008,-0.014,0.021,0.03,-0.027,0.04,-0.044),
(25,-0.036,0.037,-0.007,0.009,-0.016,0.023,0.031,-0.028,0.04,-0.044),
(26,-0.037,0.037,-0.008,0.01,-0.019,0.025,0.033,-0.029,0.04,-0.044),
(27,-0.038,0.038,-0.008,0.01,-0.021,0.027,0.034,-0.03,0.04,-0.044),
(28,-0.039,0.039,-0.009,0.011,-0.023,0.029,0.035,-0.032,0.04,-0.044),
(29,-0.04,0.04,-0.01,0.012,-0.025,0.031,0.036,-0.033,0.04,-0.044),
(30,-0.041,0.041,-0.01,0.012,-0.027,0.033,0.038,-0.034,0.04,-0.044),
(31,-0.041,0.042,-0.011,0.013,-0.03,0.035,0.039,-0.036,0.04,-0.044),
(32,-0.042,0.043,-0.012,0.014,-0.032,0.037,0.041,-0.037,0.04,-0.044),
(33,-0.043,0.043,-0.012,0.014,-0.034,0.039,0.042,-0.038,0.039,-0.044),
(34,-0.044,0.044,-0.013,0.015,-0.036,0.041,0.044,-0.04,0.039,-0.044),
(35,-0.045,0.045,-0.013,0.015,-0.038,0.043,0.045,-0.041,0.039,-0.043),
(36,-0.046,0.046,-0.014,0.016,-0.04,0.045,0.047,-0.043,0.038,-0.043),
(37,-0.047,0.047,-0.014,0.016,-0.042,0.047,0.048,-0.044,0.038,-0.043),
(38,-0.048,0.048,-0.015,0.017,-0.043,0.048,0.05,-0.046,0.037,-0.043),
(39,-0.048,0.049,-0.015,0.017,-0.045,0.05,0.052,-0.047,0.037,-0.042),
(40,-0.049,0.05,-0.016,0.018,-0.047,0.052,0.053,-0.049,0.036,-0.042),
(41,-0.05,0.051,-0.016,0.018,-0.049,0.053,0.055,-0.05,0.036,-0.042),
(42,-0.051,0.051,-0.017,0.019,-0.05,0.055,0.057,-0.052,0.035,-0.041),
(43,-0.052,0.052,-0.017,0.019,-0.052,0.056,0.059,-0.054,0.035,-0.041),
(44,-0.053,0.053,-0.018,0.019,-0.053,0.058,0.061,-0.055,0.034,-0.041),
(45,-0.054,0.054,-0.018,0.02,-0.055,0.059,0.062,-0.057,0.034,-0.04),
(46,-0.055,0.055,-0.018,0.02,-0.056,0.061,0.064,-0.059,0.033,-0.04),
(47,-0.056,0.056,-0.019,0.02,-0.058,0.062,0.066,-0.061,0.033,-0.04),
(48,-0.057,0.057,-0.019,0.02,-0.059,0.063,0.068,-0.062,0.032,-0.039),
(49,-0.058,0.059,-0.019,0.021,-0.061,0.065,0.07,-0.064,0.032,-0.039),
(50,-0.059,0.06,-0.02,0.021,-0.062,0.066,0.072,-0.066,0.031,-0.039),
(51,-0.061,0.061,-0.02,0.021,-0.063,0.067,0.074,-0.068,0.031,-0.038),
(52,-0.062,0.062,-0.02,0.021,-0.064,0.068,0.076,-0.07,0.031,-0.038),
(53,-0.063,0.063,-0.02,0.021,-0.065,0.07,0.078,-0.071,0.03,-0.038),
(54,-0.064,0.064,-0.02,0.021,-0.067,0.071,0.08,-0.073,0.03,-0.038),
(55,-0.065,0.065,-0.02,0.021,-0.068,0.072,0.082,-0.075,0.029,-0.038),
(56,-0.066,0.066,-0.02,0.021,-0.069,0.073,0.084,-0.077,0.029,-0.038),
(57,-0.067,0.067,-0.021,0.021,-0.07,0.074,0.086,-0.079,0.029,-0.038),
(58,-0.069,0.069,-0.021,0.021,-0.071,0.075,0.088,-0.081,0.029,-0.038),
(59,-0.07,0.07,-0.02,0.021,-0.072,0.076,0.09,-0.082,0.029,-0.038),
(60,-0.071,0.071,-0.02,0.021,-0.073,0.077,0.092,-0.084,0.029,-0.038),
(61,-0.072,0.072,-0.02,0.02,-0.074,0.078,0.094,-0.086,0.029,-0.039),
(62,-0.074,0.074,-0.02,0.02,-0.075,0.079,0.096,-0.088,0.03,-0.039),
(63,-0.075,0.075,-0.02,0.02,-0.076,0.08,0.098,-0.09,0.03,-0.04),
(64,-0.076,0.076,-0.02,0.02,-0.076,0.081,0.1,-0.092,0.031,-0.04),
(65,-0.077,0.078,-0.019,0.019,-0.077,0.082,0.102,-0.093,0.032,-0.042),
(66,-0.079,0.079,-0.019,0.019,-0.078,0.083,0.104,-0.095,0.033,-0.043),
(67,-0.08,0.08,-0.019,0.019,-0.079,0.084,0.106,-0.097,0.035,-0.045),
(68,-0.082,0.082,-0.018,0.019,-0.079,0.085,0.108,-0.099,0.038,-0.048),
(69,-0.083,0.083,-0.018,0.022,-0.079,0.085,0.11,-0.101,0.044,-0.053),
(70,-0.085,0.086,-0.019,0.043,-0.078,0.084,0.114,-0.104,0.058,-0.065),
(71,-0.09,0.09,-0.023,0.043,-0.073,0.08,0.123,-0.11,0.103,-0.1),
(72,-0.093,0.09,-0.033,0.043,-0.073,0.072,0.123,-0.121,0.103,-0.162),
(73,-0.105,0.09,-0.047,0.043,-0.073,0.054,0.123,-0.142,0.103,-0.284),


    ]


            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)


            # print("Tapa Saniyesi :", tapasaniyesi)

            if deltaV_hiz_farki < 0:
                for ihz_eksilme in interpolasyonlu_veri1:
                    if ihz_eksilme[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_eksilme[1], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Eksilme):", ilkhizda1mildegisim_eksilme)
            if deltaV_hiz_farki > 0:
                for ihz_artma in interpolasyonlu_veri1:
                    if ihz_artma[0] == tapasaniyesi:
                        ilkhizda1mildegisim = round(ihz_artma[2], 3)
                        # print("İlk Hızda 1 m/sn lik Değişim Miktarı (Artma):", ilkhizda1mildegisim_artma)
            if mesafe_ruzgari_durum == "BAŞ":
                for musrus_bas in interpolasyonlu_veri1:
                    if musrus_bas[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(musrus_bas[3], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Baş):", mesaferuzgaridegisimi_bas)
            if mesafe_ruzgari_durum == "ARKA":
                for mesruz_arka in interpolasyonlu_veri1:
                    if mesruz_arka[0] == tapasaniyesi:
                        mesaferuzgaridegisimi = round(mesruz_arka[4], 3)
                        # print("Mesafe Rüzgarı Değişimi Miktarı (Arka):", mesaferuzgaridegisimi_arka)
            if hava_sicakligi_durum == "EKSİLME":
                for hadese_eksilme in interpolasyonlu_veri1:
                    if hadese_eksilme[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_eksilme[5], 3)
                        # print("Hava Isısı Değişim Miktarı (Eksilme):", havaisisidegisim_eksilme)
            if hava_sicakligi_durum == "ARTMA":
                for hadese_artma in interpolasyonlu_veri1:
                    if hadese_artma[0] == tapasaniyesi:
                        havaisisidegisim = round(hadese_artma[6], 3)
                        # print("Hava Isısı Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == "EKSİLME":
                for hayog_eksilme in interpolasyonlu_veri1:
                    if hayog_eksilme[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_eksilme[7], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Eksilme):", havayogunlugudegisim_eksilme)
            if hava_yogunlugu_durum == "ARTMA":
                for hayog_artma in interpolasyonlu_veri1:
                    if hayog_artma[0] == tapasaniyesi:
                        havayogunlugudegisim = round(hayog_artma[8], 3)
                        # print("Hava Yoğunluğu Değişim Miktarı (Artma):", havaisisidegisim_artma)
            if hava_yogunlugu_durum == 0:
                havayogunlugudegisim = 0
            if merkar_durum == "EKSİLME":
                for merkar_eksilme in interpolasyonlu_veri1:
                    if merkar_eksilme[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_eksilme[9], 3)
                        # print("Mermi Kare Ağırlığı Değişim Miktarı (Eksilme):", mermikareagirligidegisim_eksilme)
            if merkar_durum == "ARTMA":
                for merkar_artma in interpolasyonlu_veri1:
                    if merkar_artma[0] == tapasaniyesi:
                        mermikareagirligidegisim = round(merkar_artma[10], 3)
            else:
                mermikareagirligidegisim = 0
                # print("Mermi Kare Ağırlığı Değişim Miktarı (Artma):", mermikareagirligidegisim_artma)
##########################################################################################################
            gac_mesafe = plan_mesafesi_2B_obüs + toplam_mesafe_duzeltmesi
            print("toplam mesafe",nisangah5bh)
# zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
               (4000,136.4,11.4,0.18,11.2,2.4,0.34,12,-11.9,3.1,-2.9,1.4,-3.4,-10.2,9.6,-12,13),
(4100,141.1,11.7,0.18,11.5,2.5,0.34,12.2,-12,3.3,-3.1,1.8,-3.8,-10.5,9.9,-12,13),
(4200,145.7,12.1,0.17,11.9,2.6,0.35,12.3,-12.2,3.5,-3.3,2.2,-4.2,-10.8,10.2,-12,13),
(4300,150.5,12.4,0.17,12.2,2.7,0.36,12.4,-12.3,3.7,-3.4,2.6,-4.6,-11,10.4,-12,13),
(4400,155.3,12.8,0.16,12.6,2.8,0.36,12.6,-12.4,3.9,-3.6,3,-5,-11.3,10.7,-12,13),
(4500,160.1,13.1,0.16,12.9,2.9,0.37,12.7,-12,4.2,-3.8,3.5,-5.4,-11.6,11,-11,13),
(4600,165,13.5,0.16,13.2,3,0.37,12.8,-12.7,4.4,-4,3.9,-5.9,-11.9,11.2,-11,13),
(4700,169.9,13.8,0.15,13.6,3.1,0.38,12.9,-12.8,4.6,-4.2,4.4,-6.3,-12.2,11.5,-11,12),
(4800,174.9,14.2,0.15,13.9,3.2,0.39,13,-12.9,4.8,-4.3,4.9,-6.7,-12.5,11.8,-11,12),
(4900,179.9,14.5,0.14,14.3,3.3,0.39,13.1,-13,5,-4.5,5.3,-7.2,-12.8,12.1,-11,12),
(5000,185,14.9,0.14,14.6,3.4,0.4,13.3,-13.1,5.3,-4.7,5.8,-7.6,-13.1,12.3,-10,12),
(5100,190.2,15.3,0.14,15,3.5,0.4,13.4,-13.3,5.5,-4.9,6.3,-8.1,-13.3,12.6,-10,12),
(5200,195.3,15.6,0.13,15.4,3.6,0.41,13.5,-13.4,5.7,-5.1,6.8,-8.5,-13.6,12.9,-10,11),
(5300,200.6,16,0.13,15.7,3.7,0.41,13.6,-13.5,5.9,-5.3,7.3,-9,-13.9,13.2,-10,11),
(5400,205.9,16.4,0.13,16.1,3.8,0.42,13.7,-13.6,6.2,-5.5,7.8,-9.5,-14.2,13.5,-9,11),
(5500,211.2,16.7,0.13,16.4,3.9,0.42,13.8,-13.7,6.4,-5.7,8.3,-9.9,-14.5,13.8,-9,11),
(5600,216.6,17.1,0.12,16.8,4,0.43,13.9,-13.8,6.6,-5.9,8.9,-10.4,-14.8,14.1,-9,10),
(5700,222,17.5,0.12,17.2,4.1,0.43,14,-13.9,6.9,-6.1,9.4,-10.8,-15.1,14.4,-8,10),
(5800,227.5,17.9,0.12,17.5,4.2,0.44,14.1,-14,7.1,-6.3,9.9,-11.3,-15.4,14.7,-8,10),
(5900,233,18.2,0.12,17.9,4.4,0.44,14.2,-14.1,7.4,-6.5,10.4,-11.7,-15.7,15,-8,10),
(6000,238.6,18,0.11,18.3,4.5,0.45,14.2,-14.2,7.6,-6.7,10.9,-12.2,-16,15.3,-8,9),
(6100,244.3,19,0.11,18.7,4.6,0.45,14.3,-14.3,7.9,-6.9,11.5,-12.7,-16.3,15.6,-7,9),
(6200,250,19.4,0.11,19.1,4.7,0.46,14.4,-14.3,8.1,-7.1,12,-13.1,-16.6,15.9,-7,9),
(6300,255.8,19.8,0.11,19.4,4.8,0.46,14.5,-14.4,8.4,-7.3,12.5,-13.6,-16.9,16.2,-7,8),
(6400,261.6,20.2,0.1,19.8,4.9,0.47,14.6,-14.5,8.6,-7.5,13,-14,-17.3,16.6,-6,8),
(6500,267.5,20.6,0.1,20.2,5.1,0.47,14.7,-14.6,8.9,-7.7,13.5,-14.5,-17.6,16.9,-6,8),
(6600,273.5,21,0.1,20.6,5.2,0.48,14.8,-14.7,9.1,-7.9,14,-14.9,-17.9,17.3,-5,7),
(6700,279.5,21.4,0.1,21,5.3,0.48,14.9,-14.8,9.4,-8.1,14.5,-15.3,-18.2,17.6,-5,7),
(6800,285.6,21.8,0.1,21.4,5.5,0.49,15,-14.9,9.6,-8.3,15.1,-15.8,-18.6,18,-5,7),
(6900,291.8,22.2,0.1,21.8,5.6,0.49,15.1,-15,9.9,-8.5,15.5,-16.2,-18.9,18.3,-4,6),
(7000,298,22.6,0.09,22.2,5.7,0.5,15.2,-15.1,10.1,-8.7,16,-16.6,-19.3,18.7,-4,6),
(7100,304.3,23,0.09,22.6,5.9,0.5,15.3,-15.1,10.4,-8.9,16.5,-17.1,-19.6,19.1,-3,5),
(7200,310.7,23.5,0.09,23,6,0.5,15.3,-15.2,10.6,-9.1,17,-17.5,-20,19.4,-3,5),
(7300,317.2,23.9,0.09,23.5,6.2,0.51,15.4,-15.3,10.9,-9.3,17.5,-17.9,-20.3,19.8,-2,5),
(7400,323.7,24.3,0.09,23.9,6.3,0.51,15.5,-15.4,11.1,-9.5,17.9,-18.3,-20.7,20.2,-2,4),
(7500,330.4,24.8,0.09,24.3,6.5,0.52,15.6,-15.5,11.4,-9.7,18.4,-18.7,-21.1,20.6,-1,4),
(7600,337.1,25.2,0.08,24.7,6.6,0.52,15.7,-15.6,11.7,-9.9,18.9,-19.1,-21.4,21,-1,3),
(7700,343.9,25.6,0.08,25.2,6.8,0.53,15.8,-15.7,11.9,-10.2,19.3,-19.5,-21.8,21.4,0,3),
(7800,350.8,26.1,0.08,25.6,6.9,0.53,15.9,-15.8,12.2,-10.4,19.8,-19.9,-22.2,21.9,0,2),
(7900,357.9,26.6,0.08,26.1,7.1,0.54,16,-15.8,12.4,-10.6,20.2,-20.3,-22.6,22.3,1,2),
(8000,365,27,0.08,26.5,7.3,0.54,16.1,-15.9,12.7,-10.8,20.6,-20.7,-23,22.7,1,1),
(8100,372.2,27.5,0.08,27,7.4,0.55,16.2,-16,13,-11,21,-21.1,-23.4,23.2,2,1),
(8200,379.6,28,0.08,27.4,7.6,0.55,16.3,-16.1,13.2,-11.2,21.5,-21.5,-23.8,23.6,2,0),
(8300,387.1,28.4,0.08,27.9,7.8,0.56,16.4,-16.2,13.5,-11.4,21.9,-21.8,-24.3,24.1,3,0),
(8400,394.7,28.9,0.07,28.4,8,0.56,16.5,-16.3,13.7,-11.6,22.3,-22.2,-24.7,24.5,3,-1),
(8500,402.4,29.4,0.07,28.9,8.2,0.57,16.6,-16.4,14,-11.8,22.7,-22.5,-25.1,25,4,-1),
(8600,410.3,29.9,0.07,29.4,8.4,0.57,16.7,-16.5,14.3,-12,23,-22.9,-25.6,25.5,5,-2),
(8700,418.4,30.4,0.07,29.9,8.6,0.58,16.8,-16.6,14.5,-12.2,23.4,-23.2,-26,26,5,-2),
(8800,426.6,30.9,0.07,30.4,8.8,0.59,16.9,-16.7,14.8,-12.4,23.8,-23.6,-26.5,26.5,6,-3),
(8900,435,31.5,0.07,30.9,9,0.59,17,-16.8,15.1,-12.6,24.1,-23.9,-26.9,27,6,-3),
(9000,443.6,32,0.07,31.4,9.2,0.6,17.1,-16.9,15.3,-12.8,24.5,-24.3,-27.4,27.5,7,-4),
(9100,452.4,32.6,0.07,31.9,9.5,0.6,17.2,-16.9,15.6,-13,24.8,-24.6,-27.9,28,8,-5),
(9200,461.4,33.1,0.07,32.5,9.7,0.61,17.3,-17,15.9,-13.3,25.2,-24.9,-28.4,28.6,8,-5),
(9300,470.6,33.7,0.06,33.1,10,0.61,17.4,-17.1,16.2,-13.5,25.5,-25.2,-28.9,29.1,9,-6),
(9400,480.1,34.3,0.06,33.6,10.2,0.62,17.5,-17.2,16.4,-13.7,25.8,-25.5,-29.4,29.7,10,-7),
(9500,489.9,34.9,0.06,34.2,10.5,0.63,17.6,-17.4,16.7,-13.9,26.1,-25.8,-29.9,30.3,11,-7),
(9600,500,35.5,0.06,34.8,10.8,0.63,17.8,-17.5,17,-14.1,26.4,-26.1,-30.4,30.9,11,-8),
(9700,510.5,36.1,0.06,35.4,11.1,0.64,17.9,-17.6,17.3,-14.3,26.7,-26.4,-30.9,31.5,12,-9),
(9800,521.4,36.8,0.06,36.1,11.4,0.65,18,-17.7,17,-14.5,26.9,-26.6,-31.5,32.1,13,-9),
(9900,532.7,37.5,0.06,36.7,11.7,0.65,18.1,-17.8,17.9,-14.7,27.2,-26.9,-32,32.7,14,-10),
(10000,544.5,38.2,0.06,37.4,12.1,0.66,18.3,-17.9,18.3,-14.9,27.4,-27.2,-32.6,33.4,14,-11),
(10100,556.8,38.9,0.06,38.2,12.5,0.67,18.4,-18,18.3,-15.1,27.6,-27.4,-33.1,34.1,15,-11),
(10200,569.9,39.7,0.06,38.9,12.9,0.68,18.5,-18.1,18.3,-15.3,27.9,-27.7,-33.7,34.8,16,-12),
(10300,583.8,40.5,0.05,39.7,13.3,0.69,18.7,-18.2,18.3,-15.5,28.1,-27.9,-34.3,35.5,17,-13),
(10400,598.6,41.3,0.05,40.5,13.8,0.7,18.8,-18.4,18.3,-15.7,28.2,-28.1,-34.9,36.3,18,-14),
(10500,614.7,42.3,0.05,41.4,14.4,0.71,18.9,-18.5,18.3,-15.9,28.4,-28.4,-35.5,37.2,19,-15),
(10600,632.5,43.3,0.05,42.4,15,0.72,19.1,-18.6,18.3,-16.1,28.5,-28.6,-36.2,38.3,20,-16),
(10700,652.5,44.4,0.05,43.5,15.7,0.74,19.3,-18.7,18.3,-16.3,28.5,-28.8,-36.8,38.3,21,-16),
(10800,676.2,45.7,0.05,44.8,16.6,0.75,19.5,-18.9,18.3,-16.5,28.5,-29,-37.5,38.3,22,-18),
(10900,706.6,47.4,0.05,46.4,17.8,0.77,19.5,-19,18.3,-16.7,28.5,-29.1,-38.2,38.3,24,-19),
(11000,764.2,50.4,0.04,49.4,20.3,0.77,19.5,-19.2,18.3,-16.9,28.5,-29.3,-38.9,38.3,24,-20),



    ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == gac_mesafe:
                    nisangah5bh = round(nsg[1], 1)
                    print("Nişangah:", nisangah5bh)
            for dgl in interpolasyonlu_veri1:
                if dgl[0] == gac_mesafe:
                    dogalyandz = round(dgl[5],3)
                    print("Doğal Yan Düzeltmesi:", dogalyandz)
            for tsg in interpolasyonlu_veri1:
                if tsg[0] == gac_mesafe:
                    m564ts = round(tsg[2], 3)
                    # print("Sıfır PY'de M564 TS:", m564ts)

            metro_yan_duzeltmesi2 = round(
                arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi, 1)
            print("Metro Yan Düzeltmesi en yakın 1 milyem99999999999999:", metro_yan_duzeltmesi2)
            print("Arzin Dnüşü:,",arzindonusuyanduzeltmesi)
            print("Rüzgaryan dzl---------------------:,",ruzgar_yan_duzeltmesi)

            
            mes_ruz_dzl = mesafe_ruzgari_fark*mesafe_ruzgari_duzeltme_birimi
            print("MESAFE RUZGARI DÜZELTMESİ", mes_ruz_dzl)
            hav_sic_dzl = round(hava_sicakligi_fark*hava_sicakligi_duzeltme_birimi, 1)
            hav_yog_dzl = round(hava_yogunlugu_fark*hava_yogunlugu_duzeltme_birimi, 1)
            print("Hava Yoğunluğu Düzeltmesi:", hav_yog_dzl)
            mer_kare_dzl = round(mermi_kare_agirligi_fark *
                                mermi_kare_agirligi_duzeltme_birimi, 1)
            print("Arzın Dönüsü Mesafe:", round(arzindonusumesafeduzeltmesi))
            dun_don_mes_dzl = round(arzindonusumesafeduzeltmesi * 0.77, 1)

            print("Dünyanın Dönmesi Mesafe Düzeltmesi:", dun_don_mes_dzl)

            metro_mesafe_duzeltmesi = (round(
                (mes_ruz_dzl) + (hav_sic_dzl) + (hav_yog_dzl) + (mer_kare_dzl)+(dun_don_mes_dzl)))
            print("Metro Mesafe Düzeltmesi:", metro_mesafe_duzeltmesi)

            toplam_yan_duzeltmesi5bh = round(metro_yan_duzeltmesi2 + mevzi_hiz_degisikligi)
            print("Toplam Yan Düzeltmesi33333333333335bh:", toplam_yan_duzeltmesi5bh)
            gac_yan_duzeltmesi2 = round(toplam_yan_duzeltmesi5bh - dogalyandz)
            print("GAC Yan Düzeltmesi:", gac_yan_duzeltmesi2)
            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu
            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]



            # Veri kümesi
            data1 = [
                (4000,181,159,0.014,-0.012),
(4500,213,212,0.019,-0.015),
(5000,246,275,0.024,-0.02),
(5500,280,347,0.031,-0.026),
(6000,316,430,0.041,-0.034),
(6500,353,525,0.054,-0.045),
(7000,393,633,0.071,-0.059),
(7500,434,757,0.094,-0.078),
(8000,479,899,0.125,-0.103),
(8500,526,1063,0.17,-0.138),
(9000,577,1255,0.237,-0.188),
(9500,633,1486,0.346,-0.263),
(10000,697,1773,0.558,-0.388),
(10500,776,2165,1.286,-0.646),
(11000,930,3054,1.286,-2.11),


]

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == gac_mesafe:
                    dusus_acisi = round(dusus_acisi[1],3)
                    # print("Düşüş Açısı:", dusus_acisi)

            for tepe_yuksekligi in interpolasyonlu_veri1:
                if tepe_yuksekligi[0] == gac_mesafe:
                    tepe_yuksekligi = round(tepe_yuksekligi[2],3)
                    # print("Tepe Yüksekliği:", tepe_yuksekligi)

            for dtac_arti1 in interpolasyonlu_veri1:
                if dtac_arti1[0] == mesafe:
                    dtac_arti1 = round(dtac_arti1[3],3)
                    print("+1 DTAÇ:", dtac_arti1)

            for dtac_eksi1 in interpolasyonlu_veri1:
                if dtac_eksi1[0] == mesafe:
                    dtac_eksi1 = round(dtac_eksi1[4],3)
                    print("-1 DTAÇ:", dtac_eksi1)
    ############################################################################################################
            """METRO DÜZELTMELERİNİ HESAPLAMAK GEREKİYOR"""
            # GTAC 1 inci oonüs için
            bt_rakimi = lne_2B_obus_rakim

            milyem_sabit_dz = 1.0186

            hedef_batarya_rakim_farki = (rakim - bt_rakimi)

            dtac = (hedef_batarya_rakim_farki / (plan_mesafesi_2B_obüs / 1000)) * milyem_sabit_dz

            # Burada dtac_arti1 ve dtac_eksi1 değerlerini sadece bir kez tanımlayın
            dtac_arti1 = 0
            dtac_eksi1 = 0

            # dtac_arti1 ve dtac_eksi1 değerlerini doğru şekilde ata
            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_arti1 = round(dusus_acisi[3], 3)

            for dusus_acisi in interpolasyonlu_veri1:
                if dusus_acisi[0] == mesafe:
                    dtac_eksi1 = round(dusus_acisi[4], 3)

            ttac = 0
            print("Mesafe",mesafe)
            print("dtacccccccccc",dtac)
            print("dtacccccccccc++++",dtac_arti1)
            print("dtacccccccccc",dtac_eksi1)
            if dtac < 0:
                ttac = int(dtac * dtac_arti1)
                tac = dtac + ttac
                #print("-TAÇ", tac)

            if dtac > 0:

                ttac = int(dtac * dtac_eksi1)
                tac = dtac + ttac
                print("+TAÇ", tac)
    #########################################################################################################
            """YÜKSELİŞİ BUL"""
            # tac = 0
            yukselis5bh = round(nisangah5bh+ + tac,1)
            self.ui.sonuc_yukselis_2B.setText(str(yukselis5bh))
    #########################################################################################################
            yan_2B_5bh = round(yan_2B+toplam_yan_duzeltmesi5bh)
            self.ui.sonuc_yan_2B.setText(str(yan_2B_5bh))
            self.ui.lne_barut_hakki.setText(str(secilen_barut_hakki))
            self.ui.sonuc_barut_hakki_2B.setText(str(secilen_barut_hakki))

            if 2200 < yan_2B_5bh < 3000:
                self.ui.sonuc_yan_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_yukselis_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_barut_hakki_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_tapa_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_istikamet_acisi_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.sonuc_mesafe_2B.setStyleSheet("color: blue; font-weight: bold")
                self.ui.lne_sonuc_mesafe.setText(str(mesafe))
                self.ui.lne_sonuc_barut_hakki.setText(str(secilen_barut_hakki))
                self.ui.lne_sonuc_yan.setText(str(yan_2B_5bh))
                self.ui.lne_sonuc_yukselis.setText(str(yukselis5bh))
                self.ui.lne_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
                # self.ui.lne_sonuc_bir_ml_degisiklil.setText(str(birmildegisim5bh)) 
                self.ui.lne_sonuc_yuz_m.setText(str(mesafe/1000)) 
                """SONRA YAPACAĞIM"""
                # self.ui.lbl_sonuc_barut_hakki_2.setText(str(secilen_barut_hakki))
                # self.ui.lbl_sonuc_yan_2.setText(str(yan_2B_5bh))
                # self.ui.lbl_sonuc_yukselis_2.setText(str(yukselis5bh))
                # self.ui.lbl_sonuc_tapa_saniyesi.setText(str(round(m564ts,1)))
            else:
                self.ui.sonuc_yan_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_yukselis_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_barut_hakki_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_tapa_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_istikamet_acisi_2B.setStyleSheet("color: grey; font-weight: normal")
                self.ui.sonuc_mesafe_2B.setStyleSheet("color: grey; font-weight: normal")
            self._doldur_sonuc("2", toplam_yan_duzeltmesi5bh, yukselis5bh, locals().get("birmildegisim5bh", 1), locals())
        

       ##################################################################################################### 
        return tac,toplam_mesafe_duzeltmesi,atis_istikameti_1
    @guvenli_slot
    def dzl_hesapla(self):
        sonuc = self.ilk_hesaplama()
        if sonuc is None:
            return None
        tac,toplam_mesafe_duzeltmesi,atis_istikameti_1 = sonuc
        #print("TAÇ",round(tac))
        print("TOPLAM MESAFE DÜZELTMESİ5BH",toplam_mesafe_duzeltmesi)

        
                    
        liste5bh = [
        (0 , 18),
        (100 , 18),
        (200 , 18),
        (300 , 18),
        (400 , 18),
        (500 , 18),
        (600 , 18),
        (700 , 18),
        (800 , 17),
        (900 , 17),
        (1000 , 17),
        (1100 , 17),
        (1200 , 17),
        (1300 , 17),
        (1400 , 17),
        (1500 , 17),
        (1600 , 16),
        (1700 , 16),
        (1800 , 16),
        (1900 , 16),
        (2000 , 16),
        (2100 , 16),
        (2200 , 16),
        (2300 , 15),
        (2400 , 15),
        (2500 , 15),
        (2600 , 15),
        (2700 , 15),
        (2800 , 15),
        (2900 , 15),
        (3000 , 14),
        (3100 , 14),
        (3200 , 14),
        (3300 , 14),
        (3400 , 14),
        (3500 , 14),
        (3600 , 13),
        (3700 , 13),
        (3800 , 13),
        (3900 , 13),
        (4000 , 13),
        (4100 , 12),
        (4200 , 12),
        (4300 , 12),
        (4400 , 12),
        (4500 , 12),
        (4600 , 11),
        (4700 , 11),
        (4800 , 11),
        (4900 , 11),
        (5000 , 11),
        (5100 , 10),
        (5200 , 10),
        (5300 , 10),
        (5400 , 10),
        (5500 , 9),
        (5600 , 9),
        (5700 , 9),
        (5800 , 8),
        (5900 , 8),
        (6000 , 8),
        (6100 , 8),
        (6200 , 7),
        (6300 , 7),
        (6400 , 7),
        (6500 , 6),
        (6600 , 6),
        (6700 , 5),
        (6800 , 5),
        (6900 , 4),
        (7000 , 4),
        (7100 , 3)]

        data5bh = []
        for i in range(len(liste5bh)-1):
            start = liste5bh[i]
            end = liste5bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data5bh.append((int(x), y))

        data5bh.append(liste5bh[-1])

        liste6bh = [
            (0 , 26),
        (100 , 26),
        (200 , 26),
        (300 , 25),
        (400 , 25),
        (500 , 24),
        (600 , 24),
        (700 , 24),
        (800 , 23),
        (900 , 23),
        (1000 , 23),
        (1100 , 22),
        (1200 , 22),
        (1300 , 22),
        (1400 , 22),
        (1500 , 21),
        (1600 , 21),
        (1700 , 21),
        (1800 , 21),
        (1900 , 21),
        (2000 , 20),
        (2100 , 20),
        (2200 , 20),
        (2300 , 20),
        (2400 , 20),
        (2500 , 19),
        (2600 , 19),
        (2700 , 19),
        (2800 , 19),
        (2900 , 19),
        (3000 , 19),
        (3100 , 19),
        (3200 , 18),
        (3300 , 18),
        (3400 , 18),
        (3500 , 18),
        (3600 , 18),
        (3700 , 18),
        (3800 , 17),
        (3900 , 17),
        (4000 , 17),
        (4100 , 17),
        (4200 , 17),
        (4300 , 17),
        (4400 , 16),
        (4500 , 16),
        (4600 , 16),
        (4700 , 16),
        (4800 , 16),
        (4900 , 16),
        (5000 , 15),
        (5100 , 15),
        (5200 , 15),
        (5300 , 15),
        (5400 , 15),
        (5500 , 14),
        (5600 , 14),
        (5700 , 14),
        (5800 , 14),
        (5900 , 14),
        (6000 , 13),
        (6100 , 13),
        (6200 , 13),
        (6300 , 13),
        (6400 , 13),
        (6500 , 12),
        (6600 , 12),
        (6700 , 12),
        (6800 , 12),
        (6900 , 11),
        (7000 , 11),
        (7100 , 11),
        (7200 , 11),
        (7300 , 10),
        (7400 , 10),
        (7500 , 10),
        (7600 , 9),
        (7700 , 9),
        (7800 , 9),
        (7900 , 8),
        (8000 , 8),
        (8100 , 8),
        (8200 , 7),
        (8300 , 7),
        (8400 , 6),
        (8500 , 6),
        (8600 , 5),
        (8700 , 5),
        (8800 , 4),
        (8900 , 3)]

        data6bh = []
        for i in range(len(liste6bh)-1):
            start = liste6bh[i]
            end = liste6bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data6bh.append((int(x), y))

        data6bh.append(liste6bh[-1])
        
        liste7bh = [
            (0 , 42),
        (100 , 42),
        (200 , 42),
        (300 , 41),
        (400 , 41),
        (500 , 40),
        (600 , 39),
        (700 , 38),
        (800 , 38),
        (900 , 37),
        (1000 , 36),
        (1100 , 36),
        (1200 , 35),
        (1300 , 34),
        (1400 , 34),
        (1500 , 33),
        (1600 , 32),
        (1700 , 32),
        (1800 , 31),
        (1900 , 30),
        (2000 , 30),
        (2100 , 29),
        (2200 , 28),
        (2300 , 28),
        (2400 , 27),
        (2500 , 27),
        (2600 , 26),
        (2700 , 26),
        (2800 , 25),
        (2900 , 25),
        (3000 , 25),
        (3100 , 24),
        (3200 , 24),
        (3300 , 24),
        (3400 , 23),
        (3500 , 23),
        (3600 , 23),
        (3700 , 22),
        (3800 , 22),
        (3900 , 22),
        (4000 , 22),
        (4100 , 21),
        (4200 , 21),
        (4300 , 21),
        (4400 , 21),
        (4500 , 21),
        (4600 , 20),
        (4700 , 20),
        (4800 , 20),
        (4900 , 20),
        (5000 , 20),
        (5100 , 19),
        (5200 , 19),
        (5300 , 19),
        (5400 , 19),
        (5500 , 19),
        (5600 , 18),
        (5700 , 18),
        (5800 , 18),
        (5900 , 18),
        (6000 , 18),
        (6100 , 18),
        (6200 , 17),
        (6300 , 17),
        (6400 , 17),
        (6500 , 17),
        (6600 , 17),
        (6700 , 16),
        (6800 , 16),
        (6900 , 16),
        (7000 , 16),
        (7100 , 16),
        (7200 , 16),
        (7300 , 15),
        (7400 , 15),
        (7500 , 15),
        (7600 , 15),
        (7700 , 15),
        (7800 , 14),
        (7900 , 14),
        (8000 , 14),
        (8100 , 14),
        (8200 , 13),
        (8300 , 13),
        (8400 , 13),
        (8500 , 13),
        (8600 , 13),
        (8700 , 12),
        (8800 , 12),
        (8900 , 12),
        (9000 , 12),
        (9100 , 11),
        (9200 , 11),
        (9300 , 11),
        (9400 , 10),
        (9500 , 10),
        (9600 , 10),
        (9700 , 9),
        (9800 , 9),
        (9900 , 9),
        (10000 , 8),
        (10100 , 8),
        (10200 , 7),
        (10300 , 7),
        (10400 , 6),
        (10500 , 6),
        (10600 , 5),
        (10700 , 5),
        (10800 , 4),
        (10900 , 2),
        (11000 , 0)]


        data7bh = []
        for i in range(len(liste7bh)-1):
            start = liste7bh[i]
            end = liste7bh[i+1]
            diff_x = end[0] - start[0]
            diff_y = end[1] - start[1]

            for j in range(diff_x):
                x = start[0] + j
                y = start[1] + (j * diff_y) / diff_x
                data7bh.append((int(x), y))

        data7bh.append(liste7bh[-1])

        mesafe = float(self.ui.lne_sonuc_mesafe.text()) if self.ui.lne_sonuc_mesafe.text() else 5000.0
        
        yan = float(self.ui.lne_sonuc_yan.text()) if self.ui.lne_sonuc_yan.text() else 2600.0
        yukselis = float(self.ui.lne_sonuc_yukselis.text()) if self.ui.lne_sonuc_yukselis.text() else 300.0
        barut_hakki = float(self.ui.lne_sonuc_barut_hakki.text()) if self.ui.lne_sonuc_barut_hakki.text() else 5.0
        tapa_saniyesi = float(self.ui.lne_sonuc_tapa_saniyesi.text()) if self.ui.lne_sonuc_tapa_saniyesi.text() else 1.0
        bir_milyemlik_degisim = float(self.ui.lne_sonuc_bir_ml_degisiklil.text()) if self.ui.lne_sonuc_bir_ml_degisiklil.text() else 1.0
        mesafe_binler_adedi = float(self.ui.lne_sonuc_yuz_m.text()) if self.ui.lne_sonuc_yuz_m.text() else mesafe/1000
        
        yan_a = float(self.ui.lne_sonuc_yan.text()) if self.ui.lne_sonuc_yan.text() else 2600.0
        yukselis_a = float(self.ui.lne_sonuc_yukselis.text()) if self.ui.lne_sonuc_yukselis.text() else 300.0
        barut_hakki_a = float(self.ui.lne_sonuc_barut_hakki.text()) if self.ui.lne_sonuc_barut_hakki.text() else 5.0
        tapa_saniyesi_a = float(self.ui.lne_sonuc_tapa_saniyesi.text()) if self.ui.lne_sonuc_tapa_saniyesi.text() else 1.0

        yan_dzl_sl = float(self.ui.lne_dzl_sola_2.text()) if self.ui.lne_dzl_sola_2.text() else 0.0
        yan_dzl_sg = float(self.ui.lne_dzl_saga_2.text()) if self.ui.lne_dzl_saga_2.text() else 0.0
        yukselis_dzl_uz = float(self.ui.lne_dzl_uzat_2.text()) if self.ui.lne_dzl_uzat_2.text() else 0.0
        yukselis_dzl_ks = float(self.ui.lne_dzl_kisalt_2.text()) if self.ui.lne_dzl_kisalt_2.text() else 0.0
        paralanma_dzl_kl = float(self.ui.lne_dzl_kaldir_2.text()) if self.ui.lne_dzl_kaldir_2.text() else 0.0
        paralanma_dzl_in = float(self.ui.lne_dzl_indir_2.text()) if self.ui.lne_dzl_indir_2.text() else 0.0
        # zemin esasları
        if barut_hakki ==5:
                # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

                # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,18),
(100,18),
(200,18),
(300,18),
(400,18),
(500,18),
(600,18),
(700,18),
(800,18),
(900,17),
(1000,17),
(1100,17),
(1200,17),
(1300,17),
(1400,17),
(1500,17),
(1600,16),
(1700,16),
(1800,16),
(1900,16),
(2000,16),
(2100,16),
(2200,16),
(2300,15),
(2400,15),
(2500,15),
(2600,15),
(2700,15),
(2800,15),
(2900,15),
(3000,14),
(3100,14),
(3200,14),
(3300,14),
(3400,14),
(3500,14),
(3600,13),
(3700,13),
(3800,13),
(3900,13),
(4000,13),
(4100,12),
(4200,12),
(4300,12),
(4400,12),
(4500,12),
(4600,11),
(4700,11),
(4800,11),
(4900,11),
(5000,11),
(5100,10),
(5200,10),
(5300,10),
(5400,10),
(5500,9),
(5600,9),
(5700,9),
(5800,8),
(5900,8),
(6000,8),
(6100,8),
(6200,7),
(6300,7),
(6400,7),
(6500,6),
(6600,6),
(6700,5),
(6800,5),
(6900,4),
(7000,4),
(7100,3),
(7200,3),
     ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == mesafe:
                    nisangah_dz = round(nsg[1],1)
                    print("Nişangah:", nisangah_dz)
        if barut_hakki ==6:
                    
                    # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
                (0,26),
(100,26),
(200,26),
(300,25),
(400,25),
(500,24),
(600,24),
(700,24),
(800,23),
(900,23),
(1000,23),
(1100,22),
(1200,22),
(1300,22),
(1400,22),
(1500,21),
(1600,21),
(1700,21),
(1800,21),
(1900,21),
(2000,20),
(2100,20),
(2200,20),
(2300,20),
(2400,20),
(2500,19),
(2600,19),
(2700,19),
(2800,19),
(2900,19),
(3000,19),
(3100,19),
(3200,18),
(3300,18),
(3400,18),
(3500,18),
(3600,18),
(3700,18),
(3800,17),
(3900,17),
(4000,17),
(4100,17),
(4200,17),
(4300,17),
(4400,16),
(4500,16),
(4600,16),
(4700,16),
(4800,16),
(4900,16),
(5000,15),
(5100,15),
(5200,15),
(5300,15),
(5400,15),
(5500,14),
(5600,14),
(5700,14),
(5800,14),
(5900,14),
(6000,13),
(6100,13),
(6200,13),
(6300,13),
(6400,13),
(6500,12),
(6600,12),
(6700,12),
(6800,12),
(6900,11),
(7000,11),
(7100,11),
(7200,11),
(7300,10),
(7400,10),
(7500,10),
(7600,9),
(7700,9),
(7800,9),
(7900,8),
(8000,8),
(8100,8),
(8200,7),
(8300,7),
(8400,6),
(8500,6),
(8600,5),
(8700,5),
(8800,4),
(8900,3),
(9000,3),

            ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)
            # print(interpolasyonlu_veri1)

            # print("Mesafe :", mesafe)


            for nsg in interpolasyonlu_veri1:
                if nsg[0] == mesafe:
                    nisangah_dz = round(nsg[1],1)
                    print("Nişangah:", nisangah_dz)
        if barut_hakki == 7:
        
        # zemin esasları

            # Lineer interpolasyon fonksiyonu
            def lineer_interpolasyon(x, x1, y1, x2, y2):
                return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

            # Veri kümesi için interpolasyon yapma fonksiyonu


            def veri_kumesi_interpolasyonu(data):
                interpolasyonlu_veri = []  # İnterpolasyon sonuçlarını tutacak boş liste

                # Veri kümesini gezerek her bir nokta arasında interpolasyon yapar
                for i in range(len(data) - 1):
                    # İlk nokta (x1) verisi ve ona karşılık gelen y değerleri (y1_degerleri)
                    x1, y1_degerleri = data[i][0], data[i][1:]
                    # İkinci nokta (x2) verisi ve ona karşılık gelen y değerleri (y2_degerleri)
                    x2, y2_degerleri = data[i + 1][0], data[i + 1][1:]

                    # x1 ve x2 arasındaki tüm değerler için interpolasyon yapar
                    for j in range(x1 + 1, x2):
                        # Her bir y değeri için lineer interpolasyon yapar ve sonuçları listeye ekler
                        interpolasyonlu_degerler = [lineer_interpolasyon(
                            j, x1, y1, x2, y2) for y1, y2 in zip(y1_degerleri, y2_degerleri)]
                        interpolasyonlu_veri.append((j,) + tuple(interpolasyonlu_degerler))

                # Sonuç listesini oluştururken orijinal verileri, interpolasyon sonuçlarını ve son noktayı birleştirir
                return data + interpolasyonlu_veri + [data[-1]]


            # Veri kümesi
            data1 = [
    (0,42),
(100,42),
(200,42),
(300,41),
(400,41),
(500,40),
(600,39),
(700,38),
(800,38),
(900,37),
(1000,36),
(1100,36),
(1200,35),
(1300,34),
(1400,34),
(1500,33),
(1600,32),
(1700,32),
(1800,31),
(1900,30),
(2000,30),
(2100,29),
(2200,28),
(2300,28),
(2400,27),
(2500,27),
(2600,26),
(2700,26),
(2800,25),
(2900,25),
(3000,25),
(3100,24),
(3200,24),
(3300,24),
(3400,23),
(3500,23),
(3600,23),
(3700,22),
(3800,22),
(3900,22),
(4000,22),
(4100,21),
(4200,21),
(4300,21),
(4400,21),
(4500,21),
(4600,20),
(4700,20),
(4800,20),
(4900,20),
(5000,20),
(5100,19),
(5200,19),
(5300,19),
(5400,19),
(5500,19),
(5600,18),
(5700,18),
(5800,18),
(5900,18),
(6000,18),
(6100,18),
(6200,17),
(6300,17),
(6400,17),
(6500,17),
(6600,17),
(6700,16),
(6800,16),
(6900,16),
(7000,16),
(7100,16),
(7200,16),
(7300,15),
(7400,15),
(7500,15),
(7600,15),
(7700,15),
(7800,14),
(7900,14),
(8000,14),
(8100,14),
(8200,13),
(8300,13),
(8400,13),
(8500,13),
(8600,13),
(8700,12),
(8800,12),
(8900,12),
(9000,12),
(9100,11),
(9200,11),
(9300,11),
(9400,10),
(9500,10),
(9600,10),
(9700,9),
(9800,9),
(9900,9),
(10000,8),
(10100,8),
(10200,7),
(10300,7),
(10400,6),
(10500,6),
(10600,5),
(10700,5),
(10800,4),
(10900,2),
(11000,2),

    ]
            # F Cetveli
            #

            # Veri kümesi için interpolasyonu yap
            interpolasyonlu_veri1 = veri_kumesi_interpolasyonu(data1)

            
            for nsg in interpolasyonlu_veri1:
                if nsg[0] == mesafe:
                    nisangah_dz = round(nsg[1], 1)
                    print("Nişangah:", nisangah_dz)
        
        if not mesafe:
            QMessageBox.warning(self, "Düzeltme", "Mesafe yok. Önce atış esaslarını hesaplayın.")
            return None
        if "nisangah_dz" not in locals() or not nisangah_dz:
            nisangah_dz = 1.0
        yan_topla = (math.atan(yan_dzl_sl / mesafe)* (180 / math.pi))*6400/360
        print("Yan Topla:",yan_topla)
        yan_cikar = (math.atan(yan_dzl_sg / mesafe)* (180 / math.pi))*6400/360
        print("Yan çıkar:",yan_cikar)
        yukselis_topla = yukselis_dzl_uz / nisangah_dz
        yukselis_cikar = yukselis_dzl_ks / nisangah_dz
        paralan_topla = (2 / tapa_saniyesi) * (paralanma_dzl_in / 10)
        paralan_cikar = (2 / tapa_saniyesi) * (paralanma_dzl_kl / 10)
        yan_sonuc = 0.0
        yukselis_sonuc = 0.0
        tapasaniye_sonuc = 0.0
        if not self.isleme_alindi:
            self.ui.lbl_sonuc_yan_2.setText(str(yan))
            self.ui.lbl_sonuc_yukselis_2.setText(str(yukselis))
            self.ui.lbl_sonuc_tapa_saniyesi.setText(str(tapa_saniyesi))
            self.ui.lbl_sonuc_barut_hakki_2.setText(str(barut_hakki))
            self.isleme_alindi = True
        else:
            yan_sonucx = float(self.ui.lbl_sonuc_yan_2.text())
            yukselis_sonucx = float(self.ui.lbl_sonuc_yukselis_2.text())
            tapasaniye_sonucx = float(self.ui.lbl_sonuc_tapa_saniyesi.text())
            
            yan_sonuc = yan_sonucx + yan_topla - yan_cikar
            print("Yan Sonucc",yan_sonuc)
            yukselis_sonuc = yukselis_sonucx + yukselis_topla - yukselis_cikar
            tapasaniye_sonuc = tapasaniye_sonucx + paralan_topla - paralan_cikar

            self.ui.lbl_sonuc_yan_2.setText(str(round((yan_sonuc))))
            self.ui.lbl_sonuc_yukselis_2.setText(str(round(yukselis_sonuc,1)))
            self.ui.lbl_sonuc_tapa_saniyesi.setText(str(round(tapasaniye_sonuc,1)))
            self.ui.lbl_sonuc_barut_hakki_2.setText(str(barut_hakki))

        return mesafe, nisangah_dz, yan_sonuc, yukselis_sonuc, tapasaniye_sonuc, barut_hakki,tapa_saniyesi

    @guvenli_slot
    def atildi(self):
        try:
            yan_sonuc = float(self.ui.lbl_sonuc_yan_2.text())
            yukselis_sonuc = float(self.ui.lbl_sonuc_yukselis_2.text())
            tapasaniye_sonuc = float(self.ui.lbl_sonuc_tapa_saniyesi.text())
            barut_hakki_sonuc = float(self.ui.lbl_sonuc_barut_hakki_2.text())
        except ValueError:
            QMessageBox.warning(self, "Atıldı", "Önce düzeltme hesaplayın.")
            return
        if self.satir_sayisi >= 20:
            QMessageBox.warning(self, "Atıldı", "Atış tablosu doldu (20 satır).")
            return

        yan_dzl_sl = float(self.ui.lne_dzl_sola_2.text()) if self.ui.lne_dzl_sola_2.text() else 0.0
        yan_dzl_sg = float(self.ui.lne_dzl_saga_2.text()) if self.ui.lne_dzl_saga_2.text() else 0.0
        yukselis_dzl_uz = float(self.ui.lne_dzl_uzat_2.text()) if self.ui.lne_dzl_uzat_2.text() else 0.0
        yukselis_dzl_ks = float(self.ui.lne_dzl_kisalt_2.text()) if self.ui.lne_dzl_kisalt_2.text() else 0.0
        paralanma_dzl_kl = float(self.ui.lne_dzl_kaldir_2.text()) if self.ui.lne_dzl_kaldir_2.text() else 0.0
        paralanma_dzl_in = float(self.ui.lne_dzl_indir_2.text()) if self.ui.lne_dzl_indir_2.text() else 0.0

        self.ui.lne_dzl_sola_2.setText("")
        self.ui.lne_dzl_saga_2.setText("")
        self.ui.lne_dzl_uzat_2.setText("")
        self.ui.lne_dzl_kisalt_2.setText("")
        self.ui.lne_dzl_kaldir_2.setText("")
        self.ui.lne_dzl_indir_2.setText("")

        datetime_str = datetime.now().strftime("%d-%m-%Y\n%H:%M:%S")
        datetime_item = QtWidgets.QTableWidgetItem(datetime_str)
        datetime_item.setTextAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignLeft)

        self.ui.tablo_atislar.setItem(self.satir_sayisi, 0, datetime_item)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 1, QtWidgets.QTableWidgetItem(str(round(yan_sonuc))) if yan_sonuc != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 2, QtWidgets.QTableWidgetItem(str(yukselis_sonuc)) if yukselis_sonuc != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 3, QtWidgets.QTableWidgetItem(str(tapasaniye_sonuc)) if tapasaniye_sonuc != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 4, QtWidgets.QTableWidgetItem(str(round(yan_dzl_sg))) if yan_dzl_sg != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 5, QtWidgets.QTableWidgetItem(str(round(yan_dzl_sl))) if yan_dzl_sl != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 6, QtWidgets.QTableWidgetItem(str(round(yukselis_dzl_uz))) if yukselis_dzl_uz != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 7, QtWidgets.QTableWidgetItem(str(round(-1*yukselis_dzl_ks))) if -1*yukselis_dzl_ks != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 8, QtWidgets.QTableWidgetItem(str(round(paralanma_dzl_kl))) if paralanma_dzl_kl != 0.0 else None)
        self.ui.tablo_atislar.setItem(self.satir_sayisi, 9, QtWidgets.QTableWidgetItem(str(round(paralanma_dzl_in))) if paralanma_dzl_in != 0.0 else None)

        self.satir_sayisi += 1
    @guvenli_slot
    def sil(self):
        self.isleme_alindi = False
        self.ui.lne_hedef_sag_deger.clear()
        self.ui.lne_hedef_yukari_deger.clear()
        self.ui.lne_hedef_rakim.clear()
        self.ui.lne_barut_hakki.clear()
        for kod in ("1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D"):
            for prefix in ("sonuc_yan", "sonuc_yukselis", "sonuc_tapa", "sonuc_mesafe", "sonuc_istikamet_acisi"):
                getattr(self.ui, "%s_%s" % (prefix, kod)).clear()
            bh = "sonuc_batur_hakki_1A" if kod == "1A" else "sonuc_barut_hakki_%s" % kod
            getattr(self.ui, bh).clear()
        for name in (
            "lbl_sonuc_yan_2", "lbl_sonuc_yukselis_2", "lbl_sonuc_barut_hakki_2", "lbl_sonuc_tapa_saniyesi",
            "lne_sonuc_mesafe", "lne_sonuc_barut_hakki", "lne_sonuc_yan", "lne_sonuc_yukselis",
            "lne_sonuc_tapa_saniyesi", "lne_sonuc_bir_ml_degisiklil", "lne_sonuc_yuz_m",
            "lne_hedef_sag_deger_sonuc", "lne_hedef_yukari_deger_sonuc", "lne_hedef_rakim_sonuc",
            "lne_dzl_sola_2", "lne_dzl_saga_2", "lne_dzl_uzat_2", "lne_dzl_kisalt_2",
            "lne_dzl_kaldir_2", "lne_dzl_indir_2",
            "lbl_sonuc_plan_yani", "lbl_sonuc_yukselis_6", "lbl_sonuc_mesafe", "lbl_sonuc_barut_hakki_20",
            "lbl_sonuc_istikamet_acisi", "lbl_sonuc_ucus_suresi", "lbl_sonuc_nisangah",
            "lbl_sonuc_dogal_yan_dzl", "lbl_sonuc_mso", "lbl_sonuc_yso", "lbl_sonuc_100_M",
            "lbl_sonuc_20_M", "lbl_sonuc_dusus_acisi", "lbl_sonuc_tepe_yuksekligi", "lbl_sonuc_dtac",
            "lbl_sonuc_tac", "lbl_sonuc_gac_yan_duzeltmesi", "lbl_sonuc_toplam_yan_duzeltmesi",
            "lbl_sonuc_toplam_mesafe_duzeltmesi", "lbl_sonuc_toplam_tapa_saniye_duzeltmesi",
            "lbl_sonuc_en_kucuk_yukselis", "lbl_sonuc_metro_yan_duzeltmesi",
            "lbl_sonuc_metro_mesafe_duzeltmesi", "lbl_gac_mesafesi",
        ):
            getattr(self.ui, name).clear()
    @guvenli_slot
    def yazdir(self):
        from openpyxl import Workbook, load_workbook
        os.makedirs(ATIS_KAYIT_DIR, exist_ok=True)
        bolgevekoordinat = (
            (self.ui.lne_hedef_bolge_nu.text() or "37")
            + (self.ui.lne_hedef_bolge_sayisi.text() or "SFA")
            + (self.ui.lne_hedef_sag_deger.text() or "")
            + (self.ui.lne_hedef_yukari_deger.text() or "")
        )
        rakim = self.ui.lne_hedef_rakim.text()
        mesafe = self.ui.lne_sonuc_mesafe.text() or self.ui.lbl_sonuc_mesafe.text()
        barut_hakki = self.ui.lbl_sonuc_barut_hakki_2.text() or self.ui.lne_barut_hakki.text()
        tapa_saniyesi = self.ui.lbl_sonuc_tapa_saniyesi.text()
        yan_sonuc = self.ui.lbl_sonuc_yan_2.text()
        tablo_model = self.ui.tablo_atislar.model()
        dosya_adi = os.path.join(BASE_DIR, "Atis_Kayit_Formu.xlsx")
        
        try:
            # Mevcut dosyayı yükle
            wb = load_workbook(dosya_adi)
        except FileNotFoundError:
            # Dosya yoksa yeni bir dosya oluştur
            wb = Workbook()

        ws = wb.active

        # Belirli hücrelere verileri yaz
        hücreler = [   #SATIR, SÜTUN İNDEKS 0 DAN BAŞLIYOR
            
            ("C15", 0, 4), #sağ                   
            ("C16", 0, 5), #sol
            ("E15", 0, 6), #uzat
            ("E16", 0, 7), #kısalt
            ("G15", 0, 8), #kaldır
            ("G16", 0, 9), #indir
            ("AH15", 0, 0), #tarih
            ("S15", 0, 1), #YAN1
            ("AE15", 0, 2), #YUKSELİS1
            ("L15", 0, 3), #TAPASANİYESİ

            ("C17", 1, 4), #sağ                   
            ("C18", 1, 5), #sol
            ("E17", 1, 6), #uzat
            ("E18", 1, 7), #kısalt
            ("G17", 1, 8), #kaldır
            ("G18", 1, 9), #indir
            ("AH17", 1, 0), #tarih
            ("S17", 1, 1), #YAN1
            ("AE17", 1, 2), #YUKSELİS1
            ("L17", 1, 3), #TAPASANİYESİ

            ("C19", 2, 4), #sağ                   
            ("C20", 2, 5), #sol
            ("E19", 2, 6), #uzat
            ("E20", 2, 7), #kısalt
            ("G19", 2, 8), #kaldır
            ("G20", 2, 9), #indir
            ("AH19", 2, 0), #tarih
            ("S19", 2, 1), #YAN1
            ("AE19", 2, 2), #YUKSELİS1
            ("L19", 2, 3), #TAPASANİYESİ

            ("C21", 3, 4), #sağ                   
            ("C22", 3, 5), #sol
            ("E21", 3, 6), #uzat
            ("E22", 3, 7), #kısalt
            ("G21", 3, 8), #kaldır
            ("G22", 3, 9), #indir
            ("AH21", 3, 0), #tarih
            ("S21", 3, 1), #YAN1
            ("AE21", 3, 2), #YUKSELİS1
            ("L21", 3, 3), #TAPASANİYESİ,

            ("C23", 4, 4), #sağ                   
            ("C24", 4, 5), #sol
            ("E23", 4, 6), #uzat
            ("E24", 4, 7), #kısalt
            ("G23", 4, 8), #kaldır
            ("G24", 4, 9), #indir
            ("AH23", 4, 0), #tarih
            ("S23", 4, 1), #YAN1
            ("AE23", 4, 2), #YUKSELİS1
            ("L23", 4, 3), #TAPASANİYESİ

            ("C25", 5, 4), #sağ                   
            ("C26", 5, 5), #sol
            ("E25", 5, 6), #uzat
            ("E26", 5, 7), #kısalt
            ("G25", 5, 8), #kaldır
            ("G26", 5, 9), #indir
            ("AH25", 5, 0), #tarih
            ("S25", 5, 1), #YAN1
            ("AE25", 5, 2), #YUKSELİS1
            ("L25", 5, 3), #TAPASANİYESİ

            ("C27", 6, 4), #sağ                   
            ("C28", 6, 5), #sol
            ("E27", 6, 6), #uzat
            ("E28", 6, 7), #kısalt
            ("G27", 6, 8), #kaldır
            ("G28", 6, 9), #indir
            ("AH27", 6, 0), #tarih
            ("S27", 6, 1), #YAN1
            ("AE27", 6, 2), #YUKSELİS1
            ("L27", 6, 3), #TAPASANİYESİ

            ("C29", 7, 4), #sağ                   
            ("C30", 7, 5), #sol
            ("E29", 7, 6), #uzat
            ("E30", 7, 7), #kısalt
            ("G29", 7, 8), #kaldır
            ("G30", 7, 9), #indir
            ("AH29", 7, 0), #tarih
            ("S29", 7, 1), #YAN1
            ("AE29", 7, 2), #YUKSELİS1
            ("L29", 7, 3), #TAPASANİYESİ

            ("C31", 8, 4), #sağ                   
            ("C32", 8, 5), #sol
            ("E31", 8, 6), #uzat
            ("E32", 8, 7), #kısalt
            ("G31", 8, 8), #kaldır
            ("G32", 8, 9), #indir
            ("AH31", 8, 0), #tarih
            ("S31", 8, 1), #YAN1
            ("AE31", 8, 2), #YUKSELİS1
            ("L31", 8, 3), #TAPASANİYESİ

            ("C33", 9, 4), #sağ                   
            ("C34", 9, 5), #sol
            ("E33", 9, 6), #uzat
            ("E34", 9, 7), #kısalt
            ("G33", 9, 8), #kaldır
            ("G34", 9, 9), #indir
            ("AH33", 9, 0), #tarih
            ("S33", 9, 1), #YAN1
            ("AE33", 9, 2), #YUKSELİS1
            ("L33", 9, 3), #TAPASANİYESİ

            ("C35", 10, 4), #sağ                   
            ("C36", 10, 5), #sol
            ("E35", 10, 6), #uzat
            ("E36", 10, 7), #kısalt
            ("G35", 10, 8), #kaldır
            ("G36", 10, 9), #indir
            ("AH35", 10, 0), #tarih
            ("S35", 10, 1), #YAN1
            ("AE35", 10, 2), #YUKSELİS1
            ("L35", 10, 3), #TAPASANİYESİ

        
            # Diğer hücreleri ekleyin...
        ]

        for hücre_konumu, satir, sutun in hücreler:
            hücre_verisi = tablo_model.index(satir, sutun).data()
            ws[hücre_konumu] = hücre_verisi
        ws['C2'] = "METE 400"
        ws['I2'] = "TD0000" # GÜNCELLENECEK
        ws['M2'] = "1 NAMLU" # GÜNCELLENECEK
        ws['C3'] = f"{bolgevekoordinat}" # GÜNCELLENECEK
        ws['K3'] = f"{rakim}" # GÜNCELLENECEK
        # ws['M3'] = f"{istikamet}" # GÜNCELLENECEK
        ws['T8'] = f"{mesafe}" # GÜNCELLENECEK
        ws['D8'] = "AAG" # GÜNCELLENECEK
        ws['Q9'] = f"{barut_hakki}" # GÜNCELLENECEK
        ws['X9'] = f"{tapa_saniyesi}" # GÜNCELLENECEK
        ws['Z9'] = f"{yan_sonuc}" # GÜNCELLENECEK








  

        # Dosyayı kaydet
        wb.save(dosya_adi)
        print(f"Veriler {dosya_adi} dosyasına eklendi.")
        # time.sleep(1)
        import os
        from datetime import datetime
        import shutil
        # Dosyanın bir kopyasını oluşturun ve C:\Users\Nurhat DUMAN\Desktop\dos1\CDRIVE dizinine kopyalayın
        destination_folder = ATIS_KAYIT_DIR
        # Saati "%d.%m.%Y %H:%M" formatında alıyoruz
        current_time = datetime.now().strftime("%d.%m.%Y %H.%M")
        new_copy_filename = os.path.join(
            destination_folder, f"ATIŞ KAYIT FORMU {current_time}.xlsx")
        try:
            shutil.copy(dosya_adi, new_copy_filename)
        except OSError as exc:
            QMessageBox.warning(self, "Excel'e Aktar", "Kopya alınamadı: %s" % exc)
            return
        QtWidgets.QMessageBox.information(
            self, "Excel'e Aktar", "Veriler '%s' dosyasına yazıldı." % dosya_adi)

   
        import time  # Eğer daha önce yapmadıysanız, time modülünü içe aktarmayı unutmayın


    
        
        #################################################################################################       
def baslat():
    sys.excepthook = _genel_hata_yakala
    try:
        import threading
        threading.excepthook = lambda args: _genel_hata_yakala(
            args.exc_type, args.exc_value, args.exc_traceback)
    except Exception:
        pass
    try:
        app = SafeApplication(sys.argv)
        app.setQuitOnLastWindowClosed(True)
        pencere = myApp()
        pencere.show()
        app.exec_()
    except Exception as exc:
        hata_goster(exc, traceback.format_exc())


if __name__ == "__main__":
    baslat()