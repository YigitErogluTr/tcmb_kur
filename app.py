# app.py — TCMB Kur Çekim Otomasyonu (Tkinter)
# Gerekli: pip install tkcalendar pandas openpyxl requests

import os
import json
import threading
import time
import datetime as dt
from tkinter import Tk, ttk, StringVar, BooleanVar, filedialog, messagebox, IntVar
from tkcalendar import DateEntry
import requests
import pandas as pd
import xml.etree.ElementTree as ET

APP_TITLE = "TCMB Kur Çekim Otomasyonu"
SETTINGS_FILE = "settings.json"

# --------------------- TCMB XML Yardımcıları ---------------------
def tcmb_xml_url(d: dt.date) -> str:
    # Format: https://www.tcmb.gov.tr/kurlar/YYYYMM/DDMMYYYY.xml
    return f"https://www.tcmb.gov.tr/kurlar/{d.year}{d.month:02d}/{d.strftime('%d%m%Y')}.xml"

def _num(text: str | None):
    if text is None:
        return None
    txt = text.strip().replace(",", ".")
    try:
        return float(txt)
    except Exception:
        return None

def parse_tcmb_xml(content: bytes, wanted_codes: set[str]) -> dict:
    """İstenen para birimlerini XML'den çıkarır."""
    root = ET.fromstring(content)
    out = {}
    for cur in root.findall("Currency"):
        code = (cur.attrib.get("Kod") or "").strip().upper()
        if not code:
            continue
        if wanted_codes and code not in wanted_codes:
            continue
        out[code] = {
            "ForexBuying":     _num((cur.findtext("ForexBuying"))),
            "ForexSelling":    _num((cur.findtext("ForexSelling"))),
            "BanknoteBuying":  _num((cur.findtext("BanknoteBuying"))),
            "BanknoteSelling": _num((cur.findtext("BanknoteSelling"))),
        }
    return out

def fetch_range(start: dt.date, end: dt.date, currency_codes: list[str], on_progress=None) -> pd.DataFrame:
    """Tarih aralığındaki günlük XML'leri indirir -> DataFrame döner (satır=index: tarih, sütunlar=MultiIndex (Kod, Alan))."""
    wanted = set([c.upper() for c in currency_codes]) if currency_codes else set()
    rows = []
    days = (end - start).days + 1
    for i in range(days):
        d = start + dt.timedelta(days=i)
        url = tcmb_xml_url(d)
        ok = False
        try:
            r = requests.get(url, timeout=15)
            if r.status_code == 200:
                parsed = parse_tcmb_xml(r.content, wanted)
                if parsed:
                    for code, vals in parsed.items():
                        rows.append({"date": d, "code": code, **vals})
                    ok = True
        except Exception:
            ok = False
        if on_progress:
            on_progress(d, ok)
        time.sleep(0.03)
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).set_index(["date", "code"]).sort_index()
    # Pivot: kolonlar (Alan, Kod) -> (Kod, Alan)
    df = df.reset_index().pivot_table(index="date", columns="code",
                                      values=["ForexBuying", "ForexSelling", "BanknoteBuying", "BanknoteSelling"])
    df = df.swaplevel(axis=1).sort_index(axis=1, level=0)
    return df

# --------------------- Frekans Dönüşümleri -----------------------
def apply_frequency(df: pd.DataFrame, freq_key: str) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df.index = pd.to_datetime(df.index)
    df = df.sort_index().ffill()  # eksik günleri son değerle doldur

    freq_map = {
        "Günlük": None,
        "İşgünü": "B",
        "Haftalık": "W-FRI",
        "Aylık": "M",
        "3 Aylık": "Q",
        "6 Aylık": "2Q",
        "Yıllık": "A",
    }

    if freq_key == "Ayda 2 Kez":
        out = []
        months = sorted({(d.year, d.month) for d in df.index})
        for y, m in months:
            for day in (1, 15):
                try:
                    t = pd.Timestamp(dt.date(y, m, day))
                except ValueError:
                    continue
                if t in df.index:
                    out.append(df.loc[[t]])
                else:
                    fwd = df.loc[df.index >= t]
                    bak = df.loc[df.index <= t]
                    if not fwd.empty:
                        out.append(fwd.iloc[[0]])
                    elif not bak.empty:
                        out.append(bak.iloc[[-1]])
        return pd.concat(out).sort_index() if out else df

    code = freq_map.get(freq_key)
    return df if code is None else df.resample(code).last()

# --------------------- Genel Ayarlar -----------------------------
ALL_CURRENCIES = [
    "USD","EUR","GBP","CHF","JPY","CAD","DKK","NOK","SEK","AUD","RUB",
    "CNY","RON","ZAR","SAR","BGN","UAH","KWD","IRR","AZN","QAR"
]
FREQUENCIES = ["Günlük","İşgünü","Haftalık","Aylık","3 Aylık","6 Aylık","Yıllık","Ayda 2 Kez"]

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_settings(data: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# İngilizce -> Türkçe sütun adları
TR_MAP = {
    "ForexBuying": "Döviz Alış",
    "ForexSelling": "Döviz Satış",
    "BanknoteBuying": "Efektif Alış",
    "BanknoteSelling": "Efektif Satış",
}

# --------------------- Arayüz -----------------------------------
class App:
    def __init__(self, master: Tk):
        self.m = master
        self.m.title(APP_TITLE)
        self.m.geometry("980x640")

        self.start_date = dt.date(dt.date.today().year, 1, 1)
        self.end_date   = dt.date.today()

        self.vars_currency = {c: BooleanVar(value=False) for c in ALL_CURRENCIES}
        self.var_freq = {f: BooleanVar(value=(f == "Günlük")) for f in FREQUENCIES}

        self.out_folder = StringVar(value="")
        self.out_format = StringVar(value="xlsx")  # xlsx / csv

        self.progress_txt = StringVar(value="Hazır.")
        self.progress_val = IntVar(value=0)
        self.total_days = 0

        self._build_ui()
        self._load_from_file()

    def _build_ui(self):
        nb = ttk.Notebook(self.m)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        frm_manual = ttk.Frame(nb)
        nb.add(frm_manual, text="Manuel Veri Çekimi")

        # Tarih aralığı
        grp_dates = ttk.LabelFrame(frm_manual, text="Tarih Aralığı")
        grp_dates.pack(fill="x", padx=8, pady=8)
        ttk.Label(grp_dates, text="Başlangıç:").pack(side="left", padx=6)
        self.dp_start = DateEntry(grp_dates, date_pattern="dd.mm.yyyy")
        self.dp_start.set_date(self.start_date)
        self.dp_start.pack(side="left", padx=6)
        ttk.Label(grp_dates, text="Bitiş:").pack(side="left", padx=6)
        self.dp_end = DateEntry(grp_dates, date_pattern="dd.mm.yyyy")
        self.dp_end.set_date(self.end_date)
        self.dp_end.pack(side="left", padx=6)

        # Döviz & Frekans
        container = ttk.Frame(frm_manual)
        container.pack(fill="both", expand=True, padx=8, pady=4)

        grp_fx = ttk.LabelFrame(container, text="Döviz Kurları")
        grp_fx.pack(side="left", fill="both", expand=True, padx=(0,8))
        left = ttk.Frame(grp_fx); left.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        right = ttk.Frame(grp_fx); right.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        half = (len(ALL_CURRENCIES)+1)//2
        for c in ALL_CURRENCIES[:half]:
            ttk.Checkbutton(left, text=c, variable=self.vars_currency[c]).pack(anchor="w")
        for c in ALL_CURRENCIES[half:]:
            ttk.Checkbutton(right, text=c, variable=self.vars_currency[c]).pack(anchor="w")
        btns_fx = ttk.Frame(grp_fx); btns_fx.pack(fill="x", pady=(6,2))
        ttk.Button(btns_fx, text="Tümü", command=self._select_all_currencies).pack(side="left", padx=4)
        ttk.Button(btns_fx, text="Hiçbiri", command=self._clear_all_currencies).pack(side="left", padx=4)

        grp_freq = ttk.LabelFrame(container, text="Frekanslar")
        grp_freq.pack(side="left", fill="both", expand=True)
        for f in FREQUENCIES:
            ttk.Checkbutton(grp_freq, text=f, variable=self.var_freq[f], command=self._ensure_single_freq).pack(anchor="w")
        ttk.Label(grp_freq, text="(Tek frekans seçilir)").pack(anchor="w", pady=(6,0))
        btns_fr = ttk.Frame(grp_freq); btns_fr.pack(fill="x", pady=(6,2))
        ttk.Button(btns_fr, text="Tümü", command=lambda: self._set_all_freq(True)).pack(side="left", padx=4)
        ttk.Button(btns_fr, text="Hiçbiri", command=lambda: self._set_all_freq(False)).pack(side="left", padx=4)

        # Çıktı ve kayıt
        grp_out = ttk.LabelFrame(frm_manual, text="Çıktı ve Kayıt")
        grp_out.pack(fill="x", padx=8, pady=8)
        ttk.Button(grp_out, text="Klasör Seç", command=self._choose_folder).pack(side="right", padx=6, pady=6)
        ttk.Label(grp_out, textvariable=self.out_folder, foreground="#444").pack(side="left", padx=8)

        frm_fmt = ttk.Frame(frm_manual); frm_fmt.pack(fill="x", padx=8)
        ttk.Label(frm_fmt, text="Format:").pack(side="left", padx=6)
        ttk.Radiobutton(frm_fmt, text="Excel (.xlsx)", variable=self.out_format, value="xlsx").pack(side="left", padx=4)
        ttk.Radiobutton(frm_fmt, text="CSV (.csv)", variable=self.out_format, value="csv").pack(side="left", padx=4)

        # Alt butonlar
        btn_row = ttk.Frame(frm_manual); btn_row.pack(fill="x", padx=8, pady=12)
        ttk.Button(btn_row, text="Ayarları Kaydet", command=self._save_to_file).pack(side="left")
        ttk.Button(btn_row, text="Manuel Veri Çekimini Başlat", command=self._start_job).pack(side="right")

        # İlerleme
        pr = ttk.Frame(frm_manual); pr.pack(fill="x", padx=8, pady=(0,10))
        self.pb = ttk.Progressbar(pr, orient="horizontal", mode="determinate", maximum=100)
        self.pb.pack(fill="x")
        ttk.Label(pr, textvariable=self.progress_txt).pack(anchor="w")

    # --- UI yardımcıları ---
    def _select_all_currencies(self):
        for v in self.vars_currency.values():
            v.set(True)

    def _clear_all_currencies(self):
        for v in self.vars_currency.values():
            v.set(False)

    def _set_all_freq(self, state: bool):
        for v in self.var_freq.values():
            v.set(state)
        if not state:
            self.var_freq["Günlük"].set(True)

    def _ensure_single_freq(self):
        active = [f for f, v in self.var_freq.items() if v.get()]
        if len(active) > 1:
            keep = active[0]
            for f, v in self.var_freq.items():
                v.set(f == keep)

    def _choose_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.out_folder.set(path)

    def _get_selected_frequency(self) -> str:
        for f, v in self.var_freq.items():
            if v.get():
                return f
        return "Günlük"

    def _save_to_file(self):
        data = {
            "start": self.dp_start.get_date().strftime("%Y-%m-%d"),
            "end": self.dp_end.get_date().strftime("%Y-%m-%d"),
            "currencies": [c for c, v in self.vars_currency.items() if v.get()],
            "frequency": self._get_selected_frequency(),
            "out_folder": self.out_folder.get(),
            "out_format": self.out_format.get(),
        }
        save_settings(data)
        messagebox.showinfo("Kaydedildi", "Ayarlar kaydedildi.")

    def _load_from_file(self):
        s = load_settings()
        if not s:
            return
        try:
            self.dp_start.set_date(dt.date.fromisoformat(s.get("start", "")))
            self.dp_end.set_date(dt.date.fromisoformat(s.get("end", "")))
        except Exception:
            pass
        for c in s.get("currencies", []):
            if c in self.vars_currency:
                self.vars_currency[c].set(True)
        self._set_all_freq(False)
        f = s.get("frequency", "Günlük")
        if f in self.var_freq:
            self.var_freq[f].set(True)
        self.out_folder.set(s.get("out_folder", ""))
        self.out_format.set(s.get("out_format", "xlsx"))

    # --- İş akışı ---
    def _start_job(self):
        start = self.dp_start.get_date()
        end = self.dp_end.get_date()
        if end < start:
            messagebox.showerror("Hata", "Bitiş tarihi başlangıçtan küçük olamaz.")
            return
        codes = [c for c, v in self.vars_currency.items() if v.get()]
        if not codes:
            if messagebox.askyesno("Uyarı", "Hiç döviz seçmediniz. Tüm dövizler çekilsin mi?"):
                codes = []
            else:
                return
        if not self.out_folder.get():
            messagebox.showerror("Hata", "Lütfen bir kayıt klasörü seçin.")
            return

        self.total_days = (end - start).days + 1
        self.pb["value"] = 0
        self.progress_txt.set("Başladı...")
        threading.Thread(target=self._run_job, args=(start, end, codes), daemon=True).start()

    def _run_job(self, start, end, codes):
        def on_progress(d, ok):
            done_days = (d - start).days + 1
            pct = int(done_days / max(self.total_days, 1) * 100)
            self.pb.after(0, lambda: self.pb.config(value=pct))
            self.progress_txt.set(f"{d.strftime('%d.%m.%Y')} {'✓' if ok else '—'}")

        try:
            df = fetch_range(start, end, codes, on_progress=on_progress)
            if df.empty:
                self.progress_txt.set("Veri bulunamadı. Tarih aralığını ve seçimleri kontrol edin.")
                return

            freq = self._get_selected_frequency()
            df2 = apply_frequency(df, freq)

            # ÇIKTI: Türkçe başlıklar
            flat_cols = [f"{code}-{TR_MAP.get(field, field)}"
                         for (code, field) in df2.columns.to_flat_index()]
            out_df = df2.copy()
            out_df.columns = flat_cols
            out_df.index.name = "Tarih"

            start_s = start.strftime("%Y%m%d")
            end_s = end.strftime("%Y%m%d")
            fname = f"TCMB_Kurlar_{start_s}_{end_s}_{freq.replace(' ', '')}.{self.out_format.get()}"
            fpath = os.path.join(self.out_folder.get(), fname)

            if self.out_format.get() == "xlsx":
                out_df.to_excel(fpath, engine="openpyxl", freeze_panes=(1,1))
            else:
                out_df.to_csv(fpath, encoding="utf-8-sig")

            self.progress_txt.set(f"Bitti. Kaydedildi: {fpath}")
            messagebox.showinfo("Tamamlandı", f"Dosya kaydedildi:\n{fpath}")
        except Exception as e:
            self.progress_txt.set(f"Hata: {e}")
            messagebox.showerror("Hata", str(e))

# --------------------- main ---------------------
if __name__ == "__main__":
    root = Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App(root)
    root.mainloop()
