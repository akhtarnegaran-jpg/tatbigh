
# -*- coding: utf-8 -*-
"""
تطبیق - ابزار آفلاین مقایسه و ادغام فایل‌های اکسل
نسخه MVP 1.0
"""

import os
import re
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from rapidfuzz import fuzz, process


APP_TITLE = "تطبیق | مقایسه و ادغام فایل‌های اکسل"
KEY_SEP = " | "


def normalize_digits(value: str) -> str:
    """تبدیل اعداد فارسی و عربی به انگلیسی."""
    trans = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    return str(value).translate(trans)


def normalize_fa(value, remove_common_words=False) -> str:
    """
    استانداردسازی متن فارسی بدون تغییر مقدار اصلی:
    - ی/ک فارسی
    - اعداد انگلیسی
    - حذف نیم‌فاصله و فاصله‌های اضافی
    - حذف علائم متداول
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s.lower() in {"nan", "none", "nat"}:
        return ""

    s = normalize_digits(s)
    replacements = {
        "ي": "ی", "ى": "ی", "ئ": "ی",
        "ك": "ک",
        "\u200c": " ", "\u200f": " ", "\u200e": " ",
        "ۀ": "ه", "ة": "ه",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)

    s = re.sub(r"[ـ]+", "", s)
    s = re.sub(r"[\t\r\n]+", " ", s)
    s = re.sub(r"[،,:؛;()\[\]{}\"'`]+", " ", s)
    s = re.sub(r"[-_/\\]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    if remove_common_words:
        common = [
            "اداره آموزش و پرورش", "آموزش و پرورش",
            "شهرستان", "منطقه", "بخش",
        ]
        for word in common:
            s = re.sub(rf"(^|\s){re.escape(word)}(\s|$)", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()


def raw_key(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s.lower() in {"nan", "none", "nat"}:
        return ""
    return s


def make_key(df: pd.DataFrame, columns, mode: str, remove_common_words=False) -> pd.Series:
    funcs = {
        "raw": lambda x: raw_key(x),
        "normalized": lambda x: normalize_fa(x, remove_common_words),
        "fuzzy": lambda x: normalize_fa(x, remove_common_words),
    }
    fn = funcs[mode]
    parts = []
    for col in columns:
        parts.append(df[col].map(fn))
    if not parts:
        raise ValueError("هیچ ستون کلیدی انتخاب نشده است.")
    out = parts[0]
    for p in parts[1:]:
        out = out + KEY_SEP + p
    # کلیدهایی که همه اجزا خالی‌اند، خالی تلقی شوند.
    all_blank = pd.Series(True, index=df.index)
    for p in parts:
        all_blank &= p.eq("")
    out = out.mask(all_blank, "")
    return out


def safe_read_table(path: str, sheet_name=None) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        # تلاش با چند انکدینگ متداول
        last_error = None
        for enc in ("utf-8-sig", "utf-8", "cp1256"):
            try:
                return pd.read_csv(path, dtype=str, keep_default_na=False, encoding=enc)
            except Exception as exc:
                last_error = exc
        raise last_error
    return pd.read_excel(path, sheet_name=sheet_name, dtype=str, keep_default_na=False)


def list_sheets(path: str):
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        return ["CSV"]
    return pd.ExcelFile(path).sheet_names


def dedupe_column_names(columns):
    result = []
    counts = {}
    for c in columns:
        name = str(c).strip() or "ستون_بدون_نام"
        if name not in counts:
            counts[name] = 1
            result.append(name)
        else:
            counts[name] += 1
            result.append(f"{name}_{counts[name]}")
    return result


@dataclass
class MatchConfig:
    mode: str
    threshold: int
    min_margin: int
    remove_common_words: bool


def prepare_output_columns(df_a, df_b, selected_a, selected_b):
    a = df_a[selected_a].copy()
    b = df_b[selected_b].copy()

    a.columns = [f"فایل اول | {c}" for c in a.columns]
    b.columns = [f"فایل دوم | {c}" for c in b.columns]
    return a, b


def run_matching(df_a, df_b, keys_a, keys_b, out_cols_a, out_cols_b, config: MatchConfig):
    df_a = df_a.copy()
    df_b = df_b.copy()
    df_a.columns = dedupe_column_names(df_a.columns)
    df_b.columns = dedupe_column_names(df_b.columns)

    for col in keys_a + out_cols_a:
        if col not in df_a.columns:
            raise KeyError(f"ستون «{col}» در فایل اول پیدا نشد.")
    for col in keys_b + out_cols_b:
        if col not in df_b.columns:
            raise KeyError(f"ستون «{col}» در فایل دوم پیدا نشد.")

    if len(keys_a) != len(keys_b):
        raise ValueError("تعداد ستون‌های کلیدی انتخاب‌شده در دو فایل باید برابر باشد.")

    df_a["_row_a"] = range(2, len(df_a) + 2)
    df_b["_row_b"] = range(2, len(df_b) + 2)

    df_a["_match_key"] = make_key(
        df_a, keys_a, config.mode, config.remove_common_words
    )
    df_b["_match_key"] = make_key(
        df_b, keys_b, config.mode, config.remove_common_words
    )

    dup_a_mask = df_a["_match_key"].ne("") & df_a["_match_key"].duplicated(keep=False)
    dup_b_mask = df_b["_match_key"].ne("") & df_b["_match_key"].duplicated(keep=False)

    dup_a = df_a.loc[dup_a_mask].copy()
    dup_b = df_b.loc[dup_b_mask].copy()

    unique_a = df_a.loc[~dup_a_mask].copy()
    unique_b = df_b.loc[~dup_b_mask].copy()

    b_key_to_index = {
        key: idx for idx, key in unique_b["_match_key"].items() if key != ""
    }

    matched_rows = []
    suspicious = []
    used_b_indices = set()

    # ابتدا تطبیق دقیق، حتی در حالت fuzzy
    for idx_a, row_a in unique_a.iterrows():
        key_a = row_a["_match_key"]
        if key_a and key_a in b_key_to_index:
            idx_b = b_key_to_index[key_a]
            matched_rows.append({
                "idx_a": idx_a,
                "idx_b": idx_b,
                "نوع تطبیق": "دقیق",
                "درصد شباهت": 100,
                "کلید فایل اول": key_a,
                "کلید فایل دوم": key_a,
            })
            used_b_indices.add(idx_b)

    exactly_matched_a = {x["idx_a"] for x in matched_rows}

    # تطبیق تقریبی برای موارد باقی‌مانده
    if config.mode == "fuzzy":
        available_b = {
            idx: key for idx, key in unique_b["_match_key"].items()
            if key and idx not in used_b_indices
        }
        choices = list(available_b.values())
        value_to_indices = {}
        for idx, val in available_b.items():
            value_to_indices.setdefault(val, []).append(idx)

        for idx_a, row_a in unique_a.iterrows():
            if idx_a in exactly_matched_a:
                continue
            key_a = row_a["_match_key"]
            if not key_a or not choices:
                continue

            candidates = process.extract(
                key_a, choices, scorer=fuzz.WRatio, limit=2
            )
            if not candidates:
                continue

            best_value, best_score, _ = candidates[0]
            second_score = candidates[1][1] if len(candidates) > 1 else 0
            margin = best_score - second_score
            candidate_indices = [
                x for x in value_to_indices.get(best_value, [])
                if x not in used_b_indices
            ]

            record = {
                "idx_a": idx_a,
                "کلید فایل اول": key_a,
                "پیشنهاد فایل دوم": best_value,
                "درصد شباهت": round(float(best_score), 1),
                "فاصله از گزینه دوم": round(float(margin), 1),
                "وضعیت": "",
            }

            if (
                best_score >= config.threshold
                and margin >= config.min_margin
                and len(candidate_indices) == 1
            ):
                idx_b = candidate_indices[0]
                matched_rows.append({
                    "idx_a": idx_a,
                    "idx_b": idx_b,
                    "نوع تطبیق": "تقریبی",
                    "درصد شباهت": round(float(best_score), 1),
                    "کلید فایل اول": key_a,
                    "کلید فایل دوم": best_value,
                })
                used_b_indices.add(idx_b)
            else:
                if best_score < config.threshold:
                    record["وضعیت"] = "کمتر از آستانه"
                elif margin < config.min_margin:
                    record["وضعیت"] = "مبهم؛ گزینه دوم بسیار نزدیک است"
                else:
                    record["وضعیت"] = "کلید مقصد تکراری یا استفاده‌شده"
                suspicious.append(record)

    matched_a_indices = {x["idx_a"] for x in matched_rows}
    matched_b_indices = {x["idx_b"] for x in matched_rows}

    only_a = unique_a.loc[~unique_a.index.isin(matched_a_indices)].copy()
    only_b = unique_b.loc[~unique_b.index.isin(matched_b_indices)].copy()

    # ساخت نتیجه نهایی
    a_out, b_out = prepare_output_columns(
        df_a, df_b, out_cols_a, out_cols_b
    )

    final_records = []
    for m in matched_rows:
        rec = {}
        rec.update(a_out.loc[m["idx_a"]].to_dict())
        rec.update(b_out.loc[m["idx_b"]].to_dict())
        rec["وضعیت تطبیق"] = m["نوع تطبیق"]
        rec["درصد شباهت"] = m["درصد شباهت"]
        rec["ردیف فایل اول"] = int(df_a.loc[m["idx_a"], "_row_a"])
        rec["ردیف فایل دوم"] = int(df_b.loc[m["idx_b"], "_row_b"])
        final_records.append(rec)

    final_df = pd.DataFrame(final_records)

    def display_df(source, selected_cols, row_col):
        cols = [c for c in selected_cols if c in source.columns]
        result = source[cols].copy()
        result.insert(0, "ردیف منبع", source[row_col].astype(int))
        result["کلید استانداردشده"] = source["_match_key"]
        return result

    only_a_df = display_df(only_a, out_cols_a, "_row_a")
    only_b_df = display_df(only_b, out_cols_b, "_row_b")
    dup_a_df = display_df(dup_a, out_cols_a, "_row_a")
    dup_b_df = display_df(dup_b, out_cols_b, "_row_b")
    suspicious_df = pd.DataFrame(suspicious)

    summary = pd.DataFrame([
        ["نام فایل اول", ""],
        ["تعداد کل ردیف‌های فایل اول", len(df_a)],
        ["تعداد کل ردیف‌های فایل دوم", len(df_b)],
        ["تطبیق موفق", len(final_df)],
        ["درصد موفقیت نسبت به فایل اول",
         round((len(final_df) / len(df_a) * 100), 2) if len(df_a) else 0],
        ["تطبیق دقیق", sum(x["نوع تطبیق"] == "دقیق" for x in matched_rows)],
        ["تطبیق تقریبی", sum(x["نوع تطبیق"] == "تقریبی" for x in matched_rows)],
        ["فقط در فایل اول", len(only_a_df)],
        ["فقط در فایل دوم", len(only_b_df)],
        ["ردیف‌های دارای کلید تکراری در فایل اول", len(dup_a_df)],
        ["ردیف‌های دارای کلید تکراری در فایل دوم", len(dup_b_df)],
        ["پیشنهادهای تقریبی نیازمند بررسی", len(suspicious_df)],
        ["نوع پردازش", {
            "raw": "تطبیق دقیق بدون پاک‌سازی",
            "normalized": "تطبیق دقیق پس از استانداردسازی",
            "fuzzy": "استانداردسازی + تطبیق تقریبی"
        }[config.mode]],
        ["آستانه شباهت", config.threshold if config.mode == "fuzzy" else "-"],
        ["زمان اجرا", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ], columns=["شاخص", "مقدار"])

    return {
        "نتیجه نهایی": final_df,
        "فقط در فایل اول": only_a_df,
        "فقط در فایل دوم": only_b_df,
        "تکراری‌های فایل اول": dup_a_df,
        "تکراری‌های فایل دوم": dup_b_df,
        "موارد مشکوک": suspicious_df,
        "گزارش": summary,
    }


def autosize_worksheet(ws, max_width=45):
    ws.freeze_panes = "A2"
    ws.sheet_view.rightToLeft = True
    ws.auto_filter.ref = ws.dimensions

    for col_cells in ws.columns:
        max_len = 0
        letter = col_cells[0].column_letter
        for cell in col_cells[:300]:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), max_width)


def export_results(results, output_path):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet, df in results.items():
            if df is None or (df.empty and len(df.columns) == 0):
                df = pd.DataFrame({"توضیح": ["موردی یافت نشد."]})
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

        for ws in writer.book.worksheets:
            autosize_worksheet(ws)
            # قالب ساده عنوان
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
                cell.alignment = cell.alignment.copy(horizontal="center")


class TatbighApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1240x780")
        self.minsize(1050, 680)
        self.option_add("*Font", ("Tahoma", 10))

        self.file_a = tk.StringVar()
        self.file_b = tk.StringVar()
        self.sheet_a = tk.StringVar()
        self.sheet_b = tk.StringVar()
        self.mode = tk.StringVar(value="normalized")
        self.threshold = tk.IntVar(value=90)
        self.min_margin = tk.IntVar(value=4)
        self.remove_common_words = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="دو فایل را انتخاب کنید.")

        self.df_a = None
        self.df_b = None

        self._build_ui()

    def _build_ui(self):
        header = ttk.Frame(self, padding=(14, 12))
        header.pack(fill="x")
        ttk.Label(header, text="تطبیق", font=("Tahoma", 20, "bold")).pack(side="right")
        ttk.Label(
            header,
            text="ابزار آفلاین مقایسه، اعلام وصول و ادغام فایل‌های اکسل",
            font=("Tahoma", 11),
        ).pack(side="right", padx=18)

        file_box = ttk.LabelFrame(self, text="۱) انتخاب فایل‌ها و شیت‌ها", padding=10)
        file_box.pack(fill="x", padx=14, pady=6)

        self._file_row(file_box, 0, "فایل اول / مرجع", self.file_a, self.sheet_a, "a")
        self._file_row(file_box, 1, "فایل دوم / دریافتی", self.file_b, self.sheet_b, "b")

        middle = ttk.Panedwindow(self, orient="horizontal")
        middle.pack(fill="both", expand=True, padx=14, pady=6)

        key_box = ttk.LabelFrame(middle, text="۲) ستون‌های کلیدی مشترک", padding=10)
        out_box = ttk.LabelFrame(middle, text="۳) ستون‌های خروجی", padding=10)
        middle.add(key_box, weight=1)
        middle.add(out_box, weight=1)

        self.key_list_a = self._dual_list_panel(key_box, "کلیدهای فایل اول", 0)
        self.key_list_b = self._dual_list_panel(key_box, "کلیدهای فایل دوم", 1)

        mode_frame = ttk.LabelFrame(key_box, text="روش تطبیق", padding=8)
        mode_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Radiobutton(
            mode_frame, text="دقیق بدون پاک‌سازی", variable=self.mode, value="raw"
        ).grid(row=0, column=2, sticky="w", padx=5)
        ttk.Radiobutton(
            mode_frame, text="دقیق پس از استانداردسازی فارسی",
            variable=self.mode, value="normalized"
        ).grid(row=0, column=1, sticky="w", padx=5)
        ttk.Radiobutton(
            mode_frame, text="تقریبی برای نام‌ها", variable=self.mode, value="fuzzy"
        ).grid(row=0, column=0, sticky="w", padx=5)

        ttk.Label(mode_frame, text="حداقل شباهت:").grid(row=1, column=2, pady=8)
        ttk.Spinbox(
            mode_frame, from_=50, to=100, width=7, textvariable=self.threshold
        ).grid(row=1, column=1, sticky="w")
        ttk.Label(mode_frame, text="حداقل فاصله از گزینه دوم:").grid(row=2, column=2)
        ttk.Spinbox(
            mode_frame, from_=0, to=30, width=7, textvariable=self.min_margin
        ).grid(row=2, column=1, sticky="w")
        ttk.Checkbutton(
            mode_frame,
            text="حذف واژه‌های عمومی مانند «منطقه» و «شهرستان» هنگام مقایسه",
            variable=self.remove_common_words,
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

        self.out_list_a = self._dual_list_panel(out_box, "ستون‌های فایل اول", 0)
        self.out_list_b = self._dual_list_panel(out_box, "ستون‌های فایل دوم", 1)

        btns = ttk.Frame(self, padding=(14, 8))
        btns.pack(fill="x")
        ttk.Button(btns, text="بارگذاری و نمایش ستون‌ها", command=self.load_tables).pack(
            side="right", padx=4
        )
        ttk.Button(btns, text="انتخاب همه ستون‌های خروجی", command=self.select_all_output).pack(
            side="right", padx=4
        )
        ttk.Button(btns, text="اجرای تطبیق و ذخیره خروجی", command=self.process).pack(
            side="right", padx=4
        )

        status_bar = ttk.Label(
            self, textvariable=self.status, relief="sunken", anchor="e", padding=6
        )
        status_bar.pack(fill="x", side="bottom")

    def _file_row(self, parent, row, label, path_var, sheet_var, side):
        ttk.Label(parent, text=label).grid(row=row, column=3, sticky="e", padx=5, pady=4)
        ttk.Entry(parent, textvariable=path_var, width=80).grid(
            row=row, column=2, sticky="ew", padx=5, pady=4
        )
        ttk.Button(
            parent, text="انتخاب فایل", command=lambda s=side: self.choose_file(s)
        ).grid(row=row, column=1, padx=5)
        combo = ttk.Combobox(parent, textvariable=sheet_var, width=20, state="readonly")
        combo.grid(row=row, column=0, padx=5)
        if side == "a":
            self.sheet_combo_a = combo
        else:
            self.sheet_combo_b = combo
        parent.columnconfigure(2, weight=1)

    def _dual_list_panel(self, parent, title, col):
        frame = ttk.Frame(parent)
        frame.grid(row=0, column=col, sticky="nsew", padx=5)
        ttk.Label(frame, text=title, font=("Tahoma", 10, "bold")).pack(anchor="e")
        lb = tk.Listbox(
            frame, selectmode=tk.MULTIPLE, exportselection=False, height=14
        )
        scroll = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=scroll.set)
        lb.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        parent.columnconfigure(col, weight=1)
        parent.rowconfigure(0, weight=1)
        return lb

    def choose_file(self, side):
        path = filedialog.askopenfilename(
            title="انتخاب فایل داده",
            filetypes=[
                ("فایل‌های اکسل و CSV", "*.xlsx *.xls *.xlsm *.csv"),
                ("Excel", "*.xlsx *.xls *.xlsm"),
                ("CSV", "*.csv"),
                ("همه فایل‌ها", "*.*"),
            ],
        )
        if not path:
            return
        try:
            sheets = list_sheets(path)
            if side == "a":
                self.file_a.set(path)
                self.sheet_combo_a["values"] = sheets
                self.sheet_a.set(sheets[0])
            else:
                self.file_b.set(path)
                self.sheet_combo_b["values"] = sheets
                self.sheet_b.set(sheets[0])
            self.status.set(f"فایل انتخاب شد: {Path(path).name}")
        except Exception as exc:
            messagebox.showerror("خطا", f"خواندن فایل ممکن نشد:\n{exc}")

    def load_tables(self):
        if not self.file_a.get() or not self.file_b.get():
            messagebox.showwarning("فایل ناقص", "هر دو فایل را انتخاب کنید.")
            return
        try:
            self.status.set("در حال خواندن فایل‌ها...")
            self.update_idletasks()

            sheet_a = None if Path(self.file_a.get()).suffix.lower() == ".csv" else self.sheet_a.get()
            sheet_b = None if Path(self.file_b.get()).suffix.lower() == ".csv" else self.sheet_b.get()

            self.df_a = safe_read_table(self.file_a.get(), sheet_a)
            self.df_b = safe_read_table(self.file_b.get(), sheet_b)
            self.df_a.columns = dedupe_column_names(self.df_a.columns)
            self.df_b.columns = dedupe_column_names(self.df_b.columns)

            self._fill_list(self.key_list_a, self.df_a.columns, select_all=False)
            self._fill_list(self.key_list_b, self.df_b.columns, select_all=False)
            self._fill_list(self.out_list_a, self.df_a.columns, select_all=True)
            self._fill_list(self.out_list_b, self.df_b.columns, select_all=True)

            self.status.set(
                f"فایل اول: {len(self.df_a):,} ردیف | "
                f"فایل دوم: {len(self.df_b):,} ردیف | ستون‌های کلیدی را انتخاب کنید."
            )
        except Exception as exc:
            messagebox.showerror("خطا در بارگذاری", f"{exc}\n\n{traceback.format_exc()}")

    def _fill_list(self, lb, values, select_all=False):
        lb.delete(0, tk.END)
        for v in values:
            lb.insert(tk.END, v)
        if select_all and values:
            lb.selection_set(0, tk.END)

    def select_all_output(self):
        for lb in (self.out_list_a, self.out_list_b):
            lb.selection_set(0, tk.END)

    @staticmethod
    def selected_values(lb):
        return [lb.get(i) for i in lb.curselection()]

    def process(self):
        if self.df_a is None or self.df_b is None:
            self.load_tables()
            if self.df_a is None or self.df_b is None:
                return

        keys_a = self.selected_values(self.key_list_a)
        keys_b = self.selected_values(self.key_list_b)
        out_a = self.selected_values(self.out_list_a)
        out_b = self.selected_values(self.out_list_b)

        if not keys_a or not keys_b:
            messagebox.showwarning(
                "کلید مشترک",
                "حداقل یک ستون کلیدی از هر فایل انتخاب کنید."
            )
            return
        if len(keys_a) != len(keys_b):
            messagebox.showwarning(
                "کلیدهای نامتوازن",
                "تعداد ستون‌های کلیدی انتخاب‌شده در دو فایل باید برابر باشد.\n"
                "ترتیب انتخاب نیز باید متناظر باشد."
            )
            return
        if not out_a and not out_b:
            messagebox.showwarning("خروجی", "حداقل یک ستون خروجی انتخاب کنید.")
            return

        default_name = f"خروجی_تطبیق_{datetime.now():%Y%m%d_%H%M}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="ذخیره نتیجه تطبیق",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not output_path:
            return

        try:
            self.status.set("در حال تطبیق داده‌ها...")
            self.update_idletasks()

            config = MatchConfig(
                mode=self.mode.get(),
                threshold=int(self.threshold.get()),
                min_margin=int(self.min_margin.get()),
                remove_common_words=bool(self.remove_common_words.get()),
            )
            results = run_matching(
                self.df_a, self.df_b,
                keys_a, keys_b,
                out_a, out_b,
                config
            )
            results["گزارش"].iloc[0, 1] = Path(self.file_a.get()).name
            export_results(results, output_path)

            matched = len(results["نتیجه نهایی"])
            only_a = len(results["فقط در فایل اول"])
            only_b = len(results["فقط در فایل دوم"])
            self.status.set(f"پایان: {matched:,} تطبیق موفق. خروجی ذخیره شد.")
            messagebox.showinfo(
                "عملیات موفق",
                f"فایل خروجی ساخته شد.\n\n"
                f"تطبیق موفق: {matched:,}\n"
                f"فقط در فایل اول: {only_a:,}\n"
                f"فقط در فایل دوم: {only_b:,}\n\n"
                f"{output_path}"
            )
        except Exception as exc:
            messagebox.showerror(
                "خطا در پردازش",
                f"{exc}\n\nجزئیات فنی:\n{traceback.format_exc()}"
            )
            self.status.set("پردازش با خطا متوقف شد.")


if __name__ == "__main__":
    app = TatbighApp()
    app.mainloop()
