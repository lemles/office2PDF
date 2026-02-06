import os
import json
import datetime
import re
import threading
import queue
import tempfile
import io
import random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from dataclasses import dataclass, asdict
from typing import List, Dict, Any, Optional, Tuple

import pythoncom
import win32com.client
from pypdf import PdfWriter, PdfReader
from PIL import Image

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

from tkinterdnd2 import DND_FILES, TkinterDnD

try:
    import winreg
except ImportError:
    winreg = None

CONFIG_FILE = "pdf_pro_config_v4.json"
WM_TEMPLATE_FILE = "watermark_templates.txt"
NM_TEMPLATE_FILE = "naming_templates.txt"


@dataclass
class AppConfig:
    output_dir: str = ""
    out_mode: str = "original"
    naming_tpl: str = "{name}"
    auto_open: bool = True
    open_folder: bool = False
    clear_after: bool = False
    compress_pdf: bool = False  # 追加

    wm1_text: str = ""
    wm1_pos: str = "None"
    wm2_text: str = ""
    wm2_pos: str = "None"
    wm_font: str = ""
    wm_size: int = 60
    wm_color: str = "#C0C0C0"
    wm_alpha: float = 0.3

    pg_enabled: bool = False
    pg_pos: str = "中央下"
    pg_format: str = "- {n} / {total} -"

    merge_all: bool = False
    split_word_page: bool = False
    split_ppt_page: bool = False
    split_pdf_page: bool = False
    split_excel_sheet: bool = False
    split_excel_page: bool = False

    password: str = ""
    excel_fit: bool = False
    excel_fit_tall: bool = False
    clear_metadata: bool = False

    def __post_init__(self):
        if not self.output_dir:
            self.output_dir = os.path.expanduser(r"~\Desktop")


class PDFUltimateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Office2PDF v4.5")
        self.root.geometry("900x920")

        self.files: List[Dict[str, Any]] = []
        self.presets: Dict[str, Any] = {}
        self.font_map: Dict[str, Tuple[str, int]] = {}
        self.processing = False
        self.cancel_flag = threading.Event()
        self.progress_queue = queue.Queue()

        self.load_config()
        self.init_templates()
        self.setup_ui()
        self.apply_config_to_ui()

        self.root.after(100, self.check_progress_queue)
        self.root.after(600, self.load_fonts_delayed)

    # --- Config & Templates ---
    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.config = AppConfig(**data.get("current", {}))
                self.presets = data.get("presets", {})
            except:
                self.config = AppConfig()
        else:
            self.config = AppConfig()

    def save_config(self):
        try:
            self.update_config_from_ui()
            data = {"current": asdict(self.config), "presets": self.presets}
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.queue_log(f"設定保存エラー: {e}")

    def init_templates(self):
        wm_content = "社外秘\nコピー厳禁\n取扱注意\nSAMPLE\nCONFIDENTIAL\nDRAFT\nCOPY\nIMPORTANT\n{date:yyyy/mm/dd}\n{date:yyyy/mm/dd hh:mm}"
        nm_content = "{name}\n{name}_{sheet}\n{seq}_{name}\n{date:yyyy-mm-dd}_{name}"
        for file, content in [(WM_TEMPLATE_FILE, wm_content), (NM_TEMPLATE_FILE, nm_content)]:
            if not os.path.exists(file):
                with open(file, "w", encoding="utf-8") as f:
                    f.write(content)
        with open(WM_TEMPLATE_FILE, "r", encoding="utf-8") as f:
            self.wm_templates = [l.strip() for l in f if l.strip()]
        with open(NM_TEMPLATE_FILE, "r", encoding="utf-8") as f:
            self.nm_templates = [l.strip() for l in f if l.strip()]

    def update_config_from_ui(self):
        self.config.wm1_text = self.wm1_val.get()
        self.config.wm1_pos = self.wm1_pos_var.get()
        self.config.wm2_text = self.wm2_val.get()
        self.config.wm2_pos = self.wm2_pos_var.get()
        self.config.wm_font = self.wm_font_combo.get()
        self.config.wm_alpha = float(self.wm_alpha_scale.get())
        self.config.wm_size = int(self.wm_size_spin.get())
        self.config.wm_color = self.wm_color_btn.cget("bg")
        self.config.pg_enabled = self.pg_en_var.get()
        self.config.pg_pos = self.pg_pos_var.get()
        self.config.pg_format = self.pg_fmt_var.get()
        self.config.naming_tpl = self.naming_var.get()
        self.config.out_mode = self.out_mode_var.get()
        self.config.output_dir = self.out_dir_var.get()
        self.config.merge_all = self.merge_var.get()
        self.config.split_word_page = self.sp_word_var.get()
        self.config.split_ppt_page = self.sp_ppt_var.get()
        self.config.split_pdf_page = self.sp_pdf_var.get()
        self.config.split_excel_sheet = self.sp_ex_sheet_var.get()
        self.config.split_excel_page = self.sp_ex_page_var.get()
        self.config.auto_open = self.open_var.get()
        self.config.open_folder = self.folder_var.get()
        self.config.clear_after = self.clear_after_var.get()
        self.config.password = self.pw_var.get()
        self.config.excel_fit = self.excel_fit_var.get()
        self.config.excel_fit_tall = self.excel_fit_tall_var.get()
        self.config.clear_metadata = self.meta_var.get()
        self.config.compress_pdf = self.compress_var.get()

    def apply_config_to_ui(self):
        self.wm1_val.set(self.config.wm1_text)
        self.wm1_pos_var.set(self.config.wm1_pos)
        self.wm2_val.set(self.config.wm2_text)
        self.wm2_pos_var.set(self.config.wm2_pos)
        self.wm_font_combo.set(self.config.wm_font)
        self.wm_alpha_scale.set(self.config.wm_alpha)
        self.wm_size_spin.delete(0, tk.END)
        self.wm_size_spin.insert(0, self.config.wm_size)
        self.wm_color_btn.config(bg=self.config.wm_color)
        self.pg_en_var.set(self.config.pg_enabled)
        self.pg_pos_var.set(self.config.pg_pos)
        self.pg_fmt_var.set(self.config.pg_format)
        self.naming_var.set(self.config.naming_tpl)
        self.out_mode_var.set(self.config.out_mode)
        self.out_dir_var.set(self.config.output_dir)
        self.merge_var.set(self.config.merge_all)
        self.sp_word_var.set(self.config.split_word_page)
        self.sp_ppt_var.set(self.config.split_ppt_page)
        self.sp_pdf_var.set(self.config.split_pdf_page)
        self.sp_ex_sheet_var.set(self.config.split_excel_sheet)
        self.sp_ex_page_var.set(self.config.split_excel_page)
        self.open_var.set(self.config.auto_open)
        self.folder_var.set(self.config.open_folder)
        self.clear_after_var.set(self.config.clear_after)
        self.pw_var.set(self.config.password)
        self.excel_fit_var.set(self.config.excel_fit)
        self.excel_fit_tall_var.set(self.config.excel_fit_tall)
        self.meta_var.set(self.config.clear_metadata)
        self.compress_var.set(self.config.compress_pdf)
        self.update_output_preview()

    # --- UI Setup ---
    def setup_ui(self):
        main_container = tk.Frame(self.root, padx=10, pady=5)
        main_container.pack(fill=tk.BOTH, expand=True)

        # File List
        file_frame = tk.LabelFrame(
            main_container, text="変換ファイル(ドロップで登録・ダブルクリックでページ指定)", padx=5, pady=5
        )
        file_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.tree = ttk.Treeview(file_frame, columns=("Type", "Name", "Range", "Out"), show="headings", height=8)
        for col, head, w in [
            ("Type", "種別", 70),
            ("Name", "ファイル名", 300),
            ("Range", "対象範囲", 150),
            ("Out", "出力プレビュー", 300),
        ]:
            self.tree.heading(col, text=head)
            self.tree.column(col, width=w)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self.handle_drop)
        self.tree.bind("<Double-1>", self.on_list_double_click)

        btn_f = tk.Frame(file_frame)
        btn_f.pack(fill=tk.X, padx=5, pady=2)
        tk.Button(btn_f, text="上へ", command=lambda: self.move_file(-1)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text="下へ", command=lambda: self.move_file(1)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text="選択削除", command=self.remove_file).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text="リスト全クリア", command=self.clear_list).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_f, text="フォルダ追加", command=self.add_folder).pack(side=tk.LEFT, padx=2)

        # Preset Area
        pre_f = tk.Frame(main_container)
        pre_f.pack(fill=tk.X, pady=5)
        tk.Label(pre_f, text="設定プリセット:").pack(side=tk.LEFT)
        self.preset_combo = ttk.Combobox(pre_f, values=list(self.presets.keys()), width=25)
        self.preset_combo.pack(side=tk.LEFT, padx=5)
        tk.Button(pre_f, text="読込", command=self.load_preset).pack(side=tk.LEFT, padx=2)
        tk.Button(pre_f, text="保存", command=self.save_preset).pack(side=tk.LEFT, padx=2)
        tk.Button(pre_f, text="削除", command=self.delete_preset).pack(side=tk.LEFT, padx=2)

        # Mid Columns
        mid_frame = tk.Frame(main_container)
        mid_frame.pack(fill=tk.X, pady=5)

        # Left: Watermark
        wm_frame = tk.LabelFrame(mid_frame, text="透かし・ページ番号設定", padx=10, pady=5)
        wm_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        f_row = tk.Frame(wm_frame)
        f_row.pack(fill=tk.X, pady=2)
        tk.Label(f_row, text="フォント:").pack(side=tk.LEFT)
        self.wm_font_combo = ttk.Combobox(f_row, width=25)
        self.wm_font_combo.pack(side=tk.LEFT, padx=5)
        tk.Label(f_row, text="サイズ:").pack(side=tk.LEFT)
        self.wm_size_spin = tk.Spinbox(f_row, from_=10, to=300, width=5)
        self.wm_size_spin.pack(side=tk.LEFT, padx=5)
        self.wm_color_btn = tk.Button(f_row, bg="#C0C0C0", width=2, command=self.choose_color)
        self.wm_color_btn.pack(side=tk.LEFT, padx=5)

        a_row = tk.Frame(wm_frame)
        a_row.pack(fill=tk.X, pady=2)
        tk.Label(a_row, text="不透明度:").pack(side=tk.LEFT)
        self.wm_alpha_scale = tk.Scale(a_row, from_=0.0, to=1.0, resolution=0.05, orient=tk.HORIZONTAL, length=120)
        self.wm_alpha_scale.pack(side=tk.LEFT, padx=5)
        # プレビューボタン追加
        tk.Button(a_row, text="選択ファイルの1頁目をプレビュー", bg="#e1e1e1", command=self.preview_watermark).pack(
            side=tk.RIGHT
        )

        pos_opts = ["None", "中央斜め", "中央大", "左上", "上中央", "右上", "左下", "下中央", "右下"]
        for i in [1, 2]:
            r = tk.Frame(wm_frame)
            r.pack(fill=tk.X, pady=2)
            tk.Label(r, text=f"透かし{i}:", width=7).pack(side=tk.LEFT)
            val = ttk.Combobox(r, values=self.wm_templates, width=20)
            val.pack(side=tk.LEFT, padx=2)
            setattr(self, f"wm{i}_val", val)
            p_var = tk.StringVar(value="None")
            ttk.Combobox(r, textvariable=p_var, values=pos_opts, width=10).pack(side=tk.LEFT, padx=2)
            setattr(self, f"wm{i}_pos_var", p_var)

        pg_row = tk.Frame(wm_frame)
        pg_row.pack(fill=tk.X, pady=5)
        self.pg_en_var = tk.BooleanVar()
        tk.Checkbutton(pg_row, text="ページ番号:", variable=self.pg_en_var).pack(side=tk.LEFT)
        self.pg_fmt_var = tk.StringVar()
        tk.Entry(pg_row, textvariable=self.pg_fmt_var, width=12).pack(side=tk.LEFT, padx=2)
        self.pg_pos_var = tk.StringVar()
        ttk.Combobox(pg_row, textvariable=self.pg_pos_var, values=["中央下", "右下"], width=8).pack(
            side=tk.LEFT, padx=2
        )

        # Right: Split Settings
        split_frame = tk.LabelFrame(mid_frame, text="出力・分割詳細設定", padx=10, pady=5)
        split_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.merge_var = tk.BooleanVar()
        tk.Checkbutton(
            split_frame,
            text="【全結合】1つのPDFにまとめる",
            variable=self.merge_var,
            font=("", 9, "bold"),
            fg="blue",
            command=self.toggle_merge_all,
        ).pack(anchor="w")

        f1 = tk.Frame(split_frame)
        f1.pack(fill=tk.X, pady=(5, 0))
        tk.Label(f1, text="Word/PPT:", width=10, anchor="w").pack(side=tk.LEFT)
        self.sp_word_var, self.sp_ppt_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(f1, text="ページ毎分割", variable=self.sp_word_var, command=self.toggle_split_box).pack(
            side=tk.LEFT
        )
        tk.Checkbutton(f1, text="(PPTも)", variable=self.sp_ppt_var, command=self.toggle_split_box).pack(side=tk.LEFT)

        f2 = tk.Frame(split_frame)
        f2.pack(fill=tk.X)
        tk.Label(f2, text="PDF:", width=10, anchor="w").pack(side=tk.LEFT)
        self.sp_pdf_var = tk.BooleanVar()
        tk.Checkbutton(f2, text="ページ毎分割", variable=self.sp_pdf_var, command=self.toggle_split_box).pack(
            side=tk.LEFT
        )

        f3 = tk.Frame(split_frame)
        f3.pack(fill=tk.X)
        tk.Label(f3, text="Excel:", width=10, anchor="w").pack(side=tk.LEFT)
        self.sp_ex_sheet_var, self.sp_ex_page_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(f3, text="シート毎", variable=self.sp_ex_sheet_var, command=self.toggle_excel_sheet).pack(
            side=tk.LEFT
        )
        tk.Checkbutton(f3, text="全ページ毎", variable=self.sp_ex_page_var, command=self.toggle_excel_page).pack(
            side=tk.LEFT
        )

        tk.Label(split_frame, text="--- Excelオプション (印刷範囲優先) ---", fg="gray", font=("", 8)).pack(
            anchor="w", pady=(5, 0)
        )
        f4 = tk.Frame(split_frame)
        f4.pack(fill=tk.X)
        self.excel_fit_var, self.excel_fit_tall_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(f4, text="横幅1Pに収める", variable=self.excel_fit_var).pack(side=tk.LEFT)
        tk.Checkbutton(f4, text="縦幅1P", variable=self.excel_fit_tall_var).pack(side=tk.LEFT)

        # Bottom
        bottom_frame = tk.LabelFrame(main_container, text="保存設定・実行", padx=10, pady=5)
        bottom_frame.pack(fill=tk.X, pady=5)

        n_row = tk.Frame(bottom_frame)
        n_row.pack(fill=tk.X, pady=2)
        tk.Label(n_row, text="ファイル名・命名ルール:").pack(side=tk.LEFT)
        self.naming_var = tk.StringVar()
        self.naming_combo = ttk.Combobox(n_row, textvariable=self.naming_var, values=self.nm_templates, width=40)
        self.naming_combo.pack(side=tk.LEFT, padx=5)

        # ヘルプボタンを追加
        tk.Button(n_row, text="タグ説明", command=self.show_naming_help, bg="#f0f0f0").pack(side=tk.LEFT)

        self.final_name_label = tk.Label(n_row, text="", fg="#0056b3", font=("", 9, "bold"))
        self.final_name_label.pack(side=tk.LEFT, padx=10)

        o_row = tk.Frame(bottom_frame)
        o_row.pack(fill=tk.X, pady=2)
        self.out_mode_var = tk.StringVar()
        tk.Radiobutton(
            o_row, text="元と同じ場所", variable=self.out_mode_var, value="original", command=self.update_output_preview
        ).pack(side=tk.LEFT)
        tk.Radiobutton(
            o_row, text="カスタム:", variable=self.out_mode_var, value="custom", command=self.update_output_preview
        ).pack(side=tk.LEFT)
        self.out_dir_var = tk.StringVar()
        tk.Entry(o_row, textvariable=self.out_dir_var, width=40).pack(side=tk.LEFT, padx=5)
        tk.Button(o_row, text="参照", command=self.browse_dir).pack(side=tk.LEFT)

        s_row = tk.Frame(bottom_frame)
        s_row.pack(fill=tk.X, pady=2)
        self.pw_var = tk.StringVar()
        tk.Label(s_row, text="PW:").pack(side=tk.LEFT)
        tk.Entry(s_row, textvariable=self.pw_var, width=12).pack(side=tk.LEFT, padx=5)
        self.meta_var, self.compress_var = tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(s_row, text="メタ削除", variable=self.meta_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(s_row, text="PDF軽量化", variable=self.compress_var).pack(side=tk.LEFT, padx=5)
        self.open_var, self.folder_var, self.clear_after_var = tk.BooleanVar(), tk.BooleanVar(), tk.BooleanVar()
        tk.Checkbutton(s_row, text="完了後開く", variable=self.open_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(s_row, text="フォルダ開く", variable=self.folder_var).pack(side=tk.LEFT, padx=5)
        tk.Checkbutton(s_row, text="リストクリア", variable=self.clear_after_var).pack(side=tk.LEFT, padx=5)

        exec_f = tk.Frame(bottom_frame)
        exec_f.pack(fill=tk.X, pady=10)
        self.btn_convert = tk.Button(
            exec_f,
            text="PDF変換開始",
            bg="#28a745",
            fg="white",
            font=("", 12, "bold"),
            height=2,
            command=self.start_thread,
        )
        self.btn_convert.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.btn_cancel = tk.Button(
            exec_f,
            text="キャンセル",
            bg="#dc3545",
            fg="white",
            font=("", 12, "bold"),
            height=2,
            command=self.cancel_process,
            state=tk.DISABLED,
        )
        self.btn_cancel.pack(side=tk.LEFT, padx=2)

        self.progress_label = tk.Label(main_container, text="待機中...", anchor="w")
        self.progress_label.pack(fill=tk.X)
        self.progress = ttk.Progressbar(main_container, orient=tk.HORIZONTAL, mode="determinate")
        self.progress.pack(fill=tk.X)

        log_f = tk.Frame(main_container)
        log_f.pack(fill=tk.BOTH, expand=False, pady=5)
        self.log_text = tk.Text(log_f, height=3, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = tk.Scrollbar(log_f, command=self.log_text.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=sb.set, state=tk.DISABLED)

        self.naming_var.trace_add("write", lambda *_: self.update_output_preview())
        self.out_dir_var.trace_add("write", lambda *_: self.update_output_preview())

    # --- Mutual Exclusion ---
    def toggle_merge_all(self):
        if self.merge_var.get():
            for v in [self.sp_word_var, self.sp_ppt_var, self.sp_pdf_var, self.sp_ex_sheet_var, self.sp_ex_page_var]:
                v.set(False)
        self.update_output_preview()

    def toggle_split_box(self):
        if any([self.sp_word_var.get(), self.sp_ppt_var.get(), self.sp_pdf_var.get()]):
            self.merge_var.set(False)
        self.update_output_preview()

    def toggle_excel_sheet(self):
        if self.sp_ex_sheet_var.get():
            self.merge_var.set(False)
            self.sp_ex_page_var.set(False)
        self.update_output_preview()

    def toggle_excel_page(self):
        if self.sp_ex_page_var.get():
            self.merge_var.set(False)
            self.sp_ex_sheet_var.set(False)
        self.update_output_preview()

    # --- Font Logic (Fixed for TTC) ---
    def build_registry_font_items(self):
        items_for_combo = []
        self.font_map = {}
        if winreg is None:
            return []
        fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        roots = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        ]
        seen_names = set()
        for root, keypath in roots:
            try:
                with winreg.OpenKey(root, keypath) as k:
                    i = 0
                    while True:
                        try:
                            name_raw, val, _ = winreg.EnumValue(k, i)
                            i += 1
                            if not isinstance(val, str):
                                continue
                            file_name = re.split(r"[,&]", val)[0].strip()
                            full_path = os.path.join(fonts_dir, file_name)
                            if not os.path.exists(full_path):
                                continue
                            clean_raw = re.sub(r"\s*\(.*?\)", "", name_raw).replace(";", "")
                            sub_names = [n.strip() for n in clean_raw.split("&")]
                            for idx, sub_name in enumerate(sub_names):
                                if sub_name and sub_name not in seen_names:
                                    seen_names.add(sub_name)
                                    self.font_map[sub_name] = (full_path, idx)
                                    items_for_combo.append(sub_name)
                        except OSError:
                            break
            except:
                continue
        items_for_combo.sort()
        return items_for_combo

    def load_fonts_delayed(self):
        try:
            names = self.build_registry_font_items()
            self.wm_font_combo["values"] = names
            if self.config.wm_font in names:
                self.wm_font_combo.set(self.config.wm_font)
            else:
                biz = [n for n in names if "BIZ" in n.upper()]
                if biz:
                    self.wm_font_combo.set(biz[0])
            self.queue_log("フォント一覧の読み込みが完了しました。")
        except Exception as e:
            self.queue_log(f"フォント取得エラー: {e}")

    def _register_reportlab_font(self) -> Tuple[str, str]:
        chosen = self.wm_font_combo.get()
        if chosen in self.font_map:
            p, idx = self.font_map[chosen]
            internal_name = f"WM_{chosen.replace(' ', '_')}_{idx}"
            try:
                if internal_name not in pdfmetrics.getRegisteredFontNames():
                    if p.lower().endswith(".ttc"):
                        pdfmetrics.registerFont(TTFont(internal_name, p, subfontIndex=idx))
                    else:
                        pdfmetrics.registerFont(TTFont(internal_name, p))
                return internal_name, f"使用フォント: {chosen}"
            except Exception as e:
                self.queue_log(f"フォント登録失敗: {e}")
        try:
            pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
            return "HeiseiKakuGo-W5", "使用フォント: 標準CIDフォント"
        except:
            return "Helvetica", "使用フォント: Helvetica (日本語不可)"

    # --- Core Processing ---
    def start_thread(self):
        if not self.files:
            messagebox.showwarning("警告", "ファイルがありません。")
            return
        self.save_config()
        self.processing = True
        self.cancel_flag.clear()
        self.btn_convert.config(state=tk.DISABLED)
        self.btn_cancel.config(state=tk.NORMAL)
        threading.Thread(target=self.main_process, daemon=True).start()

    def main_process(self):
        pythoncom.CoInitialize()
        # 一時ディレクトリ管理の改善
        with tempfile.TemporaryDirectory() as tmp_dir:
            try:
                temp_units = []
                self.queue_progress(max=len(self.files), progress=0, label="変換中...")

                for i, f in enumerate(self.files):
                    if self.cancel_flag.is_set():
                        break
                    self.queue_progress(progress=i + 1, label=f"変換中: {os.path.basename(f['path'])}")

                    if f["type"] == "Excel":
                        units = self.cv_excel_units(f, tmp_dir)
                        for u_path, s_name in units:
                            temp_units.append({"path": u_path, "orig": f, "sheet": s_name, "fseq": i + 1})
                    else:
                        tmp_p = os.path.join(tmp_dir, f"conv_{i}.pdf")
                        ok = False
                        if f["type"] == "Word":
                            ok = self.cv_word(f, tmp_p)
                        elif f["type"] == "PowerPoint":
                            ok = self.cv_ppt(f, tmp_p)
                        elif f["type"] == "Image":
                            ok = self.cv_img(f, tmp_p)
                        elif f["type"] == "PDF":
                            ok = self.cv_pdf(f, tmp_p)

                        if ok:
                            temp_units.append({"path": tmp_p, "orig": f, "sheet": "", "fseq": i + 1})

                if not temp_units or self.cancel_flag.is_set():
                    return

                # 最終PDF生成
                self.queue_progress(label="最終処理中...")
                global_seq = 1
                dest_to_open = ""

                if self.config.merge_all:
                    dest = self.get_final_dest(temp_units[0], 1, 1, 1)
                    if dest:
                        self.finalize_pdfs([u["path"] for u in temp_units], dest, temp_units)
                        dest_to_open = dest
                else:
                    for i, f_orig in enumerate(self.files):
                        if self.cancel_flag.is_set():
                            break
                        u_list = [u for u in temp_units if u["orig"] == f_orig]
                        if not u_list:
                            continue

                        t = f_orig["type"]
                        do_pg = (
                            (t == "Word" and self.config.split_word_page)
                            or (t == "PowerPoint" and self.config.split_ppt_page)
                            or (t == "PDF" and self.config.split_pdf_page)
                            or (t == "Excel" and self.config.split_excel_page)
                        )
                        do_sh = t == "Excel" and self.config.split_excel_sheet

                        if do_pg:
                            for u in u_list:
                                reader = PdfReader(u["path"])
                                p_total = len(reader.pages)
                                for p_idx in range(p_total):
                                    dest = self.get_final_dest(u, global_seq, i + 1, p_idx + 1, p_total)
                                    if dest:
                                        writer = PdfWriter()
                                        writer.add_page(reader.pages[p_idx])
                                        tmp_s = os.path.join(tmp_dir, "split.pdf")
                                        with open(tmp_s, "wb") as fs:
                                            writer.write(fs)
                                        self.finalize_pdfs([tmp_s], dest, [u], p_idx + 1, p_total)
                                        global_seq += 1
                                        dest_to_open = dest
                        elif do_sh:
                            for u in u_list:
                                dest = self.get_final_dest(u, global_seq, i + 1, 1)
                                if dest:
                                    self.finalize_pdfs([u["path"]], dest, [u])
                                    global_seq += 1
                                    dest_to_open = dest
                        else:
                            dest = self.get_final_dest(u_list[0], global_seq, i + 1, 1)
                            if dest:
                                self.finalize_pdfs([u["path"] for u in u_list], dest, u_list)
                                global_seq += 1
                                dest_to_open = dest

                self.finish_action(dest_to_open)

            except Exception as e:
                self.queue_log(f"致命的エラー: {e}")
            finally:
                self.processing = False
                self.root.after(0, lambda: self.btn_convert.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.btn_cancel.config(state=tk.DISABLED))
                pythoncom.CoUninitialize()

    # --- Office Converters (with COM cleanup & Excel Print Area) ---
    def cv_excel_units(self, f: dict, tmp_dir: str) -> List[Tuple[str, str]]:
        units = []
        excel = None
        wb = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(os.path.abspath(f["path"]), ReadOnly=True)

            target_sheets = [s.strip() for s in f.get("range", "").split(",") if s.strip()]
            if not target_sheets or target_sheets == ["全ページ"]:
                target_sheets = [s.Name for s in wb.Sheets]

            for name in target_sheets:
                try:
                    ws = wb.Worksheets(name)
                    if ws.Visible != -1:
                        continue

                    # 印刷設定
                    ps = ws.PageSetup
                    if self.config.excel_fit or self.config.excel_fit_tall:
                        ps.Zoom = False
                        if self.config.excel_fit:
                            ps.FitToPagesWide = 1
                        if self.config.excel_fit_tall:
                            ps.FitToPagesTall = 1

                    # 印刷範囲の処理: 標準機能として組込
                    # ws.ExportAsFixedFormat(0, ...) はデフォルトで PrintArea を尊重する
                    tmp_p = os.path.join(tmp_dir, f"ex_{len(units)}.pdf")
                    ws.ExportAsFixedFormat(0, tmp_p)
                    units.append((tmp_p, name))
                except:
                    continue
            return units
        finally:
            if wb:
                wb.Close(False)
                wb = None
            if excel:
                excel.Quit()
                excel = None

    def cv_word(self, f, out):
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            doc = word.Documents.Open(os.path.abspath(f["path"]), ReadOnly=True)
            doc.ExportAsFixedFormat(os.path.abspath(out), 17)
            return True
        except:
            return False
        finally:
            if doc:
                doc.Close(False)
                doc = None
            if word:
                word.Quit()
                word = None

    def cv_ppt(self, f, out):
        ppt = None
        pres = None
        try:
            ppt = win32com.client.DispatchEx("PowerPoint.Application")

            # ★PowerPointは「非表示(Visible=False)」が禁止の環境があるので触らないのが安全
            # ppt.Visible = 1  # ←必要なら True のみ（Falseは不可）

            abs_path = os.path.abspath(f["path"])
            abs_out = os.path.abspath(out)

            # Open(FileName, ReadOnly, Untitled, WithWindow)
            pres = ppt.Presentations.Open(abs_path, True, False, False)

            # ★ExportAsFixedFormat の既知回避（PrintRange=None）
            pres.ExportAsFixedFormat(abs_out, 2, PrintRange=None)

            return True

        except Exception as e:
            self.queue_log(f"PPT変換エラー ({os.path.basename(f['path'])}): {e}")
            return False

        finally:
            try:
                if pres:
                    pres.Close()
            finally:
                if ppt:
                    ppt.Quit()

    def cv_img(self, f, out):
        try:
            with Image.open(f["path"]) as img:
                img.convert("RGB").save(out, "PDF")
            return True
        except:
            return False

    def cv_pdf(self, f, out):
        try:
            reader = PdfReader(f["path"])
            writer = PdfWriter()
            for p in reader.pages:
                writer.add_page(p)
            with open(out, "wb") as fs:
                writer.write(fs)
            return True
        except:
            return False

    # --- Watermark & Finalize ---
    def finalize_pdfs(
        self, src_list: List[str], dest: str, units: List[dict], page_offset: int = 1, total_override: int = 0
    ):
        writer = PdfWriter()
        font_name, _ = self._register_reportlab_font()
        readers = [PdfReader(s) for s in src_list]
        total_p = total_override if total_override > 0 else sum(len(r.pages) for r in readers)
        curr_p = page_offset

        for r in readers:
            for page in r.pages:
                page.transfer_rotation_to_content()
                has_wm = any(getattr(self, f"wm{i}_pos_var").get() != "None" for i in [1, 2])
                has_pg = self.pg_en_var.get()

                if has_wm or has_pg:
                    w, h = float(page.mediabox.width), float(page.mediabox.height)
                    packet = io.BytesIO()
                    c = canvas.Canvas(packet, pagesize=(w, h))
                    # 透かし描画
                    for i in [1, 2]:
                        txt_raw = getattr(self, f"wm{i}_val").get()
                        pos = getattr(self, f"wm{i}_pos_var").get()
                        if pos == "None" or not txt_raw:
                            continue
                        txt = self.apply_tags(txt_raw, units[0], curr_p, 1, curr_p, total_p)
                        c.saveState()
                        c.setFont(font_name, int(self.wm_size_spin.get()))
                        rgb = [int(self.wm_color_btn.cget("bg").lstrip("#")[j : j + 2], 16) / 255 for j in (0, 2, 4)]
                        c.setFillColorRGB(*rgb, alpha=float(self.wm_alpha_scale.get()))
                        if pos == "中央斜め":
                            c.translate(w / 2, h / 2)
                            c.rotate(45)
                            c.drawCentredString(0, 0, txt)
                        else:
                            tw = c.stringWidth(txt, font_name, int(self.wm_size_spin.get()))
                            tx = (w - tw) / 2 if "中央" in pos else (20 if "左" in pos else w - tw - 20)
                            ty = (h - 40) if "上" in pos else (20 if "下" in pos else h / 2)
                            c.drawString(tx, ty, txt)
                        c.restoreState()
                    # ページ番号
                    if has_pg:
                        p_str = self.pg_fmt_var.get().replace("{n}", str(curr_p)).replace("{total}", str(total_p))
                        c.setFont(font_name, 10)
                        c.setFillColorRGB(0.3, 0.3, 0.3)
                        if self.pg_pos_var.get() == "中央下":
                            c.drawCentredString(w / 2, 20, p_str)
                        else:
                            c.drawRightString(w - 20, 20, p_str)
                    c.showPage()
                    c.save()
                    packet.seek(0)
                    page.merge_page(PdfReader(packet).pages[0])

                writer.add_page(page)
                curr_p += 1

        if self.meta_var.get():
            writer.add_metadata({})
        if self.pw_var.get():
            writer.encrypt(self.pw_var.get())
        if self.compress_var.get():
            # 新旧どちらのメソッド名でも動くようにチェック
            if hasattr(writer, "compress_contents"):
                writer.compress_contents()
            elif hasattr(writer, "compress_content_streams"):
                writer.compress_content_streams()

        with open(dest, "wb") as f:
            writer.write(f)

    # --- Preview Feature ---
    def preview_watermark(self):
        if not self.files:
            messagebox.showinfo("情報", "プレビューするファイルがありません。")
            return

        self.update_config_from_ui()
        f = self.files[0]

        def _task():
            pythoncom.CoInitialize()
            with tempfile.TemporaryDirectory() as tmp_dir:
                try:
                    self.queue_log(f"プレビュー生成中: {os.path.basename(f['path'])}")
                    tmp_pdf = os.path.join(tmp_dir, "preview_base.pdf")
                    # 最初の一部だけ取得するように工夫
                    if f["type"] == "Excel":
                        res = self.cv_excel_units(f, tmp_dir)
                        if not res:
                            return
                        tmp_pdf = res[0][0]
                    else:
                        ok = False
                        if f["type"] == "Word":
                            ok = self.cv_word(f, tmp_pdf)
                        elif f["type"] == "PowerPoint":
                            ok = self.cv_ppt(f, tmp_pdf)

                        # 以下のログ出力を追加
                        if ok:
                            temp_units.append({"path": tmp_p, "orig": f, "sheet": "", "fseq": i + 1})
                        else:
                            self.queue_log(f"❌ 変換に失敗しました: {os.path.basename(f['path'])}")

                        if not temp_units:  # 変換されたファイルがゼロの場合
                            self.queue_log("⚠️ 処理対象のファイルが生成されなかったため、終了します。")
                            return

                        elif f["type"] == "Image":
                            ok = self.cv_img(f, tmp_pdf)
                        elif f["type"] == "PDF":
                            ok = self.cv_pdf(f, tmp_pdf)
                        if not ok:
                            return

                    # 1ページ目だけに絞る
                    reader = PdfReader(tmp_pdf)
                    writer = PdfWriter()
                    writer.add_page(reader.pages[0])
                    tmp_one = os.path.join(tmp_dir, "one.pdf")
                    with open(tmp_one, "wb") as fs:
                        writer.write(fs)

                    # 透かし適用
                    out_p = os.path.join(tempfile.gettempdir(), "PDFPro_Preview.pdf")
                    self.finalize_pdfs([tmp_one], out_p, [{"orig": f, "sheet": "Preview"}], 1, 1)
                    os.startfile(out_p)
                    self.queue_log("プレビューを表示しました。")
                except Exception as e:
                    self.queue_log(f"プレビュー失敗: {e}")
                finally:
                    pythoncom.CoUninitialize()

        threading.Thread(target=_task, daemon=True).start()

    # --- Utils ---
    def apply_tags(self, tpl: str, u_info: dict, seq: int, fseq: int, pseq: int, ptotal: int = 1) -> str:
        now = datetime.datetime.now()
        path = u_info["orig"]["path"]
        name = os.path.splitext(os.path.basename(path))[0]
        sheet = u_info.get("sheet", "")
        parent = os.path.basename(os.path.dirname(path))  # 親フォルダ名

        res = tpl
        res = res.replace("{name}", name)
        res = res.replace("{sheet}", sheet)
        res = res.replace("{parent}", parent)
        res = res.replace("{seq}", str(seq))
        res = res.replace("{fseq}", str(fseq))
        res = res.replace("{pseq}", str(pseq))
        res = res.replace("{total}", str(len(self.files)))
        res = res.replace("{ptotal}", str(ptotal))
        res = res.replace("{username}", os.getlogin())
        res = res.replace("{rand}", f"{random.randint(0, 9999):04d}")

        # 日付・時刻タグの処理 {date:yyyy-mm-dd HH:MM:SS} 等
        def _repl(m):
            fmt = m.group(1)
            # Pythonのstrftime形式に変換
            fmt = fmt.replace("yyyy", "%Y").replace("mm", "%m").replace("dd", "%d")
            fmt = fmt.replace("HH", "%H").replace("MM", "%M").replace("SS", "%S")
            return now.strftime(fmt)

        res = re.sub(r"{date:(.*?)}", _repl, res)
        return res

    def show_naming_help(self):
        h = (
            "【利用可能なタグ】\n\n"
            "{name} : 元のファイル名\n"
            "{sheet} : Excelシート名\n"
            "{parent} : 親フォルダの名前\n"
            "{seq} : 全体の通し番号\n"
            "{fseq} : ファイル毎の番号\n"
            "{pseq} : ページ毎の番号\n"
            "{total} : 全ファイル数\n"
            "{ptotal} : ファイル内の総ページ数\n"
            "{username} : PCユーザー名\n"
            "{rand} : 4桁のランダム数字\n\n"
            "【日付・時刻】\n"
            "{date:yyyy-mm-dd} -> 2024-02-06\n"
            "{date:yyyy年mm月dd日 HH時MM分} \n"
            "※HH:時, MM:分, SS:秒"
        )
        messagebox.showinfo("命名ルールのタグ説明", h)

    def get_final_dest(self, u_info, seq, fseq, pseq, ptotal=1):
        out_name = self.apply_tags(self.naming_var.get(), u_info, seq, fseq, pseq, ptotal)
        base = (
            os.path.dirname(u_info["orig"]["path"]) if self.out_mode_var.get() == "original" else self.out_dir_var.get()
        )
        if not os.path.exists(base):
            os.makedirs(base, exist_ok=True)
        dest = os.path.join(base, re.sub(r'[\\/:*?"<>|]+', "_", out_name) + ".pdf")
        return self.confirm_overwrite_or_rename(dest)

    def confirm_overwrite_or_rename(self, dest):
        if not os.path.exists(dest):
            return dest
        ans = self._ui_call_sync(
            messagebox.askyesnocancel,
            "上書き確認",
            f"存在します: {os.path.basename(dest)}\nYes:上書き, No:連番, Cancel:中止",
        )
        if ans is None:
            return None
        if ans:
            return dest
        b, e = os.path.splitext(dest)
        i = 1
        while os.path.exists(f"{b}_{i}{e}"):
            i += 1
        return f"{b}_{i}{e}"

    def _ui_call_sync(self, func, *args, **kwargs):
        ev = threading.Event()
        slot = {"res": None}
        self.root.after(0, lambda: [slot.update({"res": func(*args, **kwargs)}), ev.set()])
        ev.wait()
        return slot["res"]

    def queue_log(self, msg: str):
        self.progress_queue.put({"log": msg})

    def queue_progress(self, **kwargs):
        self.progress_queue.put(kwargs)

    def check_progress_queue(self):
        try:
            while True:
                m = self.progress_queue.get_nowait()
                if "log" in m:
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.insert(tk.END, f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {m['log']}\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                if "progress" in m:
                    self.progress["value"] = m["progress"]
                if "max" in m:
                    self.progress["maximum"] = m["max"]
                if "label" in m:
                    self.progress_label.config(text=m["label"])
        except queue.Empty:
            pass
        self.root.after(100, self.check_progress_queue)

    def finish_action(self, path):
        self.queue_log("✅ すべての処理が完了しました。")
        if self.open_var.get() and path and os.path.exists(path):
            os.startfile(path)
        if self.folder_var.get() and path:
            os.startfile(os.path.dirname(path))
        if self.clear_after_var.get():
            self.root.after(0, self.clear_list)

    # --- File Handlers ---
    def handle_drop(self, event):
        paths = self.root.tk.splitlist(event.data)
        threading.Thread(target=self.add_files_worker, args=(paths,), daemon=True).start()

    def add_files_worker(self, paths):
        for p in paths:
            ext = os.path.splitext(p)[1].lower()  # 拡張子を小文字で取得
            if ext in [".docx", ".doc", ".xlsx", ".xls", ".xlsm", ".pdf", ".pptx", ".ppt", ".jpg", ".png", ".jpeg"]:
                if any(f["path"] == p for f in self.files):
                    continue

                # 文字列検索ではなく、拡張子で正確に判定する
                if ext in [".docx", ".doc"]:
                    t = "Word"
                elif ext in [".xlsx", ".xls", ".xlsm"]:
                    t = "Excel"
                elif ext in [".pptx", ".ppt"]:
                    t = "PowerPoint"
                elif ext in [".pdf"]:
                    t = "PDF"
                else:
                    t = "Image"

                info = {"path": p, "type": t, "range": "全ページ", "sheets": []}
                if t == "Excel":
                    info["sheets"] = self.get_excel_sheets(p)
                self.files.append(info)
        self.root.after(0, self.update_tree)

    def get_excel_sheets(self, p):
        try:
            pythoncom.CoInitialize()
            ex = win32com.client.DispatchEx("Excel.Application")
            wb = ex.Workbooks.Open(os.path.abspath(p), ReadOnly=True)
            names = [s.Name for s in wb.Sheets]
            wb.Close(False)
            ex.Quit()
            return names
        except:
            return []

    def update_tree(self):
        self.tree.delete(*self.tree.get_children())
        for f in self.files:
            idx = self.files.index(f) + 1
            u = {"orig": f, "sheet": f["sheets"][0] if f["sheets"] else "", "fseq": idx}
            prev = f"{self.apply_tags(self.naming_var.get(), u, idx, idx, 1)}.pdf"
            self.tree.insert("", tk.END, values=(f["type"], os.path.basename(f["path"]), f["range"], prev))

    def update_output_preview(self):
        if not self.files:
            self.final_name_label.config(text="")
            return
        f0 = self.files[0]
        u0 = {"orig": f0, "sheet": f0["sheets"][0] if f0["sheets"] else "Sheet1", "fseq": 1}
        p_name = self.apply_tags(self.naming_var.get(), u0, 1, 1, 1)
        self.final_name_label.config(text=f"出力例: {p_name}.pdf")
        self.update_tree()

    def load_preset(self):
        name = self.preset_combo.get()
        if name in self.presets:
            self.config = AppConfig(**self.presets[name])
            self.apply_config_to_ui()
            self.queue_log(f"プリセット '{name}' を読込。")

    def save_preset(self):
        import tkinter.simpledialog as sd

        name = sd.askstring("保存", "プリセット名:")
        if name:
            self.update_config_from_ui()
            self.presets[name] = asdict(self.config)
            self.preset_combo["values"] = list(self.presets.keys())
            self.save_config()

    def delete_preset(self):
        name = self.preset_combo.get()
        if name in self.presets and messagebox.askyesno("確認", "削除しますか？"):
            del self.presets[name]
            self.preset_combo["values"] = list(self.presets.keys())
            self.save_config()

    def choose_color(self):
        c = colorchooser.askcolor()[1]
        if c:
            self.wm_color_btn.config(bg=c)

    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.out_dir_var.set(d)

    def move_file(self, d):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        n = idx + d
        if 0 <= n < len(self.files):
            self.files[idx], self.files[n] = self.files[n], self.files[idx]
            self.update_output_preview()

    def remove_file(self):
        for s in self.tree.selection():
            del self.files[self.tree.index(s)]
        self.update_output_preview()

    def clear_list(self):
        self.files = []
        self.update_output_preview()

    def add_folder(self):
        d = filedialog.askdirectory()
        if d:
            self.add_files_worker([os.path.join(d, f) for f in os.listdir(d)])

    def on_list_double_click(self, e):
        item = self.tree.identify_row(e.y)
        if not item:
            return
        idx = self.tree.index(item)
        f = self.files[idx]
        win = tk.Toplevel(self.root)
        win.title("範囲編集")
        win.geometry("300x400")
        if f["type"] == "Excel":
            lb = tk.Listbox(win, selectmode=tk.MULTIPLE)
            lb.pack(fill=tk.BOTH, expand=True)
            for s in f["sheets"]:
                lb.insert(tk.END, s)
            tk.Button(
                win,
                text="保存",
                command=lambda: [
                    f.update({"range": ",".join([lb.get(i) for i in lb.curselection()])}),
                    self.update_output_preview(),
                    win.destroy(),
                ],
            ).pack()
        else:
            ent = tk.Entry(win)
            ent.insert(0, f["range"])
            ent.pack(pady=20)
            tk.Button(
                win,
                text="保存",
                command=lambda: [f.update({"range": ent.get()}), self.update_output_preview(), win.destroy()],
            ).pack()

    def cancel_process(self):
        self.cancel_flag.set()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFUltimateApp(root)
    root.mainloop()
