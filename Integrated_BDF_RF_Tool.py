#!/usr/bin/env python3
"""
Integrated BDF Tool v14.0
=========================
Refactored BDF Tool - Tab 3 (Offset) merged into Tab 2, Tab 4 (RF Check) removed.

Tab 1: BDF Merge Preparation
Tab 2: BDF Post-Process (with integrated offset calculation & application)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import threading
import csv
import shutil
import tempfile
import pandas as pd
from pyNastran.bdf.bdf import BDF
from pyNastran.op2.op2 import OP2
import numpy as np


class IntegratedBDFRFTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Integrated BDF Tool v14.0")
        self.root.geometry("1100x950")
        
        # Tab 1 variables
        self.thermal_bdfs = []
        self.maneuver_bdfs = []
        self.excel_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_thermal_name = tk.StringVar(value="merged_thermal.bdf")
        self.output_maneuver_name = tk.StringVar(value="merged_maneuver.bdf")
        self.set_id = tk.StringVar(value="99")
        self.temp_initial = tk.StringVar(value="10")
        
        # Tab 2 variables
        self.run_bdfs = []
        self.property_excel_path = tk.StringVar()
        self.nastran_path = tk.StringVar()
        self.run_output_folder = tk.StringVar()
        self.csv_output_name = tk.StringVar(value="bar_stress_results.csv")
        self.combined_csv_name = tk.StringVar(value="combined_stress_results.csv")
        self.tab2_offset_csv = tk.StringVar()  # Offset CSV for Tab 2
        
        self.bar_properties = {}
        self.skin_properties = {}
        self.residual_strength_df = None
        
        # Tab 2 Offset variables (merged from Tab 3)
        self.offset_element_excel = tk.StringVar()
        self.offset_csv_name = tk.StringVar(value="calculated_offsets.csv")

        self.setup_ui()
    
    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="BDF Merge Preparation")
        self.setup_tab1()
        
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="BDF Post-Process")
        self.setup_tab2()
    
    def setup_tab1(self):
        main = ttk.Frame(self.tab1, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main, text="BDF Merge Preparation v8", font=('Helvetica', 14, 'bold')).pack(pady=(0,10))
        
        # === THERMAL SECTION ===
        thermal_main = ttk.LabelFrame(main, text="THERMAL", padding="5")
        thermal_main.pack(fill=tk.X, pady=5)
        
        thm_master_f = ttk.Frame(thermal_main)
        thm_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(thm_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(thm_master_f, text="Add...", command=self.add_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(thm_master_f, text="Clear", command=self.clear_thermal_bdfs).pack(side=tk.LEFT, padx=2)
        self.thermal_count = tk.StringVar(value="0 files")
        ttk.Label(thm_master_f, textvariable=self.thermal_count).pack(side=tk.LEFT, padx=5)
        
        self.thermal_listbox = tk.Listbox(thermal_main, height=3, width=100)
        self.thermal_listbox.pack(fill=tk.X, pady=2)
        
        # === MANEUVER SECTION ===
        maneuver_main = ttk.LabelFrame(main, text="MANEUVER", padding="5")
        maneuver_main.pack(fill=tk.X, pady=5)
        
        man_master_f = ttk.Frame(maneuver_main)
        man_master_f.pack(fill=tk.X, pady=2)
        ttk.Label(man_master_f, text="MASTER BDFs:", width=12).pack(side=tk.LEFT)
        ttk.Button(man_master_f, text="Add...", command=self.add_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        ttk.Button(man_master_f, text="Clear", command=self.clear_maneuver_bdfs).pack(side=tk.LEFT, padx=2)
        self.maneuver_count = tk.StringVar(value="0 files")
        ttk.Label(man_master_f, textvariable=self.maneuver_count).pack(side=tk.LEFT, padx=5)
        
        self.maneuver_listbox = tk.Listbox(maneuver_main, height=3, width=100)
        self.maneuver_listbox.pack(fill=tk.X, pady=2)
        
        # === SETTINGS ===
        sf = ttk.LabelFrame(main, text="Settings", padding="10")
        sf.pack(fill=tk.X, pady=5)
        ttk.Label(sf, text="Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.excel_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_excel).grid(row=0, column=2)
        ttk.Label(sf, text="Output:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.output_folder, width=70).grid(row=1, column=1, padx=5)
        ttk.Button(sf, text="Browse", command=self.browse_output).grid(row=1, column=2)
        ttk.Label(sf, text="SET ID:").grid(row=2, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.set_id, width=10).grid(row=2, column=1, sticky=tk.W, padx=5)
        ttk.Label(sf, text="TEMP(INIT):").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(sf, textvariable=self.temp_initial, width=10).grid(row=3, column=1, sticky=tk.W, padx=5)
        
        bf = ttk.Frame(main)
        bf.pack(fill=tk.X, pady=10)
        self.process_btn = ttk.Button(bf, text=">>> PROCESS & MERGE <<<", command=self.start_processing)
        self.process_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Clear Log", command=self.clear_log1).pack(side=tk.LEFT)
        
        self.progress1 = ttk.Progressbar(main, mode='indeterminate')
        self.progress1.pack(fill=tk.X, pady=5)
        
        lf = ttk.LabelFrame(main, text="Log", padding="10")
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text1 = scrolledtext.ScrolledText(lf, height=15)
        self.log_text1.pack(fill=tk.BOTH, expand=True)
    
    def setup_tab2(self):
        main = ttk.Frame(self.tab2, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main, text="BDF Post-Process", font=('Helvetica', 14, 'bold')).pack(pady=(0,10))
        
        bf = ttk.LabelFrame(main, text="BDF Files", padding="10")
        bf.pack(fill=tk.X, pady=5)
        bb = ttk.Frame(bf)
        bb.pack(fill=tk.X)
        ttk.Button(bb, text="Add...", command=self.add_run_bdfs).pack(side=tk.LEFT, padx=5)
        ttk.Button(bb, text="Clear", command=self.clear_run_bdfs).pack(side=tk.LEFT)
        self.run_listbox = tk.Listbox(bf, height=3, width=100)
        self.run_listbox.pack(fill=tk.X, pady=5)
        self.run_count = tk.StringVar(value="0 files")
        ttk.Label(bf, textvariable=self.run_count).pack(anchor=tk.W)
        
        pf = ttk.LabelFrame(main, text="Property Excel", padding="10")
        pf.pack(fill=tk.X, pady=5)
        ttk.Label(pf, text="Excel:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(pf, textvariable=self.property_excel_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(pf, text="Browse", command=self.browse_property_excel).grid(row=0, column=2)
        ttk.Button(pf, text="Load Properties", command=self.load_properties).grid(row=1, column=0, pady=5)
        
        pvf = ttk.Frame(pf)
        pvf.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=5)
        self.bar_prop_text = tk.Text(pvf, height=2, width=35)
        self.bar_prop_text.pack(side=tk.LEFT, padx=3)
        self.bar_prop_text.insert(tk.END, "Bar: Not loaded")
        self.skin_prop_text = tk.Text(pvf, height=2, width=35)
        self.skin_prop_text.pack(side=tk.LEFT, padx=3)
        self.skin_prop_text.insert(tk.END, "Skin: Not loaded")
        self.resid_text = tk.Text(pvf, height=2, width=30)
        self.resid_text.pack(side=tk.LEFT, padx=3)
        self.resid_text.insert(tk.END, "Residual: Not loaded")
        
        nf = ttk.LabelFrame(main, text="Nastran", padding="10")
        nf.pack(fill=tk.X, pady=5)
        ttk.Label(nf, text="Path:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(nf, textvariable=self.nastran_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(nf, text="Browse", command=self.browse_nastran).grid(row=0, column=2)
        
        of = ttk.LabelFrame(main, text="Output", padding="10")
        of.pack(fill=tk.X, pady=5)
        ttk.Label(of, text="Folder:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(of, textvariable=self.run_output_folder, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(of, text="Browse", command=self.browse_run_output).grid(row=0, column=2)
        ttk.Label(of, text="CSV:").grid(row=1, column=0, sticky=tk.W)
        ttk.Entry(of, textvariable=self.csv_output_name, width=25).grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # === OFFSET SECTION (merged from Tab 3) ===
        off = ttk.LabelFrame(main, text="Offset Calculation & Application", padding="10")
        off.pack(fill=tk.X, pady=5)

        ttk.Label(off, text="Element Excel:").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Entry(off, textvariable=self.offset_element_excel, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(off, text="Browse", command=self.browse_offset_element_excel).grid(row=0, column=2, padx=5)
        ttk.Label(off, text="Sheets: 'Landing_Offset' (Col A), 'Bar_Offset' (Col A)",
                 font=('Helvetica', 8, 'italic')).grid(row=0, column=3, padx=5)

        self.btn_calc_offset = ttk.Button(off, text=">>> CALCULATE OFFSETS <<<",
                                          command=self.start_calculate_offsets)
        self.btn_calc_offset.grid(row=1, column=0, columnspan=2, pady=5)

        ttk.Label(off, text="Offset CSV:").grid(row=2, column=0, sticky=tk.W, padx=5)
        ttk.Entry(off, textvariable=self.tab2_offset_csv, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(off, text="Browse", command=self.browse_tab2_offset_csv).grid(row=2, column=2, padx=5)

        self.btn_apply_offset_tab2 = ttk.Button(off, text=">>> APPLY OFFSETS TO BDFs <<<",
                                                command=self.start_apply_offsets_tab2)
        self.btn_apply_offset_tab2.grid(row=3, column=0, columnspan=2, pady=5)
        
        af = ttk.Frame(main)
        af.pack(fill=tk.X, pady=10)
        self.btn1 = ttk.Button(af, text="1.Update Props", command=self.start_update_properties, width=14)
        self.btn1.pack(side=tk.LEFT, padx=2)
        self.btn2 = ttk.Button(af, text="2.Run Nastran", command=self.start_run_nastran, width=14)
        self.btn2.pack(side=tk.LEFT, padx=2)
        self.btn3 = ttk.Button(af, text="3.Post-Process", command=self.start_postprocess, width=14)
        self.btn3.pack(side=tk.LEFT, padx=2)
        self.btn4 = ttk.Button(af, text="4.Combine", command=self.start_combine_stress, width=12)
        self.btn4.pack(side=tk.LEFT, padx=2)
        self.btn_full = ttk.Button(af, text=">>> FULL <<<", command=self.start_full_run, width=12)
        self.btn_full.pack(side=tk.LEFT, padx=2)
        ttk.Button(af, text="Clear", command=self.clear_log2).pack(side=tk.LEFT, padx=2)
        
        self.progress2 = ttk.Progressbar(main, mode='indeterminate')
        self.progress2.pack(fill=tk.X, pady=5)
        
        lf = ttk.LabelFrame(main, text="Log", padding="10")
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_text2 = scrolledtext.ScrolledText(lf, height=12)
        self.log_text2.pack(fill=tk.BOTH, expand=True)

    # ============= TAB 1 HELPERS =============
    def add_thermal_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.thermal_bdfs:
                self.thermal_bdfs.append(f)
                self.thermal_listbox.insert(tk.END, f)
        self.thermal_count.set(f"{len(self.thermal_bdfs)} files")
    
    def clear_thermal_bdfs(self):
        self.thermal_bdfs.clear()
        self.thermal_listbox.delete(0, tk.END)
        self.thermal_count.set("0 files")
    
    def add_maneuver_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.maneuver_bdfs:
                self.maneuver_bdfs.append(f)
                self.maneuver_listbox.insert(tk.END, f)
        self.maneuver_count.set(f"{len(self.maneuver_bdfs)} files")
    
    def clear_maneuver_bdfs(self):
        self.maneuver_bdfs.clear()
        self.maneuver_listbox.delete(0, tk.END)
        self.maneuver_count.set("0 files")
    
    def browse_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f: self.excel_path.set(f)
    
    def browse_output(self):
        f = filedialog.askdirectory()
        if f: self.output_folder.set(f)
    
    def log1(self, msg):
        self.log_text1.insert(tk.END, msg + "\n")
        self.log_text1.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log1(self):
        self.log_text1.delete(1.0, tk.END)
    
    def format_include_nastran(self, abs_path):
        """Uzun INCLUDE path'lerini Nastran formatına uygun böler."""
        include_line = f"INCLUDE '{abs_path}'"
        if len(include_line) <= 72:
            return [include_line]
        parts = abs_path.split('/')
        lines = []
        current_line = "INCLUDE '"
        for i, part in enumerate(parts):
            is_last = (i == len(parts) - 1)
            segment = part if is_last else part + '/'
            if len(current_line + segment) <= 72:
                current_line += segment
            else:
                if current_line != "INCLUDE '":
                    lines.append(current_line)
                current_line = segment
        if current_line:
            current_line += "'"
            lines.append(current_line)
        return lines
    
    def read_file_safe(self, fpath):
        """Dosyayı güvenli şekilde oku."""
        for enc in ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']:
            try:
                with open(fpath, 'r', encoding=enc, errors='replace') as f:
                    return f.read()
            except:
                continue
        return ""
    
    def extract_subcase_load_info(self, bdf_content):
        """
        BDF içeriğinden TÜM SUBCASE bilgilerini çıkarır.
        Bir dosyada birden fazla SUBCASE olabilir (MASTER_GUST.BDF gibi).
        
        Returns: list of dicts, her biri:
            - subcase_id
            - load_id
            - temp_load_id
            - subtitle
        """
        results = []
        lines = bdf_content.split('\n')
        
        current_subcase = None
        current_load = None
        current_temp_load = None
        current_subtitle = None
        
        for line in lines:
            upper = line.upper().strip()
            original = line.strip()
            
            # SUBCASE satırı
            if upper.startswith('SUBCASE'):
                # Önceki subcase'i kaydet
                if current_subcase is not None:
                    results.append({
                        'subcase_id': current_subcase,
                        'load_id': current_load,
                        'temp_load_id': current_temp_load,
                        'subtitle': current_subtitle
                    })
                
                # Yeni subcase başlat
                parts = upper.split()
                if len(parts) >= 2:
                    try:
                        current_subcase = int(parts[1])
                        current_load = None
                        current_temp_load = None
                        current_subtitle = None
                    except:
                        pass
            
            # LOAD = satırı
            elif current_subcase and upper.startswith('LOAD') and '=' in upper:
                match = re.search(r'LOAD\s*=\s*(\d+)', upper)
                if match:
                    current_load = int(match.group(1))
            
            # TEMPERATURE(LOAD) = satırı
            elif current_subcase and 'TEMPERATURE' in upper and 'LOAD' in upper and '=' in upper:
                match = re.search(r'TEMPERATURE\s*\(\s*LOAD\s*\)\s*=\s*(\d+)', upper)
                if match:
                    current_temp_load = int(match.group(1))
            
            # SUBTITLE satırı
            elif current_subcase and upper.startswith('SUBTITLE'):
                m = re.search(r'SUBTITLE\s+(.+)', original, re.IGNORECASE)
                if m:
                    current_subtitle = m.group(1).strip()
            
            # BEGIN BULK'a ulaştıysak case control bitti
            elif upper.startswith('BEGIN') and 'BULK' in upper:
                break
        
        # Son subcase'i kaydet
        if current_subcase is not None:
            results.append({
                'subcase_id': current_subcase,
                'load_id': current_load,
                'temp_load_id': current_temp_load,
                'subtitle': current_subtitle
            })
        
        return results
    
    def parse_multiline_includes(self, content, bdf_dir):
        """
        Çok satırlı INCLUDE'ları parse eder.
        Nastran formatında INCLUDE şöyle olabilir:
        
        INCLUDE '../../../_COMMON_/INTERFACE/STRUCTURAL/INTER_AF_HT/
        INTER_AF_HT_STRU.BDF'
        
        Returns: list of dict with keys: lines, full_text, abs_path, start_idx, end_idx
        """
        lines = content.split('\n')
        includes = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            upper = line.upper().strip()
            
            if upper.startswith('INCLUDE'):
                # INCLUDE başladı - tırnak içindeki path'i bul
                include_lines = [line]
                full_text = line
                
                # Tırnak sayısını kontrol et - tek tırnak veya çift tırnak
                quote_char = None
                if "'" in line:
                    quote_char = "'"
                elif '"' in line:
                    quote_char = '"'
                
                if quote_char:
                    # Tırnak sayısını say
                    quote_count = full_text.count(quote_char)
                    
                    # Eğer tek tırnak varsa (açılmış ama kapanmamış), devam satırlarını oku
                    j = i + 1
                    while quote_count % 2 != 0 and j < len(lines):
                        next_line = lines[j]
                        include_lines.append(next_line)
                        full_text += '\n' + next_line
                        quote_count = full_text.count(quote_char)
                        j += 1
                    
                    # Path'i çıkar - newline'ları temizle
                    clean_text = full_text.replace('\n', '')
                    match = re.search(rf"INCLUDE\s*{quote_char}([^{quote_char}]*){quote_char}", 
                                     clean_text, re.IGNORECASE)
                    if match:
                        inc_path = match.group(1).strip()
                        # Absolute path'e çevir
                        if not os.path.isabs(inc_path):
                            abs_path = os.path.normpath(os.path.join(bdf_dir, inc_path))
                        else:
                            abs_path = os.path.normpath(inc_path)
                        abs_path = abs_path.replace('\\', '/')
                        
                        includes.append({
                            'lines': include_lines,
                            'full_text': full_text,
                            'abs_path': abs_path,
                            'start_idx': i,
                            'end_idx': j - 1 if j > i + 1 else i
                        })
                    
                    i = j
                else:
                    i += 1
            else:
                i += 1
        
        return includes
    
    def collect_all_lines_from_masters(self, bdf_files, load_case_set, common_type):
        """
        Tüm MASTER BDF'lerden TÜM SATIRLARI toplar.
        
        1. Excel'deki SUBCASE ID'lere uyan MASTER BDF'leri bulur
        2. Her birinden TÜM satırları alır ve alt alta yapıştırır
        3. INCLUDE'ları kategorize et:
           - COMMON LOAD/THERMAL → Ayrı tut (sonra INCLUDE olarak eklenecek)
           - Diğerleri (STRUCTURE, INTERFACE, vs.) → Satır olarak ekle (pyNastran açacak)
        4. Duplicate satırları çıkar
        
        NOT: Bir BDF birden fazla SUBCASE içerebilir (MASTER_GUST.BDF gibi)
        
        common_type: 'LOAD' veya 'THERMAL'
        
        Returns: (all_lines, common_includes, subcase_info_map)
        """
        all_lines_raw = []  # Tüm satırlar (INCLUDE satırları dahil - COMMON hariç)
        common_includes_raw = []  # COMMON INCLUDE path'leri
        subcase_info_map = {}
        processed_files = set()  # Aynı dosyayı birden fazla kez işlememek için
        
        self.log1(f"    Reading ALL lines from {len(bdf_files)} MASTER BDFs...")
        matched_count = 0
        matched_subcases = 0
        
        for bdf_path in bdf_files:
            bdf_dir = os.path.dirname(os.path.abspath(bdf_path))
            content = self.read_file_safe(bdf_path)
            
            # INREL dosyalarını atla
            bdf_basename = os.path.basename(bdf_path).upper()
            if 'INREL' in bdf_basename:
                self.log1(f"      SKIP (INREL): {os.path.basename(bdf_path)}")
                continue
            
            # TÜM SUBCASE bilgilerini al (birden fazla olabilir)
            all_subcases = self.extract_subcase_load_info(content)
            
            # Bu dosyadaki hangi subcase'ler Excel listesinde?
            matching_subcases = []
            for sc_info in all_subcases:
                sc_id = sc_info['subcase_id']
                if sc_id and sc_id in load_case_set:
                    matching_subcases.append(sc_info)
            
            # Eşleşen subcase varsa bu dosyayı işle
            if matching_subcases:
                # Dosya daha önce işlendiyse sadece subcase info'ları ekle
                if bdf_path in processed_files:
                    for sc_info in matching_subcases:
                        sc_id = sc_info['subcase_id']
                        if sc_id not in subcase_info_map:
                            if common_type == 'THERMAL':
                                subcase_info_map[sc_id] = {
                                    'temp_load_id': sc_info['temp_load_id'] if sc_info['temp_load_id'] else sc_id,
                                    'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Thermal Case {sc_id}"
                                }
                            else:
                                subcase_info_map[sc_id] = {
                                    'load_id': sc_info['load_id'] if sc_info['load_id'] else sc_id,
                                    'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Manoeuvre {sc_id}"
                                }
                            matched_subcases += 1
                    continue
                
                processed_files.add(bdf_path)
                matched_count += 1
                
                # Log - kaç subcase eşleşti
                sc_ids = [str(sc['subcase_id']) for sc in matching_subcases]
                self.log1(f"      MATCH: {os.path.basename(bdf_path)} ({len(matching_subcases)} subcases: {', '.join(sc_ids[:5])}{'...' if len(sc_ids) > 5 else ''})")
                
                # Subcase info'ları kaydet
                for sc_info in matching_subcases:
                    sc_id = sc_info['subcase_id']
                    matched_subcases += 1
                    if common_type == 'THERMAL':
                        subcase_info_map[sc_id] = {
                            'temp_load_id': sc_info['temp_load_id'] if sc_info['temp_load_id'] else sc_id,
                            'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Thermal Case {sc_id}"
                        }
                    else:
                        subcase_info_map[sc_id] = {
                            'load_id': sc_info['load_id'] if sc_info['load_id'] else sc_id,
                            'subtitle': sc_info['subtitle'] if sc_info['subtitle'] else f"Manoeuvre {sc_id}"
                        }
                
                # Önce tüm INCLUDE'ları parse et (çok satırlı dahil)
                all_includes = self.parse_multiline_includes(content, bdf_dir)
                
                # INCLUDE'ları kategorize et
                common_include_indices = set()  # COMMON INCLUDE satır indeksleri
                structure_include_count = 0
                common_include_count = 0
                
                for inc in all_includes:
                    abs_path_upper = inc['abs_path'].upper()
                    
                    # Bu INCLUDE COMMON LOAD/THERMAL mı?
                    is_common = False
                    if common_type == 'LOAD':
                        if '_COMMON_/LOAD' in abs_path_upper or '/COMMON/LOAD' in abs_path_upper:
                            is_common = True
                    elif common_type == 'THERMAL':
                        if '_COMMON_/THERMAL' in abs_path_upper or '/COMMON/THERMAL' in abs_path_upper:
                            is_common = True
                    
                    if is_common:
                        # COMMON INCLUDE - path'i kaydet, satırları atla
                        common_includes_raw.append(inc['abs_path'])
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        common_include_count += 1
                    else:
                        # Structure/Interface/vs INCLUDE - tek satır INCLUDE olarak ekle (pyNastran açacak)
                        # Absolute path ile yeni INCLUDE satırı oluştur
                        include_line = f"INCLUDE '{inc['abs_path']}'"
                        all_lines_raw.append(include_line)
                        # Orijinal satırları atla (common_include_indices'e ekle)
                        for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                            common_include_indices.add(idx)
                        structure_include_count += 1
                
                # TÜM SATIRLARI oku (INCLUDE satırları hariç - hem COMMON hem diğerleri)
                lines = content.split('\n')
                line_count = 0
                
                for idx, line in enumerate(lines):
                    # INCLUDE satırı mı? Atla (zaten yukarıda işledik)
                    if idx in common_include_indices:
                        continue
                    
                    stripped = line.strip()
                    if not stripped:
                        continue
                    
                    # Case control satırlarını atla
                    upper = stripped.upper()
                    skip_keywords = ['SOL ', 'SOL\t', 'CEND', 'TITLE', 'SUBTITLE', 'ECHO', 
                                    'SUBCASE', 'LOAD =', 'LOAD=', 'SPC =', 'SPC=', 
                                    'TEMPERATURE', 'DISPLACEMENT', 'FORCE', 'GPFORCE', 
                                    'OLOAD', 'SPCFORCE', 'SET ', 'BEGIN BULK', 
                                    'BEGIN,BULK', 'ENDDATA']
                    is_skip = any(upper.startswith(kw) for kw in skip_keywords)
                    
                    if not is_skip:
                        all_lines_raw.append(line)  # Orijinal satırı koru
                        line_count += 1
                
                self.log1(f"        -> {line_count} data lines, {structure_include_count} structure includes, {common_include_count} common includes")
        
        self.log1(f"    Matched {matched_count} MASTER BDFs with {matched_subcases} total subcases")
        self.log1(f"    Total raw lines (including structure includes): {len(all_lines_raw)}")
        self.log1(f"    Total raw COMMON includes: {len(common_includes_raw)}")
        
        # Duplicate satırları çıkar
        self.log1("    Removing duplicate lines...")
        seen_lines = set()
        unique_lines = []
        for line in all_lines_raw:
            # Normalize et (boşlukları düzenle)
            normalized = ' '.join(line.split())
            if normalized not in seen_lines:
                seen_lines.add(normalized)
                unique_lines.append(line)
        
        # Duplicate COMMON INCLUDE'ları çıkar
        seen_includes = set()
        unique_common_includes = []
        for inc_path in common_includes_raw:
            normalized = inc_path.lower().replace('\\', '/')
            if normalized not in seen_includes:
                seen_includes.add(normalized)
                unique_common_includes.append(inc_path)
        
        self.log1(f"    After removing duplicates:")
        self.log1(f"      Unique lines: {len(unique_lines)} (removed {len(all_lines_raw) - len(unique_lines)} duplicates)")
        self.log1(f"      Unique COMMON includes: {len(unique_common_includes)}")
        
        return unique_lines, unique_common_includes, subcase_info_map
    
    def extract_param_cards(self, bulk_data):
        """
        Bulk data'dan PARAM kartlarını ayırır.
        Returns: (param_lines, remaining_bulk_data)
        """
        lines = bulk_data.split('\n')
        param_lines = []
        other_lines = []
        
        seen_params = set()  # Duplicate PARAM kontrolü
        
        i = 0
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            upper = stripped.upper()
            
            if upper.startswith('PARAM'):
                # PARAM kartı - continuation satırlarını da al
                param_block = [line]
                
                # PARAM name'ini çıkar (duplicate kontrolü için)
                param_name = None
                try:
                    if ',' in line:
                        parts = line.split(',')
                        if len(parts) > 1:
                            param_name = parts[1].strip().upper()
                    else:
                        if len(line) >= 16:
                            param_name = line[8:16].strip().upper()
                except:
                    pass
                
                i += 1
                # Continuation satırlarını kontrol et
                while i < len(lines):
                    next_line = lines[i]
                    next_stripped = next_line.strip()
                    is_cont = (next_line.startswith('+') or 
                              next_line.startswith('*') or
                              (next_line.startswith('        ') and next_stripped and 
                               not next_stripped.startswith('$') and
                               not any(next_stripped.upper().startswith(card) for card in 
                                      ['PARAM', 'GRID', 'CBAR', 'CBEAM', 'CQUAD', 'CTRIA', 
                                       'MAT', 'PBAR', 'PSHELL', 'FORCE', 'MOMENT', 'RBE',
                                       'CORD', 'SPC', 'MPC', 'INCLUDE', 'ENDDATA'])))
                    if is_cont and next_stripped:
                        param_block.append(next_line)
                        i += 1
                    else:
                        break
                
                # Duplicate kontrolü
                if param_name and param_name not in seen_params:
                    seen_params.add(param_name)
                    param_lines.extend(param_block)
            else:
                other_lines.append(line)
                i += 1
        
        return param_lines, '\n'.join(other_lines)
    
    def check_and_remove_duplicates(self, bulk_data):
        """
        Bulk data içindeki duplicate kartları tespit edip kaldırır.
        
        - Element/Property/Material kartları: ID bazlı kontrol (aynı ID → duplicate)
        - SPC/FORCE/MOMENT kartları: Tüm satır bazlı kontrol (birebir aynıysa → duplicate)
        """
        self.log1("    Checking for duplicate entries...")
        
        lines = bulk_data.split('\n')
        
        # ID bazlı kontrol yapılacak kartlar
        id_based_cards = {
            'GRID': set(),
            'CBAR': set(),
            'CBEAM': set(),
            'CROD': set(),
            'CONROD': set(),
            'CQUAD4': set(),
            'CQUAD8': set(),
            'CTRIA3': set(),
            'CTRIA6': set(),
            'CHEXA': set(),
            'CPENTA': set(),
            'CTETRA': set(),
            'CBUSH': set(),
            'CELAS1': set(),
            'CELAS2': set(),
            'CDAMP1': set(),
            'CDAMP2': set(),
            'CMASS1': set(),
            'CMASS2': set(),
            'RBE2': set(),
            'RBE3': set(),
            'PBAR': set(),
            'PBARL': set(),
            'PBEAM': set(),
            'PBEAML': set(),
            'PROD': set(),
            'PSHELL': set(),
            'PCOMP': set(),
            'PCOMPG': set(),
            'PSOLID': set(),
            'PBUSH': set(),
            'PELAS': set(),
            'PDAMP': set(),
            'PMASS': set(),
            'PTUBE': set(),
            'PVISC': set(),
            'PGAP': set(),
            'PWELD': set(),
            'MAT1': set(),
            'MAT2': set(),
            'MAT8': set(),
            'MAT9': set(),
            'MATS1': set(),
            'CORD1R': set(),
            'CORD2R': set(),
            'CORD1C': set(),
            'CORD2C': set(),
            'CORD1S': set(),
            'CORD2S': set(),
        }
        
        # Tüm satır bazlı kontrol yapılacak kartlar (SPC, FORCE, MOMENT, MPC)
        line_based_cards = ['SPC', 'SPC1', 'FORCE', 'MOMENT', 'MPC', 'LOAD', 'TEMP', 'TEMPD']
        seen_full_lines = set()  # Tüm satır için
        
        # İstatistikler
        duplicate_counts = {}
        
        result_lines = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            stripped = line.strip()
            
            # Boş veya comment satırı
            if not stripped or stripped.startswith('$'):
                result_lines.append(line)
                i += 1
                continue
            
            upper = stripped.upper()
            
            # Hangi tip kart?
            card_type = None
            is_line_based = False
            
            # Önce line-based kartları kontrol et
            for ctype in line_based_cards:
                if upper.startswith(ctype) and (len(upper) == len(ctype) or 
                    upper[len(ctype)] in ' ,\t*'):
                    card_type = ctype
                    is_line_based = True
                    break
            
            # Sonra ID-based kartları kontrol et
            if not card_type:
                for ctype in id_based_cards.keys():
                    if upper.startswith(ctype) and (len(upper) == len(ctype) or 
                        upper[len(ctype)] in ' ,\t*'):
                        card_type = ctype
                        break
            
            if card_type:
                if is_line_based:
                    # LINE-BASED: Tüm satırı (ve continuation'ları) karşılaştır
                    full_card = [line]
                    j = i + 1
                    
                    # Continuation satırlarını topla
                    while j < len(lines):
                        next_line = lines[j]
                        next_stripped = next_line.strip()
                        is_cont = (next_line.startswith('+') or 
                                  next_line.startswith('*') or
                                  (next_line.startswith('        ') and next_stripped and 
                                   not next_stripped.startswith('$') and
                                   not any(next_stripped.upper().startswith(ct) for ct in 
                                          list(id_based_cards.keys()) + line_based_cards)))
                        if is_cont and next_stripped:
                            full_card.append(next_line)
                            j += 1
                        else:
                            break
                    
                    # Normalize edilmiş tam kart
                    normalized_card = '|'.join(' '.join(l.split()) for l in full_card)
                    
                    if normalized_card in seen_full_lines:
                        # Duplicate - atla
                        if card_type not in duplicate_counts:
                            duplicate_counts[card_type] = 0
                        duplicate_counts[card_type] += 1
                        i = j
                        continue
                    else:
                        seen_full_lines.add(normalized_card)
                        result_lines.extend(full_card)
                        i = j
                        continue
                else:
                    # ID-BASED: Sadece ID'ye bak
                    card_id = None
                    try:
                        if ',' in line:
                            parts = line.split(',')
                            if len(parts) > 1:
                                id_str = parts[1].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                        else:
                            if len(line) >= 16:
                                id_str = line[8:16].strip()
                                if id_str:
                                    card_id = int(float(id_str))
                    except:
                        pass
                    
                    if card_id is not None:
                        if card_id in id_based_cards[card_type]:
                            # Duplicate - atla
                            if card_type not in duplicate_counts:
                                duplicate_counts[card_type] = 0
                            duplicate_counts[card_type] += 1
                            
                            # Continuation satırlarını da atla
                            i += 1
                            while i < len(lines):
                                next_line = lines[i]
                                next_stripped = next_line.strip()
                                is_cont = (next_line.startswith('+') or 
                                          next_line.startswith('*') or
                                          (next_line.startswith('        ') and next_stripped and 
                                           not next_stripped.startswith('$') and
                                           not any(next_stripped.upper().startswith(ct) for ct in 
                                                  list(id_based_cards.keys()) + line_based_cards)))
                                if is_cont and next_stripped:
                                    i += 1
                                else:
                                    break
                            continue
                        else:
                            id_based_cards[card_type].add(card_id)
            
            result_lines.append(line)
            i += 1
        
        # Rapor
        if duplicate_counts:
            self.log1("    Removed duplicates:")
            for ctype, count in sorted(duplicate_counts.items()):
                self.log1(f"      {ctype}: {count}")
            total = sum(duplicate_counts.values())
            self.log1(f"      TOTAL: {total} duplicate entries removed")
        else:
            self.log1("    No duplicates found")
        
        return '\n'.join(result_lines)
    
    def merge_lines_with_pynastran(self, lines):
        """
        Satırları temp BDF'e yazıp pyNastran ile merge eder.
        Duplicate ID hatası olursa, satırları direkt yazar.
        Returns: merged bulk data string
        """
        if not lines:
            self.log1("    WARNING: No lines to merge!")
            return ""
        
        temp_bdf_path = os.path.join(tempfile.gettempdir(), "_temp_lines_to_merge.bdf")
        
        self.log1(f"    Writing {len(lines)} lines to temp BDF...")
        
        try:
            # Temp BDF oluştur
            with open(temp_bdf_path, 'w', encoding='utf-8') as f:
                f.write("$ Temporary BDF for merging\n")
                f.write("SOL 101\n")
                f.write("CEND\n")
                f.write("BEGIN BULK\n")
                for line in lines:
                    f.write(line + "\n")
                f.write("ENDDATA\n")
            
            self.log1(f"    Reading temp BDF with pyNastran (following includes)...")
            
            try:
                bdf = BDF(debug=False)
                # allow_duplicate_ids ile dene
                bdf.read_bdf(temp_bdf_path, validate=False, xref=False, 
                            read_includes=True, save_file_structure=False)
                
                self.log1(f"      Loaded: {len(bdf.nodes)} nodes, {len(bdf.elements)} elements")
                self.log1(f"      Properties: {len(bdf.properties)}, Materials: {len(bdf.materials)}")
                self.log1(f"      Coords: {len(bdf.coords)}, MPCs: {len(bdf.mpcs)}, SPCs: {len(bdf.spcs)}")
                
                # Merge edilmiş BDF'i yaz
                merged_temp_path = os.path.join(tempfile.gettempdir(), "_temp_merged.bdf")
                bdf.write_bdf(merged_temp_path, size=8, is_double=False)
                
                with open(merged_temp_path, 'r', errors='ignore') as f:
                    merged_content = f.read()
                
                if os.path.exists(merged_temp_path): os.remove(merged_temp_path)
                
            except Exception as e:
                self.log1(f"    pyNastran failed: {str(e)[:100]}")
                self.log1(f"    Falling back to direct file reading...")
                
                # pyNastran başarısız oldu - INCLUDE'ları manuel aç
                merged_content = self.expand_includes_manually(temp_bdf_path)
            
            # Temizlik
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)
            
            # BEGIN BULK sonrasını al
            bulk_match = re.search(r'BEGIN\s*,?\s*BULK', merged_content, re.IGNORECASE)
            if bulk_match:
                bulk_data = merged_content[bulk_match.end():]
            else:
                bulk_data = merged_content
            
            # Gereksiz satırları temizle
            result_lines = []
            for l in bulk_data.split('\n'):
                if l.startswith('$pyNastran'): continue
                if l.strip().upper().startswith('ENDDATA'): continue
                if l.strip().upper().startswith('INCLUDE'): continue
                if l.strip().upper().startswith('SOL '): continue
                if l.strip().upper().startswith('CEND'): continue
                if l.strip().upper().startswith('BEGIN'): continue
                result_lines.append(l)
            
            result = '\n'.join(result_lines)
            self.log1(f"      Merged bulk data: {len(result)} characters")
            return result
            
        except Exception as e:
            self.log1(f"    ERROR merging: {e}")
            import traceback
            self.log1(traceback.format_exc())
            if os.path.exists(temp_bdf_path): os.remove(temp_bdf_path)
            return ""
    
    def expand_includes_manually(self, bdf_path):
        """
        INCLUDE'ları manuel olarak açar (pyNastran başarısız olduğunda).
        """
        self.log1("    Expanding includes manually...")
        
        content = self.read_file_safe(bdf_path)
        bdf_dir = os.path.dirname(os.path.abspath(bdf_path))
        
        # INCLUDE'ları bul ve aç
        all_includes = self.parse_multiline_includes(content, bdf_dir)
        
        # Satırları işle
        lines = content.split('\n')
        result_lines = []
        
        # INCLUDE satır indekslerini topla
        include_indices = {}
        for inc in all_includes:
            for idx in range(inc['start_idx'], inc['end_idx'] + 1):
                include_indices[idx] = inc
        
        processed_includes = set()
        
        for idx, line in enumerate(lines):
            if idx in include_indices:
                inc = include_indices[idx]
                # Sadece start_idx'te işle (continuation satırlarını atla)
                if idx == inc['start_idx']:
                    inc_path = inc['abs_path']
                    if inc_path not in processed_includes:
                        processed_includes.add(inc_path)
                        # INCLUDE dosyasını oku
                        if os.path.exists(inc_path):
                            try:
                                inc_content = self.read_file_safe(inc_path)
                                result_lines.append(f"$ === EXPANDED: {os.path.basename(inc_path)} ===")
                                for inc_line in inc_content.split('\n'):
                                    # Recursive INCLUDE'ları atla (basit tutuyoruz)
                                    if not inc_line.strip().upper().startswith('INCLUDE'):
                                        result_lines.append(inc_line)
                            except Exception as e:
                                result_lines.append(f"$ ERROR reading {inc_path}: {e}")
                        else:
                            result_lines.append(f"$ FILE NOT FOUND: {inc_path}")
            else:
                result_lines.append(line)
        
        self.log1(f"      Expanded {len(processed_includes)} includes")
        return '\n'.join(result_lines)
    
    def start_processing(self):
        if not self.thermal_bdfs and not self.maneuver_bdfs:
            messagebox.showerror("Error","Add BDF files"); return
        if not self.excel_path.get():
            messagebox.showerror("Error","Select Excel"); return
        if not self.output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.process_btn.config(state=tk.DISABLED)
        self.progress1.start()
        threading.Thread(target=self.process_merge, daemon=True).start()
    
    def process_merge(self):
        try:
            self.log1("="*70)
            self.log1("BDF Merger Tool v8")
            self.log1("="*70)
            
            self.log1("\n[1] Reading Excel...")
            xl = pd.ExcelFile(self.excel_path.get())
            sheets = xl.sheet_names
            self.log1(f"    Available sheets: {sheets}")

            thermal_sh = maneuver_sh = element_sh = None

            # First pass: look for exact or specific matches
            for s in sheets:
                sl = s.lower()
                # Element_Set exact match (priority)
                if sl == 'element_set' or sl == 'elementset':
                    element_sh = s
                # Thermal
                elif 'thermal' in sl and not thermal_sh:
                    thermal_sh = s
                # Maneuver
                elif ('maneuver' in sl or 'manevra' in sl) and not maneuver_sh:
                    maneuver_sh = s

            # Second pass: if element_sh not found, look for partial matches
            if not element_sh:
                for s in sheets:
                    sl = s.lower()
                    if 'element' in sl and 'set' in sl:
                        element_sh = s
                        break

            # Fallback to index-based if still not found
            if not thermal_sh and len(sheets) > 0: thermal_sh = sheets[0]
            if not maneuver_sh and len(sheets) > 1: maneuver_sh = sheets[1]
            if not element_sh and len(sheets) > 2: element_sh = sheets[2]

            self.log1(f"    Using sheets -> Thermal: '{thermal_sh}', Maneuver: '{maneuver_sh}', Element_Set: '{element_sh}'")
            
            thermal_cases = pd.read_excel(xl, sheet_name=thermal_sh).iloc[:,0].dropna().astype(int).tolist() if thermal_sh else []
            maneuver_cases = pd.read_excel(xl, sheet_name=maneuver_sh).iloc[:,0].dropna().astype(int).tolist() if maneuver_sh else []
            element_ids = sorted(pd.read_excel(xl, sheet_name=element_sh).iloc[:,0].dropna().astype(int).tolist()) if element_sh else []
            
            self.log1(f"    Thermal cases: {len(thermal_cases)}")
            self.log1(f"    Maneuver cases: {len(maneuver_cases)}")
            self.log1(f"    Element IDs: {len(element_ids)}")
            
            set_id = int(self.set_id.get())
            temp_initial = self.temp_initial.get()
            out_dir = self.output_folder.get()
            os.makedirs(out_dir, exist_ok=True)
            
            if self.thermal_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[2] Processing THERMAL...")
                self.log1("="*70)
                self.process_thermal_bdf(self.thermal_bdfs, thermal_cases, element_ids, set_id,
                    temp_initial, os.path.join(out_dir, self.output_thermal_name.get()))
            
            if self.maneuver_bdfs:
                self.log1("\n" + "="*70)
                self.log1("[3] Processing MANEUVER...")
                self.log1("="*70)
                self.process_maneuver_bdf(self.maneuver_bdfs, maneuver_cases, element_ids, set_id,
                    os.path.join(out_dir, self.output_maneuver_name.get()))
            
            self.log1("\n" + "="*70)
            self.log1("COMPLETED!")
            self.log1("="*70)
            self.root.after(0, lambda: messagebox.showinfo("Done","Merge completed!"))
        except Exception as e:
            self.log1(f"\nERROR: {e}")
            import traceback
            self.log1(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error",str(e)))
        finally:
            self.root.after(0, lambda: [self.progress1.stop(), self.process_btn.config(state=tk.NORMAL)])
    
    def process_thermal_bdf(self, bdf_files, load_cases, element_ids, set_id, temp_initial, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        load_case_set = set(load_cases)
        
        # Step 1: Tüm satırları topla
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'THERMAL'
        )
        
        # Step 2: pyNastran ile merge et
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)
        
        # Step 2.5: Duplicate kontrolü
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)
        
        # Step 3: Output dosyası oluştur
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        out.append(f"$ {'='*60}")
        out.append(f"$ THERMAL - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = THERMAL ANALYSIS")
        out.append(f"TEMPERATURE(INITIAL) = {temp_initial}")
        out.append("$")
        
        # SET definition
        chunks = []
        current = ""
        for eid in element_ids:
            test = f"{current},{eid}" if current else str(eid)
            if len(test) > 60 and current:
                chunks.append(current)
                current = str(eid)
            else:
                current = test
        if current: chunks.append(current)
        
        for i, chunk in enumerate(chunks):
            if i == 0:
                out.append(f"SET {set_id} = {chunk}" + ("," if len(chunks) > 1 else ""))
            elif i == len(chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")
        
        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append("GPFORCE(PLOT)=ALL")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")
        
        # SUBCASE definitions
        for lc in load_cases:
            if lc in subcase_info_map:
                info = subcase_info_map[lc]
                temp_load = info['temp_load_id']
                subtitle = info['subtitle']
            else:
                temp_load = lc
                subtitle = f"Thermal Case {lc}"
            out.append(f"SUBCASE {lc}")
            out.append(f"SUBTITLE {subtitle}")
            out.append("SPC = 1")
            out.append(f"TEMPERATURE(LOAD) = {temp_load}")
            out.append("$")
        
        out.append("BEGIN BULK")
        
        # PARAM kartlarını ayır ve BEGIN BULK'tan hemen sonra yaz
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)
        
        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")
        
        # Merged bulk data
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)
        
        # Common Thermal INCLUDE'ları
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON THERMAL INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")
        
        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)
        
        out.append("$")
        out.append("ENDDATA")
        
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))
        
        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON THERMAL INCLUDES: {len(common_includes)}")
    
    def process_maneuver_bdf(self, bdf_files, load_cases, element_ids, set_id, output_path):
        self.log1(f"    MASTER BDFs: {len(bdf_files)}")
        self.log1(f"    Load cases to match: {len(load_cases)}")
        load_case_set = set(load_cases)
        
        # Step 1: Tüm satırları topla
        self.log1("\n    === Step 1: Collecting all lines from MASTER BDFs ===")
        all_lines, common_includes, subcase_info_map = self.collect_all_lines_from_masters(
            bdf_files, load_case_set, 'LOAD'
        )
        
        # Step 2: pyNastran ile merge et
        self.log1("\n    === Step 2: Merging with pyNastran ===")
        merged_bulk_data = self.merge_lines_with_pynastran(all_lines)
        
        # Step 2.5: Duplicate kontrolü
        if merged_bulk_data:
            self.log1("\n    === Step 2.5: Checking duplicates ===")
            merged_bulk_data = self.check_and_remove_duplicates(merged_bulk_data)
        
        # Step 3: Output dosyası oluştur
        self.log1("\n    === Step 3: Writing output BDF ===")
        out = []
        out.append(f"$ {'='*60}")
        out.append(f"$ MANEUVER - MERGED BDF (v8)")
        out.append(f"$ {'='*60}")
        out.append("SOL 101")
        out.append("CEND")
        out.append("ECHO=NONE")
        out.append(f"TITLE = MANEUVER ANALYSIS")
        out.append("$")
        
        # SET definition
        chunks = []
        current = ""
        for eid in element_ids:
            test = f"{current},{eid}" if current else str(eid)
            if len(test) > 60 and current:
                chunks.append(current)
                current = str(eid)
            else:
                current = test
        if current: chunks.append(current)
        
        for i, chunk in enumerate(chunks):
            if i == 0:
                out.append(f"SET {set_id} = {chunk}" + ("," if len(chunks) > 1 else ""))
            elif i == len(chunks) - 1:
                out.append(f"         {chunk}")
            else:
                out.append(f"         {chunk},")
        
        out.append("$")
        out.append("DISPLACEMENT(SORT1,PLOT,REAL)=ALL")
        out.append(f"FORCE(SORT1,PLOT,REAL,CENTER)={set_id}")
        out.append("GPFORCE(PLOT)=ALL")
        out.append("OLOAD(PLOT)=ALL")
        out.append("SPCFORCE(SORT1,PLOT)=ALL")
        out.append("$")
        
        # SUBCASE definitions
        for lc in load_cases:
            if lc in subcase_info_map:
                info = subcase_info_map[lc]
                load_id = info['load_id']
                subtitle = info['subtitle']
            else:
                load_id = lc
                subtitle = f"Manoeuvre {lc}"
            out.append(f"SUBCASE {lc}")
            out.append(f"SUBTITLE {subtitle}")
            out.append("SPC = 1")
            out.append(f"LOAD = {load_id}")
            out.append("$")
        
        out.append("BEGIN BULK")
        
        # PARAM kartlarını ayır ve BEGIN BULK'tan hemen sonra yaz
        param_lines = []
        if merged_bulk_data:
            param_lines, merged_bulk_data = self.extract_param_cards(merged_bulk_data)
        
        if param_lines:
            out.append("$ --- PARAM CARDS ---")
            out.extend(param_lines)
            out.append("$")
        
        # Merged bulk data
        if merged_bulk_data:
            out.append(f"$ {'='*60}")
            out.append(f"$ MERGED STRUCTURE DATA")
            out.append(f"$ {'='*60}")
            out.append(merged_bulk_data)
        
        # Common Load INCLUDE'ları
        out.append("$")
        out.append(f"$ {'='*60}")
        out.append(f"$ COMMON LOAD INCLUDES ({len(common_includes)} files)")
        out.append(f"$ {'='*60}")
        
        for abs_path in sorted(common_includes):
            include_lines = self.format_include_nastran(abs_path)
            out.extend(include_lines)
        
        out.append("$")
        out.append("ENDDATA")
        
        with open(output_path, 'w') as f:
            f.write('\n'.join(out))
        
        self.log1(f"    Output: {os.path.basename(output_path)}")
        self.log1(f"    COMMON LOAD INCLUDES: {len(common_includes)}")

    # ============= TAB 2 HELPERS =============
    def add_run_bdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("BDF","*.bdf *.dat *.nas"),("All","*.*")])
        for f in files:
            if f not in self.run_bdfs:
                self.run_bdfs.append(f)
                self.run_listbox.insert(tk.END, f)
        self.run_count.set(f"{len(self.run_bdfs)} files")
    
    def clear_run_bdfs(self):
        self.run_bdfs.clear()
        self.run_listbox.delete(0, tk.END)
        self.run_count.set("0 files")
    
    def browse_property_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if f: self.property_excel_path.set(f)
    
    def browse_nastran(self):
        f = filedialog.askopenfilename(filetypes=[("All","*.*")])
        if f: self.nastran_path.set(f)
    
    def browse_run_output(self):
        f = filedialog.askdirectory()
        if f: self.run_output_folder.set(f)
    
    def log2(self, msg):
        self.log_text2.insert(tk.END, msg + "\n")
        self.log_text2.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log2(self):
        self.log_text2.delete(1.0, tk.END)
    
    def load_properties(self):
        if not self.property_excel_path.get():
            messagebox.showerror("Error","Select Excel"); return
        try:
            xl = pd.ExcelFile(self.property_excel_path.get())
            sheets = xl.sheet_names
            bar_sh = skin_sh = res_sh = None

            # First pass: exact matches (priority)
            for s in sheets:
                sl = s.lower().replace('_',' ').replace('-',' ')
                if sl == 'skin property' or sl == 'skinproperty':
                    skin_sh = s
                elif sl == 'bar property' or sl == 'barproperty':
                    bar_sh = s
                elif sl == 'residual strength' or sl == 'residualstrength':
                    res_sh = s

            # Second pass: partial matches
            for s in sheets:
                sl = s.lower().replace('_',' ')
                if not bar_sh and 'bar' in sl and 'prop' in sl:
                    bar_sh = s
                elif not skin_sh and 'skin' in sl and 'prop' in sl:
                    skin_sh = s
                elif not res_sh and ('residual' in sl or 'strength' in sl):
                    res_sh = s

            self.bar_properties.clear()
            self.skin_properties.clear()

            if bar_sh:
                df = pd.read_excel(xl, sheet_name=bar_sh)
                for _, row in df.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        d1 = float(row.iloc[1]) if len(df.columns)>1 else 0
                        d2 = float(row.iloc[2]) if len(df.columns)>2 else 0
                        self.bar_properties[pid] = {'dim1':d1, 'dim2':d2}
                    except: pass

            if skin_sh:
                df = pd.read_excel(xl, sheet_name=skin_sh)
                for _, row in df.iterrows():
                    try:
                        pid = int(row.iloc[0])
                        t = float(row.iloc[1])
                        self.skin_properties[pid] = {'thickness':t}
                    except: pass

            if res_sh:
                self.residual_strength_df = pd.read_excel(xl, sheet_name=res_sh)

            self.bar_prop_text.delete(1.0, tk.END)
            self.bar_prop_text.insert(tk.END, f"Bar: {len(self.bar_properties)} loaded")
            self.skin_prop_text.delete(1.0, tk.END)
            self.skin_prop_text.insert(tk.END, f"Skin: {len(self.skin_properties)} loaded")
            self.resid_text.delete(1.0, tk.END)
            self.resid_text.insert(tk.END, f"Residual: {'Yes' if self.residual_strength_df is not None else 'No'}")

            # Log which sheets were used
            print(f"[Load Properties] Bar sheet: {bar_sh}, Skin sheet: {skin_sh}")
            print(f"[Load Properties] Bar PIDs: {len(self.bar_properties)}, Skin PIDs: {len(self.skin_properties)}")
            if self.skin_properties:
                sample_pids = list(self.skin_properties.keys())[:5]
                print(f"[Load Properties] Sample Skin PIDs: {sample_pids}")

            messagebox.showinfo("OK", f"Bar:{len(self.bar_properties)} Skin:{len(self.skin_properties)}\n\nSheets used:\nBar: {bar_sh}\nSkin: {skin_sh}")
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", str(e))
    
    def read_file(self, path):
        for enc in ['utf-8','latin-1','cp1252']:
            try:
                with open(path, 'r', encoding=enc) as f:
                    return f.read()
            except: pass
        return ""
    
    def copy_bdf_to_output(self, bdf_path, output_folder):
        bdf_name = os.path.basename(bdf_path)
        out_bdf = os.path.join(output_folder, bdf_name)
        shutil.copy2(bdf_path, out_bdf)
        return out_bdf
    
    def count_pcomp_plies(self, lines, start_idx):
        ply_count = 0
        j = start_idx + 1
        while j < len(lines):
            line = lines[j]
            stripped = line.strip()
            is_continuation = (
                line.startswith('+') or line.startswith('*') or
                (line.startswith(' ') and stripped and not stripped.startswith('$') and
                 not any(stripped.upper().startswith(card) for card in 
                        ['PSHELL','PCOMP','PBAR','PBARL','CBAR','CQUAD','CTRIA',
                         'GRID','MAT','FORCE','INCLUDE','END','SOL','CEND','BEGIN']))
            )
            if not is_continuation:
                break
            ply_count += 1
            j += 1
        return ply_count, j
    
    def update_properties_in_file(self, filepath):
        content = self.read_file(filepath)
        lines = content.split('\n')
        new_lines = []
        i = 0
        stats = {'pbarl': 0, 'pbar': 0, 'pshell': 0, 'pcomp': 0}
        warnings = []
        pshell_found = []
        pcomp_found = []
        
        while i < len(lines):
            line = lines[i]
            upper = line.upper().strip()
            
            if upper.startswith('PBARL'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    if i+1 < len(lines) and pid in self.bar_properties:
                        d1 = self.bar_properties[pid]['dim1']
                        d2 = self.bar_properties[pid]['dim2']
                        new_lines.append(line)
                        next_line = lines[i+1]
                        if ',' in next_line:
                            parts = next_line.split(',')
                            start_idx = 1 if parts[0].strip().startswith('+') else 0
                            if len(parts) > start_idx: parts[start_idx] = f"{d1}."
                            if len(parts) > start_idx + 1: parts[start_idx + 1] = f"{d2}."
                            new_lines.append(','.join(parts))
                        else:
                            cont = next_line[:8]
                            rest = next_line[24:] if len(next_line) > 24 else ""
                            d1_str = f"{d1:<8.6g}".rstrip()
                            if '.' not in d1_str and 'E' not in d1_str.upper(): d1_str += '.'
                            d2_str = f"{d2:<8.6g}".rstrip()
                            if '.' not in d2_str and 'E' not in d2_str.upper(): d2_str += '.'
                            new_lines.append(f"{cont}{d1_str:>8}{d2_str:>8}{rest}")
                        stats['pbarl'] += 1
                        i += 2
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            
            elif upper.startswith('PBAR') and not upper.startswith('PBARL'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    if pid in self.bar_properties:
                        d1 = self.bar_properties[pid]['dim1']
                        d2 = self.bar_properties[pid]['dim2']
                        area = d1 * d2
                        if ',' in line:
                            parts = line.split(',')
                            parts[3] = str(area)
                            new_lines.append(','.join(parts))
                        else:
                            new_lines.append(line[:24] + f"{area:8.4g}" + line[32:])
                        stats['pbar'] += 1
                        i += 1
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            
            elif upper.startswith('PSHELL'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    pshell_found.append(pid)
                    if pid in self.skin_properties:
                        t = self.skin_properties[pid]['thickness']
                        if ',' in line:
                            parts = line.split(',')
                            t_str = f"{t}"
                            if '.' not in t_str and 'E' not in t_str.upper():
                                t_str += '.'
                            parts[3] = t_str
                            new_lines.append(','.join(parts))
                        else:
                            t_str = f"{t:<8.6g}".rstrip()
                            if '.' not in t_str and 'E' not in t_str.upper():
                                t_str += '.'
                            new_lines.append(line[:24] + f"{t_str:>8}" + line[32:])
                        stats['pshell'] += 1
                        i += 1
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            
            elif upper.startswith('PCOMP'):
                try:
                    if ',' in line:
                        pid = int(float(line.split(',')[1].strip()))
                    else:
                        pid = int(float(line[8:16].strip()))
                    pcomp_found.append(pid)
                    ply_count, end_idx = self.count_pcomp_plies(lines, i)
                    if pid in self.skin_properties:
                        if ply_count > 1:
                            warnings.append(f"PCOMP {pid}: {ply_count} plies - SKIPPED")
                            for k in range(i, end_idx):
                                new_lines.append(lines[k])
                            i = end_idx
                            continue
                        else:
                            t = self.skin_properties[pid]['thickness']
                            new_lines.append(line)
                            if i + 1 < end_idx:
                                cont_line = lines[i + 1]
                                if ',' in cont_line:
                                    parts = cont_line.split(',')
                                    if len(parts) >= 3:
                                        t_str = f"{t}"
                                        if '.' not in t_str and 'E' not in t_str.upper():
                                            t_str += '.'
                                        parts[2] = t_str
                                    new_lines.append(','.join(parts))
                                else:
                                    cont = cont_line[:8]
                                    mid = cont_line[8:16] if len(cont_line) > 8 else "        "
                                    rest = cont_line[24:] if len(cont_line) > 24 else ""
                                    t_str = f"{t:<8.6g}".rstrip()
                                    if '.' not in t_str and 'E' not in t_str.upper():
                                        t_str += '.'
                                    new_lines.append(f"{cont}{mid}{t_str:>8}{rest}")
                                for k in range(i + 2, end_idx):
                                    new_lines.append(lines[k])
                            stats['pcomp'] += 1
                            i = end_idx
                            continue
                    else:
                        for k in range(i, end_idx):
                            new_lines.append(lines[k])
                        i = end_idx
                        continue
                except: pass
                new_lines.append(line)
                i += 1
            else:
                new_lines.append(line)
                i += 1
        
        # Debug logging
        print(f"[Update Props] File: {os.path.basename(filepath)}")
        print(f"[Update Props] PSHELL found in BDF: {len(pshell_found)}, PIDs: {pshell_found[:10]}...")
        print(f"[Update Props] PCOMP found in BDF: {len(pcomp_found)}, PIDs: {pcomp_found[:10]}...")
        print(f"[Update Props] Skin properties loaded: {len(self.skin_properties)}")
        if pshell_found:
            matched = [p for p in pshell_found if p in self.skin_properties]
            not_matched = [p for p in pshell_found if p not in self.skin_properties]
            print(f"[Update Props] PSHELL matched: {len(matched)}, not matched: {len(not_matched)}")
            if not_matched:
                print(f"[Update Props] Sample unmatched PSHELL PIDs: {not_matched[:5]}")

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(new_lines))
        return stats, warnings
    
    def start_update_properties(self):
        if not self.run_bdfs:
            messagebox.showerror("Error","Add BDF files"); return
        if not self.bar_properties and not self.skin_properties:
            messagebox.showerror("Error","Load properties first"); return
        if not self.run_output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.btn1.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_update_properties, daemon=True).start()
    
    def do_update_properties(self):
        try:
            self.log2("="*60)
            self.log2("STEP 1: Update Properties")
            self.log2("="*60)
            self.log2(f"  Loaded Bar properties: {len(self.bar_properties)}")
            self.log2(f"  Loaded Skin properties: {len(self.skin_properties)}")
            if self.skin_properties:
                sample_pids = list(self.skin_properties.keys())[:5]
                self.log2(f"  Sample Skin PIDs: {sample_pids}")
            out_folder = self.run_output_folder.get()
            os.makedirs(out_folder, exist_ok=True)
            all_warnings = []
            for bdf_path in self.run_bdfs:
                self.log2(f"\n  Processing: {os.path.basename(bdf_path)}")
                self.log2("    Copying BDF to output folder...")
                out_bdf = self.copy_bdf_to_output(bdf_path, out_folder)
                self.log2("    Updating properties...")
                stats, warnings = self.update_properties_in_file(out_bdf)
                self.log2(f"    Updated: PBARL={stats['pbarl']} PBAR={stats['pbar']} PSHELL={stats['pshell']} PCOMP={stats['pcomp']}")
                all_warnings.extend(warnings)
            if all_warnings:
                self.log2("\n  Warnings:")
                for w in all_warnings[:10]:
                    self.log2(f"    {w}")
            self.log2("\n" + "="*60)
            self.log2("COMPLETED!")
            self.root.after(0, lambda: messagebox.showinfo("Done","Properties updated!"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error",str(e)))
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn1.config(state=tk.NORMAL)])
    
    def start_run_nastran(self):
        if not self.nastran_path.get():
            messagebox.showerror("Error","Select Nastran path"); return
        if not self.run_output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.btn2.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_run_nastran, daemon=True).start()
    
    def do_run_nastran(self):
        try:
            self.log2("="*60)
            self.log2("STEP 2: Run Nastran")
            self.log2("="*60)
            out_folder = self.run_output_folder.get()
            nastran = self.nastran_path.get()
            bdfs = [os.path.join(out_folder, f) for f in os.listdir(out_folder) if f.lower().endswith(('.bdf','.dat','.nas'))]
            for bdf in bdfs:
                self.log2(f"\n  Running: {os.path.basename(bdf)}")
                cmd = f'"{nastran}" "{bdf}" out="{out_folder}" scratch=yes batch=no'
                os.popen(cmd)
            self.log2("\nJobs submitted!")
            self.root.after(0, lambda: messagebox.showinfo("Done","Nastran jobs submitted!"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn2.config(state=tk.NORMAL)])
    
    def start_postprocess(self):
        if not self.run_output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        self.btn3.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_postprocess, daemon=True).start()
    
    def do_postprocess(self):
        try:
            self.log2("="*60)
            self.log2("STEP 3: Post-Process OP2")
            self.log2("="*60)
            out_folder = self.run_output_folder.get()
            op2_files = [os.path.join(out_folder, f) for f in os.listdir(out_folder) if f.lower().endswith('.op2')]
            if not op2_files:
                self.log2("No OP2 files found!")
                return
            
            # Element -> Property mapping ve PBARL dimensions
            elem_prop = {}
            pbarl_dims = {}  # {pid: {'dim1': x, 'dim2': y, 'type': 'ROD'/'BAR'}}
            
            # BDF'lerden oku
            bdf_files_to_read = list(self.run_bdfs)
            # Output folder'daki BDF'leri de ekle
            for f in os.listdir(out_folder):
                if f.lower().endswith(('.bdf', '.dat', '.nas')):
                    bdf_files_to_read.append(os.path.join(out_folder, f))
            
            for bdf_path in bdf_files_to_read:
                try:
                    self.log2(f"  Reading BDF: {os.path.basename(bdf_path)}")
                    bdf = BDF(debug=False)
                    bdf.read_bdf(bdf_path, validate=False, xref=False, read_includes=True)
                    
                    # Element -> Property mapping
                    for eid, el in bdf.elements.items():
                        if hasattr(el, 'pid'):
                            elem_prop[eid] = el.pid
                    
                    # PBARL dimensions
                    for pid, prop in bdf.properties.items():
                        prop_type = prop.type
                        if prop_type == 'PBARL':
                            # PBARL formatı: dim listesi var
                            dims = prop.dim if hasattr(prop, 'dim') else []
                            bar_type = prop.bar_type if hasattr(prop, 'bar_type') else 'UNKNOWN'
                            
                            if len(dims) >= 2:
                                pbarl_dims[pid] = {
                                    'dim1': dims[0],
                                    'dim2': dims[1],
                                    'type': bar_type
                                }
                            elif len(dims) == 1:
                                # ROD gibi tek dimension
                                pbarl_dims[pid] = {
                                    'dim1': dims[0],
                                    'dim2': dims[0],
                                    'type': bar_type
                                }
                        elif prop_type == 'PBAR':
                            # PBAR: A (area) var
                            area = prop.A if hasattr(prop, 'A') else None
                            if area:
                                # Area'dan approximate dims (kare varsayımı)
                                import math
                                side = math.sqrt(area) if area > 0 else 0
                                pbarl_dims[pid] = {
                                    'dim1': side,
                                    'dim2': side,
                                    'type': 'PBAR',
                                    'area': area
                                }
                    
                    self.log2(f"    Elements: {len(elem_prop)}, PBARL/PBAR props: {len(pbarl_dims)}")
                except Exception as e:
                    self.log2(f"    Warning reading BDF: {e}")
            
            self.log2(f"\n  Total: {len(elem_prop)} elements, {len(pbarl_dims)} bar properties from BDF")
            
            results = []
            for op2_path in op2_files:
                self.log2(f"\n  Processing: {os.path.basename(op2_path)}")
                try:
                    op2 = OP2(debug=False)
                    op2.read_op2(op2_path)
                    if hasattr(op2, 'cbar_force') and op2.cbar_force:
                        for sc_id, force in op2.cbar_force.items():
                            for i, eid in enumerate(force.element):
                                axial = force.data[0,i,6] if len(force.data.shape)==3 else force.data[i,6]
                                pid = elem_prop.get(eid)
                                d1 = d2 = area = stress = None
                                
                                # Önce Excel'den yüklenen properties'e bak
                                if pid and pid in self.bar_properties:
                                    d1 = self.bar_properties[pid]['dim1']
                                    d2 = self.bar_properties[pid]['dim2']
                                    area = d1 * d2
                                    if area > 0: stress = axial / area
                                # Excel'de yoksa BDF'den okunan PBARL'a bak
                                elif pid and pid in pbarl_dims:
                                    prop_info = pbarl_dims[pid]
                                    d1 = prop_info['dim1']
                                    d2 = prop_info['dim2']
                                    if 'area' in prop_info:
                                        area = prop_info['area']
                                    else:
                                        area = d1 * d2
                                    if area > 0: stress = axial / area
                                
                                results.append({'OP2': os.path.basename(op2_path), 'Subcase': sc_id, 'Element': eid,
                                    'Property': pid, 'Axial': axial, 'Dim1': d1, 'Dim2': d2, 'Area': area, 'Stress': stress})
                except Exception as e:
                    self.log2(f"    ERROR: {e}")
            
            csv_path = os.path.join(out_folder, self.csv_output_name.get())
            with open(csv_path, 'w', newline='') as f:
                w = csv.DictWriter(f, fieldnames=['OP2','Subcase','Element','Property','Axial','Dim1','Dim2','Area','Stress'])
                w.writeheader()
                w.writerows(results)
            self.log2(f"\n  Saved: {csv_path} ({len(results)} rows)")
            self.root.after(0, lambda: messagebox.showinfo("Done","Results saved"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn3.config(state=tk.NORMAL)])
    
    def start_combine_stress(self):
        if not self.run_output_folder.get():
            messagebox.showerror("Error","Select output folder"); return
        if self.residual_strength_df is None:
            messagebox.showerror("Error","Load Residual Strength data"); return
        self.btn4.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_combine_stress, daemon=True).start()
    
    def do_combine_stress(self):
        try:
            self.log2("="*60)
            self.log2("STEP 4: Combine Stress")
            self.log2("="*60)
            out_folder = self.run_output_folder.get()
            stress_csv = os.path.join(out_folder, self.csv_output_name.get())
            if not os.path.exists(stress_csv):
                self.log2("Stress CSV not found")
                return
            
            stress_df = pd.read_csv(stress_csv)
            lookup = {}
            for _, row in stress_df.iterrows():
                key = (int(row['Subcase']), int(row['Element']))
                lookup[key] = row['Stress'] if pd.notna(row['Stress']) else 0
            elements = stress_df['Element'].unique()
            
            rs_df = self.residual_strength_df
            cols = rs_df.columns.tolist()
            comb_col = cols[0]
            
            comp_cols = []
            i = 1
            while i < len(cols) - 1:
                col_name = str(cols[i]).upper()
                next_col_name = str(cols[i+1]).upper()
                if ('CASE' in col_name or 'ID' in col_name) and 'MULT' in next_col_name:
                    comp_cols.append((cols[i], cols[i+1]))
                    i += 2
                else:
                    i += 1
            
            results = []
            for _, rs_row in rs_df.iterrows():
                comb_lc = rs_row[comb_col]
                if pd.isna(comb_lc): continue
                comb_lc = int(comb_lc)
                
                for eid in elements:
                    total_stress = 0.0
                    components = []
                    for case_col, mult_col in comp_cols:
                        case_id = rs_row[case_col]
                        multiplier = rs_row[mult_col]
                        if pd.isna(case_id) or pd.isna(multiplier): continue
                        case_id = int(case_id)
                        multiplier = float(multiplier)
                        key = (case_id, int(eid))
                        if key in lookup:
                            stress = lookup[key]
                            if stress is not None:
                                total_stress += stress * multiplier
                                components.append(f"{case_id}*{multiplier}")
                    if components:
                        results.append({'Combined_LC': comb_lc, 'Element': eid, 
                            'Combined_Stress': total_stress, 'Components': ' + '.join(components)})
            
            comb_csv = os.path.join(out_folder, self.combined_csv_name.get())
            with open(comb_csv, 'w', newline='') as f:
                w = csv.DictWriter(f, fieldnames=['Combined_LC','Element','Combined_Stress','Components'])
                w.writeheader()
                w.writerows(results)
            self.log2(f"\n  Saved: {comb_csv}")
            self.root.after(0, lambda: messagebox.showinfo("Done","Combined stress saved"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn4.config(state=tk.NORMAL)])
    
    # ============= TAB 2 - OFFSET APPLICATION =============
    def browse_tab2_offset_csv(self):
        f = filedialog.askopenfilename(
            title="Select Offset CSV File",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if f:
            self.tab2_offset_csv.set(f)
            self.log2(f"Selected offset CSV: {os.path.basename(f)}")
    
    def start_apply_offsets_tab2(self):
        """Apply offsets to BDF files in Tab 2"""
        if not self.run_bdfs:
            messagebox.showerror("Error", "Please add BDF files first")
            return
        if not self.tab2_offset_csv.get():
            messagebox.showerror("Error", "Please select offset CSV file")
            return
        
        self.btn_apply_offset_tab2.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.apply_offsets_tab2, daemon=True).start()
    
    def apply_offsets_tab2(self):
        """Apply offsets from CSV to all BDF files - TEXT BASED (preserves INCLUDEs)"""
        try:
            self.log2("\n" + "="*60)
            self.log2("APPLYING OFFSETS TO BDF FILES (Text-based)")
            self.log2("="*60)

            # Read offset CSV
            self.log2("\nReading offset CSV...")
            landing_offsets = {}  # {eid: zoffset}
            bar_offsets = {}      # {eid: (offset_x, offset_y, offset_z)}

            with open(self.tab2_offset_csv.get(), 'r') as f:
                reader = csv.reader(f)
                section = None

                for row in reader:
                    if not row or not row[0]:
                        continue

                    if 'LANDING OFFSETS' in row[0]:
                        section = 'landing'
                        next(reader)  # Skip header
                        continue
                    elif 'BAR OFFSETS' in row[0]:
                        section = 'bar'
                        next(reader)  # Skip header
                        continue

                    if section == 'landing':
                        try:
                            eid = int(row[0])
                            zoffset = float(row[5])
                            landing_offsets[eid] = zoffset
                        except:
                            pass
                    elif section == 'bar':
                        try:
                            eid = int(row[0])
                            offset_x = float(row[8])
                            offset_y = float(row[9])
                            offset_z = float(row[10])
                            bar_offsets[eid] = (offset_x, offset_y, offset_z)
                        except:
                            pass

            self.log2(f"  Loaded {len(landing_offsets)} landing offsets")
            self.log2(f"  Loaded {len(bar_offsets)} bar offsets")

            # Helper function to format field (8-char fixed width)
            def fmt_field(value, width=8):
                if isinstance(value, float):
                    s = f"{value:.4f}"
                    if len(s) > width:
                        s = f"{value:.2E}"
                    return s[:width].ljust(width)
                return str(value)[:width].ljust(width)

            # Apply to each BDF file using text replacement
            total_landing = 0
            total_bar = 0

            for bdf_path in self.run_bdfs:
                self.log2(f"\n  Processing: {os.path.basename(bdf_path)}")

                # Read file as text
                with open(bdf_path, 'r', encoding='latin-1') as f:
                    lines = f.readlines()

                new_lines = []
                i = 0
                landing_modified = 0
                bar_modified = 0

                while i < len(lines):
                    line = lines[i]

                    # Check for CQUAD4 (landing/skin elements with zoffset)
                    if line.startswith('CQUAD4'):
                        # Parse element ID from field 2 (chars 8-16)
                        try:
                            eid = int(line[8:16].strip())
                            if eid in landing_offsets:
                                zoff = landing_offsets[eid]
                                # CQUAD4 format: CQUAD4, EID, PID, G1, G2, G3, G4, THETA, ZOFFS
                                # Fields are 8 chars each
                                # Field 8 (chars 56-64) is THETA/MCID, Field 9 (chars 64-72) is ZOFFS
                                if len(line) >= 64:
                                    # Modify ZOFFS field (field 9, chars 64-72)
                                    new_line = line[:64] + fmt_field(zoff) + line[72:] if len(line) > 72 else line[:64] + fmt_field(zoff) + '\n'
                                    new_lines.append(new_line)
                                    landing_modified += 1
                                    i += 1
                                    continue
                        except:
                            pass
                        new_lines.append(line)
                        i += 1
                        continue

                    # Check for CBAR elements
                    elif line.startswith('CBAR'):
                        # Parse element ID from field 2 (chars 8-16)
                        try:
                            eid = int(line[8:16].strip())
                            if eid in bar_offsets:
                                offset_vec = bar_offsets[eid]
                                # CBAR can be single line or multi-line (with continuation)
                                # Format: CBAR, EID, PID, GA, GB, X1, X2, X3, OFFT (line 1)
                                #         +, PA, PB, W1A, W2A, W3A, W1B, W2B, W3B (line 2)

                                # Check if next line is continuation
                                if i + 1 < len(lines) and (lines[i+1].startswith('+') or lines[i+1].startswith('*') or lines[i+1].startswith(' ')):
                                    # Multi-line CBAR - modify continuation line
                                    cont_line = lines[i+1]
                                    # Fields on continuation: +, PA, PB, W1A, W2A, W3A, W1B, W2B, W3B
                                    # W1A at field 4 (chars 24-32), W2A at field 5 (chars 32-40), W3A at field 6 (chars 40-48)
                                    # W1B at field 7 (chars 48-56), W2B at field 8 (chars 56-64), W3B at field 9 (chars 64-72)

                                    # Keep first part (continuation marker, PA, PB)
                                    new_cont = cont_line[:24]
                                    # Add W1A, W2A, W3A
                                    new_cont += fmt_field(offset_vec[0])
                                    new_cont += fmt_field(offset_vec[1])
                                    new_cont += fmt_field(offset_vec[2])
                                    # Add W1B, W2B, W3B (same as WA)
                                    new_cont += fmt_field(offset_vec[0])
                                    new_cont += fmt_field(offset_vec[1])
                                    new_cont += fmt_field(offset_vec[2])
                                    new_cont += '\n'

                                    new_lines.append(line)
                                    new_lines.append(new_cont)
                                    bar_modified += 1
                                    i += 2
                                    continue
                                else:
                                    # Single line CBAR - need to add continuation for offsets
                                    new_lines.append(line.rstrip() + '+CB' + str(eid)[-4:] + '\n')
                                    # Create continuation line with offsets
                                    cont_name = '+CB' + str(eid)[-4:]
                                    new_cont = cont_name.ljust(8)
                                    new_cont += '        '  # PA (blank)
                                    new_cont += '        '  # PB (blank)
                                    new_cont += fmt_field(offset_vec[0])  # W1A
                                    new_cont += fmt_field(offset_vec[1])  # W2A
                                    new_cont += fmt_field(offset_vec[2])  # W3A
                                    new_cont += fmt_field(offset_vec[0])  # W1B
                                    new_cont += fmt_field(offset_vec[1])  # W2B
                                    new_cont += fmt_field(offset_vec[2])  # W3B
                                    new_cont += '\n'
                                    new_lines.append(new_cont)
                                    bar_modified += 1
                                    i += 1
                                    continue
                        except:
                            pass
                        new_lines.append(line)
                        i += 1
                        continue

                    else:
                        new_lines.append(line)
                        i += 1

                # Save with "_offset" suffix
                base_name = os.path.basename(bdf_path)
                name_parts = os.path.splitext(base_name)
                output_name = name_parts[0] + "_offset" + name_parts[1]
                output_path = os.path.join(os.path.dirname(bdf_path), output_name)

                with open(output_path, 'w', encoding='latin-1') as f:
                    f.writelines(new_lines)

                self.log2(f"    Landing (ZOFFS): {landing_modified}, Bar (WA/WB): {bar_modified}")
                self.log2(f"    Saved: {output_name}")

                total_landing += landing_modified
                total_bar += bar_modified

            self.log2("\n" + "="*60)
            self.log2("OFFSET APPLICATION COMPLETED!")
            self.log2("="*60)
            self.log2(f"\nTotal across all files:")
            self.log2(f"  Landing elements (ZOFFS): {total_landing}")
            self.log2(f"  Bar elements (WA/WB): {total_bar}")
            self.log2(f"  Files processed: {len(self.run_bdfs)}")
            self.log2("  * INCLUDEs preserved (text-based modification)")
            self.log2("="*60)

            self.root.after(0, lambda: messagebox.showinfo("Success",
                f"Offsets applied to {len(self.run_bdfs)} BDF files!\n\n"
                f"Landing (ZOFFS): {total_landing}\nBar (WA/WB): {total_bar}\n\n"
                f"Files saved with '_offset' suffix\n"
                f"INCLUDEs preserved!"))

        except Exception as e:
            self.log2(f"\nERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: [self.progress2.stop(),
                                       self.btn_apply_offset_tab2.config(state=tk.NORMAL)])


    def browse_offset_element_excel(self):
        f = filedialog.askopenfilename(
            title="Select Element IDs Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        if f:
            self.offset_element_excel.set(f)
            self.log2(f"Selected Element Excel: {os.path.basename(f)}")

    def start_calculate_offsets(self):
        """Thread starter for offset calculation (merged from Tab 3)"""
        if not self.run_bdfs:
            messagebox.showerror("Error", "Please add BDF files first")
            return
        if not self.offset_element_excel.get():
            messagebox.showerror("Error", "Please select Excel file with element IDs")
            return
        if not self.run_output_folder.get():
            messagebox.showerror("Error", "Please select output folder")
            return

        self.btn_calc_offset.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.calculate_offsets, daemon=True).start()

    def calculate_offsets(self):
        """Calculate offsets using BDF files and save to CSV (merged from Tab 3)"""
        try:
            self.log2("\n" + "="*70)
            self.log2("CALCULATING OFFSETS (from BDF files)")
            self.log2("="*70)

            # Read Excel for element IDs
            self.log2("\nReading element IDs from Excel...")
            xl = pd.ExcelFile(self.offset_element_excel.get())
            sheets = xl.sheet_names
            self.log2(f"Available sheets: {', '.join(sheets)}")

            landing_sheet = bar_sheet = None
            for s in sheets:
                s_lower = s.lower().replace('_', '').replace(' ', '')
                if 'landing' in s_lower and 'offset' in s_lower:
                    landing_sheet = s
                elif 'bar' in s_lower and 'offset' in s_lower:
                    bar_sheet = s

            landing_elem_ids = []
            bar_elem_ids = []

            if landing_sheet:
                self.log2(f"\nReading '{landing_sheet}'...")
                df = pd.read_excel(xl, sheet_name=landing_sheet)
                landing_elem_ids = df.iloc[:,0].dropna().astype(int).tolist()
                self.log2(f"  Found {len(landing_elem_ids)} landing element IDs")

            if bar_sheet:
                self.log2(f"\nReading '{bar_sheet}'...")
                df = pd.read_excel(xl, sheet_name=bar_sheet)
                bar_elem_ids = df.iloc[:,0].dropna().astype(int).tolist()
                self.log2(f"  Found {len(bar_elem_ids)} bar element IDs")

            # Read first BDF with pyNastran for element/property info
            bdf_path = self.run_bdfs[0]
            self.log2(f"\nReading BDF with pyNastran: {os.path.basename(bdf_path)}")

            bdf = BDF(debug=False)
            try:
                bdf.read_bdf(bdf_path, validate=False, xref=False,
                            read_includes=True, encoding='latin-1')
            except Exception:
                self.log2("  Standard read failed, retrying with punch=True...")
                bdf = BDF(debug=False)
                bdf.read_bdf(bdf_path, validate=False, xref=False,
                            read_includes=True, encoding='latin-1', punch=True)

            self.log2(f"  Nodes: {len(bdf.nodes)}")
            self.log2(f"  Elements: {len(bdf.elements)}")
            self.log2(f"  Properties: {len(bdf.properties)}")

            # Calculate Landing offsets
            self.log2("\nCalculating Landing Offsets...")

            landing_results = []
            landing_thickness = {}
            landing_normals = {}

            for eid in landing_elem_ids:
                if eid in bdf.elements:
                    elem = bdf.elements[eid]
                    if hasattr(elem, 'pid') and elem.pid in bdf.properties:
                        prop = bdf.properties[elem.pid]

                        thickness = None
                        if hasattr(prop, 't'):
                            thickness = prop.t
                        elif hasattr(prop, 'total_thickness'):
                            thickness = prop.total_thickness()

                        if thickness:
                            zoffset = -thickness / 2.0
                            landing_thickness[eid] = thickness

                            landing_results.append({
                                'Element_ID': eid,
                                'Element_Type': elem.type,
                                'Property_ID': elem.pid,
                                'Property_Type': prop.type,
                                'Thickness': thickness,
                                'Zoffset': zoffset
                            })

                            if elem.type in ['CQUAD4', 'CTRIA3', 'CQUAD8', 'CTRIA6']:
                                node_ids = elem.node_ids[:4] if elem.type.startswith('CQUAD') else elem.node_ids[:3]
                                nodes = [bdf.nodes[nid] for nid in node_ids if nid in bdf.nodes]

                                if len(nodes) >= 3:
                                    p1 = np.array(nodes[0].xyz)
                                    p2 = np.array(nodes[1].xyz)
                                    p3 = np.array(nodes[2].xyz)

                                    v1 = p2 - p1
                                    v2 = p3 - p1
                                    normal = np.cross(v1, v2)
                                    normal_len = np.linalg.norm(normal)

                                    if normal_len > 1e-10:
                                        landing_normals[eid] = normal / normal_len

            self.log2(f"  Calculated offsets for {len(landing_results)} landing elements")

            # Build node -> shell mapping
            self.log2("Building node-to-shell mapping...")

            node_to_shells = {}
            for eid, elem in bdf.elements.items():
                if elem.type in ['CQUAD4', 'CTRIA3', 'CQUAD8', 'CTRIA6']:
                    for nid in elem.node_ids:
                        if nid not in node_to_shells:
                            node_to_shells[nid] = []
                        node_to_shells[nid].append(eid)

            self.log2(f"  Mapped {len(node_to_shells)} nodes to shell elements")

            # Calculate Bar offsets
            self.log2("Calculating Bar Offsets...")

            bar_results = []
            bar_no_landing = 0

            for eid in bar_elem_ids:
                if eid in bdf.elements:
                    elem = bdf.elements[eid]
                    if elem.type == 'CBAR' and hasattr(elem, 'pid') and elem.pid in bdf.properties:
                        prop = bdf.properties[elem.pid]

                        thickness = None
                        if prop.type == 'PBARL':
                            if hasattr(prop, 'dim') and len(prop.dim) > 0:
                                thickness = prop.dim[0]
                        elif prop.type == 'PBAR':
                            if hasattr(prop, 'A') and prop.A > 0:
                                thickness = np.sqrt(prop.A)

                        if thickness:
                            bar_nodes = elem.node_ids[:2]

                            if bar_nodes[0] in node_to_shells and bar_nodes[1] in node_to_shells:
                                shells_n1 = set(node_to_shells[bar_nodes[0]])
                                shells_n2 = set(node_to_shells[bar_nodes[1]])
                                connected_shells = shells_n1.intersection(shells_n2)

                                max_landing_thick = 0
                                landing_normal = None
                                connected_landing_id = None

                                for shell_eid in connected_shells:
                                    if shell_eid in landing_thickness:
                                        t = landing_thickness[shell_eid]
                                        if t > max_landing_thick:
                                            max_landing_thick = t
                                            connected_landing_id = shell_eid
                                            if shell_eid in landing_normals:
                                                landing_normal = landing_normals[shell_eid]

                                if landing_normal is not None and max_landing_thick > 0:
                                    offset_magnitude = max_landing_thick + (thickness / 2.0)
                                    offset_vector = -landing_normal * offset_magnitude

                                    bar_results.append({
                                        'Element_ID': eid,
                                        'Element_Type': elem.type,
                                        'Property_ID': elem.pid,
                                        'Property_Type': prop.type,
                                        'Bar_Thickness': thickness,
                                        'Connected_Landing_ID': connected_landing_id,
                                        'Landing_Thickness': max_landing_thick,
                                        'Offset_Magnitude': offset_magnitude,
                                        'Offset_X': offset_vector[0],
                                        'Offset_Y': offset_vector[1],
                                        'Offset_Z': offset_vector[2]
                                    })
                                else:
                                    bar_no_landing += 1
                            else:
                                bar_no_landing += 1

            self.log2(f"  Calculated offsets for {len(bar_results)} bar elements")
            if bar_no_landing > 0:
                self.log2(f"  Skipped {bar_no_landing} bars (no landing connection)")

            # Save to CSV
            output_dir = self.run_output_folder.get()
            os.makedirs(output_dir, exist_ok=True)
            csv_path = os.path.join(output_dir, self.offset_csv_name.get())

            with open(csv_path, 'w', newline='') as f:
                writer = csv.writer(f)

                writer.writerow(['LANDING OFFSETS'])
                writer.writerow(['Element_ID', 'Element_Type', 'Property_ID', 'Property_Type',
                                'Thickness', 'Zoffset'])
                for row in landing_results:
                    writer.writerow([row['Element_ID'], row['Element_Type'], row['Property_ID'],
                                   row['Property_Type'], row['Thickness'], row['Zoffset']])

                writer.writerow([])

                writer.writerow(['BAR OFFSETS'])
                writer.writerow(['Element_ID', 'Element_Type', 'Property_ID', 'Property_Type',
                                'Bar_Thickness', 'Connected_Landing_ID', 'Landing_Thickness',
                                'Offset_Magnitude', 'Offset_X', 'Offset_Y', 'Offset_Z'])
                for row in bar_results:
                    writer.writerow([row['Element_ID'], row['Element_Type'], row['Property_ID'],
                                   row['Property_Type'], row['Bar_Thickness'],
                                   row['Connected_Landing_ID'], row['Landing_Thickness'],
                                   row['Offset_Magnitude'], row['Offset_X'],
                                   row['Offset_Y'], row['Offset_Z']])

            self.log2(f"\nSaved: {csv_path}")

            # Auto-set the offset CSV path for apply step
            self.tab2_offset_csv.set(csv_path)

            self.log2(f"\nSummary:")
            self.log2(f"  Landing elements: {len(landing_results)}")
            self.log2(f"  Bar elements: {len(bar_results)}")
            self.log2("="*70)

            self.root.after(0, lambda: messagebox.showinfo("Success",
                f"Offsets calculated!\n\nLanding: {len(landing_results)}\nBar: {len(bar_results)}\n\nCSV: {self.offset_csv_name.get()}"))

        except Exception as e:
            self.log2(f"\nERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn_calc_offset.config(state=tk.NORMAL)])


    def start_full_run(self):
        if not self.run_bdfs:
            messagebox.showerror("Error","Add BDF files"); return
        self.btn_full.config(state=tk.DISABLED)
        self.progress2.start()
        threading.Thread(target=self.do_full_run, daemon=True).start()
    
    def do_full_run(self):
        try:
            self.log2("="*60)
            self.log2("FULL RUN (All 4 Steps)")
            self.log2("="*60)
            out_folder = self.run_output_folder.get()
            os.makedirs(out_folder, exist_ok=True)
            
            # === STEP 1: Update Properties ===
            self.log2("\n>>> STEP 1: Update Properties")
            for bdf_path in self.run_bdfs:
                self.log2(f"  {os.path.basename(bdf_path)}")
                out_bdf = self.copy_bdf_to_output(bdf_path, out_folder)
                stats, _ = self.update_properties_in_file(out_bdf)
                self.log2(f"    PBARL={stats['pbarl']} PBAR={stats['pbar']} PSHELL={stats['pshell']} PCOMP={stats['pcomp']}")
            
            # === STEP 2: Run Nastran ===
            if self.nastran_path.get():
                self.log2("\n>>> STEP 2: Run Nastran")
                nastran = self.nastran_path.get()
                import subprocess
                import time
                
                bdf_files_in_output = [f for f in os.listdir(out_folder) if f.lower().endswith(('.bdf','.dat','.nas'))]
                
                for f in bdf_files_in_output:
                    bdf_full_path = os.path.join(out_folder, f)
                    self.log2(f"  Running: {f}")
                    
                    # Nastran'ı çalıştır ve bitene kadar bekle
                    try:
                        cmd = f'"{nastran}" "{bdf_full_path}" out="{out_folder}" scratch=yes batch=no'
                        process = subprocess.Popen(cmd, shell=True)
                        process.wait()  # Bitene kadar bekle
                        self.log2(f"    Completed: {f}")
                    except Exception as e:
                        self.log2(f"    Error running {f}: {e}")
                
                # OP2 dosyalarının oluşmasını bekle
                self.log2("  Waiting for OP2 files...")
                time.sleep(2)
            else:
                self.log2("\n>>> STEP 2: SKIPPED (No Nastran path)")
            
            # === STEP 3: Post-Process OP2 ===
            self.log2("\n>>> STEP 3: Post-Process OP2")
            op2_files = [os.path.join(out_folder, f) for f in os.listdir(out_folder) if f.lower().endswith('.op2')]
            
            if op2_files:
                # Element -> Property mapping ve PBARL dimensions
                elem_prop = {}
                pbarl_dims = {}
                
                # BDF'lerden oku
                bdf_files_to_read = list(self.run_bdfs)
                for f in os.listdir(out_folder):
                    if f.lower().endswith(('.bdf', '.dat', '.nas')):
                        bdf_files_to_read.append(os.path.join(out_folder, f))
                
                for bdf_path in bdf_files_to_read:
                    try:
                        bdf = BDF(debug=False)
                        bdf.read_bdf(bdf_path, validate=False, xref=False, read_includes=True)
                        
                        for eid, el in bdf.elements.items():
                            if hasattr(el, 'pid'):
                                elem_prop[eid] = el.pid
                        
                        for pid, prop in bdf.properties.items():
                            prop_type = prop.type
                            if prop_type == 'PBARL':
                                dims = prop.dim if hasattr(prop, 'dim') else []
                                bar_type = prop.bar_type if hasattr(prop, 'bar_type') else 'UNKNOWN'
                                if len(dims) >= 2:
                                    pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[1], 'type': bar_type}
                                elif len(dims) == 1:
                                    pbarl_dims[pid] = {'dim1': dims[0], 'dim2': dims[0], 'type': bar_type}
                            elif prop_type == 'PBAR':
                                area = prop.A if hasattr(prop, 'A') else None
                                if area:
                                    import math
                                    side = math.sqrt(area) if area > 0 else 0
                                    pbarl_dims[pid] = {'dim1': side, 'dim2': side, 'type': 'PBAR', 'area': area}
                    except:
                        pass
                
                self.log2(f"  Elements: {len(elem_prop)}, Bar properties: {len(pbarl_dims)}")
                
                results = []
                for op2_path in op2_files:
                    self.log2(f"  Processing: {os.path.basename(op2_path)}")
                    try:
                        op2 = OP2(debug=False)
                        op2.read_op2(op2_path)
                        if hasattr(op2, 'cbar_force') and op2.cbar_force:
                            for sc_id, force in op2.cbar_force.items():
                                for i, eid in enumerate(force.element):
                                    axial = force.data[0,i,6] if len(force.data.shape)==3 else force.data[i,6]
                                    pid = elem_prop.get(eid)
                                    d1 = d2 = area = stress = None
                                    
                                    if pid and pid in self.bar_properties:
                                        d1 = self.bar_properties[pid]['dim1']
                                        d2 = self.bar_properties[pid]['dim2']
                                        area = d1 * d2
                                        if area > 0: stress = axial / area
                                    elif pid and pid in pbarl_dims:
                                        prop_info = pbarl_dims[pid]
                                        d1 = prop_info['dim1']
                                        d2 = prop_info['dim2']
                                        area = prop_info.get('area', d1 * d2)
                                        if area > 0: stress = axial / area
                                    
                                    results.append({'OP2': os.path.basename(op2_path), 'Subcase': sc_id, 'Element': eid,
                                        'Property': pid, 'Axial': axial, 'Dim1': d1, 'Dim2': d2, 'Area': area, 'Stress': stress})
                    except Exception as e:
                        self.log2(f"    Error: {e}")
                
                csv_path = os.path.join(out_folder, self.csv_output_name.get())
                with open(csv_path, 'w', newline='') as f:
                    w = csv.DictWriter(f, fieldnames=['OP2','Subcase','Element','Property','Axial','Dim1','Dim2','Area','Stress'])
                    w.writeheader()
                    w.writerows(results)
                self.log2(f"  Saved: {self.csv_output_name.get()} ({len(results)} rows)")
            else:
                self.log2("  No OP2 files found!")
            
            # === STEP 4: Combine Stress ===
            self.log2("\n>>> STEP 4: Combine Stress")
            if self.residual_strength_df is not None:
                stress_csv = os.path.join(out_folder, self.csv_output_name.get())
                if os.path.exists(stress_csv):
                    stress_df = pd.read_csv(stress_csv)
                    lookup = {}
                    for _, row in stress_df.iterrows():
                        key = (int(row['Subcase']), int(row['Element']))
                        lookup[key] = row['Stress'] if pd.notna(row['Stress']) else 0
                    elements = stress_df['Element'].unique()
                    
                    rs_df = self.residual_strength_df
                    cols = rs_df.columns.tolist()
                    comb_col = cols[0]
                    
                    comp_cols = []
                    idx = 1
                    while idx < len(cols) - 1:
                        col_name = str(cols[idx]).upper()
                        next_col_name = str(cols[idx+1]).upper()
                        if ('CASE' in col_name or 'ID' in col_name) and 'MULT' in next_col_name:
                            comp_cols.append((cols[idx], cols[idx+1]))
                            idx += 2
                        else:
                            idx += 1
                    
                    comb_results = []
                    for _, rs_row in rs_df.iterrows():
                        comb_lc = rs_row[comb_col]
                        if pd.isna(comb_lc): continue
                        comb_lc = int(comb_lc)
                        
                        for eid in elements:
                            total_stress = 0.0
                            components = []
                            for case_col, mult_col in comp_cols:
                                case_id = rs_row[case_col]
                                multiplier = rs_row[mult_col]
                                if pd.isna(case_id) or pd.isna(multiplier): continue
                                case_id = int(case_id)
                                multiplier = float(multiplier)
                                key = (case_id, int(eid))
                                if key in lookup:
                                    stress_val = lookup[key]
                                    if stress_val is not None:
                                        total_stress += stress_val * multiplier
                                        components.append(f"{case_id}*{multiplier}")
                            if components:
                                comb_results.append({'Combined_LC': comb_lc, 'Element': eid, 
                                    'Combined_Stress': total_stress, 'Components': ' + '.join(components)})
                    
                    comb_csv = os.path.join(out_folder, self.combined_csv_name.get())
                    with open(comb_csv, 'w', newline='') as f:
                        w = csv.DictWriter(f, fieldnames=['Combined_LC','Element','Combined_Stress','Components'])
                        w.writeheader()
                        w.writerows(comb_results)
                    self.log2(f"  Saved: {self.combined_csv_name.get()} ({len(comb_results)} rows)")
                else:
                    self.log2("  Stress CSV not found!")
            else:
                self.log2("  SKIPPED (No Residual Strength data loaded)")
            
            self.log2("\n" + "="*60)
            self.log2("FULL RUN COMPLETED!")
            self.log2("="*60)
            self.root.after(0, lambda: messagebox.showinfo("Done","Full run completed!"))
        except Exception as e:
            self.log2(f"ERROR: {e}")
            import traceback
            self.log2(traceback.format_exc())
        finally:
            self.root.after(0, lambda: [self.progress2.stop(), self.btn_full.config(state=tk.NORMAL)])



def main():
    root = tk.Tk()
    app = IntegratedBDFRFTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
