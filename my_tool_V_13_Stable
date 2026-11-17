# All_In_One_Network_Tool_merged.py
# For EXE (example):
#   pyinstaller --onefile --noconsole All_In_One_Network_Tool_merged.py
#   python -m PyInstaller --onefile --noconsole --clean --strip --upx-dir C:\upx-5.0.2-win64 Master_tool.py

#py -m PyInstaller --onefile --noconsole --collect-all paramiko --collect-all netmiko --collect-all openpyxl my_tool_V_7_Stable.py
import sys
import os
import re
import subprocess
import threading
import time
import json
import queue
import threading
import platform
import matplotlib.pyplot as plt
import winsound

import tkinter as tk
import paramiko
from tkinter import ttk, filedialog, messagebox, scrolledtext
from concurrent.futures import ThreadPoolExecutor, as_completed
from netmiko import ConnectHandler
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook, load_workbook
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import ipaddress
from email.mime.text import MIMEText
from openpyxl.styles import Alignment, Font, PatternFill
import threading, time, paramiko, re, os, smtplib
import pystray
from PIL import Image, ImageDraw
import threading
from tkinter import simpledialog, messagebox


TRAY_PASSWORD = "love"

def create_tray_icon():
    image = Image.new('RGB', (64, 64), "black")
    d = ImageDraw.Draw(image)
    d.text((10, 20), "CP", fill="white")
    return image


def setup_tray(root):

    def show_window(icon, item):

        def ask_password():
            pwd = simpledialog.askstring("Password", "Enter password:", show="*")

            if pwd == TRAY_PASSWORD:
                root.deiconify()
                icon.stop()
            else:
                messagebox.showerror("Error", "Wrong Password")

        # ✅ Run in main thread
        root.after(0, ask_password)


    def exit_app(icon, item):
        root.after(0, root.destroy)
        icon.stop()


    icon = pystray.Icon(
        "CPMonitor",
        create_tray_icon(),
        "My Tool v5",
        menu=pystray.Menu(
            pystray.MenuItem("Show", show_window),
            pystray.MenuItem("Exit", exit_app)
        )
    )

    threading.Thread(target=icon.run, daemon=True).start()


def hide_to_tray(root):
    root.withdraw()
    setup_tray(root)


  #self.email_enabled = True
# optional imports (fail gracefully for Tkinter parts)
try:
    import pandas as pd
except Exception:
    pd = None

try:
    from netmiko import ConnectHandler
except Exception:
    ConnectHandler = None

# ----- Optional PyQt5 imports for VPRN Collector -----
try:
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
    from PyQt5.QtWidgets import (
        QApplication, QWidget, QVBoxLayout, QHBoxLayout,
        QLabel, QPushButton, QTableWidget, QTableWidgetItem,
        QPlainTextEdit, QFileDialog as QtFileDialog, QMessageBox as QtMessageBox,
        QProgressBar as QtProgressBar
    )
    HAS_PYQT5 = True
except Exception:
    HAS_PYQT5 = False

#---------------- App settings ----------------
APP_TITLE = "All-In-One Network Tools_(FOR_NOKIA_ROUTERS_7750 &7250_Version_10.1.0)"
APP_GEOMETRY = "1200x820"
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "settings.json")

# =====================================================
# ========== SHARED UTILITY FUNCTIONS (TK) ============
# =====================================================

def load_settings():
    try:
        with open(SETTINGS_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {"dark_mode": False}

def save_settings(data: dict):
    try:
        with open(SETTINGS_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def save_text(parent, content: str, defname: str):
    if not content.strip():
        messagebox.showinfo("Nothing to save", "No content to save.", parent=parent)
        return
    path = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".txt",
        initialfile=defname,
        filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    if path:
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo("Saved", f"Saved to {path}", parent=parent)
        except Exception as e:
            messagebox.showerror("Error", f"Failed saving file:\n{e}", parent=parent)

def save_single_sheet_excel(parent, sheet_name: str, content: str, defname: str):
    if pd is None:
        messagebox.showwarning("Excel export unavailable",
                               "Install pandas + openpyxl to export Excel.",
                               parent=parent)
        return
    path = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".xlsx",
        initialfile=defname,
        filetypes=[("Excel", "*.xlsx")]
    )
    if not path:
        return
    try:
        lines = [l for l in content.splitlines()]
        df = pd.DataFrame({"Output": lines})
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=(sheet_name or "Sheet")[:31])
        messagebox.showinfo("Saved", f"Excel saved to {path}", parent=parent)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write Excel file:\n{e}", parent=parent)

def export_all_excel(parent, sheets_dict: dict):
    if pd is None:
        messagebox.showwarning("Excel export unavailable",
                               "Install pandas + openpyxl to export Excel.",
                               parent=parent)
        return
    path = filedialog.asksaveasfilename(
        parent=parent,
        defaultextension=".xlsx",
        initialfile="nokia_tools_export.xlsx",
        filetypes=[("Excel", "*.xlsx")]
    )
    if not path:
        return
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for sheet, content in sheets_dict.items():
                lines = [l for l in content.splitlines()]
                df = pd.DataFrame({"Output": lines})
                safe_sheet = (sheet or "Sheet")[:31] or "Sheet"
                df.to_excel(writer, index=False, sheet_name=safe_sheet)
        messagebox.showinfo("Saved", f"Excel saved to {path}", parent=parent)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write Excel file:\n{e}", parent=parent)

# ---------------- Theme Manager ----------------


class ThemeManager:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.dark = True

        # ☀️ Soft Colorful Light Theme
        self.colors_light = {
            "bg": "#685B5B",        # soft background
            "fg": "#0A0A0A",        # dark grey text
            "accent": "#5E81AC",    # muted blue

            "text_bg": "#020202",
            "text_fg": "#A2F10E",

            "entry_bg": "#020202",
            "entry_fg": "#A2F10E",

            "button_bg": "#4C9223", # soft green
            "button_fg": "#020202"
        }

        # 🌙 Colorful Dark Theme (eye comfort)
        self.colors_dark = {
            "bg": "#1E1E2E",        # dark bluish bg
            "fg": "#CDD6F4",        # soft text
            "accent": "#89B4FA",    # pastel blue

            "text_bg": "#252538",
            "text_fg": "#CDD6F4",

            "entry_bg": "#313244",
            "entry_fg": "#F5E0DC",

            "button_bg": "#A6E3A1", # calm green
            "button_fg": "#1E1E2E"
        }

        # Active theme
        self.colors = self.colors_dark if self.dark else self.colors_light
        self.style = ttk.Style(self.root)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass
        self.apply(False)

    def apply(self, dark: bool):
        self.dark = dark
        c = self.colors_dark if dark else self.colors_light
        self.root.configure(bg=c["bg"])
        self.style.configure("TFrame", background=c["bg"])
        self.style.configure("TLabelframe", background=c["bg"], foreground=c["fg"])
        self.style.configure("TLabelframe.Label", background=c["bg"], foreground=c["fg"])
        self.style.configure("TLabel", background=c["bg"], foreground=c["fg"])
        self.style.configure("TButton", foreground=c["fg"])
        self.style.map("TButton", background=[("active", c["accent"])]
                       )
        self.style.configure("TEntry", fieldbackground=c["entry_bg"], foreground=c["entry_fg"])
        self.style.configure("TNotebook", background=c["bg"])
        self.style.configure("TNotebook.Tab", background=c["bg"], foreground=c["fg"])

        def recolor_texts(widget):
            for w in widget.winfo_children():
                recolor_texts(w)
            if isinstance(widget, (tk.Text, scrolledtext.ScrolledText)):
                widget.configure(bg=c["text_bg"],
                                 fg=c["text_fg"],
                                 insertbackground=c["text_fg"])
        recolor_texts(self.root)

#class for vprn Genrator 

OUTPUT_DIR = "configs_output"


class VPRNConfigTab(ttk.Frame):
    def get_tab(self):
        return self.frame
    def __init__(self, parent):
        super().__init__(parent)
        self.entries = {}

        self.build_ui()

    # -------------------------------
    # Validation
    # -------------------------------
    def validate_ip(self, ip):
        try:
            ipaddress.ip_interface(ip)
            return True
        except:
            return False

    def validate_peer_ip(self, ip):
        try:
            ipaddress.ip_address(ip)
            return True
        except:
            return False

    def validate_sap(self, sap):
        return ":" in sap and all(part.isdigit() for part in sap.split(":"))

    # -------------------------------
    # Config Generator
    # -------------------------------
    def generate_config(self, data):
        service_id = int(data['service_id'])
        name = data['name'].strip()
        interface_name = data['interface_name'].strip()
        local_ip = data['local_ip'].strip()
        peer_ip = data['peer_ip'].strip()
        peer_as = int(data['peer_as'])
        local_as = int(data['local_as'])
        sap = data['sap'].strip()
        rd_ip = data['rd_ip'].strip()

        NAME = name.upper()
        
        # #name = ['cups-cp','inet,bsnl.in','ftth.bsnl.in','lfmt.in','ftth.lfmt.in','operation.in'
        # ,'voip-cdot','voip-ims','tpartyoltmgmt.in','bbnlinband-ll','wbpostal.in','bsnlintranet.in'
        # ,'cbecapis.in','cbec.com','corpbank.in','districtlibrarywb.in'
        # ,'ecourt.nic.in','enfr.in','erpbsnl.in','cdreastzone.com','idbi.com','iob.com','lic.com'
        # ,'muthoot.in','nacilbb.in','uiicadsl.in','wb.cctns.in','wbrto.nic.in','WIFI-offload','wimax'
        # ,'sancharsofteast.in','intranet.etp.in','kolpostal.in','kmcgov.in','mcx.com','globalcomn.com','tr69.in']

        return f'''/configure service vprn {service_id} name "{name}" customer 1 create
    autonomous-system {local_as}

    interface "{interface_name}" create
        address {local_ip}
        sap lag-{sap} create
    exit
    bgp-ipvpn
        mpls
            auto-bind-tunnel
                resolution-filter
                    ldp
                exit
                resolution any
            exit
            route-distinguisher {rd_ip}:{service_id}
            vrf-import "{NAME}-IMPORT"
            vrf-export "{NAME}-EXPORT"
            no shutdown
        exit
    exit

    group "EBGP-MPLS-{NAME}"
        family ipv4
        type external
        export "BGP-EXPORT-{NAME}"
        peer-as {peer_as}
        local-as {local_as}
        neighbor {peer_ip}
            description {interface_name}
        exit
    exit
    no shutdown
exit
no shutdown

/configure router policy-options 
begin
    community "{NAME}"
    members "target:{local_as}:{service_id}"
    exit

policy-statement "{NAME}-EXPORT"
    entry 10
        from
            protocol direct
                exit
                action accept
                    community add "{NAME}"
                    exit
                exit
                entry 20
                    from
                        protocol managed
                    exit
                    action accept
                        community add "{NAME}"
                    exit
                exit
                entry 30
                    from
                        protocol bgp
                    exit
                    action accept
                        community add "{NAME}"
                    exit
                exit
policy-statement "{NAME}-EXPORT"
                entry 10
                    from
                        protocol direct
                    exit
                    action accept
                        community add "{NAME}"
                    exit
                exit
                entry 20
                    from
                        protocol managed
                    exit
                    action accept
                        community add "{NAME}"
                    exit
                exit
                entry 30
                    from
                        protocol bgp
                    exit
                    action accept
                        community add "{NAME}"
                    exit
                exit
              exit    
policy-statement "BGP-EXPORT-{NAME}"
                entry 10
                    from
                        protocol direct
                    exit
                    action accept
                    exit
                exit
                entry 20
                    from
                        protocol managed
                    exit
                    action accept
                    exit
                exit
                entry 30
                    from
                        protocol bgp-vpn
                    exit
                    action accept
                        as-path-prepend {local_as} 2
                    exit
                exit
                entry 40
                    from
                        prefix-list "TEJAS-POOL"
                    exit
                    action accept
                    exit
                exit
                default-action drop
                exit                
'''

    # -------------------------------
    # Preview Window
    # -------------------------------
    def show_preview(self, text, filename):
        win = tk.Toplevel()
        win.title("Config Preview")
        win.geometry("800x500")

        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        txt = tk.Text(frame, wrap="none")
        txt.insert("1.0", text)

        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=txt.xview)

        txt.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        txt.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        def save():
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            with open(os.path.join(OUTPUT_DIR, filename), "w") as f:
                f.write(text)
            messagebox.showinfo("Saved", f"Saved as {filename}")

        ttk.Button(win, text="💾 Save", command=save).pack(pady=5)

    # -------------------------------
    # Single Mode
    # -------------------------------
    def preview_single(self):
        data = {k: self.entries[k].get() for k in self.entries}

        if not self.validate_ip(data['local_ip']):
            messagebox.showerror("Error", "Invalid Local IP")
            return

        if not self.validate_peer_ip(data['peer_ip']):
            messagebox.showerror("Error", "Invalid Peer IP")
            return

        if not self.validate_sap(data['sap']):
            messagebox.showerror("Error", "Invalid SAP (use 20:4027)")
            return

        config = self.generate_config(data)
        filename = f"{data['name']}_{data['service_id']}.txt"

        self.show_preview(config, filename)

    # -------------------------------
    # Bulk Mode
    # -------------------------------
    def browse(self):
        file = filedialog.askopenfilename(filetypes=[("CSV/Excel", "*.csv *.xlsx")])
        self.entry_file.delete(0, tk.END)
        self.entry_file.insert(0, file)

    def preview_bulk(self):
        file = self.entry_file.get()
        if not file:
            messagebox.showerror("Error", "Select file")
            return

        df = pd.read_csv(file) if file.endswith(".csv") else pd.read_excel(file)

        text = ""
        for _, row in df.head(5).iterrows():
            if not self.validate_sap(str(row.get("sap", ""))):
                continue
            text += self.generate_config(row) + "\n" + "-" * 40 + "\n"

        self.show_preview(text, "bulk_preview.txt")

    # -------------------------------
    # UI Builder
    # -------------------------------
    def build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        # Bulk Tab
        bulk_tab = ttk.Frame(notebook)
        notebook.add(bulk_tab, text="📂 Bulk Mode")

        self.entry_file = ttk.Entry(bulk_tab, width=50)
        self.entry_file.grid(row=0, column=0, padx=10, pady=10)

        ttk.Button(bulk_tab, text="Browse", command=self.browse).grid(row=0, column=1)
        ttk.Button(bulk_tab, text="Preview", command=self.preview_bulk).grid(row=1, column=0, pady=10)

        # Single Tab
        single_tab = ttk.Frame(notebook)
        notebook.add(single_tab, text="⚙️ Single Mode")

        fields = [
            'service_id', 'name', 'interface_name',
            'local_ip', 'peer_ip', 'peer_as', 'local_as', 'sap', 'rd_ip'
        ]

        for i, field in enumerate(fields):
            ttk.Label(single_tab, text=field).grid(row=i, column=0, padx=10, pady=5, sticky="w")
            e = ttk.Entry(single_tab)
            e.grid(row=i, column=1, padx=10, pady=5)
            self.entries[field] = e

        ttk.Button(single_tab, text="Preview Config", command=self.preview_single)\
            .grid(row=len(fields), column=0, columnspan=2, pady=15)




# ---------------- SimpleInOutTab base ----------------
class SimpleInOutTab(ttk.Frame):
    SHEET_NAME = ""

    def __init__(self, master, title):
        super().__init__(master)
        self.title = title
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        left = ttk.Frame(self)
        left.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Label(left, text=f"{self.title} – Input").pack(anchor="w")
        self.input = scrolledtext.ScrolledText(left, wrap="word", height=20)
        self.input.pack(fill="both", expand=True)

        right = ttk.Frame(self)
        right.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)
        ttk.Label(right, text=f"{self.title} – Output").pack(anchor="w")
        self.output = scrolledtext.ScrolledText(right, wrap="word", height=20)
        self.output.pack(fill="both", expand=True)

        row = ttk.Frame(self)
        row.grid(row=1, column=0, columnspan=2, pady=6)
        ttk.Button(row, text="Generate", command=self.generate).pack(side="left", padx=4)
        ttk.Button(row, text="Save TXT", command=self.save_txt).pack(side="left", padx=4)
        ttk.Button(row, text="Save Excel", command=self.save_excel).pack(side="left", padx=4)

    def generate(self):
        pass

    def save_txt(self):
        save_text(self, self.output.get("1.0", tk.END),
                  f"{self.SHEET_NAME or 'output'}.txt")

    def save_excel(self):
        save_single_sheet_excel(self, self.SHEET_NAME or "Sheet",
                                self.output.get("1.0", tk.END),
                                f"{self.SHEET_NAME or 'output'}.xlsx")

# =====================================================
# ============== O&M / SERVICE TABS (TK) ==============
# (Same as your Master_tool.py)
# =====================================================

class VPLSGenTab(ttk.Frame):
    SHEET_NAME = "VPLS Generator"

    def __init__(self, master):
        super().__init__(master)
        self._build()

    def _build(self):
        box = ttk.Labelframe(self, text="Service / VPLS Generator", padding=8)
        box.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        for i in range(4):
            box.columnconfigure(i, weight=1)

        def add(label, r, c, key):
            ttk.Label(box, text=label).grid(row=r, column=2*c,
                                            sticky="w", padx=4, pady=2)
            e = tk.Entry(box, width=22, bg = "#1E1E1E", fg = "#D4D4D4", insertbackground="white")
            e.grid(row=r, column=2*c+1, sticky="ew", padx=6, pady=2)
            setattr(self, key, e)

        add("OLT VLAN:", 0, 0, "e_vlan")
        add("First Downlink Port:", 0, 1, "e_first")
        add("Second Downlink Port:", 1, 0, "e_second")
        add("Spoke SDP ID:", 1, 1, "e_spoke")

        rowb = ttk.Frame(self)
        rowb.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        ttk.Button(rowb, text="Generate Config",
                   command=self.generate).pack(side="left", padx=4)
        ttk.Button(rowb, text="Save TXT",
                   command=self.save_txt).pack(side="left", padx=4)
        ttk.Button(rowb, text="Save Excel",
                   command=self.save_excel).pack(side="left", padx=4)

        ttk.Label(self, text="Preview:").grid(row=2, column=0,
                                              sticky="w", padx=6)
        self.preview = scrolledtext.ScrolledText(self, height=18,bg = "#1E1E1E",fg = "#D4D4D4")
        self.preview.grid(row=3, column=0, sticky="nsew", padx=6, pady=4,)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)

    def build_config(self):
        olt_vlan = self.e_vlan.get().strip()
        first_port = self.e_first.get().strip()
        second_port = self.e_second.get().strip()
        spoke_sdp = self.e_spoke.get().strip()

        if not all([olt_vlan, first_port, second_port, spoke_sdp]):
            raise ValueError("Please fill VLAN, First/Second Downlink Port and Spoke SDP ID.")

        vpls_name = f"1{first_port}{olt_vlan}"
        service_name = f"vpls-{olt_vlan}"
        split_group = service_name

        config = f"""
/configure service vpls {vpls_name} name "{service_name}" customer 1 create
    description "{service_name}"
    service-mtu 9014
    split-horizon-group "{split_group}" create
    exit
    stp
        shutdown
    exit
    sap lag-10:{olt_vlan}.cp-1 create
        no shutdown
    exit
    sap lag-{first_port}:{olt_vlan}.cp-11{first_port} split-horizon-group "{split_group}" create
        ingress
            qos 10
        exit
        no shutdown
    exit
    sap lag-{second_port}:{olt_vlan}.cp-11{second_port} split-horizon-group "{split_group}" create
        ingress
            qos 10
        exit
        no shutdown
    exit
    spoke-sdp 100:{spoke_sdp} create
        no shutdown
    exit
    no shutdown
    info
    exit all 
    admin save 
""".strip("\n")
        return config

    def generate(self):
        try:
            cfg = self.build_config()
            self.preview.delete("1.0", tk.END)
            self.preview.insert(tk.END, cfg)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def save_txt(self):
        save_text(self, self.preview.get("1.0", tk.END), "VPLS_Generator.txt")

    def save_excel(self):
        save_single_sheet_excel(self, self.SHEET_NAME,
                                self.preview.get("1.0", tk.END),
                                "VPLS_Generator.xlsx")

    def get_output_text(self):
        return self.preview.get("1.0", tk.END)

class SDPFilterTab(SimpleInOutTab):
    SHEET_NAME = "SDPFilter"

    def __init__(self, master):
        super().__init__(master, "Unused SDP Finder")

        # (Sort option kept but forced to ASC for this use-case)
        self.sort_order = tk.StringVar(value="ASC")

        option_frame = ttk.Frame(self)
        option_frame.grid(row=1, column=0, sticky="w", padx=6, pady=4)

        ttk.Label(option_frame, text="Mode:").grid(row=0, column=0, padx=4)

        ttk.Label(
            option_frame,
            text="Finding UNUSED SDP IDs (Low → High)",
            foreground="blue"
        ).grid(row=0, column=1, padx=4)

        # Precompile regex
        self.sdp_pattern = re.compile(r"^\d+:\d+$")

    # -------------------- MAIN --------------------
    def generate(self):
        raw_data = self.input.get("1.0", tk.END).strip()

        if not raw_data:
            self.output.delete("1.0", tk.END)
            self.output.insert(tk.END, "⚠️ No input provided")
            return

        self.output.delete("1.0", tk.END)

        used_sdps = {}

        # ------------------ PARSE INPUT ------------------
        for line in raw_data.splitlines():
            line = line.strip()

            # Skip unwanted lines
            if not line or line.startswith(("=", "-", "SvcId", "Number")):
                continue

            parts = line.split()

            if len(parts) >= 2:
                sdp = parts[1]

                if self.sdp_pattern.match(sdp):
                    a, b = map(int, sdp.split(":"))
                    used_sdps.setdefault(a, set()).add(b)

        if not used_sdps:
            self.output.insert(tk.END, "⚠️ No SDP IDs found")
            return

        # ------------------ FIND UNUSED ------------------
        unused_list = []

        for a in used_sdps:
            nums = sorted(used_sdps[a])

            min_val = nums[0]
            max_val = nums[-1]

            full_range = set(range(min_val, max_val + 1))
            missing = sorted(full_range - set(nums))

            for m in missing:
                unused_list.append(f"{a}:{m}")

        if not unused_list:
            self.output.insert(tk.END, "✅ No unused SDP IDs (continuous range)")
            return

        # ------------------ SORT (LOW → HIGH ONLY) ------------------
        def sort_key(x):
            a, b = x.split(":")
            return int(a), int(b)

        unused_list = sorted(unused_list, key=sort_key)

        # ------------------ OUTPUT ------------------
        final_output = (
            f"Unused SdpId (Total: {len(unused_list)})\n"
            + "\n".join(unused_list)
        )

        self.output.insert(tk.END, final_output)

class LocalASFinderTab(ttk.Frame):

    SHEET_NAME = "LocalAS"

    def __init__(self, master):
        super().__init__(master)

        self.excel_path = ""
        self.results = []
        self.ui_queue = queue.Queue()

        # ---------------- TITLE ----------------
        tk.Label(self, text="Local-AS Finder (Parallel SSH)", 
                 font=("Arial", 12, "bold")).pack(pady=5)

        # ---------------- LOAD EXCEL ----------------
        tk.Button(self, text="Load Excel File", bg="blue", fg="white",
                  command=self.load_excel).pack(pady=5)

        self.excel_label = tk.Label(self, text="No file loaded")
        self.excel_label.pack()

        # ---------------- COMMAND ----------------
        tk.Label(self, text="\nCommand to Execute on Devices:").pack(anchor="w")

        self.command_text = scrolledtext.ScrolledText(self, height=4)
        self.command_text.pack(fill="both", padx=5)

        default_cmd = 'show router bgp summary | match context all "Local AS:"'
        self.command_text.insert(tk.END, default_cmd)

        tk.Button(self, text="Run Command", bg="green", fg="white",
                  command=self.run_command).pack(pady=10)

        # ---------------- SAVE BUTTONS ----------------
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=5)

        tk.Button(btn_frame, text="Save Output As...",
                  bg="orange", command=self.save_output_as).pack(side="left", padx=10)

        tk.Button(btn_frame, text="Save Results as Excel",
                  bg="purple", fg="white",
                  command=self.save_results_as_excel).pack(side="left", padx=10)

        # ---------------- OUTPUT ----------------
        tk.Label(self, text="Output:").pack(anchor="w")

        self.output_box = scrolledtext.ScrolledText(self, height=20)
        self.output_box.pack(fill="both", padx=5, pady=5)

        self.process_ui_queue()

    # ---------------- SAFE PRINT ----------------
    def safe_print(self, msg):
        self.ui_queue.put(msg)

    def process_ui_queue(self):
        while not self.ui_queue.empty():
            msg = self.ui_queue.get()
            self.output_box.insert(tk.END, msg)
            self.output_box.see(tk.END)
        self.after(100, self.process_ui_queue)

    # ---------------- LOAD EXCEL ----------------
    def load_excel(self):
        self.excel_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if self.excel_path:
            self.excel_label.config(text=f"Loaded: {self.excel_path}")

    # ---------------- EXTRACT LOCAL AS ----------------
    def extract_local_as(self, output):
        match = re.search(r"Local AS[: ]+(\d+)", output, re.IGNORECASE)
        return match.group(1) if match else "NOT FOUND"

    # ---------------- RUN COMMAND ----------------
    def run_command(self):

        if not self.excel_path:
            messagebox.showerror("Error", "Please load an Excel file first.")
            return

        command = self.command_text.get("1.0", tk.END).strip()
        if not command:
            messagebox.showerror("Error", "Please enter a command.")
            return

        try:
            df = pd.read_excel(self.excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read Excel: {e}")
            return

        required_cols = {"device_ip", "username", "password", "device_name"}
        if not required_cols.issubset(df.columns):
            messagebox.showerror(
                "Error",
                f"Excel must include: {', '.join(required_cols)}"
            )
            return

        if "hostname" not in df.columns:
            df["hostname"] = ""
        else:
            df["hostname"] = df["hostname"].fillna("")

        self.results.clear()
        self.safe_print("\n=== Starting Parallel Execution (20 Threads) ===\n")

        def connect_and_run(row):
            ip = str(row["device_ip"]).strip()
            username = str(row["username"]).strip()
            password = str(row["password"]).strip()
            device_name = str(row["device_name"]).strip()
            hostname = str(row["hostname"]).strip() or ip

            self.safe_print(f"\n[{ip}] Connecting... Hostname='{hostname}'\n")

            try:
                ssh = paramiko.SSHClient()
                ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                ssh.connect(ip, username=username, password=password,
                            timeout=10, allow_agent=False, look_for_keys=False)

                chan = ssh.invoke_shell()
                time.sleep(0.5)
                chan.send(command + "\n")
                time.sleep(2)

                output = ""
                while chan.recv_ready():
                    output += chan.recv(65535).decode(errors="ignore")

                ssh.close()
                self.safe_print(f"--- {ip} OUTPUT ---\n{output}\n")
                local_as = self.extract_local_as(output)

            except Exception as e:
                self.safe_print(f"[{ip}] ERROR: {e}\n")
                local_as = "ERROR"

            return {
                "device_ip": ip,
                "device_name": device_name,
                "hostname": hostname,
                "local_as": local_as
            }

        with ThreadPoolExecutor(max_workers=20) as executor:
            futures = [executor.submit(connect_and_run, row)
                       for _, row in df.iterrows()]

            for future in as_completed(futures):
                self.results.append(future.result())

        unique_df = pd.DataFrame(self.results)\
            .drop_duplicates(subset=["local_as"])

        unique_df.to_excel("output_local_as.xlsx", index=False)

        self.safe_print(
            "\n=== Completed. Saved unique Local AS to output_local_as.xlsx ===\n"
        )

    # ---------------- SAVE OUTPUT ----------------
    def save_output_as(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")]
        )
        if file_path:
            content = self.output_box.get("1.0", tk.END)
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(content)

    # ---------------- SAVE RESULTS ----------------
    def save_results_as_excel(self):
        if not self.results:
            messagebox.showwarning("No Data", "Run command first.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            pd.DataFrame(self.results)\
                .drop_duplicates(subset=["local_as"])\
                .to_excel(file_path, index=False)

class PingMonitorTab(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)

        self.devices = []
        self.results = []
        self.stop_monitor = False
        self.history_up = []
        self.history_down = []

        # ================= UI =================
        top = tk.Frame(self)
        top.pack(pady=10)

        tk.Button(top, text="Load Excel", command=self.load_excel)\
            .pack(side=tk.LEFT, padx=5)

        tk.Button(top, text="Start", bg="#27AE60", fg="white",
                  command=self.start_monitor)\
            .pack(side=tk.LEFT, padx=5)

        tk.Button(top, text="Stop", bg="#E74C3C", fg="white",
                  command=self.stop)\
            .pack(side=tk.LEFT, padx=5)

        tk.Button(top, text="📊 Pie Chart",
                  command=self.show_pie_chart)\
            .pack(side=tk.LEFT, padx=5)

        tk.Button(top, text="📈 Live Graph",
                  command=self.show_live_graph)\
            .pack(side=tk.LEFT, padx=5)

        tk.Label(top, text="Refresh(sec)").pack(side=tk.LEFT, padx=5)

        self.refresh = tk.IntVar(value=60)
        tk.Entry(top, width=5, textvariable=self.refresh)\
            .pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(self)
        self.progress.pack(fill=tk.X, padx=20, pady=10)

        self.output = scrolledtext.ScrolledText(
            self, font=("Consolas", 10)
        )
        self.output.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.output.tag_config("up", foreground="green")
        self.output.tag_config("down", foreground="red")

        # ================= DASHBOARD =================
        dash_frame = tk.LabelFrame(self, text="Dashboard")
        dash_frame.pack(fill=tk.X, padx=20, pady=10)

        self.total_var = tk.IntVar(value=0)
        self.up_var = tk.IntVar(value=0)
        self.down_var = tk.IntVar(value=0)
        self.up_percent = tk.DoubleVar(value=0)
        self.down_percent = tk.DoubleVar(value=0)

        tk.Label(dash_frame, text="Total Devices:")\
            .grid(row=0, column=0, padx=10, pady=5)
        tk.Label(dash_frame, textvariable=self.total_var)\
            .grid(row=0, column=1)

        tk.Label(dash_frame, text="UP Devices:")\
            .grid(row=1, column=0)
        tk.Label(dash_frame, textvariable=self.up_var)\
            .grid(row=1, column=1)

        tk.Label(dash_frame, text="DOWN Devices:")\
            .grid(row=2, column=0)
        tk.Label(dash_frame, textvariable=self.down_var)\
            .grid(row=2, column=1)

        tk.Label(dash_frame, text="UP %:")\
            .grid(row=0, column=2)
        tk.Label(dash_frame, textvariable=self.up_percent)\
            .grid(row=0, column=3)

        tk.Label(dash_frame, text="DOWN %:")\
            .grid(row=1, column=2)
        tk.Label(dash_frame, textvariable=self.down_percent)\
            .grid(row=1, column=3)

    # ---------- Load Excel ----------
    def load_excel(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file:
            return

        df = pd.read_excel(file)
        if not {"Device Name", "IP Address"}.issubset(df.columns):
            messagebox.showerror(
                "Error",
                "Excel must have Device Name & IP Address"
            )
            return

        self.devices = df.to_dict("records")
        self.total_var.set(len(self.devices))
        messagebox.showinfo("Loaded",
                            f"{len(self.devices)} devices loaded")

    # ---------- Ping (NO CMD POPUP) ----------
    def ping_device(self, ip):
        param = "-n" if platform.system().lower() == "windows" else "-c"
        cmd = ["ping", param, "1", ip]

        if platform.system().lower() == "windows":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            out = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                startupinfo=startupinfo
            )
        else:
            out = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )

        if out.returncode == 0:
            match = re.search(r"time[=<]\s*(\d+)", out.stdout)
            latency = match.group(1) if match else "N/A"
            return "UP", latency

        return "DOWN", "N/A"

    # ---------- Start Monitoring ----------
    def start_monitor(self):

        if not self.devices:
            messagebox.showwarning("Warning", "Load Excel first")
            return

        self.stop_monitor = False
        self.output.delete("1.0", tk.END)
        self.progress["maximum"] = len(self.devices)

        def run():
            while not self.stop_monitor:

                self.results.clear()
                self.progress["value"] = 0

                for d in self.devices:

                    if self.stop_monitor:
                        break

                    status, latency = self.ping_device(
                        d["IP Address"]
                    )

                    line = f'{d["Device Name"]:15} {d["IP Address"]:15} {status:5} {latency} ms\n'
                    tag = "up" if status == "UP" else "down"

                    self.output.insert(tk.END, line, tag)
                    self.output.see(tk.END)

                    if status == "DOWN":
                        try:
                            winsound.Beep(1000, 300)
                        except:
                            pass

                    self.results.append({
                        "Device Name": d["Device Name"],
                        "IP Address": d["IP Address"],
                        "Status": status,
                        "Latency(ms)": latency
                    })

                    self.progress["value"] += 1

                self.update_dashboard()
                time.sleep(self.refresh.get())

        threading.Thread(target=run, daemon=True).start()

    # ---------- Stop ----------
    def stop(self):
        self.stop_monitor = True

    # ---------- Dashboard ----------
    def update_dashboard(self):
        up_count = sum(1 for r in self.results if r["Status"] == "UP")
        down_count = sum(1 for r in self.results if r["Status"] == "DOWN")
        total = len(self.results)

        self.up_var.set(up_count)
        self.down_var.set(down_count)

        if total > 0:
            self.up_percent.set(round((up_count / total) * 100, 2))
            self.down_percent.set(round((down_count / total) * 100, 2))

        self.history_up.append(up_count)
        self.history_down.append(down_count)

    # ---------- Pie Chart ----------
    def show_pie_chart(self):
        up = self.up_var.get()
        down = self.down_var.get()

        if up + down == 0:
            messagebox.showwarning("Warning", "No data yet")
            return

        plt.figure()
        plt.pie([up, down], labels=["UP", "DOWN"], autopct="%1.1f%%")
        plt.title("Device Status Distribution")
        plt.show()

    # ---------- Live Graph ----------
    def show_live_graph(self):
        if not self.history_up:
            messagebox.showwarning("Warning", "No monitoring data yet")
            return

        plt.figure()
        plt.plot(self.history_up)
        plt.title("UP Devices Over Time")
        plt.xlabel("Cycle")
        plt.ylabel("UP Count")
        plt.show()

#ScriptRunnerTab ##############################################
class ScriptRunnerTab(ttk.Frame):

    MAX_THREADS = 5

    DEVICE_TYPES = {
        "Nokia SR OS": "nokia_sros",
        "Cisco IOS": "cisco_ios",
        "Juniper JunOS": "juniper"
    }

    def __init__(self, master):
        super().__init__(master)

        self.devices_df = None
        self.stop_flag = False
        self.completed = 0
        self.failed = 0
        self.log_queue = queue.Queue()

        # ================= MAIN LAYOUT =================
        main = tk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=10)

        # ================= INPUT SIDE =================
        input_frame = tk.LabelFrame(main, text="INPUT",
                                    font=("Segoe UI", 10, "bold"))
        input_frame.pack(side="left", fill="y", padx=(0, 10))

        tk.Label(input_frame, text="Device Type").pack(anchor="w")

        self.device_type_var = tk.StringVar(value="Nokia SR OS")

        ttk.Combobox(
            input_frame,
            textvariable=self.device_type_var,
            values=list(self.DEVICE_TYPES.keys()),
            state="readonly"
        ).pack(fill="x", pady=3)

        self.preview_table = ttk.Treeview(
            input_frame,
            columns=("Router", "IP", "User"),
            show="headings",
            height=6
        )

        for col in ("Router", "IP", "User"):
            self.preview_table.heading(col, text=col)

        self.preview_table.pack(fill="x", pady=5)

        tk.Button(input_frame, text="Load Device Excel",
                  command=self.load_excel)\
            .pack(fill="x")

        tk.Label(input_frame, text="Commands").pack(anchor="w")

        self.text_command = scrolledtext.ScrolledText(
            input_frame, height=10
        )
        self.text_command.pack(fill="x")

        self.config_var = tk.BooleanVar()
        tk.Checkbutton(input_frame, text="Config Mode",
                       variable=self.config_var)\
            .pack(anchor="w")

        self.admin_var = tk.BooleanVar()
        tk.Checkbutton(input_frame,
                       text="Admin CLI Mode (Nokia)",
                       variable=self.admin_var)\
            .pack(anchor="w")

        tk.Button(input_frame, text="RUN",
                  bg="#16a34a", fg="white",
                  command=self.run_script)\
            .pack(fill="x", pady=5)

        tk.Button(input_frame, text="STOP",
                  bg="#dc2626", fg="white",
                  command=self.stop_execution)\
            .pack(fill="x")

        # ================= OUTPUT SIDE =================
        output_frame = tk.LabelFrame(main, text="OUTPUT",
                                     font=("Segoe UI", 10, "bold"))
        output_frame.pack(side="right", fill="both", expand=True)

        self.output_box = scrolledtext.ScrolledText(output_frame)
        self.output_box.pack(fill="both", expand=True)

        self.progress_var = tk.IntVar()
        self.progress_bar = ttk.Progressbar(
            output_frame, variable=self.progress_var
        )
        self.progress_bar.pack(fill="x", pady=5)

        self.summary_var = tk.StringVar()
        tk.Label(output_frame,
                 textvariable=self.summary_var)\
            .pack(anchor="w")

        self.after(200, self.process_queue)

    # ================= QUEUE =================
    def process_queue(self):
        while not self.log_queue.empty():
            msg, status = self.log_queue.get()

            self.output_box.insert(tk.END, msg)
            self.output_box.see(tk.END)

            if status == "ok":
                self.completed += 1
            elif status == "fail":
                self.failed += 1

            self.progress_var.set(self.completed + self.failed)
            self.update_summary()

        self.after(200, self.process_queue)

    # ================= HELPERS =================
    def update_summary(self):
        total = len(self.devices_df) if self.devices_df is not None else 0
        self.summary_var.set(
            f"Total: {total} | Completed: {self.completed} | Failed: {self.failed}"
        )

    def write_log(self, router, ip, output):
        os.makedirs("logs", exist_ok=True)
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        with open(f"logs/{router}_{ip}_{timestamp}.txt",
                  "w", encoding="utf-8") as f:
            f.write(output)

    # ================= LOAD EXCEL =================
    def load_excel(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not file:
            return

        self.devices_df = pd.read_excel(file)

        required = {"ip", "username", "password", "router_name"}
        if not required.issubset(self.devices_df.columns):
            messagebox.showerror(
                "Error",
                "Excel must contain ip, username, password, router_name"
            )
            self.devices_df = None
            return

        self.preview_table.delete(
            *self.preview_table.get_children()
        )

        for _, row in self.devices_df.iterrows():
            self.preview_table.insert(
                "", "end",
                values=(row["router_name"],
                        row["ip"],
                        row["username"])
            )

        self.progress_bar["maximum"] = len(self.devices_df)
        self.update_summary()

    # ================= EXECUTION =================
    def run_on_device(self, row, commands,
                      is_config, is_admin):

        if self.stop_flag:
            return

        router = row["router_name"]
        ip = row["ip"]

        try:
            device_type = self.DEVICE_TYPES.get(
                self.device_type_var.get(),
                "nokia_sros"
            )

            conn = ConnectHandler(
                device_type=device_type,
                host=ip,
                username=row["username"],
                password=row["password"],
                timeout=90
            )

            output = ""

            if is_config:
                for cmd in commands:
                    if self.stop_flag:
                        break
                    cmd_output = conn.send_config_set([cmd])
                    output += cmd_output
                    self.log_queue.put(
                        (f"\n[{router} - {ip}]# {cmd}\n{cmd_output}\n", None)
                    )

            elif is_admin and device_type == "nokia_sros":
                conn.send_command_timing("admin")

                for cmd in commands:
                    if self.stop_flag:
                        break
                    cmd_output = conn.send_command_timing(cmd)
                    output += cmd_output
                    self.log_queue.put(
                        (f"\n[{router} - {ip}]# {cmd}\n{cmd_output}\n", None)
                    )

                conn.send_command_timing("exit")

            else:
                for cmd in commands:
                    if self.stop_flag:
                        break
                    cmd_output = conn.send_command(cmd)
                    output += cmd_output
                    self.log_queue.put(
                        (f"\n[{router} - {ip}]# {cmd}\n{cmd_output}\n", None)
                    )

            conn.disconnect()
            self.write_log(router, ip, output)

            self.log_queue.put(
                (f"\n✅ {router} ({ip}) Completed\n", "ok")
            )

        except Exception as e:
            self.log_queue.put(
                (f"\n❌ {router} ({ip}) FAILED: {e}\n", "fail")
            )

    def run_script(self):

        if self.devices_df is None:
            messagebox.showerror("Error", "Load Excel first")
            return

        commands = self.text_command.get(
            "1.0", tk.END
        ).strip().splitlines()

        if not commands:
            messagebox.showerror("Error", "No commands found")
            return

        if self.config_var.get() and self.admin_var.get():
            messagebox.showerror(
                "Error",
                "Do not enable Config and Admin together"
            )
            return

        self.completed = self.failed = 0
        self.stop_flag = False
        self.progress_var.set(0)
        self.output_box.delete("1.0", tk.END)

        executor = ThreadPoolExecutor(
            max_workers=self.MAX_THREADS
        )

        for _, row in self.devices_df.iterrows():
            executor.submit(
                self.run_on_device,
                row,
                commands,
                self.config_var.get(),
                self.admin_var.get()
            )

    def stop_execution(self):
        self.stop_flag = True
        self.output_box.insert(
            tk.END, "\n⛔ Execution stopped\n")
###VPLS PORT MODIFY #####################
class LagReplaceTab(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.build_ui()

    def parse_and_generate(self, cfg, old_lag, new_lag, new_cp):
        import re

        vpls_list = []
        current = None
        lines = cfg.splitlines()

        for i, line in enumerate(lines):
            line = line.strip()

            match = re.search(r'vpls\s+(\d+)', line)
            if match:
                if current:
                    vpls_list.append(current)
                current = {"vpls_id": match.group(1), "saps": []}
                continue

            if not current:
                continue

            match = re.search(rf'sap\s+(lag-{old_lag}:\d+\.cp-\d+)', line)
            if match:
                sap_value = match.group(1)

                qos = "10"
                for j in range(i+1, min(i+6, len(lines))):
                    qmatch = re.search(r'qos\s+(\d+)', lines[j])
                    if qmatch:
                        qos = qmatch.group(1)
                        break

                current["saps"].append((sap_value, qos))

        if current:
            vpls_list.append(current)

        output = []

        for vpls in vpls_list:
            if not vpls["saps"]:
                continue

            vpls_id = vpls["vpls_id"]
            last4 = vpls_id[-4:]
            shg = f"OLT-{last4}"

            block = [f"/configure service vpls {vpls_id}"]

            for sap, qos in vpls["saps"]:
                new_sap = f"lag-{new_lag}:{last4}.cp-11{new_cp}"

                block.append(f"    sap {sap} shutdown")
                block.append(f"    no sap {sap}")
                block.append(f"    sap {new_sap} split-horizon-group \"{shg}\" create")
                block.append(f"        ingress")
                block.append(f"            qos {qos}")
                block.append(f"        exit")
                block.append(f"        no shutdown")
                block.append(f"    exit")
            block.append(f"    no shutdown")
            output.append("\n".join(block))

        return "\n\n".join(output)

    def generate_output(self):
        cfg = self.input_text.get("1.0", tk.END).strip()
        old_lag = self.oldlag_entry.get().strip()
        new_lag = self.newlag_entry.get().strip()
        new_cp = self.cp_entry.get().strip()

        if not cfg or not old_lag or not new_lag or not new_cp:
            messagebox.showerror("Error", "All fields required!")
            return

        result = self.parse_and_generate(cfg, old_lag, new_lag, new_cp)

        if not result:
            messagebox.showinfo("Result", "No matching OLD LAG found!")
            return

        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, result)

    def copy_output(self):
        self.clipboard_clear()
        self.clipboard_append(self.output_text.get("1.0", tk.END))
        messagebox.showinfo("Copied", "Output copied!")

    def build_ui(self):
        frame = ttk.Frame(self)
        frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame, text="Old LAG:").grid(row=0, column=0)
        self.oldlag_entry = ttk.Entry(frame, width=10)
        self.oldlag_entry.grid(row=0, column=1)

        ttk.Label(frame, text="New LAG:").grid(row=0, column=2)
        self.newlag_entry = ttk.Entry(frame, width=10)
        self.newlag_entry.grid(row=0, column=3)

        ttk.Label(frame, text="New CP:").grid(row=0, column=4)
        self.cp_entry = ttk.Entry(frame, width=10)
        self.cp_entry.grid(row=0, column=5)

        ttk.Button(frame, text="Generate",
                   command=self.generate_output).grid(row=0, column=6, padx=10)

        ttk.Button(frame, text="📋 Copy",
                   command=self.copy_output).grid(row=0, column=7)

        ttk.Label(self, text="Paste Config:").pack(anchor="w", padx=10)
        self.input_text = tk.Text(self, height=18)
        self.input_text.pack(fill="both", padx=10, pady=5)

        ttk.Label(self, text="Output (Only Changed SAPs):").pack(anchor="w", padx=10)
        self.output_text = tk.Text(self, height=18)
        self.output_text.pack(fill="both", padx=10, pady=5)


#EpipeGen######################################################
class EpipeGenTab(ttk.Frame):
    SHEET_NAME = "EpipeGen"

    def __init__(self, master):
        super().__init__(master)
        self._build()

    def _build(self):
        svc = ttk.Labelframe(self, text="Service / Epipe fields", padding=8)
        svc.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        for i in range(6):
            svc.columnconfigure(i, weight=1)

        def add(label, r, c, key, width=22, default=""):
            ttk.Label(svc, text=label).grid(row=r, column=2*c,
                                            sticky="w", padx=4, pady=2)
            e = ttk.Entry(svc, width=width)
            e.grid(row=r, column=2*c+1, sticky="ew", padx=6, pady=2)
            if default != "":
                e.insert(0, default)
            setattr(self, key, e)

        add("VLAN:",                  0, 0, "e_vlan")
        add("LAG (endpoint X):",      0, 1, "e_lagx")
        add("Spoke-SDP (endpoint Y):",1, 0, "e_spy")
        add("Spoke-SDP (endpoint X):",1, 1, "e_spx")

        rowb = ttk.Frame(self)
        rowb.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        ttk.Button(rowb, text="Generate Config",
                   command=self.generate).pack(side="left", padx=4)
        ttk.Button(rowb, text="Save TXT",
                   command=self.save_txt).pack(side="left", padx=4)
        ttk.Button(rowb, text="Save Excel",
                   command=self.save_excel).pack(side="left", padx=4)

        ttk.Label(self, text="Preview:").grid(row=2, column=0,
                                              sticky="w", padx=6)
        self.preview = scrolledtext.ScrolledText(self, height=12)
        self.preview.grid(row=3, column=0, sticky="nsew", padx=6, pady=4)

        dev = ttk.Labelframe(self, text="Device / SSH", padding=8)
        dev.grid(row=4, column=0, sticky="ew", padx=6, pady=6)
        for i in range(6):
            dev.columnconfigure(i, weight=1)

        ttk.Label(dev, text="Host/IP:").grid(row=0, column=0, sticky="w")
        self.d_host = ttk.Entry(dev)
        self.d_host.grid(row=0, column=1, sticky="ew", padx=4)

        ttk.Label(dev, text="Port:").grid(row=0, column=2, sticky="w")
        self.d_port = ttk.Entry(dev, width=6)
        self.d_port.grid(row=0, column=3, sticky="w", padx=4)
        self.d_port.insert(0, "22")

        ttk.Label(dev, text="Username:").grid(row=1, column=0, sticky="w")
        self.d_user = ttk.Entry(dev)
        self.d_user.grid(row=1, column=1, sticky="ew", padx=4)

        ttk.Label(dev, text="Password:").grid(row=1, column=2, sticky="w")
        self.d_pass = ttk.Entry(dev, show="*")
        self.d_pass.grid(row=1, column=3, sticky="ew", padx=4)

        ttk.Label(dev, text="Per-line delay (s):").grid(row=2, column=0, sticky="w")
        self.e_delay = ttk.Entry(dev, width=8)
        self.e_delay.grid(row=2, column=1, sticky="w", padx=4)
        self.e_delay.insert(0, "0.5")

        self.var_dry = tk.BooleanVar(value=True)
        ttk.Checkbutton(dev, text="Dry Run (do not connect)",
                        variable=self.var_dry).grid(row=2, column=2,
                                                    columnspan=2,
                                                    sticky="w", pady=4)

        send = ttk.Frame(self)
        send.grid(row=5, column=0, sticky="ew", padx=6, pady=4)
        ttk.Button(send, text="Send All-At-Once",
                   command=lambda: self.send(True)).pack(side="left", padx=4)
        ttk.Button(send, text="Send Line-By-Line",
                   command=lambda: self.send(False)).pack(side="left", padx=4)

        ttk.Label(self, text="Log / Device Output:").grid(row=6, column=0,
                                                          sticky="w", padx=6)
        self.log = scrolledtext.ScrolledText(self, height=10, state="disabled")
        self.log.grid(row=7, column=0, sticky="nsew", padx=6, pady=6)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=2)
        self.rowconfigure(7, weight=2)

    def build_config(self):
        vlan = self.e_vlan.get().strip()
        lagx = self.e_lagx.get().strip()
        spy = self.e_spy.get().strip()
        spx = self.e_spx.get().strip()
        if not (vlan and lagx and spy and spx):
            raise ValueError("Fill VLAN, LAGX, Spoke-SDP X and Y.")

        lines = [
            f'/configure service epipe 1{lagx}{vlan} name "OLT_{vlan}" customer 1 create',
            f'    description "OLT_{vlan}"',
            f'    endpoint "x" create',
            f'    exit',
            f'    endpoint "y" create',
            f'    exit',
            f'    service-mtu 9014',
            f'    sap lag-10:{vlan}.cp-1 endpoint "y" create',
            f'        no shutdown',
            f'    exit',
            f'    sap lag-{lagx}:{vlan}.cp-11{lagx} endpoint "x" create',
            f'        ingress',
            f'            qos 10',
            f'        exit',
            f'        no shutdown',
            f'    exit',
            f'    spoke-sdp 100:{spy} endpoint "y" icb create',
            f'        no shutdown',
            f'    exit',
            f'    spoke-sdp 100:{spx} endpoint "x" icb create',
            f'        no shutdown',
            f'    exit',
            f'    no shutdown',
            f'    info',
            f'    exit all',
            f'    admin save'
        ]
        return "\n".join(lines)

    def generate(self):
        try:
            cfg = self.build_config()
            self.preview.delete("1.0", tk.END)
            self.preview.insert(tk.END, cfg)
        except Exception as e:
            messagebox.showwarning("Missing fields", str(e), parent=self)

    def save_txt(self):
        save_text(self, self.preview.get("1.0", tk.END), "EpipeGen.txt")

    def save_excel(self):
        save_single_sheet_excel(self, "EpipeGen",
                                self.preview.get("1.0", tk.END),
                                "EpipeGen.xlsx")

    def append_log(self, text, tag=None):
        def inner():
            self.log.configure(state="normal")
            prefix = {
                "error": "[ERROR] ",
                "info": "[INFO] ",
                "cmd": "[CMD] ",
                "resp": "[RESP] "
            }.get(tag, "")
            self.log.insert(tk.END, prefix + str(text) + "\n")
            self.log.see(tk.END)
            self.log.configure(state="disabled")
        self.log.after(0, inner)

    def send(self, all_at_once=True):
        try:
            delay = float(self.e_delay.get())
        except Exception:
            delay = 0.5
        dry = self.var_dry.get()
        threading.Thread(target=self._send_worker,
                         args=(all_at_once, delay, dry),
                         daemon=True).start()

    def _send_worker(self, all_at_once, line_delay, dry_run):
        cfg_text = self.preview.get("1.0", tk.END).strip()
        if not cfg_text:
            self.append_log("No config to send. Generate config first.", tag="error")
            return
        host = self.d_host.get().strip()
        user = self.d_user.get().strip()
        pwd = self.d_pass.get().strip()
        try:
            port = int(self.d_port.get().strip() or 22)
        except Exception:
            port = 22

        if dry_run:
            self.append_log("DRY RUN: Not connecting. Below is the config that WOULD be sent:", tag="info")
            self.append_log(cfg_text, tag="cmd")
            return
        if ConnectHandler is None:
            self.append_log("Netmiko not installed. Install with: pip install netmiko", tag="error")
            return
        if not (host and user and pwd):
            self.append_log("Missing device credentials (host/username/password).", tag="error")
            return

        device = {
            "device_type": "nokia_sros",
            "host": host,
            "username": user,
            "password": pwd,
            "port": port
        }
        self.append_log(f"Connecting to {host}...", tag="info")
        try:
            conn = ConnectHandler(**device)
        except Exception as e:
            self.append_log(f"Connection failed: {e}", tag="error")
            return

        self.append_log("Connected. Sending configuration...", tag="info")
        commands = cfg_text.splitlines()
        try:
            if all_at_once:
                out = conn.send_config_set(commands)
                self.append_log(out, tag="resp")
            else:
                for i, cmd in enumerate(commands, 1):
                    if not cmd.strip():
                        continue
                    self.append_log(f"→ [{i}/{len(commands)}] {cmd}", tag="cmd")
                    out = conn.send_config_set([cmd], exit_config_mode=False)
                    self.append_log(out, tag="resp")
                    time.sleep(line_delay)
                try:
                    conn.exit_config_mode()
                except Exception:
                    pass
            self.append_log("Configuration push finished. Closing.", tag="info")
        except Exception as e:
            self.append_log(f"Error while sending: {e}", tag="error")
        finally:
            try:
                conn.disconnect()
            except Exception:
                pass

    def get_output_text(self):
        return self.preview.get("1.0", tk.END)

class SAPDeleteTab(SimpleInOutTab):
    SHEET_NAME = "SAPDelete"

    def __init__(self, master):
        super().__init__(master, "SAP Delete Config Generator")

    def generate(self):
        src = self.input.get("1.0", tk.END).strip()
        out = []
        current_vpls = None
        for line in src.splitlines():
            s = line.strip()
            if s.startswith("vpls "):
                parts = s.split()
                if len(parts) >= 2:
                    current_vpls = parts[1]
            elif s.startswith("sap ") and "create" in s and current_vpls:
                sap_full = s.split("create")[0].strip()
                sap_id = sap_full.split()[1]
                out.append(f"/configure service vpls {current_vpls}")
                out.append(f"    sap {sap_id} shutdown")
                out.append(f"    no sap {sap_id}")
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, "\n".join(out))
        
 #################VPLS config TAB ###################
 
 
class VPLSConfigTab1(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)

        import tkinter as tk
        from tkinter import ttk, messagebox, filedialog
        import pandas as pd

        # =========================================================
        # MAIN NOTEBOOK
        # =========================================================

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        # =========================================================
        # VPLS CONFIG GENERATOR TAB
        # =========================================================

        config_tab = tk.Frame(notebook)
        notebook.add(config_tab, text="VPLS Config Generator")

        # =========================================================
        # INPUT FRAME
        # =========================================================

        input_frame = tk.LabelFrame(
            config_tab,
            text="VPLS Configuration Input",
            padx=10,
            pady=10
        )
        input_frame.pack(fill="x", padx=10, pady=10)

        # =========================================================
        # VARIABLES
        # =========================================================

        vpls_id_var = tk.StringVar()
        sap_id_var = tk.StringVar()
        olt_vlan_var = tk.StringVar()
        mgmt_vlan_var = tk.StringVar()
        olt_ip_var = tk.StringVar()
        description_var = tk.StringVar()
        horizon_group_var = tk.StringVar()
        vpls_name_var = tk.StringVar()

        # =========================================================
        # INPUT ROW 1
        # =========================================================

        tk.Label(
            input_frame,
            text="VPLS ID:"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=vpls_id_var,
            width=25
        ).grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        tk.Label(
            input_frame,
            text="VPLS NAME:"
        ).grid(
            row=0,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=vpls_name_var,
            width=25
        ).grid(
            row=0,
            column=3,
            padx=5,
            pady=5
        )

        # =========================================================
        # INPUT ROW 2
        # =========================================================

        tk.Label(
            input_frame,
            text="OLT VLAN:"
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=olt_vlan_var,
            width=25
        ).grid(
            row=1,
            column=1,
            padx=5,
            pady=5
        )

        tk.Label(
            input_frame,
            text="MGMT / Voice VLAN:"
        ).grid(
            row=1,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=mgmt_vlan_var,
            width=25
        ).grid(
            row=1,
            column=3,
            padx=5,
            pady=5
        )

        # =========================================================
        # INPUT ROW 3
        # =========================================================

        tk.Label(
            input_frame,
            text="OLT IP (Optional):"
        ).grid(
            row=2,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=olt_ip_var,
            width=25
        ).grid(
            row=2,
            column=1,
            padx=5,
            pady=5
        )

        tk.Label(
            input_frame,
            text="Description:"
        ).grid(
            row=2,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=description_var,
            width=25
        ).grid(
            row=2,
            column=3,
            padx=5,
            pady=5
        )


        tk.Label(
            input_frame,
            text="SAP ID:"
        ).grid(
            row=3,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )
        tk.Entry(
            input_frame,
            textvariable=sap_id_var,
            width=25
        ).grid(
            row=3,
            column=3,
            padx=5,
            pady=5
        )
        # =========================================================
        # INPUT ROW 4 - HORIZON GROUP
        # =========================================================

        tk.Label(
            input_frame,
            text="Horizon Group:"
        ).grid(
            row=3,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=horizon_group_var,
            width=25
        ).grid(
            row=3,
            column=1,
            padx=5,
            pady=5
        )

        # =========================================================
        # CONFIG GENERATOR FUNCTION
        # =========================================================

        def generate_config():

            vpls_id = vpls_id_var.get().strip()
            sap_id = sap_id_var.get().strip()
            olt_vlan = olt_vlan_var.get().strip()
            mgmt_vlan = mgmt_vlan_var.get().strip()
            olt_ip = olt_ip_var.get().strip()
            description = description_var.get().strip()
            horizon_group = horizon_group_var.get().strip()
           

            # -----------------------------------------------------
            # Validation
            # -----------------------------------------------------

            if not vpls_id:
                messagebox.showerror(
                    "Error",
                    "Please enter VPLS ID."
                )
                return

            if not sap_id:
                messagebox.showerror(
                    "Error",
                    "Please enter SAP ID."
                )
                return

            if not olt_vlan:
                messagebox.showerror(
                    "Error",
                    "Please enter OLT VLAN."
                )
                return

            if not mgmt_vlan:
                messagebox.showerror(
                    "Error",
                    "Please enter MGMT / Voice VLAN."
                )
                return

            if not description:
                messagebox.showerror(
                    "Error",
                    "Please enter Description."
                )
                return

            if not horizon_group:
                messagebox.showerror(
                    "Error",
                    "Please enter Horizon Group."
                )
                return

            # -----------------------------------------------------
            # SAP
            # -----------------------------------------------------

            sap = f"{sap_id}:{olt_vlan}.{mgmt_vlan}"

            # -----------------------------------------------------
            # Start configuration
            # -----------------------------------------------------

            config = []

            config.append(
                f'/configure service vpls {vpls_id} '
                f'name "{horizon_group.lower()}" customer 1 create'
            )

            config.append(
                f'    sap {sap} '
                f'split-horizon-group "{horizon_group}" create'
            )

            # -----------------------------------------------------
            # Description
            # -----------------------------------------------------

            config.append(
                f'        description "{description}"'
            )

            # -----------------------------------------------------
            # OLT IP is optional
            # -----------------------------------------------------

            if olt_ip:

                config.append(
                    f'        static-host ip {olt_ip} create'
                )

                config.append(
                    '            no shutdown'
                )

                config.append(
                    '        exit'
                )

            # -----------------------------------------------------
            # Common configuration
            # -----------------------------------------------------

            config.append(
                '        dist-cpu-protection '
                '"dist-cpu-arp-access"'
            )

            config.append(
                '        no shutdown'
            )

            config.append(
                '    exit'
            )
            config.append('')

            config.append(
                '/admin save'
            )

            # -----------------------------------------------------
            # Display
            # -----------------------------------------------------

            output.delete(
                "1.0",
                tk.END
            )

            output.insert(
                tk.END,
                "\n".join(config)
            )

        # =========================================================
        # BUTTON FRAME
        # =========================================================

        button_frame = tk.Frame(config_tab)
        button_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        tk.Button(
            button_frame,
            text="Generate Config",
            command=generate_config,
            width=18
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # COPY FUNCTION
        # =========================================================

        def copy_config():

            config = output.get(
                "1.0",
                tk.END
            ).strip()

            if not config:

                messagebox.showwarning(
                    "Warning",
                    "No configuration available."
                )

                return

            self.clipboard_clear()
            self.clipboard_append(config)
            self.update()

            messagebox.showinfo(
                "Copied",
                "Configuration copied to clipboard."
            )

        tk.Button(
            button_frame,
            text="Copy Config",
            command=copy_config,
            width=18
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # SAVE CONFIG
        # =========================================================

        def save_config():

            config = output.get(
                "1.0",
                tk.END
            ).strip()

            if not config:

                messagebox.showwarning(
                    "Warning",
                    "No configuration available."
                )

                return

            filename = filedialog.asksaveasfilename(
                title="Save VPLS Configuration",
                defaultextension=".txt",
                filetypes=[
                    ("Text Files", "*.txt"),
                    ("All Files", "*.*")
                ]
            )

            if not filename:
                return

            try:

                with open(
                    filename,
                    "w",
                    encoding="utf-8"
                ) as file:

                    file.write(config)

                messagebox.showinfo(
                    "Success",
                    "Configuration saved successfully."
                )

            except Exception as e:

                messagebox.showerror(
                    "Error",
                    f"Unable to save file:\n{e}"
                )

        tk.Button(
            button_frame,
            text="Save Config",
            command=save_config,
            width=18
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # EXCEL IMPORT
        # =========================================================

        excel_frame = tk.LabelFrame(
            config_tab,
            text="Excel Import",
            padx=10,
            pady=10
        )

        excel_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        # =========================================================
        # EXCEL IMPORT FUNCTION
        # =========================================================

        def import_excel():

            filename = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[
                    ("Excel Files", "*.xlsx *.xls"),
                    ("All Files", "*.*")
                ]
            )

            if not filename:
                return

            try:

                # -------------------------------------------------
                # Read Excel
                # -------------------------------------------------

                df = pd.read_excel(filename)

                # -------------------------------------------------
                # Required columns
                # -------------------------------------------------

                required_columns = [
                    "vpls id",
                    "sap id",
                    "OLT_vlan",
                    "mgmt_vlan/voice",
                    "OLT ip",
                    "description",
                    "horizon group",
                    "vpls_name"
                ]

                # -------------------------------------------------
                # Create case-insensitive column map
                # -------------------------------------------------

                column_map = {
                    str(col).strip().lower(): col
                    for col in df.columns
                }

                # -------------------------------------------------
                # Check missing columns
                # -------------------------------------------------

                missing = []

                for column in required_columns:

                    if column.lower() not in column_map:

                        missing.append(column)

                if missing:

                    messagebox.showerror(
                        "Excel Error",
                        "Missing columns:\n\n" +
                        "\n".join(missing)
                    )

                    return

                # -------------------------------------------------
                # Check empty Excel
                # -------------------------------------------------

                if df.empty:

                    messagebox.showwarning(
                        "Excel Warning",
                        "The Excel file contains no data."
                    )

                    return

                # =================================================
                # GROUP BY VPLS ID
                # =================================================

                grouped = df.groupby(
                    column_map["vpls id"],
                    sort=False
                )

                all_configs = []

                # =================================================
                # PROCESS EACH VPLS
                # =================================================

                for vpls_id, group in grouped:

                    # -------------------------------------------------
                    # Clean VPLS ID
                    # -------------------------------------------------

                    if pd.isna(vpls_id):

                        continue

                    vpls_id = str(
                        vpls_id
                    ).strip()

                    if not vpls_id:

                        continue

                    # -------------------------------------------------
                    # VPLS NAME
                    # -------------------------------------------------

                    vpls_name = "tparty-olt-mgmt"

                    # -------------------------------------------------
                    # Start VPLS
                    # -------------------------------------------------

                    config = []

                    config.append(
                        f'/configure service vpls {vpls_id} '
                        f'name "{vpls_name}" customer 1 create'
                    )

                    # =================================================
                    # PROCESS SAPS
                    # =================================================

                    for index, row in group.iterrows():

                        # -------------------------------------------------
                        # Read values
                        # -------------------------------------------------

                        sap_id = row[
                            column_map["sap id"]
                        ]

                        olt_vlan = row[
                            column_map["olt_vlan"]
                        ]

                        mgmt_vlan = row[
                            column_map["mgmt_vlan/voice"]
                        ]

                        olt_ip = row[
                            column_map["olt ip"]
                        ]

                        description = row[
                            column_map["description"]
                        ]

                        horizon_group = row[
                            column_map["horizon group"]
                        ]

                        # -------------------------------------------------
                        # Convert values to string safely
                        # -------------------------------------------------

                        def clean_value(value):

                            if pd.isna(value):
                                return ""

                            return str(value).strip()

                        sap_id = clean_value(sap_id)
                        olt_vlan = clean_value(olt_vlan)
                        mgmt_vlan = clean_value(mgmt_vlan)
                        olt_ip = clean_value(olt_ip)
                        description = clean_value(description)
                        horizon_group = clean_value(horizon_group)

                        # -------------------------------------------------
                        # Validate required SAP fields
                        # -------------------------------------------------

                        if not sap_id:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"SAP ID is empty. Row skipped."
                            )

                            continue

                        if not olt_vlan:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"OLT VLAN is empty. Row skipped."
                            )

                            continue

                        if not mgmt_vlan:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"MGMT / Voice VLAN is empty. Row skipped."
                            )

                            continue

                        if not horizon_group:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"Horizon Group is empty. Row skipped."
                            )

                            continue

                        # -------------------------------------------------
                        # SAP
                        # -------------------------------------------------

                        sap = (
                            f"{sap_id}:"
                            f"{olt_vlan}."
                            f"{mgmt_vlan}"
                        )

                        # -------------------------------------------------
                        # SAP configuration
                        # -------------------------------------------------

                        config.append(
                            f'    sap {sap} '
                            f'split-horizon-group '
                            f'"{horizon_group}" create'
                        )

                        # -------------------------------------------------
                        # Description
                        # -------------------------------------------------

                        if description:

                            config.append(
                                f'        description '
                                f'"{description}"'
                            )

                        # -------------------------------------------------
                        # Optional OLT IP
                        # -------------------------------------------------

                        if olt_ip:

                            config.append(
                                f'        static-host ip '
                                f'{olt_ip} create'
                            )

                            config.append(
                                '            no shutdown'
                            )

                            config.append(
                                '        exit'
                            )

                        # -------------------------------------------------
                        # Distribution CPU protection
                        # -------------------------------------------------

                        config.append(
                            '        dist-cpu-protection '
                            '"dist-cpu-arp-access"'
                        )

                        config.append(
                            '        no shutdown'
                        )

                        # -------------------------------------------------
                        # Close SAP
                        # -------------------------------------------------

                        config.append(
                            '    exit'
                        )

                    # =================================================
                    # CLOSE VPLS
                    # =================================================

                    config.append(
                        'exit all'
                    )

                    config.append('')

                    all_configs.append(
                        "\n".join(config)
                    )

                # =================================================
                # CHECK RESULT
                # =================================================

                if not all_configs:

                    messagebox.showwarning(
                        "Excel Warning",
                        "No valid configuration could be generated."
                    )

                    return

                # =================================================
                # ADD ADMIN SAVE ONLY ONCE
                # =================================================

                all_configs.append(
                    "/admin save"
                )

                # =================================================
                # DISPLAY
                # =================================================

                output.delete(
                    "1.0",
                    tk.END
                )

                output.insert(
                    tk.END,
                    "\n\n".join(all_configs)
                )

                # =================================================
                # SUCCESS MESSAGE
                # =================================================

                valid_vpls = len(grouped)

                messagebox.showinfo(
                    "Success",
                    f"Excel imported successfully.\n\n"
                    f"VPLS groups: {valid_vpls}\n"
                    f"Rows processed: {len(df)}"
                )

            except Exception as e:

                messagebox.showerror(
                    "Excel Error",
                    f"Unable to import Excel file:\n\n{e}"
                )

        # =========================================================
        # IMPORT EXCEL BUTTON
        # =========================================================

        tk.Button(
            excel_frame,
            text="Import Excel & Generate",
            command=import_excel,
            width=25
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # OUTPUT
        # =========================================================

        output_frame = tk.LabelFrame(
            config_tab,
            text="Generated Nokia SR OS Configuration",
            padx=5,
            pady=5
        )

        output_frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        output = tk.Text(
            output_frame,
            height=20,
            wrap="none"
        )

        output.pack(
            side="left",
            fill="both",
            expand=True
        )

        # =========================================================
        # VERTICAL SCROLLBAR
        # =========================================================

        scrollbar_y = ttk.Scrollbar(
            output_frame,
            orient="vertical",
            command=output.yview
        )

        scrollbar_y.pack(
            side="right",
            fill="y"
        )

        output.configure(
            yscrollcommand=scrollbar_y.set
        )

        # =========================================================
        # HORIZONTAL SCROLLBAR
        # =========================================================

        scrollbar_x = ttk.Scrollbar(
            config_tab,
            orient="horizontal",
            command=output.xview
        )

        scrollbar_x.pack(
            fill="x",
            padx=10
        )

        output.configure(
            xscrollcommand=scrollbar_x.set
        )      
       


    def __init__(self, master):
        super().__init__(master)

        import tkinter as tk
        from tkinter import ttk, messagebox, filedialog
        import pandas as pd

        # =========================================================
        # MAIN NOTEBOOK
        # =========================================================

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        # =========================================================
        # VPLS CONFIG GENERATOR TAB
        # =========================================================

        config_tab = tk.Frame(notebook)
        notebook.add(config_tab, text="VPLS Config Generator")

        # =========================================================
        # INPUT FRAME
        # =========================================================

        input_frame = tk.LabelFrame(
            config_tab,
            text="VPLS Configuration Input",
            padx=10,
            pady=10
        )
        input_frame.pack(fill="x", padx=10, pady=10)

        # =========================================================
        # VARIABLES
        # =========================================================

        vpls_id_var = tk.StringVar()
        sap_id_var = tk.StringVar()
        olt_vlan_var = tk.StringVar()
        mgmt_vlan_var = tk.StringVar()
        olt_ip_var = tk.StringVar()
        description_var = tk.StringVar()
        horizon_group_var = tk.StringVar()
        vpls_name_var = tk.StringVar()

        # =========================================================
        # INPUT ROW 1
        # =========================================================

        tk.Label(
            input_frame,
            text="VPLS ID:"
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=vpls_id_var,
            width=25
        ).grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        tk.Label(
            input_frame,
            text="VPLS NAME:"
        ).grid(
            row=0,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=vpls_name_var,
            width=25
        ).grid(
            row=0,
            column=3,
            padx=5,
            pady=5
        )

        # =========================================================
        # INPUT ROW 2
        # =========================================================

        tk.Label(
            input_frame,
            text="OLT VLAN:"
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=olt_vlan_var,
            width=25
        ).grid(
            row=1,
            column=1,
            padx=5,
            pady=5
        )

        tk.Label(
            input_frame,
            text="MGMT / Voice VLAN:"
        ).grid(
            row=1,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=mgmt_vlan_var,
            width=25
        ).grid(
            row=1,
            column=3,
            padx=5,
            pady=5
        )

        # =========================================================
        # INPUT ROW 3
        # =========================================================

        tk.Label(
            input_frame,
            text="OLT IP (Optional):"
        ).grid(
            row=2,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=olt_ip_var,
            width=25
        ).grid(
            row=2,
            column=1,
            padx=5,
            pady=5
        )

        tk.Label(
            input_frame,
            text="Description:"
        ).grid(
            row=2,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=description_var,
            width=25
        ).grid(
            row=2,
            column=3,
            padx=5,
            pady=5
        )


        tk.Label(
            input_frame,
            text="SAP ID:"
        ).grid(
            row=3,
            column=2,
            sticky="w",
            padx=5,
            pady=5
        )
        tk.Entry(
            input_frame,
            textvariable=sap_id_var,
            width=25
        ).grid(
            row=3,
            column=3,
            padx=5,
            pady=5
        )
        # =========================================================
        # INPUT ROW 4 - HORIZON GROUP
        # =========================================================

        tk.Label(
            input_frame,
            text="Horizon Group:"
        ).grid(
            row=3,
            column=0,
            sticky="w",
            padx=5,
            pady=5
        )

        tk.Entry(
            input_frame,
            textvariable=horizon_group_var,
            width=25
        ).grid(
            row=3,
            column=1,
            padx=5,
            pady=5
        )

        # =========================================================
        # CONFIG GENERATOR FUNCTION
        # =========================================================

        def generate_config():

            vpls_id = vpls_id_var.get().strip()
            sap_id = sap_id_var.get().strip()
            olt_vlan = olt_vlan_var.get().strip()
            mgmt_vlan = mgmt_vlan_var.get().strip()
            olt_ip = olt_ip_var.get().strip()
            description = description_var.get().strip()
            horizon_group = horizon_group_var.get().strip()
            vpls_name = vpls_name_var.get().strip()

            # -----------------------------------------------------
            # Validation
            # -----------------------------------------------------

            if not vpls_id:
                messagebox.showerror(
                    "Error",
                    "Please enter VPLS ID."
                )
                return

            if not sap_id:
                messagebox.showerror(
                    "Error",
                    "Please enter SAP ID."
                )
                return

            if not olt_vlan:
                messagebox.showerror(
                    "Error",
                    "Please enter OLT VLAN."
                )
                return

            if not mgmt_vlan:
                messagebox.showerror(
                    "Error",
                    "Please enter MGMT / Voice VLAN."
                )
                return

            if not description:
                messagebox.showerror(
                    "Error",
                    "Please enter Description."
                )
                return

            if not horizon_group:
                messagebox.showerror(
                    "Error",
                    "Please enter Horizon Group."
                )
                return
            
            if not vpls_name:
                
                messagebox.showerror(
                    "Error",
                     "Please enter vpls_name."
                            )
                return

            # -----------------------------------------------------
            # SAP
            # -----------------------------------------------------

            sap = f"{sap_id}:{olt_vlan}.{mgmt_vlan}"

            # -----------------------------------------------------
            # Start configuration
            # -----------------------------------------------------

            config = []

            config.append(
                f'/configure service vpls {vpls_id} '
                f'name "{vpls_name}" customer 1 create'
            )

            config.append(
                f'    sap {sap} '
                f'split-horizon-group "{horizon_group}" create'
            )

            # -----------------------------------------------------
            # Description
            # -----------------------------------------------------

            config.append(
                f'        description "{description}"'
            )

            # -----------------------------------------------------
            # OLT IP is optional
            # -----------------------------------------------------

            if olt_ip:

                config.append(
                    f'        static-host ip {olt_ip} create'
                )

                config.append(
                    '            no shutdown'
                )

                config.append(
                    '        exit'
                )

            # -----------------------------------------------------
            # Common configuration
            # -----------------------------------------------------

            config.append(
                '        dist-cpu-protection '
                '"dist-cpu-arp-access"'
            )

            config.append(
                '        no shutdown'
            )

            config.append(
                '    exit'
            )
            config.append('')

            config.append(
                '/admin save'
            )

            # -----------------------------------------------------
            # Display
            # -----------------------------------------------------

            output.delete(
                "1.0",
                tk.END
            )

            output.insert(
                tk.END,
                "\n".join(config)
            )

        # =========================================================
        # BUTTON FRAME
        # =========================================================

        button_frame = tk.Frame(config_tab)
        button_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        tk.Button(
            button_frame,
            text="Generate Config",
            command=generate_config,
            width=18
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # COPY FUNCTION
        # =========================================================

        def copy_config():

            config = output.get(
                "1.0",
                tk.END
            ).strip()

            if not config:

                messagebox.showwarning(
                    "Warning",
                    "No configuration available."
                )

                return

            self.clipboard_clear()
            self.clipboard_append(config)
            self.update()

            messagebox.showinfo(
                "Copied",
                "Configuration copied to clipboard."
            )

        tk.Button(
            button_frame,
            text="Copy Config",
            command=copy_config,
            width=18
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # SAVE CONFIG
        # =========================================================

        def save_config():

            config = output.get(
                "1.0",
                tk.END
            ).strip()

            if not config:

                messagebox.showwarning(
                    "Warning",
                    "No configuration available."
                )

                return

            filename = filedialog.asksaveasfilename(
                title="Save VPLS Configuration",
                defaultextension=".txt",
                filetypes=[
                    ("Text Files", "*.txt"),
                    ("All Files", "*.*")
                ]
            )

            if not filename:
                return

            try:

                with open(
                    filename,
                    "w",
                    encoding="utf-8"
                ) as file:

                    file.write(config)

                messagebox.showinfo(
                    "Success",
                    "Configuration saved successfully."
                )

            except Exception as e:

                messagebox.showerror(
                    "Error",
                    f"Unable to save file:\n{e}"
                )

        tk.Button(
            button_frame,
            text="Save Config",
            command=save_config,
            width=18
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # EXCEL IMPORT
        # =========================================================

        excel_frame = tk.LabelFrame(
            config_tab,
            text="Excel Import",
            padx=10,
            pady=10
        )

        excel_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        # =========================================================
        # EXCEL IMPORT FUNCTION
        # =========================================================

        def import_excel():

            filename = filedialog.askopenfilename(
                title="Select Excel File",
                filetypes=[
                    ("Excel Files", "*.xlsx *.xls"),
                    ("All Files", "*.*")
                ]
            )

            if not filename:
                return

            try:

                # -------------------------------------------------
                # Read Excel
                # -------------------------------------------------

                df = pd.read_excel(filename)

                # -------------------------------------------------
                # Required columns
                # -------------------------------------------------

                required_columns = [
                    "vpls id",
                    "sap id",
                    "OLT_vlan",
                    "mgmt_vlan/voice",
                    "OLT ip",
                    "description",
                    "horizon group",
                    "vpls_name"
                ]

                # -------------------------------------------------
                # Create case-insensitive column map
                # -------------------------------------------------

                column_map = {
                    str(col).strip().lower(): col
                    for col in df.columns
                }

                # -------------------------------------------------
                # Check missing columns
                # -------------------------------------------------

                missing = []

                for column in required_columns:

                    if column.lower() not in column_map:

                        missing.append(column)

                if missing:

                    messagebox.showerror(
                        "Excel Error",
                        "Missing columns:\n\n" +
                        "\n".join(missing)
                    )

                    return

                # -------------------------------------------------
                # Check empty Excel
                # -------------------------------------------------

                if df.empty:

                    messagebox.showwarning(
                        "Excel Warning",
                        "The Excel file contains no data."
                    )

                    return

                # =================================================
                # GROUP BY VPLS ID
                # =================================================

                grouped = df.groupby(
                    column_map["vpls id"],
                    sort=False
                )
                
                vpls_name = str(
                group.iloc[0][
                    column_map["vpls_name"]
                ]
            ).strip()
                
                vpls_name = "1"
                
                all_configs = []

                # =================================================
                # PROCESS EACH VPLS
                # =================================================

                for vpls_id, group in grouped:
                    

                    # -------------------------------------------------
                    # Clean VPLS ID
                    # -------------------------------------------------

                    if pd.isna(vpls_id):

                        continue

                    vpls_id = str(
                        vpls_id
                    ).strip()

                    if not vpls_id:

                        continue
                    
                   


                    # -------------------------------------------------
                    # Start VPLS
                    # -------------------------------------------------

                    config = []

                    config.append(
                        f'/configure service vpls {vpls_id} '
                        f'name "{vpls_name}" customer 1 create'
                    )

                    # =================================================
                    # PROCESS SAPS
                    # =================================================

                    for index, row in group.iterrows():

                        # -------------------------------------------------
                        # Read values
                        # -------------------------------------------------

                        sap_id = row[
                            column_map["sap id"]
                        ]

                        olt_vlan = row[
                            column_map["olt_vlan"]
                        ]

                        mgmt_vlan = row[
                            column_map["mgmt_vlan/voice"]
                        ]

                        olt_ip = row[
                            column_map["olt ip"]
                        ]

                        description = row[
                            column_map["description"]
                        ]

                        horizon_group = row[
                            column_map["horizon group"]
                        ]

                        # -------------------------------------------------
                        # Convert values to string safely
                        # -------------------------------------------------

                        def clean_value(value):

                            if pd.isna(value):
                                return ""

                            return str(value).strip()

                        sap_id = clean_value(sap_id)
                        olt_vlan = clean_value(olt_vlan)
                        mgmt_vlan = clean_value(mgmt_vlan)
                        olt_ip = clean_value(olt_ip)
                        description = clean_value(description)
                        horizon_group = clean_value(horizon_group)

                        # -------------------------------------------------
                        # Validate required SAP fields
                        # -------------------------------------------------

                        if not sap_id:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"SAP ID is empty. Row skipped."
                            )

                            continue

                        if not olt_vlan:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"OLT VLAN is empty. Row skipped."
                            )

                            continue

                        if not mgmt_vlan:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"MGMT / Voice VLAN is empty. Row skipped."
                            )

                            continue

                        if not horizon_group:

                            messagebox.showwarning(
                                "Excel Warning",
                                f"Row {index + 2}: "
                                f"Horizon Group is empty. Row skipped."
                            )

                            continue

                        # -------------------------------------------------
                        # SAP
                        # -------------------------------------------------

                        sap = (
                            f"{sap_id}:"
                            f"{olt_vlan}."
                            f"{mgmt_vlan}"
                        )

                        # -------------------------------------------------
                        # SAP configuration
                        # -------------------------------------------------

                        config.append(
                            f'    sap {sap} '
                            f'split-horizon-group '
                            f'"{horizon_group}" create'
                        )

                        # -------------------------------------------------
                        # Description
                        # -------------------------------------------------

                        if description:

                            config.append(
                                f'        description '
                                f'"{description}"'
                            )

                        # -------------------------------------------------
                        # Optional OLT IP
                        # -------------------------------------------------

                        if olt_ip:

                            config.append(
                                f'        static-host ip '
                                f'{olt_ip} create'
                            )

                            config.append(
                                '            no shutdown'
                            )

                            config.append(
                                '        exit'
                            )

                        # -------------------------------------------------
                        # Distribution CPU protection
                        # -------------------------------------------------

                        config.append(
                            '        dist-cpu-protection '
                            '"dist-cpu-arp-access"'
                        )

                        config.append(
                            '        no shutdown'
                        )

                        # -------------------------------------------------
                        # Close SAP
                        # -------------------------------------------------

                        config.append(
                            '    exit'
                        )

                    # =================================================
                    # CLOSE VPLS
                    # =================================================

                    config.append(
                        'exit all'
                    )

                    config.append('')

                    all_configs.append(
                        "\n".join(config)
                    )

                # =================================================
                # CHECK RESULT
                # =================================================

                if not all_configs:

                    messagebox.showwarning(
                        "Excel Warning",
                        "No valid configuration could be generated."
                    )

                    return

                # =================================================
                # ADD ADMIN SAVE ONLY ONCE
                # =================================================

                all_configs.append(
                    "/admin save"
                )

                # =================================================
                # DISPLAY
                # =================================================

                output.delete(
                    "1.0",
                    tk.END
                )

                output.insert(
                    tk.END,
                    "\n\n".join(all_configs)
                )

                # =================================================
                # SUCCESS MESSAGE
                # =================================================

                valid_vpls = len(grouped)

                messagebox.showinfo(
                    "Success",
                    f"Excel imported successfully.\n\n"
                    f"VPLS groups: {valid_vpls}\n"
                    f"Rows processed: {len(df)}"
                )

            except Exception as e:

                messagebox.showerror(
                    "Excel Error",
                    f"Unable to import Excel file:\n\n{e}"
                )

        # =========================================================
        # IMPORT EXCEL BUTTON
        # =========================================================

        tk.Button(
            excel_frame,
            text="Import Excel & Generate",
            command=import_excel,
            width=25
        ).pack(
            side="left",
            padx=5
        )

        # =========================================================
        # OUTPUT
        # =========================================================

        output_frame = tk.LabelFrame(
            config_tab,
            text="Generated Nokia SR OS Configuration",
            padx=5,
            pady=5
        )

        output_frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

        output = tk.Text(
            output_frame,
            height=20,
            wrap="none"
        )

        output.pack(
            side="left",
            fill="both",
            expand=True
        )

        # =========================================================
        # VERTICAL SCROLLBAR
        # =========================================================

        scrollbar_y = ttk.Scrollbar(
            output_frame,
            orient="vertical",
            command=output.yview
        )

        scrollbar_y.pack(
            side="right",
            fill="y"
        )

        output.configure(
            yscrollcommand=scrollbar_y.set
        )

        # =========================================================
        # HORIZONTAL SCROLLBAR
        # =========================================================

        scrollbar_x = ttk.Scrollbar(
            config_tab,
            orient="horizontal",
            command=output.xview
        )

        scrollbar_x.pack(
            fill="x",
            padx=10
        )

        output.configure(
            xscrollcommand=scrollbar_x.set
        )      


########################## BBNL Config genrater ####################

class BbnlinbandTab(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.build_ui()

    def build_ui(self):

        # ================= HEADER =================
        header = ttk.Frame(self)
        header.pack(fill="x", pady=8)

        title = tk.Label(
            header,
            text="⚙ BBNLINBAND_GEN",
            font=("Segoe UI", 16, "bold"),
            fg="#2C3E50"
        )
        title.pack(side=tk.LEFT, padx=10)

        # ================= BUTTON BAR =================
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", pady=5)

        self.import_btn = tk.Button(
            btn_frame,
            text="📂 Import Excel",
            command=self.load_excel,
            bg="#2E86C1",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=5
        )
        self.import_btn.pack(side=tk.LEFT, padx=5)

        self.save_btn = tk.Button(
            btn_frame,
            text="💾 Save Script",
            command=self.save_output,
            bg="#27AE60",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=5
        )
        self.save_btn.pack(side=tk.LEFT, padx=5)

        # ================= OUTPUT FRAME =================
        output_frame = ttk.LabelFrame(self, text="Generated Script Output")
        output_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.output_text = scrolledtext.ScrolledText(
            output_frame,
            width=100,
            height=30,
            font=("Consolas", 10),
            bg="#1E1E1E",
            fg="#00FF88",
            insertbackground="white"
        )
        self.output_text.pack(fill="both", expand=True, padx=5, pady=5)

    # ---------------- LOAD EXCEL ----------------
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            self.generate_scripts(df)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------------- GENERATE SCRIPT ----------------
    def generate_scripts(self, df):

        self.output_text.delete(1.0, tk.END)

        scripts = []

        for _, row in df.iterrows():

            script = f'''
interface "{row['interface']}" create
    description "{row['description']}"
    address {row['address']}/29
    sap {row['sap_port']}:{row['olt_vlan']}.{row['ill_vlan']} create
      static-host ip {row['olt_ip']} create
        no shutdown
      exit
      dist-cpu-protection "dist-cpu-arp-access"
    exit
exit
'''
            scripts.append(script)

        self.output_text.insert(tk.END, "\n".join(scripts))

    # ---------------- SAVE OUTPUT ----------------
    def save_output(self):

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt")]
        )

        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.output_text.get(1.0, tk.END))
                
################### Epipe replace tab ######################


class EpipeReplaceGenerateTab(ttk.Frame):

    SHEET_NAME = "EpipeReplaceGenerate"

    def __init__(self, master):
        super().__init__(master)

        # ===== TOP FRAME =====
        frame_top = ttk.Frame(self)
        frame_top.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_top, text="LAG ID:").grid(row=0, column=0, padx=5)
        self.lag_entry = ttk.Entry(frame_top, width=10)
        self.lag_entry.grid(row=0, column=1)

        ttk.Label(frame_top, text="CP Suffix:").grid(row=0, column=2, padx=5)
        self.cp_entry = ttk.Entry(frame_top, width=10)
        self.cp_entry.grid(row=0, column=3)

        ttk.Button(frame_top, text="Generate Config",
                   command=self.generate_output).grid(row=0, column=4, padx=10)

        ttk.Button(frame_top, text="Replace Only",
                   command=self.replace_only).grid(row=0, column=5, padx=10)

        ttk.Button(frame_top, text="📋 Copy",
                   command=self.copy_output).grid(row=0, column=6, padx=10)

        # ===== INPUT =====
        ttk.Label(self, text="Paste Input Config:").pack(anchor="w", padx=10)
        self.input_text = tk.Text(self, height=15)
        self.input_text.pack(fill="both", padx=10, pady=5)

        # ===== OUTPUT =====
        ttk.Label(self, text="Output:").pack(anchor="w", padx=10)
        self.output_text = tk.Text(self, height=15)
        self.output_text.pack(fill="both", padx=10, pady=5)

    # ================= ORIGINAL LOGIC =================
    def parse_and_generate(self, cfg, lag_input, cp_input):
        epipes = []
        current_epipe = None

        for line in cfg.splitlines():
            line = line.strip()

            match = re.search(r'epipe\s+(\d+)\s+name\s+"(.+?)"\s+customer\s+(\d+)', line)
            if match:
                if current_epipe:
                    epipes.append(current_epipe)

                current_epipe = {
                    "epipe_id": match.group(1),
                    "name": match.group(2),
                    "customer": match.group(3),
                    "old_sap": None,
                    "qos": "10"
                }
                continue

            if not current_epipe:
                continue

            match = re.search(r'sap\s+(\S+)\s+endpoint\s+"x"\s+create', line)
            if match:
                current_epipe["old_sap"] = match.group(1)

            match = re.search(r'qos\s+(\d+)', line)
            if match:
                current_epipe["qos"] = match.group(1)

        if current_epipe:
            epipes.append(current_epipe)

        output = []

        for ep in epipes:
            if not ep["old_sap"]:
                continue

            epipe_num = ep["epipe_id"][-4:]
            new_sap = f"lag-{lag_input}:{epipe_num}.cp-11{cp_input}"
            old_sap = ep["old_sap"]

            block = f'''
/configure service epipe {ep["epipe_id"]} name "{ep["name"]}" customer {ep["customer"]} create
    sap {old_sap} shutdown
    no sap {old_sap}
    sap {new_sap} endpoint "x" create
        ingress
            qos {ep["qos"]}
        exit
        no shutdown
    exit
    no shutdown
'''.strip()

            output.append(block)

        return "\n\n".join(output)

    # ================= BUTTON ACTIONS =================
    def generate_output(self):
        cfg = self.input_text.get("1.0", tk.END).strip()
        lag_input = self.lag_entry.get().strip()
        cp_input = self.cp_entry.get().strip()

        if not cfg or not lag_input or not cp_input:
            messagebox.showerror("Error", "All fields are required!", parent=self)
            return

        result = self.parse_and_generate(cfg, lag_input, cp_input)

        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, result)

    def copy_output(self):
        self.clipboard_clear()
        self.clipboard_append(self.output_text.get("1.0", tk.END))
        messagebox.showinfo("Copied", "Output copied!", parent=self)

    def replace_only(self):
        cfg = self.input_text.get("1.0", tk.END)

        lag_input = self.lag_entry.get().strip()
        cp_input = self.cp_entry.get().strip()

        if not lag_input or not cp_input:
            messagebox.showerror("Error", "Enter LAG ID & CP Suffix", parent=self)
            return

        def replacer(match):
            old = match.group(0)
            num_match = re.search(r':(\d+)\.cp-', old)
            if num_match:
                num = num_match.group(1)
                return f"lag-{lag_input}:{num}.cp-11{cp_input}"
            return old

        new_cfg = re.sub(r'lag-\d+:\d+\.cp-\d+', replacer, cfg)

        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, new_cfg)

class EpipeRemoverTab(ttk.Frame):
    SHEET_NAME = "EpipeRemover"

    def __init__(self, master):
        super().__init__(master)
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        ttk.Button(top, text="Load Config File",
                   command=self.load_file).pack(side="left", padx=4)
        ttk.Button(top, text="Generate Remove Commands",
                   command=self.generate).pack(side="left", padx=4)
        ttk.Button(top, text="Save TXT",
                   command=self.save_txt).pack(side="left", padx=4)
        ttk.Button(top, text="Save Excel",
                   command=self.save_excel).pack(side="left", padx=4)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        pan = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        pan.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)

        left = ttk.Labelframe(pan, text="Input Config", padding=6)
        right = ttk.Labelframe(pan, text="Removal Commands", padding=6)
        pan.add(left, weight=1)
        pan.add(right, weight=1)

        self.input = scrolledtext.ScrolledText(left, wrap="word")
        self.input.pack(fill="both", expand=True)
        self.output = scrolledtext.ScrolledText(right, wrap="word")
        self.output.pack(fill="both", expand=True)

    def load_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if path:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                self.input.delete("1.0", tk.END)
                self.input.insert(tk.END, f.read())

    def generate(self):
        text = self.input.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("Warning", "Please paste the configuration first.", parent=self)
            return
        m = re.search(r"epipe\s+(\d+)", text)
        if not m:
            messagebox.showerror("Error", "Could not find service ID in config.", parent=self)
            return
        service_id = m.group(1)
        sap_pat = re.compile(r"sap\s+([^\s]+)")
        sdp_pat = re.compile(r"spoke-sdp\s+([^\s]+)")
        saps, sdps = [], []
        for line in text.splitlines():
            s = line.strip()
            sm = sdp_pat.search(s)
            pm = sap_pat.search(s)
            if sm:
                sdps.append(sm.group(1))
            if pm:
                saps.append(pm.group(1))
        out = []
        out.append(f"/configure service epipe {service_id}")
        out.append("    shutdown")
        for sdp in sdps:
            out.append(f"    spoke-sdp {sdp} shutdown")
            out.append(f"    no spoke-sdp {sdp}")
        for sap in saps:
            out.append(f"    sap {sap} shutdown")
            out.append(f"    no sap {sap}")
        out.append("    back")
        out.append(f"    no epipe {service_id}")
        out.append("exit all")
        out.append(f"/admin display-config | match context all {service_id}")
        out.append("admin save ")
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, "\n".join(out))

    def save_txt(self):
        save_text(self, self.output.get("1.0", tk.END), "EpipeRemover.txt")

    def save_excel(self):
        save_single_sheet_excel(self, "EpipeRemover",
                                self.output.get("1.0", tk.END),
                                "EpipeRemover.xlsx")

    def get_output_text(self):
        return self.output.get("1.0", tk.END)
    



class CPReportGen(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        self.df = pd.DataFrame()

        # -------- INPUT BOX --------
        self.input_text = tk.Text(self, height=12)
        self.input_text.pack(fill="x", padx=10, pady=5)

        # -------- BUTTONS --------
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)

        ttk.Button(btn_frame, text="Parse",
                   command=self.parse_data).grid(row=0, column=0, padx=5)

        ttk.Button(btn_frame, text="Copy Output",
                   command=self.copy_output).grid(row=0, column=1, padx=5)

        ttk.Button(btn_frame, text="Export TXT",
                   command=self.export_txt).grid(row=0, column=2, padx=5)

        ttk.Button(btn_frame, text="Export CSV",
                   command=self.export_csv).grid(row=0, column=3, padx=5)

        ttk.Button(btn_frame, text="Export Excel",
                   command=self.export_excel).grid(row=0, column=4, padx=5)

        # -------- TABLE --------
        self.columns = [
            "Node", "Date", "Card1", "Card2", "Card3", "Card4",
            "CPM A", "CPM B", "Subscribers Count",
            "Redundancy State", "CP Master",
            "CPU Usage", "Last Peer/BFD"
        ]

        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(
            table_frame,
            columns=self.columns,
            show="headings"
        )

        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="center")

        yscroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        xscroll = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.df = pd.DataFrame(columns=self.columns)

    # ---------------- PARSE ----------------
    def parse_data(self):
        raw_data = self.input_text.get("1.0", tk.END)

        if not raw_data.strip():
            messagebox.showwarning("Warning", "Input is empty")
            return

        blocks = re.split(r'\n#+\n', raw_data)
        rows = []

        for block in blocks:
            node = re.search(r'A:(\S+)#', block)
            node = node.group(1) if node else ""

            date = re.search(r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun).*?\d{4}', block)
            date = date.group(0) if date else ""

            cpu = ""
            
            all_cpu = re.findall(r'Usage\s+[\d,]+\s+([\d.]+%)', block)
            if all_cpu:
               cpu = all_cpu[0]


            sessions = re.search(r'Total number of sessions\s*:\s*(\d+)', block)
            sessions = sessions.group(1) if sessions else ""

            role = re.search(r'Oper Role\s*:\s*(\w+)', block)
            role = role.group(1) if role else ""

            geo = re.search(r'Geo-Redundancy State\s*:\s*(\w+)', block)
            geo = geo.group(1) if geo else ""

            peer = re.search(r'Last Time Peer Connected\s*:\s*(.*)', block)
            peer = peer.group(1).strip() if peer else ""

            boots = re.findall(r'Time of last boot\s*:\s*(.*)', block)

            row = {
                "Node": node,
                "Date": date,
                "Card1": boots[0] if len(boots) > 0 else "",
                "Card2": boots[1] if len(boots) > 1 else "",
                "Card3": boots[2] if len(boots) > 2 else "",
                "Card4": boots[3] if len(boots) > 3 else "",
                "CPM A": boots[4] if len(boots) > 4 else "",
                "CPM B": boots[5] if len(boots) > 5 else "",
                "Subscribers Count": sessions,
                "Redundancy State": geo,
                "CP Master": role,
                "CPU Usage": cpu,
                "Last Peer/BFD": peer
            }

            if any(row.values()):
                rows.append(row)

        if not rows:
            messagebox.showerror("Error", "No valid data found")
            return

        self.df = pd.DataFrame(rows, columns=self.columns)
        self.update_table()

        messagebox.showinfo("Success", "Parsing Completed")

    # ---------------- TABLE ----------------
    def update_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row))

    # ---------------- COPY ----------------
    def copy_output(self):
        if self.df.empty:
            messagebox.showwarning("Warning", "No data to copy")
            return

        text_data = self.df.to_string(index=False)

        self.clipboard_clear()
        self.clipboard_append(text_data)
        self.update()

        messagebox.showinfo("Copied", "Output copied to clipboard")

    # ---------------- EXPORTS ----------------
    def export_txt(self):
        if self.df.empty:
            messagebox.showwarning("Warning", "No data to export")
            return

        file = filedialog.asksaveasfilename(defaultextension=".txt")
        if file:
            with open(file, "w") as f:
                f.write(self.df.to_string(index=False))
            messagebox.showinfo("Saved", "TXT file saved")

    def export_csv(self):
        if self.df.empty:
            messagebox.showwarning("Warning", "No data to export")
            return

        file = filedialog.asksaveasfilename(defaultextension=".csv")
        if file:
            self.df.to_csv(file, index=False)
            messagebox.showinfo("Saved", "CSV file saved")

    def export_excel(self):
        if self.df.empty:
            messagebox.showwarning("Warning", "No data to export")
            return

        file = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if file:
            self.df.to_excel(file, index=False)
            messagebox.showinfo("Saved", "Excel file saved")



class CPMonitorTab(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)

        self.lock = threading.Lock()
        self.previous_data = {}

        # Control flags
        self.stop_flag = False
        self.scheduler_running = True
        self.is_running = False

        # Hide/Unhide
        self.hide_var = tk.BooleanVar()
        self.unhide_password = "love"
        self.tab_hidden = False

        # Save path
        self.save_path = r"E:\ \Don't"
        os.makedirs(self.save_path, exist_ok=True)

        # Routers
        self.routers = [
            {"node": "kol-cpf-dc-01", "ip": "172.30.189.4", "peer": "172.30.145.5",
             "user": "suprem", "pass": "Ssplbng#$789"},
            {"node": "kol-cpf-dr-02", "ip": "172.30.145.5", "peer": "172.30.189.4",
             "user": "suprem", "pass": "Ssplbng#$789"},
        ]

        self.create_ui()
        self.start_scheduler()

        self.winfo_toplevel().bind("<Control-Shift-U>", self.restore_tab)
        self.create_notebook_menu()

    # ================= UI =================
    def create_ui(self):

        cols = ("Node","Date","Time","Card1","Card2","Card3","Card4",
                "CPM A","CPM B","Subs","Geo","Role","CPU","Peer Time")

        self.table = ttk.Treeview(self, columns=cols, show="headings")

        for c in cols:
            self.table.heading(c, text=c)
            self.table.column(c, width=110, anchor="center")

        self.table.pack(fill="both", expand=True)

        frame = tk.Frame(self)
        frame.pack(pady=5)

        ttk.Button(frame, text="Run Now", command=self.manual_run).pack(side="left", padx=5)
        ttk.Button(frame, text="Stop", command=self.stop_run).pack(side="left", padx=5)
        ttk.Button(frame, text="Copy", command=self.copy_selected).pack(side="left", padx=5)
        ttk.Button(frame, text="Save Now", command=self.manual_save).pack(side="left", padx=5)

        ttk.Checkbutton(frame, text="Hide Tab",
                        variable=self.hide_var,
                        command=self.toggle_tab).pack(side="left", padx=10)

        self.status = tk.Label(self, text="Waiting...", fg="blue")
        self.status.pack()

    # ================= MENU =================
    def create_notebook_menu(self):
        notebook = self.master
        self.menu = tk.Menu(notebook, tearoff=0)
        self.menu.add_command(label="Restore CP Monitor Tab", command=self.restore_tab)
        notebook.bind("<Button-3>", self.show_menu)

    def show_menu(self, event):
        self.menu.tk_popup(event.x_root, event.y_root)

    # ================= HIDE =================
    def toggle_tab(self):
        if self.hide_var.get():
            self.master.tab(self, state="hidden")
            self.tab_hidden = True

    # ================= RESTORE =================
    def restore_tab(self, event=None):

        if not self.tab_hidden:
            return

        pwd = tk.simpledialog.askstring("Password Required",
                                       "Enter password to restore tab:",
                                       show="*")

        if pwd == self.unhide_password:
            self.master.tab(self, state="normal")
            self.hide_var.set(False)
            self.tab_hidden = False
        else:
            messagebox.showerror("Error", "Wrong Password")

    # ================= COPY =================
    def copy_selected(self):
        items = self.table.selection()
        if not items:
            return

        text = []
        for i in items:
            row = self.table.item(i)["values"]
            text.append("\t".join(map(str, row)))

        self.clipboard_clear()
        self.clipboard_append("\n".join(text))

    # ================= SSH =================
    def get_router_data(self, router):

        if self.stop_flag:
            return None

        cmds = [
            "environment time-stamp",
            "environment no more",
            "show system cpu | match Usage",
            f"show redundancy multi-chassis mc-mobile peer {router['peer']}",
            "show mobile-gateway bng session count",
            'show card detail 1 | match context all "Time of last boot"',
            'show card detail 2 | match context all "Time of last boot"',
            'show card detail 3 | match context all "Time of last boot"',
            'show card detail 4 | match context all "Time of last boot"',
            'show card detail A | match context all "Time of last boot"',
            'show card detail B | match context all "Time of last boot"',
        ]

        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            ssh.connect(router["ip"],
                        username=router["user"],
                        password=router["pass"],
                        timeout=10)

            shell = ssh.invoke_shell()
            time.sleep(2)

            output = ""

            for cmd in cmds:
                if self.stop_flag:
                    break

                shell.send(cmd + "\n")
                time.sleep(10)

                while shell.recv_ready():
                    output += shell.recv(65535).decode(errors="ignore")

            ssh.close()

            if self.stop_flag:
                return None

            boots = re.findall(r"Time of last boot\s*:\s*([^\n]+)", output)

            subs = re.search(r"Total number of sessions\s*:\s*(\d+)", output)
            subs = int(subs.group(1)) if subs else ""

            cpu = re.search(r"Usage.*?(\d+(\.\d+)?%)", output)
            cpu = cpu.group(1) if cpu else ""

            role = re.search(r"Oper\s*Role\s*:\s*([^\n]+)", output)
            role = role.group(1).strip() if role else ""

            geo = re.search(r"Geo-Redundancy State\s*:\s*([^\n]+)", output)
            geo = geo.group(1).strip() if geo else ""

            peer = re.search(r"Last Time Peer Connected\s*:\s*([^\n]+)", output)
            peer = peer.group(1).strip() if peer else ""

            now = datetime.now()

            return [
                router["node"],
                now.strftime("%m-%d-%Y"),
                now.strftime("%H:%M"),
                *(boots + [""]*6)[:6],
                subs,
                geo,
                role,
                cpu,
                peer
            ]

        except Exception as e:
            print("ERROR:", e)
            return None

    # ================= EXCEL =================
    def export_excel(self, rows):

        file = os.path.join(
            self.save_path,
            f"BNG_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )

        if os.path.exists(file):
            wb = load_workbook(file)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append([
                "Node","Date","Time","Card1","Card2","Card3","Card4",
                "CPM A","CPM B","Subs","Geo","Role","CPU","Peer Time"
            ])

        for r in rows:
            ws.append(r)

        wb.save(file)

    # ================= CORE RUN =================
    def run_check(self):

        if self.is_running:
            print("Already running, skipping...")
            return

        self.is_running = True

        try:
            self.stop_flag = False
            rows = []

            for r in self.routers:
                if self.stop_flag:
                    break

                data = self.get_router_data(r)
                if data:
                    rows.append(data)

            if rows:
                self.export_excel(rows)

            self.after(0, self.update_table, rows)

        finally:
            self.is_running = False

    def update_table(self, rows):
        for data in rows:
            self.table.insert("", "end", values=data)

        self.status.config(
            text=f"Waiting... ({datetime.now().strftime('%H:%M:%S')})",
            fg="green"
        )

    # ================= CONTROL =================
    def run_in_thread(self):
        threading.Thread(target=self.run_check, daemon=True).start()

    def manual_run(self):
        print("[MANUAL] Run triggered")

        self.status.config(
            text=f"Running Manual ({datetime.now().strftime('%H:%M:%S')})",
            fg="orange"
        )

        self.run_in_thread()

    def stop_run(self):
        self.stop_flag = True
        self.scheduler_running = False
        self.status.config(text="Stopped", fg="red")

    def manual_save(self):
        rows = [self.table.item(i)["values"] for i in self.table.get_children()]
        if rows:
            self.export_excel(rows)
            messagebox.showinfo("Saved", "Data saved successfully!")

    # ================= SCHEDULER =================
    def scheduler(self):

        last_run_hour = None

        while self.scheduler_running:

            now = datetime.now()
            current_hour = now.strftime("%H")
            current_minute = now.strftime("%M")

            # 🟡 Waiting status
            self.after(0, lambda: self.status.config(
                text=f"Waiting... ({now.strftime('%H:%M:%S')})",
                fg="blue"
            ))

            # ✅ Run within first 2 minutes
            if current_minute in ["00", "01"] and current_hour != last_run_hour:

                self.after(0, lambda: self.status.config(
                    text=f"Running Auto ({now.strftime('%H:%M:%S')})",
                    fg="orange"
                ))

                print(f"[AUTO] Running at {now.strftime('%H:%M:%S')}")
                self.run_in_thread()
                last_run_hour = current_hour

            time.sleep(10)

    def start_scheduler(self):
        threading.Thread(target=self.scheduler, daemon=True).start()

class EpipeModifyTab(SimpleInOutTab):
    SHEET_NAME = "EpipeModify"

    def __init__(self, master):
        super().__init__(master, "EPipe Modifier")

        # Top input fields (LAG + CP)
        top = ttk.Frame(self)
        top.grid(row=2, column=0, columnspan=2, pady=5)

        ttk.Label(top, text="LAG ID:").pack(side="left", padx=5)
        self.lag_entry = ttk.Entry(top, width=10)
        self.lag_entry.pack(side="left", padx=5)

        ttk.Label(top, text="CP Suffix:").pack(side="left", padx=5)
        self.cp_entry = ttk.Entry(top, width=10)
        self.cp_entry.pack(side="left", padx=5)

        ttk.Button(top, text="📋 Copy Output", command=self.copy_output)\
            .pack(side="left", padx=10)

    # ================= CORE LOGIC =================
    def parse_and_generate(self, cfg, lag_input, cp_input):

        epipes = []
        current = None
        current_sap = None

        for line in cfg.splitlines():
            line = line.strip()

            # Detect EPipe
            match = re.search(r'epipe\s+(\d+)', line)
            if match:
                if current:
                    epipes.append(current)

                current = {
                    "epipe_id": match.group(1),
                    "sap": None,
                    "qos": "10",
                    "spokes": []
                }
                current_sap = None
                continue

            if not current:
                continue

            # SAP
            match = re.search(r'sap\s+(\S+)\s+endpoint\s+"(.+?)"', line)
            if match:
                sap_id = match.group(1)
                endpoint = match.group(2)

                if endpoint == "x":
                    current["sap"] = sap_id
                    current_sap = sap_id
                else:
                    current_sap = None
                continue

            # QoS
            match = re.search(r'qos\s+(\d+)', line)
            if match and current_sap:
                current["qos"] = match.group(1)

            # Spoke SDP
            match = re.search(r'spoke-sdp\s+(\S+)\s+endpoint\s+"(.+?)"', line)
            if match:
                current["spokes"].append((match.group(1), match.group(2)))

        if current:
            epipes.append(current)

        # ===== Generate Output =====
        output = []

        for ep in epipes:
            if not ep["sap"]:
                continue

            epipe_last4 = ep["epipe_id"][-4:]
            new_sap = f"lag-{lag_input}:{epipe_last4}.cp-11{cp_input}"
            old_sap = ep["sap"]

            block = []
            block.append(f"/configure service epipe {ep['epipe_id']}")
            block.append("    shutdown")

            for sp, _ in ep["spokes"]:
                block.append(f"    spoke-sdp {sp} shutdown")
                block.append(f"    no spoke-sdp {sp}")

            block.append(f"    sap {old_sap} shutdown")
            block.append(f"    no sap {old_sap}")
            block.append(f"    sap {new_sap} endpoint \"x\" create")
            block.append("        ingress")
            block.append(f"            qos {ep['qos']}")
            block.append("        exit")
            block.append("        no shutdown")
            block.append("    exit")

            for sp, endpoint in ep["spokes"]:
                block.append(f"    spoke-sdp {sp} endpoint \"{endpoint}\" icb create")
                block.append("        no shutdown")
                block.append("    exit")

            block.append("    no shutdown")

            output.append("\n".join(block))

        return "\n\n".join(output)

    # ================= BUTTON ACTION =================
    def generate(self):
        cfg = self.input.get("1.0", tk.END).strip()
        lag_input = self.lag_entry.get().strip()
        cp_input = self.cp_entry.get().strip()

        if not cfg or not lag_input or not cp_input:
            messagebox.showerror("Error", "All fields required!", parent=self)
            return

        result = self.parse_and_generate(cfg, lag_input, cp_input)

        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, result)

    def copy_output(self):
        self.clipboard_clear()
        self.clipboard_append(self.output.get("1.0", tk.END))
        messagebox.showinfo("Copied", "Output copied!", parent=self)

class VPLSConvertTab(SimpleInOutTab):
    SHEET_NAME = "VPLS-Convert-Delete"

    def __init__(self, master):
        super().__init__(master, "VPLS Config → Delete Config")

    def generate(self):
        raw = self.input.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showerror("Error",
                                 "Please paste the full VPLS config(s).",
                                 parent=self)
            return

        try:
            vpls_blocks = re.split(r"\bvpls\s+\d+", raw)
            vpls_ids = re.findall(r"\bvpls\s+(\d+)", raw)

            if not vpls_ids:
                messagebox.showerror("Error", "No VPLS configs found.", parent=self)
                return

            configs = []
            for idx, vpls_id in enumerate(vpls_ids):
                block = "vpls " + vpls_id
                if idx + 1 < len(vpls_blocks):
                    block += vpls_blocks[idx + 1]
                sdp_match = re.search(r"\bspoke-sdp\s+([\d:]+)", block)
                sdp = sdp_match.group(1) if sdp_match else None
                sap_matches = re.findall(r"\bsap\s+([^\s]+)", block)
                config = []
                config.append(f"/configure service vpls {vpls_id}")
                config.append("    shutdown")
                if sdp:
                    config.append(f"    spoke-sdp {sdp} shutdown")
                    config.append(f"    no spoke-sdp {sdp}")
                for sap in sap_matches:
                    config.append(f"    sap {sap} shutdown")
                    config.append(f"    no sap {sap}")
                config.append("    back")
                config.append(f"    no vpls {vpls_id}")
                config.append(f"exit all ")
                config.append(f"/admin display-config | match context all {vpls_id}")
                config.append(f"admin save")
                configs.append("\n".join(config))

            self.output.delete("1.0", tk.END)
            self.output.insert(tk.END, "\n\n".join(configs))

        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

######## Epipe Migration TAB class #############

class NokiaRouter7250Tab(ttk.Frame):

    def __init__(self, parent):

        super().__init__(parent)


        # ======================================================
        # TITLE
        # ======================================================

        title = tk.Label(
            self,
            text="***NOKIA_Router_7250_Epipe_Migration Tool***",
            font=("Arial", 18, "bold")
        )
        title.pack(pady=10)

        # ======================================================
        # ROUTER LOGIN FRAME
        # ======================================================

        router_frame = tk.LabelFrame(
            self,
            text="Router Login Details",
            font=("Arial", 10, "bold")
        )
        router_frame.pack(fill="x", padx=10, pady=5)

        tk.Label(router_frame, text="Router IP").grid(
            row=0, column=0, padx=5, pady=5
        )

        self.ip_entry = tk.Entry(router_frame, width=20)
        self.ip_entry.grid(row=0, column=1, padx=5)

        tk.Label(router_frame, text="Username").grid(
            row=0, column=2, padx=5
        )

        self.user_entry = tk.Entry(router_frame, width=20)
        self.user_entry.grid(row=0, column=3, padx=5)

        tk.Label(router_frame, text="Password").grid(
            row=0, column=4, padx=5
        )

        self.pass_entry = tk.Entry(
            router_frame,
            show="*",
            width=20
        )
        self.pass_entry.grid(row=0, column=5, padx=5)

        tk.Label(router_frame, text="Port").grid(
            row=0,
            column=6,
            padx=5
        )

        self.port_entry = tk.Entry(
            router_frame,
            width=10
        )
        self.port_entry.insert(0, "22")
        self.port_entry.grid(
            row=0,
            column=7,
            padx=5
        )

        # ======================================================
        # BUTTON FRAME
        # ======================================================

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="Import Excel",
            width=18,
            bg="lightblue",
            command=self.import_excel
        ).grid(row=0, column=0, padx=5)

        tk.Button(
            btn_frame,
            text="Generate Config",
            width=18,
            bg="lightgreen",
            command=self.generate_config
        ).grid(row=0, column=1, padx=5)

        tk.Button(
            btn_frame,
            text="Copy Output",
            width=18,
            bg="orange",
            command=self.copy_output
        ).grid(row=0, column=2, padx=5)

        tk.Button(
            btn_frame,
            text="Export TXT",
            width=18,
            bg="yellow",
            command=self.export_txt
        ).grid(row=0, column=3, padx=5)

        tk.Button(
            btn_frame,
            text="Export Excel",
            width=18,
            bg="pink",
            command=self.export_excel
        ).grid(row=0, column=4, padx=5)

        tk.Button(
            btn_frame,
            text="Push To Router",
            width=18,
            bg="lightcoral",
            command=self.start_push_thread
        ).grid(row=0, column=5, padx=5)

        tk.Button(
            btn_frame,
            text="Clear",
            width=18,
            bg="tomato",
            command=self.clear_output
        ).grid(row=0, column=6, padx=5)

        # ======================================================
        # OUTPUT BOX
        # ======================================================

        self.text_box = scrolledtext.ScrolledText(
            self,
            wrap="none",
            font=("Consolas", 10)
        )

        self.text_box.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=10
        )

    # ======================================================
    # IMPORT EXCEL
    # ======================================================

    def import_excel(self):

        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:

            self.df = pd.read_excel(file_path)
            self.df.columns = self.df.columns.str.strip()

            messagebox.showinfo(
                "Success",
                "Excel imported successfully."
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e)
            )

    # ======================================================
    # GENERATE CONFIG
    # ======================================================

    def generate_config(self):

        if self.df is None:

            messagebox.showwarning(
                "Warning",
                "Import Excel first."
            )
            return

        try:

            output = ""

            for _, row in self.df.iterrows():

                olt_vlan = str(
                    row.get("OLT VLAN", "")
                ).strip()

                olt_name = str(
                    row.get("OLT name", "")
                ).strip()

                downlink = str(
                    row.get("down link port", "")
                ).strip()

                spoke_x = str(
                    row.get('spoke-sdp"x"', "")
                ).strip()

                spoke_y = str(
                    row.get('spoke-sdp"y"', "")
                ).strip()

                if not olt_vlan:
                    continue

                config = f'''/configure service epipe 1{downlink}{olt_vlan} name "{olt_name}" customer 1 create
    description "{olt_name}"
    endpoint "x" create
    exit
    endpoint "y" create
    exit
    service-mtu 9014
    sap lag-10:{olt_vlan}.cp-1 endpoint "y" create
        no shutdown
    exit
    sap lag-{downlink}:{olt_vlan}.cp-11{downlink} endpoint "x" create
        ingress
            qos 10
        exit
        no shutdown
    exit
    spoke-sdp 100:{spoke_y} endpoint "y" icb create
        no shutdown
    exit
    spoke-sdp 100:{spoke_x} endpoint "x" icb create
        no shutdown
    exit
    no shutdown

'''

                output += config

            self.generated_output = output

            self.text_box.delete(
                1.0,
                tk.END
            )

            self.text_box.insert(
                tk.END,
                output
            )

            messagebox.showinfo(
                "Success",
                "Configuration generated successfully."
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                str(e)
            )

    # ======================================================
    # COPY OUTPUT
    # ======================================================

    def copy_output(self):

        if not self.generated_output:
            return

        self.clipboard_clear()
        self.clipboard_append(
            self.generated_output
        )

        messagebox.showinfo(
            "Success",
            "Output copied successfully."
        )

    # ======================================================
    # EXPORT TXT
    # ======================================================

    def export_txt(self):

        if not self.generated_output:
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")]
        )

        if not file_path:
            return

        with open(
            file_path,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(self.generated_output)

        messagebox.showinfo(
            "Success",
            "TXT exported successfully."
        )

    # ======================================================
    # EXPORT EXCEL
    # ======================================================

    def export_excel(self):

        if not self.generated_output:
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )

        if not file_path:
            return

        lines = self.generated_output.splitlines()

        pd.DataFrame(
            lines,
            columns=["Configuration Output"]
        ).to_excel(
            file_path,
            index=False
        )

        messagebox.showinfo(
            "Success",
            "Excel exported successfully."
        )

    # ======================================================
    # PUSH CONFIG
    # ======================================================

    def start_push_thread(self):

        threading.Thread(
            target=self.push_to_router,
            daemon=True
        ).start()

    def push_to_router(self):

        if not self.generated_output:

            messagebox.showwarning(
                "Warning",
                "Generate config first."
            )
            return

        try:

            ip = self.ip_entry.get().strip()
            username = self.user_entry.get().strip()
            password = self.pass_entry.get().strip()
            port = int(self.port_entry.get().strip())

            self.text_box.insert(
                tk.END,
                "\nConnecting to router...\n"
            )

            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(
                paramiko.AutoAddPolicy()
            )

            ssh.connect(
                hostname=ip,
                username=username,
                password=password,
                port=port,
                timeout=10
            )

            shell = ssh.invoke_shell()

            for cmd in self.generated_output.splitlines():

                if not cmd.strip():
                    continue

                shell.send(cmd + "\n")

                self.text_box.insert(
                    tk.END,
                    f"Sent: {cmd}\n"
                )

                self.text_box.see(tk.END)

            ssh.close()

            self.text_box.insert(
                tk.END,
                "\nConfiguration pushed successfully.\n"
            )

            messagebox.showinfo(
                "Success",
                "Configuration pushed successfully."
            )

        except Exception as e:

            messagebox.showerror(
                "Router Error",
                str(e)
            )

    # ======================================================
    # CLEAR
    # ======================================================

    def clear_output(self):

        self.generated_output = ""

        self.text_box.delete(
            1.0,
            tk.END
        )


# =====================================================
# ================= IP SUBNET TAB =====================
# =====================================================

class IPSubnetTab(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)

        import ipaddress

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        # ================= IPv4 TAB =================
        ipv4_tab = tk.Frame(notebook)
        notebook.add(ipv4_tab, text="IPv4 Calculator")

        def calculate_ipv4():
            try:
                network = ipaddress.ip_network(ipv4_entry.get(), strict=False)
                ipv4_output.delete("1.0", tk.END)

                ipv4_output.insert(tk.END, f"Network: {network.network_address}\n")
                ipv4_output.insert(tk.END, f"Broadcast: {network.broadcast_address}\n")
                ipv4_output.insert(tk.END, f"Subnet Mask: {network.netmask}\n")
                ipv4_output.insert(tk.END, f"Wildcard Mask: {network.hostmask}\n")
                ipv4_output.insert(tk.END, f"Total Hosts: {network.num_addresses - 2}\n")

                hosts = list(network.hosts())
                if hosts:
                    ipv4_output.insert(tk.END, f"First Host: {hosts[0]}\n")
                    ipv4_output.insert(tk.END, f"Last Host: {hosts[-1]}\n")

            except:
                messagebox.showerror("Error", "Invalid IPv4 Input")

        tk.Label(ipv4_tab, text="Enter IPv4 with Prefix (192.168.1.1/24)").pack(pady=10)
        ipv4_entry = tk.Entry(ipv4_tab, width=40)
        ipv4_entry.pack()

        tk.Button(ipv4_tab, text="Calculate", command=calculate_ipv4).pack(pady=5)

        ipv4_output = tk.Text(ipv4_tab, height=15)
        ipv4_output.pack(fill="both", expand=True, pady=10)

        # ================= IPv6 TAB =================
        ipv6_tab = tk.Frame(notebook)
        notebook.add(ipv6_tab, text="IPv6 Calculator")

        def calculate_ipv6():
            try:
                network = ipaddress.ip_network(ipv6_entry.get(), strict=False)
                ipv6_output.delete("1.0", tk.END)

                ipv6_output.insert(tk.END, f"Network: {network.network_address}\n")
                ipv6_output.insert(tk.END, f"Prefix: /{network.prefixlen}\n")
                ipv6_output.insert(tk.END, f"Total Addresses: {network.num_addresses}\n")

            except:
                messagebox.showerror("Error", "Invalid IPv6 Input")

        tk.Label(ipv6_tab, text="Enter IPv6 with Prefix (2001:db8::1/64)").pack(pady=10)
        ipv6_entry = tk.Entry(ipv6_tab, width=50)
        ipv6_entry.pack()

        tk.Button(ipv6_tab, text="Calculate", command=calculate_ipv6).pack(pady=5)

        ipv6_output = tk.Text(ipv6_tab, height=15)
        ipv6_output.pack(fill="both", expand=True, pady=10)

        # ================= CIDR TAB =================
        cidr_tab = tk.Frame(notebook)
        notebook.add(cidr_tab, text="CIDR Converter")

        def convert_cidr():
            try:
                prefix = int(cidr_entry.get())
                net = ipaddress.IPv4Network(f"0.0.0.0/{prefix}")
                cidr_output.delete("1.0", tk.END)
                cidr_output.insert(tk.END, f"Subnet Mask: {net.netmask}\n")
                cidr_output.insert(tk.END, f"Wildcard Mask: {net.hostmask}\n")
            except:
                messagebox.showerror("Error", "Enter valid prefix (0-32)")

        tk.Label(cidr_tab, text="Enter CIDR Prefix (Example: 24)").pack(pady=10)
        cidr_entry = tk.Entry(cidr_tab, width=20)
        cidr_entry.pack()

        tk.Button(cidr_tab, text="Convert", command=convert_cidr).pack(pady=5)

        cidr_output = tk.Text(cidr_tab, height=15)
        cidr_output.pack(fill="both", expand=True, pady=10)

# =====================================================
# =========== MULTI DEVICE & PFCP TABS (TK) ===========
# =====================================================

class MultiDeviceRunnerTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.devices = []
        self.results = []
        self.stop_flag = False
        self.output_queue = queue.Queue()
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill="x", pady=6)
        ttk.Button(top, text="📁 Import Devices from Excel",
                   command=self.import_excel).pack(side="left", padx=6)
        ttk.Button(top, text="💾 Save Results (Excel)",
                   command=self.save_results).pack(side="left", padx=6)

        self.device_tree = ttk.Treeview(
            self,
            columns=("Name", "IP", "Type", "Username"),
            show="headings",
            height=5
        )
        for col in ("Name", "IP", "Type", "Username"):
            self.device_tree.heading(col, text=col)
            self.device_tree.column(col, width=150)
        self.device_tree.pack(fill="x", padx=6, pady=6)

        ttk.Label(self, text="Enter Commands (one per line):").pack(anchor="w", padx=6)
        self.script_box = scrolledtext.ScrolledText(self, height=5)
        self.script_box.insert(tk.END, "show version")
        self.script_box.pack(fill="x", padx=6, pady=4)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=6)
        self.run_button = ttk.Button(btn_frame,
                                     text="▶ Run on All Devices (Parallel)",
                                     command=self.run_script_all)
        self.run_button.grid(row=0, column=0, padx=6)
        self.stop_button = ttk.Button(btn_frame,
                                      text="⛔ Stop Execution",
                                      command=self.stop_execution,
                                      state="disabled")
        self.stop_button.grid(row=0, column=1, padx=6)

        ttk.Label(self, text="Execution Output:").pack(anchor="w", padx=6)
        self.output_box = scrolledtext.ScrolledText(self, height=15)
        self.output_box.pack(fill="both", expand=True, padx=6, pady=6)

    def import_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        try:
            df = pd.read_excel(path)
            self.devices = df.to_dict("records")
            self.update_device_table()
            messagebox.showinfo("Imported", f"Imported {len(self.devices)} devices.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_device_table(self):
        for r in self.device_tree.get_children():
            self.device_tree.delete(r)
        for i, d in enumerate(self.devices):
            self.device_tree.insert(
                "",
                "end",
                iid=i,
                values=(d.get("Name", ""),
                        d.get("IP", ""),
                        d.get("Type", ""),
                        d.get("Username", ""))
            )

    def run_script_all(self):
        if not self.devices:
            messagebox.showwarning("No Devices", "Please import devices first.")
            return
        script = self.script_box.get("1.0", tk.END).strip()
        if not script:
            messagebox.showwarning("No Command", "Please enter command(s).")
            return
        self.stop_flag = False
        self.output_box.delete("1.0", tk.END)
        self.run_button.config(state="disabled")
        self.stop_button.config(state="normal")
        threading.Thread(target=self._run_parallel,
                         args=(script,), daemon=True).start()
        self.after(100, self._update_output_box)

    def stop_execution(self):
        self.stop_flag = True
        self.output_queue.put("\n⛔ Stopping execution... please wait.\n")
        self.stop_button.config(state="disabled")

    def _run_parallel(self, script):
        self.results.clear()
        threads = []
        script_lines = [l.strip() for l in script.splitlines() if l.strip()]
        log_dir = "Logs"
        os.makedirs(log_dir, exist_ok=True)
        for dev in self.devices:
            t = threading.Thread(target=self._exec_device,
                                 args=(dev, script_lines, log_dir))
            t.start()
            threads.append(t)
        for t in threads:
            t.join()
        self.run_button.config(state="normal")
        self.stop_button.config(state="disabled")
        if self.stop_flag:
            messagebox.showinfo("Stopped", "Execution stopped by user.")
        else:
            messagebox.showinfo("Completed",
                                "Execution completed on all devices.")

    def _exec_device(self, dev, script_lines, log_dir):
        if self.stop_flag:
            return
        self.output_queue.put(
            f"\n=== Connecting to {dev.get('Name')} ({dev.get('IP')}) ===\n"
        )
        full_output = ""
        try:
            if ConnectHandler is None:
                self.output_queue.put(
                    "Netmiko not installed. Install with: pip install netmiko\n"
                )
                return
            conn = ConnectHandler(
                device_type=dev["Type"],
                host=dev["IP"],
                username=dev["Username"],
                password=dev["Password"]
            )
            self.output_queue.put(f"{dev['IP']}: Connected ✅\n")
            for cmd in script_lines:
                if self.stop_flag:
                    break
                self.output_queue.put(f"\n{dev.get('Name')} > {cmd}\n")
                result = conn.send_command(cmd, expect_string=r"#|\>",
                                           read_timeout=20)
                self.output_queue.put(result + "\n")
                full_output += f"\n\n# {cmd}\n{result}\n"
            try:
                conn.disconnect()
            except Exception:
                pass
            log_filename = os.path.join(
                log_dir,
                f"{dev['Name']}_{dev['IP']}.txt"
            )
            with open(log_filename, "w", encoding="utf-8") as f:
                f.write(full_output.strip())
            self.results.append({
                "Date": datetime.date.today(),
                "Time": datetime.datetime.now().strftime("%H:%M:%S"),
                "Name": dev["Name"],
                "IP": dev["IP"],
                "Output_File": log_filename
            })
        except Exception as e:
            self.output_queue.put(f"❌ {dev.get('IP')} Failed: {e}\n")

    def _update_output_box(self):
        try:
            while True:
                line = self.output_queue.get_nowait()
                self.output_box.insert(tk.END, line)
                self.output_box.see(tk.END)
        except queue.Empty:
            pass
        if self.run_button["state"] == "disabled":
            self.after(100, self._update_output_box)

    def save_results(self):
        if not self.results:
            messagebox.showwarning("No Results", "No data to save.")
            return
        df = pd.DataFrame(self.results)
        filename = f"MultiDevice_Report_{datetime.date.today()}.xlsx"
        if os.path.exists(filename):
            old_df = pd.read_excel(filename)
            new_df = pd.concat([old_df, df], ignore_index=True)
        else:
            new_df = df
        new_df.to_excel(filename, index=False)
        messagebox.showinfo("Saved", f"✅ Results saved to {filename}")

        
###############ILL Genrater class ################

class InterfaceScriptTab(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self.build_ui()

    def build_ui(self):

        # ================= HEADER =================
        header = ttk.Frame(self)
        header.pack(fill="x", pady=8)

        title = tk.Label(
            header,
            text="⚙ Interface Script Generator",
            font=("Segoe UI", 16, "bold"),
            fg="#2C3E50"
        )
        title.pack(side=tk.LEFT, padx=10)

        # ================= BUTTON BAR =================
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", pady=5)

        self.import_btn = tk.Button(
            btn_frame,
            text="📂 Import Excel",
            command=self.load_excel,
            bg="#2E86C1",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=5
        )
        self.import_btn.pack(side=tk.LEFT, padx=5)

        self.save_btn = tk.Button(
            btn_frame,
            text="💾 Save Script",
            command=self.save_output,
            bg="#27AE60",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=5
        )
        self.save_btn.pack(side=tk.LEFT, padx=5)

        # ================= OUTPUT FRAME =================
        output_frame = ttk.LabelFrame(self, text="Generated Script Output")
        output_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.output_text = scrolledtext.ScrolledText(
            output_frame,
            width=100,
            height=30,
            font=("Consolas", 10),
            bg="#1E1E1E",
            fg="#00FF88",
            insertbackground="white"
        )
        self.output_text.pack(fill="both", expand=True, padx=5, pady=5)

    # ---------------- LOAD EXCEL ----------------
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            self.generate_scripts(df)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------------- GENERATE SCRIPT ----------------
    def generate_scripts(self, df):

        self.output_text.delete(1.0, tk.END)

        scripts = []

        for _, row in df.iterrows():

            script = f'''
interface "{row['interface']}" create
    description "{row['description']}"
    address {row['address']}
    sap 1/1/c1/1:{row['olt_vlan']}.{row['ill_vlan']} create
        ingress
            qos {row['qos']}
        exit
        egress
            qos {row['qos']}
        exit
        dist-cpu-protection "dist-cpu-arp-access"
    exit
exit
'''
            scripts.append(script)

        self.output_text.insert(tk.END, "\n".join(scripts))

    # ---------------- SAVE OUTPUT ----------------
    def save_output(self):

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt")]
        )

        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.output_text.get(1.0, tk.END))


    """
    Separate Tab: Replace Outer VLAN and static-host IP in Nokia SR OS config
    Rules:
    - Replace outer VLAN only in 'sap' lines (e.g., 3206.x -> <user>.x)
    - Keep inner VLAN unchanged
    - Replace 'static-host ip' with user-provided IP
    - Leave all other lines untouched
    """
class OuterVLANReplaceTab(ttk.Frame):

    def __init__(self, master):
        super().__init__(master)
        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # Top input (VLAN + IP)
        top = ttk.Frame(main)
        top.pack(fill="x", pady=5)

        ttk.Label(top, text="Outer VLAN:").pack(side="left")
        self.vlan_entry = ttk.Entry(top, width=12)
        self.vlan_entry.pack(side="left", padx=5)

        ttk.Label(top, text="Static IP:").pack(side="left", padx=10)
        self.ip_entry = ttk.Entry(top, width=18)
        self.ip_entry.pack(side="left", padx=5)

        # Input
        inp_frame = ttk.LabelFrame(main, text="Input Config", padding=10)
        inp_frame.pack(fill="both", expand=True, pady=5)

        self.input_text = scrolledtext.ScrolledText(inp_frame, height=12)
        self.input_text.pack(fill="both", expand=True)

        # Buttons
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=5)

        ttk.Button(btn_frame, text="Process", command=self.run_task).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Clear", command=self.clear_all).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Save", command=self.save_output).pack(side="right", padx=5)

        # Output
        out_frame = ttk.LabelFrame(main, text="Output", padding=10)
        out_frame.pack(fill="both", expand=True, pady=5)

        self.output_text = scrolledtext.ScrolledText(out_frame, height=12)
        self.output_text.pack(fill="both", expand=True)

    def run_task(self):
        data = self.input_text.get("1.0", tk.END).strip()
        vlan = self.vlan_entry.get().strip()
        ip = self.ip_entry.get().strip()

        if not data:
            messagebox.showwarning("Warning", "Paste config first")
            return

        if not vlan.isdigit():
            messagebox.showerror("Error", "Invalid VLAN")
            return

        if not ip:
            messagebox.showerror("Error", "Invalid IP")
            return

        output = []

        for line in data.splitlines():

            # ✅ Add /configure service before vpls
            if re.match(r'^\s*vpls', line):
                line = re.sub(r'^\s*vpls', lambda m: m.group(0).replace("vpls", "/configure service vpls", 1), line)

            # ✅ Replace outer VLAN only in SAP lines
            if "sap" in line:
                line = re.sub(r'(\d+)\.(\d+)',
                              lambda m: vlan + "." + m.group(2),
                              line)

            # ✅ Replace static-host IP
            if "static-host ip" in line:
                line = re.sub(r'static-host ip\s+\d+\.\d+\.\d+\.\d+',
                              f'static-host ip {ip}',
                              line)

            output.append(line)

        final = "\n".join(output)

        self.output_text.delete("1.0", tk.END)
        self.output_text.insert(tk.END, final)

    def clear_all(self):
        self.input_text.delete("1.0", tk.END)
        self.output_text.delete("1.0", tk.END)
        self.vlan_entry.delete(0, tk.END)
        self.ip_entry.delete(0, tk.END)

    def save_output(self):
        data = self.output_text.get("1.0", tk.END).strip()

        if not data:
            messagebox.showwarning("Warning", "No output to save")
            return

        file = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")]
        )

        if file:
            with open(file, "w") as f:
                f.write(data)

            messagebox.showinfo("Success", "File saved successfully")


# ✅ Add this in your main tool
def add_outer_vlan_tab(notebook):
    tab = OuterVLANReplaceTab(notebook)
    notebook.add(tab, text="Outer VLAN Replace")


class PFCPMonitorTab:

    def __init__(self, notebook):

        import tkinter as tk
        from tkinter import ttk, messagebox, scrolledtext

        self.tk = tk
        self.ttk = ttk
        self.messagebox = messagebox
        self.scrolledtext = scrolledtext

        from netmiko import ConnectHandler
        self.ConnectHandler = ConnectHandler

        import pandas as pd
        self.pd = pd

        import threading
        self.threading = threading

        import datetime
        self.datetime = datetime

        import re
        self.re = re

        import os
        self.os = os

        from concurrent.futures import ThreadPoolExecutor
        self.ThreadPoolExecutor = ThreadPoolExecutor

        # ======================================================
        # Variables
        # ======================================================

        self.devices = []

        self.results = []

        self.stop_flag = False

        self.device_file_path = (
            r"D:\input file\PFCP_Session.xlsx"
        )

        # ======================================================
        # Main Frame
        # ======================================================

        self.frame = tk.Frame(
            notebook,
            bg="#e6f2ff"
        )

        # ======================================================
        # Title
        # ======================================================

        tk.Label(
            self.frame,
            text="Nokia PFCP Session Monitoring Tool",
            font=("Helvetica", 16, "bold"),
            bg="#e6f2ff"
        ).pack(pady=10)

        # ======================================================
        # Device Table
        # ======================================================

        self.device_tree = ttk.Treeview(
            self.frame,
            columns=("Name", "IP"),
            show="headings",
            height=6
        )

        self.device_tree.heading("Name", text="Name")
        self.device_tree.heading("IP", text="IP")

        self.device_tree.column("Name", width=250)
        self.device_tree.column("IP", width=250)

        self.device_tree.pack(
            fill="x",
            padx=10,
            pady=10
        )

        # ======================================================
        # Command Box
        # ======================================================

        tk.Label(
            self.frame,
            text="Enter Command(s)",
            bg="#e6f2ff",
            font=("Arial", 12, "bold")
        ).pack(pady=5)

        self.script_box = scrolledtext.ScrolledText(
            self.frame,
            width=130,
            height=5
        )

        self.script_box.insert(
            tk.END,
            "show subscriber-mgmt pfcp summary"
        )

        self.script_box.pack(pady=5)

        # ======================================================
        # Buttons
        # ======================================================

        button_frame = tk.Frame(
            self.frame,
            bg="#e6f2ff"
        )

        button_frame.pack(pady=10)

        self.run_button = tk.Button(
            button_frame,
            text="▶ Run",
            bg="#f57c00",
            fg="white",
            font=("Arial", 12, "bold"),
            command=self.run_script_all
        )

        self.run_button.grid(
            row=0,
            column=0,
            padx=10
        )

        self.stop_button = tk.Button(
            button_frame,
            text="🛑 Stop",
            bg="#d32f2f",
            fg="white",
            font=("Arial", 12, "bold"),
            command=self.stop_script
        )

        self.stop_button.grid(
            row=0,
            column=1,
            padx=10
        )

        # ======================================================
        # Progress Bar
        # ======================================================

        progress_frame = tk.Frame(
            self.frame,
            bg="#e6f2ff"
        )

        progress_frame.pack(pady=5)

        style = ttk.Style()

        style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor="#f0f0f0",
            background="#4CAF50",
            thickness=20
        )

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            length=600,
            mode="determinate",
            style="Green.Horizontal.TProgressbar"
        )

        self.progress_bar.pack(
            side="left",
            padx=10
        )

        self.progress_label = tk.Label(
            progress_frame,
            text="Progress: 0%",
            bg="#e6f2ff"
        )

        self.progress_label.pack(side="left")

        # ======================================================
        # Output Box
        # ======================================================

        tk.Label(
            self.frame,
            text="Execution Output",
            bg="#e6f2ff",
            font=("Arial", 12, "bold")
        ).pack(pady=5)

        self.output_box = scrolledtext.ScrolledText(
            self.frame,
            width=130,
            height=15
        )

        self.output_box.pack(
            pady=5,
            padx=10,
            fill="both",
            expand=True
        )

        # ======================================================
        # Initial Load
        # ======================================================

        self.load_devices_from_excel()

        self.scheduler()

    # ==========================================================
    # Load Devices
    # ==========================================================

    def load_devices_from_excel(self):

        try:

            df = self.pd.read_excel(
                self.device_file_path
            )

            self.devices = df.to_dict("records")

            self.update_device_table()

            self.output_box.insert(
                self.tk.END,
                f"\nLoaded {len(self.devices)} devices.\n"
            )

            self.output_box.see(self.tk.END)

        except Exception as e:

            self.messagebox.showerror(
                "Error",
                str(e)
            )

    # ==========================================================
    # Update Device Table
    # ==========================================================

    def update_device_table(self):

        for row in self.device_tree.get_children():

            self.device_tree.delete(row)

        for i, dev in enumerate(self.devices):

            self.device_tree.insert(
                "",
                "end",
                iid=i,
                values=(
                    dev.get("Name", ""),
                    dev.get("IP", "")
                )
            )

    # ==========================================================
    # Parse Output
    # ==========================================================

    def parse_pfcp_summary(self, output):

        data = {
            "PFCP_Sessions": 0,
            "PFCP_Sessions_in_setup": 0
        }

        total = self.re.search(
            r"Number of PFCP sessions\s*:\s*(\d+)",
            output
        )

        setup = self.re.search(
            r"Number of PFCP sessions in setup\s*:\s*(\d+)",
            output
        )

        if total:

            data["PFCP_Sessions"] = int(total.group(1))

        if setup:

            data["PFCP_Sessions_in_setup"] = int(setup.group(1))

        return data

    # ==========================================================
    # Run Script
    # ==========================================================

    def run_script_all(self):

        self.stop_flag = False

        if not self.devices:

            self.messagebox.showwarning(
                "Warning",
                "No devices loaded."
            )

            return

        script = self.script_box.get(
            "1.0",
            self.tk.END
        ).strip()

        if not script:

            self.messagebox.showwarning(
                "Warning",
                "No command entered."
            )

            return

        script_lines = [
            line.strip()
            for line in script.splitlines()
            if line.strip()
        ]

        self.results.clear()

        self.output_box.insert(
            self.tk.END,
            "\nStarting execution...\n"
        )

        self.output_box.see(self.tk.END)

        self.threading.Thread(
            target=self.execute_parallel,
            args=(script_lines,),
            daemon=True
        ).start()

    # ==========================================================
    # Stop
    # ==========================================================

    def stop_script(self):

        self.stop_flag = True

        self.output_box.insert(
            self.tk.END,
            "\nStop requested...\n"
        )

    # ==========================================================
    # Execute Parallel
    # ==========================================================

    def execute_parallel(self, script_lines):

        self.run_button.config(state="disabled")

        total = len(self.devices)

        completed = 0

        with self.ThreadPoolExecutor(max_workers=10) as executor:

            futures = []

            for dev in self.devices:

                futures.append(
                    executor.submit(
                        self.connect_and_collect,
                        dev,
                        script_lines
                    )
                )

            for future in futures:

                if self.stop_flag:
                    break

                result = future.result()

                if result:

                    self.results.append(result)

                completed += 1

                self.update_progress(
                    completed,
                    total
                )

        self.progress_bar["value"] = 100

        self.progress_label.config(
            text="Progress: 100%"
        )

        if self.results:

            self.save_results()

        self.run_button.config(state="normal")

        self.messagebox.showinfo(
            "Completed",
            "Execution Completed"
        )

    # ==========================================================
    # Connect Device
    # ==========================================================

    def connect_and_collect(self, dev, script_lines):

        if self.stop_flag:

            return None

        name = dev.get("Name", "")

        ip = dev.get("IP", "")

        dtype = dev.get("Type", "")

        username = dev.get("Username", "")

        password = dev.get("Password", "")

        try:

            self.output_box.insert(
                self.tk.END,
                f"\nConnecting : {name} ({ip})\n"
            )

            conn = self.ConnectHandler(
                device_type=dtype,
                host=ip,
                username=username,
                password=password,
                timeout=15,
                banner_timeout=15
            )

            conn.send_command(
                "environment no more"
            )

            full_output = ""

            for cmd in script_lines:

                if self.stop_flag:
                    break

                output = conn.send_command(
                    cmd,
                    expect_string=r"#|\>",
                    read_timeout=20
                )

                full_output += output + "\n"

            conn.disconnect()

            parsed = self.parse_pfcp_summary(
                full_output
            )

            if parsed["PFCP_Sessions_in_setup"] <= 0:

                return None

            return {

                "Date":
                    self.datetime.date.today().strftime("%Y-%m-%d"),

                "Time":
                    self.datetime.datetime.now().strftime("%H:%M:%S"),

                "Name":
                    name,

                "IP":
                    ip,

                "Number of PFCP sessions":
                    parsed["PFCP_Sessions"],

                "Number of PFCP sessions in setup":
                    parsed["PFCP_Sessions_in_setup"]
            }

        except Exception as e:

            self.output_box.insert(
                self.tk.END,
                f"\n{name} Failed : {e}\n"
            )

            return None

    # ==========================================================
    # Update Progress
    # ==========================================================

    def update_progress(self, completed, total):

        progress = (completed / total) * 100

        self.progress_bar["value"] = progress

        self.progress_label.config(
            text=f"Progress: {progress:.1f}%"
        )

        self.frame.update_idletasks()

    # ==========================================================
    # Save Results
    # ==========================================================

    def save_results(self):

        report_folder = r"D:\PFCP_Reports"

        self.os.makedirs(
            report_folder,
            exist_ok=True
        )

        timestamp = (
            self.datetime.datetime.now().strftime(
                "%Y-%m-%d_%H-%M-%S"
            )
        )

        filename = self.os.path.join(
            report_folder,
            f"PFCP_Report_{timestamp}.xlsx"
        )

        df = self.pd.DataFrame(self.results)

        # Sorting
        df = df.sort_values(
            by=[
                "Number of PFCP sessions",
                "Number of PFCP sessions in setup"
            ],
            ascending=[False, False]
        )

        # Serial Number
        df.index = range(1, len(df) + 1)

        df.reset_index(inplace=True)

        df.rename(
            columns={"index": "S.NO"},
            inplace=True
        )

        # Column Order
        df = df[
            [
                "S.NO",
                "Date",
                "Time",
                "Name",
                "IP",
                "Number of PFCP sessions",
                "Number of PFCP sessions in setup"
            ]
        ]

        df.to_excel(
            filename,
            index=False
        )

        self.output_box.insert(
            self.tk.END,
            f"\nSaved : {filename}\n"
        )

    # ==========================================================
    # Scheduler
    # ==========================================================

    def scheduler(self):

        current_time = (
            self.datetime.datetime.now().strftime("%M:%S")
        )

        if current_time == "00:00":

            self.output_box.insert(
                self.tk.END,
                "\nScheduled Run Started...\n"
            )

            self.run_script_all()

        self.frame.after(
            1000,
            self.scheduler
        )




# =====================================================
# ====== PYQT5 VPRN COLLECTOR (FROM VOIP FILE) ========
#   - launched in a separate process via --vprn
# =====================================================

# Helpers for parsing Nokia outputs (same logic as VOIP_CDOT_IMS_v2.py)

class VprnCollectorTab(ttk.Frame):

    MAX_THREADS = 5

    def __init__(self, master):
        super().__init__(master)

        self.router_df = None
        self.interfaces_data = []

        self.task_queue = queue.Queue()
        self.result_queue = queue.Queue()

        self.total = 0
        self.done = 0
        self.active_threads = 0

        self._build_ui()
        self.after(100, self._process_queue)

    # ================= UI ================= #
    def _build_ui(self):

        top = ttk.Frame(self)
        top.pack(fill="x", pady=6)

        ttk.Button(top, text="Import Excel",
                   command=self.load_excel).pack(side="left", padx=5)

        self.run_btn = ttk.Button(top, text="Run",
                                 command=self.start)
        self.run_btn.pack(side="left", padx=5)

        ttk.Button(top, text="Save Excel",
                   command=self.save_excel).pack(side="left", padx=5)

        ttk.Button(top, text="Save CSV",
                   command=self.save_csv).pack(side="left", padx=5)

        self.status_var = tk.StringVar(value="Load Excel to start")
        ttk.Label(self, textvariable=self.status_var).pack(anchor="w", padx=5)

        self.progress = ttk.Progressbar(self, length=400)
        self.progress.pack(pady=5)

        # Table
        cols = ["IP", "Router", "Hostname", "VPRN",
                "Interface", "Admin", "Oper",
                "Mode", "Port/SAP", "IPs"]

        self.table = ttk.Treeview(self, columns=cols, show="headings", height=10)

        for c in cols:
            self.table.heading(c, text=c)
            self.table.column(c, width=120)

        self.table.pack(fill="both", expand=True, padx=5, pady=5)

        # Logs
        self.output_box = scrolledtext.ScrolledText(self, height=10)
        self.output_box.pack(fill="both", expand=True, padx=5, pady=5)

    # ================= LOG ================= #
    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.output_box.insert(tk.END, f"[{ts}] {msg}\n")
        self.output_box.see(tk.END)

    # ================= HELPERS ================= #
    def has_srrp_master(self, txt):
        return "master" in txt.lower()

    def extract_vprn(self, txt):
        return list(set(re.findall(r"^\s*(\d+)\s+vprn", txt, re.M | re.I)))

    def parse_interfaces(self, text):
        lines = text.splitlines()
        data = []
        i = 0

        while i < len(lines):
            line = lines[i].strip()

            if not line or line.startswith(("A:", "Interface", "=")):
                i += 1
                continue

            if not lines[i].startswith(" "):
                p = line.split()
                if len(p) >= 4:
                    name, admin, oper, mode = p[:4]
                    port = p[4] if len(p) > 4 else ""

                    ips = []
                    j = i + 1
                    while j < len(lines) and lines[j].startswith(" "):
                        if "/" in lines[j]:
                            ips.append(lines[j].split()[0])
                        j += 1

                    data.append((name, admin, oper, mode, port, ";".join(ips)))
                    i = j
                    continue
            i += 1

        return data

    # ================= LOAD EXCEL ================= #
    def load_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            df = pd.read_excel(path)
            df.columns = df.columns.str.lower()

            required = {"ip", "username", "password", "router_name"}
            if not required.issubset(df.columns):
                raise ValueError("Excel must have: ip, username, password, router_name")

            self.router_df = df
            self.status_var.set(f"Loaded {len(df)} routers")

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ================= START ================= #
    def start(self):

        if self.router_df is None:
            messagebox.showwarning("Warning", "Load Excel first")
            return

        if ConnectHandler is None:
            messagebox.showerror("Error", "Install netmiko")
            return

        self.interfaces_data.clear()
        self.table.delete(*self.table.get_children())
        self.output_box.delete("1.0", tk.END)

        self.total = len(self.router_df)
        self.done = 0
        self.active_threads = 0
        self.progress["value"] = 0

        self.run_btn.config(state="disabled")

        for _, row in self.router_df.iterrows():
            self.task_queue.put(row)

        self._start_threads()

    def _start_threads(self):
        while self.active_threads < self.MAX_THREADS and not self.task_queue.empty():
            row = self.task_queue.get()
            t = threading.Thread(target=self.worker, args=(row,))
            t.daemon = True
            t.start()
            self.active_threads += 1

    # ================= WORKER ================= #
    def worker(self, row):

        ip = str(row["ip"])
        user = str(row["username"])
        pwd = str(row["password"])
        name = str(row["router_name"])

        try:
            self.result_queue.put(("log", f"{ip} connecting..."))

            conn = ConnectHandler(
                device_type="alcatel_sros",
                ip=ip,
                username=user,
                password=pwd
            )

            hostname = ""
            sysinfo = conn.send_command("show system information")
            m = re.search(r"System Name\s*:\s*(\S+)", sysinfo)
            if m:
                hostname = m.group(1)

            srrp = conn.send_command("show srrp")

            if not self.has_srrp_master(srrp):
                self.result_queue.put(("log", f"{ip} skipped (no master)"))
                conn.disconnect()
                self.result_queue.put(("done",))
                return

            svc = conn.send_command("show service service-using vprn")
            vprns = self.extract_vprn(svc)

            for v in vprns:
                out = conn.send_command(f"show router {v} interface")
                parsed = self.parse_interfaces(out)

                for p in parsed:
                    row_data = [ip, name, hostname, v] + list(p)
                    self.result_queue.put(("row", row_data))

            conn.disconnect()
            self.result_queue.put(("log", f"{ip} done"))

        except Exception as e:
            self.result_queue.put(("log", f"{ip} ERROR: {e}"))

        self.result_queue.put(("done",))

    # ================= QUEUE ================= #
    def _process_queue(self):

        try:
            while True:
                item = self.result_queue.get_nowait()

                if item[0] == "log":
                    self.log(item[1])

                elif item[0] == "row":
                    self.interfaces_data.append(item[1])
                    self.table.insert("", "end", values=item[1])

                elif item[0] == "done":
                    self.done += 1
                    self.active_threads -= 1

                    pct = int((self.done / self.total) * 100)
                    self.progress["value"] = pct

                    if self.done >= self.total:
                        self.run_btn.config(state="normal")
                        self.status_var.set("Completed")
                    else:
                        self._start_threads()

        except queue.Empty:
            pass

        self.after(100, self._process_queue)

    # ================= SAVE ================= #
    def save_excel(self):
        if not self.interfaces_data:
            return

        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path:
            pd.DataFrame(self.interfaces_data).to_excel(path, index=False)

    def save_csv(self):
        if not self.interfaces_data:
            return

        path = filedialog.asksaveasfilename(defaultextension=".csv")
        if path:
            pd.DataFrame(self.interfaces_data).to_csv(path, index=False)
            
            
#OLT_details_finder_on_IXR
class NokiaServiceParserTab(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)

        self.data = []
        self.path = tk.StringVar()

        # ================= Top Frame =================
        top = ttk.Frame(self)
        top.pack(fill="x", padx=5, pady=5)

        ttk.Label(top, text="Config File :").pack(side="left", padx=5)

        ttk.Entry(
            top,
            textvariable=self.path,
            width=70
        ).pack(side="left", padx=5, fill="x", expand=True)

        ttk.Button(
            top,
            text="Browse",
            command=self.browse
        ).pack(side="left", padx=5)

        ttk.Button(
            top,
            text="Parse",
            command=self.parse
        ).pack(side="left", padx=5)

        ttk.Button(
            top,
            text="Export Excel",
            command=self.export
        ).pack(side="left", padx=5)

        # ================= Treeview =================
        table = ttk.Frame(self)
        table.pack(fill="both", expand=True, padx=5, pady=5)

        columns = (
            "Type",
            "Service ID",
            "VLAN",
            "Uplink",
            "Downlink"
        )

        self.tree = ttk.Treeview(
            table,
            columns=columns,
            show="headings"
        )

        for col in columns:
            self.tree.heading(col, text=col)

        self.tree.column("Type", width=90, anchor="center")
        self.tree.column("Service ID", width=100, anchor="center")
        self.tree.column("VLAN", width=100, anchor="center")
        self.tree.column("Uplink", width=120, anchor="center")
        self.tree.column("Downlink", width=350)

        yscroll = ttk.Scrollbar(
            table,
            orient="vertical",
            command=self.tree.yview
        )

        xscroll = ttk.Scrollbar(
            table,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

    # ======================================================
    def browse(self):

        filename = filedialog.askopenfilename(
            filetypes=[
                ("Text Files", "*.txt"),
                ("All Files", "*.*")
            ]
        )

        if filename:
            self.path.set(filename)

    # ======================================================
    def parse(self):

        if not self.path.get():
            messagebox.showwarning(
                "Warning",
                "Please select a config file."
            )
            return

        try:

            self.data.clear()
            self.tree.delete(*self.tree.get_children())

            with open(
                self.path.get(),
                encoding="utf-8",
                errors="ignore"
            ) as f:
                txt = f.read()

            starts = list(
                re.finditer(
                    r'^\s*(epipe|vpls)\s+\d+',
                    txt,
                    re.M | re.I
                )
            )

            for i, m in enumerate(starts):

                block = txt[
                    m.start():
                    (
                        starts[i + 1].start()
                        if i + 1 < len(starts)
                        else len(txt)
                    )
                ]

                header = re.search(
                    r'^\s*(epipe|vpls)\s+(\d+)',
                    block,
                    re.M | re.I
                )

                if not header:
                    continue

                service_type = header.group(1).upper()
                service_id = header.group(2)

                vlan = ""

                patterns = [
                    r'name\s+"vlan[-_](\d+)"',
                    r'lag-\d+:(\d+)\.cp-',
                    r'lag-\d+:(\d+)\.\d+',
                    r'lag-\d+:(\d+)'
                ]

                for pattern in patterns:
                    match = re.search(
                        pattern,
                        block,
                        re.I
                    )

                    if match:
                        vlan = match.group(1)
                        break

                lags = []

                for line in block.splitlines():

                    match = re.search(
                        r'^\s*sap\s+lag-(\d+):',
                        line,
                        re.I
                    )

                    if match:

                        if match.group(1) not in lags:
                            lags.append(match.group(1))

                if not lags:
                    continue

                row = (
                    service_type,
                    service_id,
                    vlan,
                    lags[0],
                    ", ".join(f"lag {x}" for x in lags[1:])
                )

                self.data.append(row)

                self.tree.insert(
                    "",
                    "end",
                    values=row
                )

            messagebox.showinfo(
                "Completed",
                f"{len(self.data)} services parsed successfully."
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

    # ======================================================
    def export(self):

        if not self.data:

            messagebox.showwarning(
                "Warning",
                "No data available."
            )

            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel File", "*.xlsx")
            ]
        )

        if not filename:
            return

        try:

            pd.DataFrame(
                self.data,
                columns=[
                    "Type",
                    "Service ID",
                    "VLAN",
                    "Uplink",
                    "Downlink"
                ]
            ).to_excel(
                filename,
                index=False
            )

            messagebox.showinfo(
                "Success",
                "Excel exported successfully."
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )

class OLTVoiceExtractorTab(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)

        self.df = None

        self.pat_vpls = re.compile(r'^\s*vpls\s+\d+\s+name\s+"([^"]+)"')
        self.pat_sap = re.compile(r'sap\s+\S+:(\d+)\.(\d+)')

        # ================= Top Frame =================
        top = ttk.Frame(self)
        top.pack(fill="x", padx=5, pady=5)

        ttk.Button(
            top,
            text="Open Config",
            command=self.load_config
        ).pack(side="left", padx=5)

        ttk.Button(
            top,
            text="Export Excel",
            command=self.export_excel
        ).pack(side="left", padx=5)

        # ================= Table Frame =================
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("olt_vlan", "mgmt_vlan", "olt_type")

        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings"
        )

        self.tree.heading("olt_vlan", text="OLT VLAN")
        self.tree.heading("mgmt_vlan", text="MGMT VLAN")
        self.tree.heading("olt_type", text="OLT Type")

        self.tree.column("olt_vlan", width=120, anchor="center")
        self.tree.column("mgmt_vlan", width=120, anchor="center")
        self.tree.column("olt_type", width=300, anchor="w")

        yscroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        xscroll = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    # =======================================================
    def parse_config(self, filename):

        data = []
        name = None

        with open(filename, encoding="utf-8", errors="ignore") as f:

            for line in f:

                m = self.pat_vpls.search(line)
                if m:
                    name = m.group(1)
                    continue

                m = self.pat_sap.search(line)

                if m and name:
                    data.append([
                        m.group(1),
                        m.group(2),
                        name
                    ])

        return pd.DataFrame(
            data,
            columns=["OLT Vlan", "MGMT Vlan", "OLT Type"]
        ).drop_duplicates()

    # =======================================================
    def load_config(self):

        filename = filedialog.askopenfilename(
            title="Select Config File",
            filetypes=[
                ("Text Files", "*.txt"),
                ("All Files", "*.*")
            ]
        )

        if not filename:
            return

        try:
            self.df = self.parse_config(filename)

            self.tree.delete(*self.tree.get_children())

            for row in self.df.itertuples(index=False):
                self.tree.insert("", "end", values=row)

            messagebox.showinfo(
                "Done",
                f"{len(self.df)} records found."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # =======================================================
    def export_excel(self):

        if self.df is None or self.df.empty:
            messagebox.showwarning(
                "Warning",
                "No data available."
            )
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel File", "*.xlsx")]
        )

        if not filename:
            return

        try:
            self.df.to_excel(filename, index=False)

            messagebox.showinfo(
                "Success",
                "Excel exported successfully."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))


class OLTExtractorTab(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)

        self.df = None

        self.pat_vpls = re.compile(r'^\s*vpls\s+\d+\s+name\s+"([^"]+)"')
        self.pat_sap = re.compile(r'sap\s+\S+:(\d+)\.(\d+)')

        # ================= Top Frame =================
        top = ttk.Frame(self)
        top.pack(fill="x", padx=5, pady=5)

        ttk.Button(
            top,
            text="Open Config",
            command=self.load_config
        ).pack(side="left", padx=5)

        ttk.Button(
            top,
            text="Export Excel",
            command=self.export_excel
        ).pack(side="left", padx=5)

        # ================= Table Frame =================
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("olt_vlan", "mgmt_vlan", "olt_type", "static_ip")

        self.tree = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings"
        )

        self.tree.heading("olt_vlan", text="OLT VLAN")
        self.tree.heading("mgmt_vlan", text="MGMT VLAN")
        self.tree.heading("olt_type", text="OLT Type")
        self.tree.heading("static_ip", text="Static Host IP")

        self.tree.column("olt_vlan", width=120, anchor="center")
        self.tree.column("mgmt_vlan", width=120, anchor="center")
        self.tree.column("olt_type", width=300, anchor="w")
        self.tree.column("static_ip", width=150, anchor="center")
        

        yscroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        xscroll = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=yscroll.set,
            xscrollcommand=xscroll.set
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

    # =======================================================
    def parse_config(self, filename):

        data = []
        current_name = ""
        current_olt_vlan = ""
        current_mgmt_vlan = ""

        with open(filename, encoding="utf-8", errors="ignore") as f:

            for line in f:

                m = self.pat_vpls.search(line)
                
                if m:
                    current_name = m.group(1).upper()
                    continue

                m = self.pat_sap.search(line)
                if m:
                    current_olt_vlan = m.group(1)
                    current_mgmt_vlan = m.group(2)
                    continue
                m = re.search(r'static-host\s+ip\s+(\d+\.\d+\.\d+\.\d+)',line,re.I)
                if m and current_olt_vlan:data.append([
                    current_olt_vlan,current_mgmt_vlan,current_name,m.group(1)])
                
                
                current_olt_vlan = ""
                current_mgmt_vlan = ""       
        return pd.DataFrame(
        data,
        columns=[
            "OLT Vlan",
            "MGMT Vlan",
            "OLT Type",
            "Static Host IP"
        ]
    ).drop_duplicates()

    # =======================================================
    def load_config(self):

        filename = filedialog.askopenfilename(
            title="Select Config File",
            filetypes=[
                ("Text Files", "*.txt"),
                ("All Files", "*.*")
            ]
        )

        if not filename:
            return

        try:
            self.df = self.parse_config(filename)

            self.tree.delete(*self.tree.get_children())

            for row in self.df.itertuples(index=False):
                self.tree.insert("", "end", values=row)

            messagebox.showinfo(
                "Done",
                f"{len(self.df)} records found."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    # =======================================================
    def export_excel(self):

        if self.df is None or self.df.empty:
            messagebox.showwarning(
                "Warning",
                "No data available."
            )
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel File", "*.xlsx")]
        )

        if not filename:
            return

        try:
            self.df.to_excel(filename, index=False)

            messagebox.showinfo(
                "Success",
                "Excel exported successfully."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

# =====================================================
# ================ MAIN TKINTER APPLICATION ===========
# =====================================================

class AllInOneApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(APP_GEOMETRY)
        self.theme = ThemeManager(self.root)
        self.settings = load_settings()
        self.theme.apply(self.settings.get("dark_mode", False))
        self._login_screen()

    def _login_screen(self):
        self.login = ttk.Frame(self.root, padding=20)
        self.login.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(self.login, text="Login",
                  font=("Arial", 18)).grid(
            row=0, column=0, columnspan=2, pady=10
        )
        ttk.Label(self.login, text="Username:").grid(
            row=1, column=0, sticky="e", padx=6, pady=6
        )
        self.u = ttk.Entry(self.login)
        self.u.grid(row=1, column=1, padx=6, pady=6)
        ttk.Label(self.login, text="Password:").grid(
            row=2, column=0, sticky="e", padx=6, pady=6
        )
        self.p = ttk.Entry(self.login, show="*")
        self.p.grid(row=2, column=1, padx=6, pady=6)
        ttk.Button(self.login, text="Login",
                   command=self._check_login).grid(
            row=3, column=0, columnspan=2, pady=10
        )
        ttk.Button(self.login, text="Exit",
                   command=self.root.destroy).grid(
            row=4, column=0, columnspan=2
        )
        self.u.focus()  # cursor starts in username
        self.u.bind('<Return>', lambda e: self.p.focus())
        self.p.unbind('<Return>')
        self.p.bind('<Return>', lambda e: self._check_login())
    def _check_login(self):

       users = {
        "Kuchendra": "Kuchendra@1234",
        #"Manjeev": "Manjeev@1234",
        "admin": "9012",
        "Madhu": "Madhu@02",
        #"subhrata": "subhrata"
     }

       username = self.u.get().strip()
       password = self.p.get().strip()

       if username in users and users[username] == password:
         self.login.destroy()
         self._main_window()
       else:
         messagebox.showerror(
            "Login Failed",
            "Invalid username or password",
            parent=self.login
        )

    def _main_window(self):
        bar = ttk.Frame(self.root)
        bar.pack(fill="x", padx=6, pady=6)
        self.mode_btn = ttk.Button(
            bar,
            text="🌙 Dark Mode" if not self.theme.dark else "☀️ Light Mode",
            command=self.toggle_mode
        )
        self.mode_btn.pack(side="left")

        # Developer Name 
        style = ttk.Style()
        style.configure("Dev.TLabel",foreground="white"
                        ,background="#2c3e50",font=("Segoe UI", 9, "bold"),padding=4)
        dev_label = ttk.Label(bar,text="👨‍💻 Developer: Kuchendra Kumar",style="Dev.TLabel",
                              anchor="center")
        dev_label.pack(side="right", padx=6)
        
        # epipe to vpls converter class 
        
        import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


class AdditionalPortRow:
    """Helper row representing a user-defined secondary or extra downlink port."""
    def __init__(self, parent, port_index, remove_callback=None, default_lag="", default_cp="", default_qos="10"):
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill="x", pady=2)
        self.port_index = port_index
        self.remove_callback = remove_callback

        self.lag_var = tk.StringVar(value=default_lag)
        self.cp_var = tk.StringVar(value=default_cp)
        self.qos_var = tk.StringVar(value=default_qos)

        title = f"Port {port_index} (User-Defined)" if port_index == 2 else f"Port {port_index} (Extra)"
        ttk.Label(self.frame, text=f"{title} -> LAG:", width=20, anchor="w").pack(side="left", padx=(0, 4))
        ttk.Entry(self.frame, textvariable=self.lag_var, width=6).pack(side="left", padx=(0, 10))

        ttk.Label(self.frame, text="CP ID:").pack(side="left", padx=(0, 4))
        ttk.Entry(self.frame, textvariable=self.cp_var, width=8).pack(side="left", padx=(0, 10))

        ttk.Label(self.frame, text="QoS:").pack(side="left", padx=(0, 4))
        ttk.Entry(self.frame, textvariable=self.qos_var, width=5).pack(side="left", padx=(0, 10))

        if remove_callback and port_index > 2:
            ttk.Button(self.frame, text="Remove", command=self.destroy).pack(side="left", padx=(5, 0))

    def get_values(self):
        lag = self.lag_var.get().strip()
        cp = self.cp_var.get().strip()
        qos = self.qos_var.get().strip()
        return lag, cp, qos

    def destroy(self):
        self.frame.destroy()
        if self.remove_callback:
            self.remove_callback(self)


class EpipeModifyTab(ttk.Frame):
    """
    Nokia SR OS Epipe-to-VPLS Bulk Converter Tab.
    Embeds directly into a master application Notebook:
        self.t_emod = EpipeModifyTab(self.nb)
        self.nb.add(self.t_emod, text="EpipeModify")
    """
    SERVICE_RE = re.compile(
        r'^\s*epipe\s+(\d+)\s+name\s+("(?:[^"]*)"|\S+)\s+customer\s+(\d+)\s+create\s*$',
        re.IGNORECASE,
    )
    SAP_LAG_RE = re.compile(
        r'sap\s+lag-(\d+):(\d+)',
        re.IGNORECASE,
    )

    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.extra_ports = []
        self._build_ui()

    def _build_ui(self):
        # Header Banner
        header = ttk.Frame(self, padding=(8, 4))
        header.pack(fill="x")
        ttk.Label(header, text="Epipe → VPLS Converter & Deletion Generator", font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Label(header, text="Port 1: Auto-detected | Port 2 & Extra: User-defined", foreground="green").pack(side="right")

        # Configuration Container
        settings = ttk.LabelFrame(self, text="Configuration & Downlink Ports", padding=6)
        settings.pack(fill="x", padx=8, pady=(0, 4))

        # Basic Prefix & Filter Controls
        row1 = ttk.Frame(settings)
        row1.pack(fill="x", pady=(0, 4))

        self.epipe_lead_digit = tk.StringVar(value="1")
        self.vpls_lead_digit = tk.StringVar(value="2")
        self.ignored_uplink_lags = tk.StringVar(value="10")

        ttk.Label(row1, text="Epipe Lead:").pack(side="left", padx=(0, 2))
        ttk.Entry(row1, textvariable=self.epipe_lead_digit, width=4).pack(side="left", padx=(0, 8))

        ttk.Label(row1, text="VPLS Lead:").pack(side="left", padx=(0, 2))
        ttk.Entry(row1, textvariable=self.vpls_lead_digit, width=4).pack(side="left", padx=(0, 8))

        ttk.Label(row1, text="Core/Uplink LAG to skip:").pack(side="left", padx=(0, 2))
        ttk.Entry(row1, textvariable=self.ignored_uplink_lags, width=6).pack(side="left", padx=(0, 12))

        ttk.Button(row1, text="+ Add Another Downlink Port", command=self._add_extra_port_row).pack(side="left", padx=(0, 8))
        ttk.Label(
            row1,
            text="Port 1 auto-extracted (e.g. lag-34 -> VPLS 234XXXX)",
            foreground="#0055aa",
            font=("Segoe UI", 8, "italic")
        ).pack(side="left")

        # Scrollable container for downlink ports
        ports_scroll_frame = ttk.Frame(settings)
        ports_scroll_frame.pack(fill="x", expand=False)

        self.ports_canvas = tk.Canvas(ports_scroll_frame, height=68, highlightthickness=0)
        ports_vbar = ttk.Scrollbar(ports_scroll_frame, orient="vertical", command=self.ports_canvas.yview)
        self.ports_inner_frame = ttk.Frame(self.ports_canvas)

        self.ports_canvas.create_window((0, 0), window=self.ports_inner_frame, anchor="nw")
        self.ports_canvas.configure(yscrollcommand=ports_vbar.set)

        self.ports_canvas.pack(side="left", fill="x", expand=True)
        ports_vbar.pack(side="right", fill="y")

        self.ports_inner_frame.bind(
            "<Configure>",
            lambda e: self.ports_canvas.configure(scrollregion=self.ports_canvas.bbox("all"))
        )
        self.ports_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # Port 2 default definition
        port2 = AdditionalPortRow(self.ports_inner_frame, port_index=2, default_lag="48", default_cp="1148", default_qos="10")
        self.extra_ports.append(port2)

        # Sub-Notebook Tabs inside this Tab
        self.sub_notebook = ttk.Notebook(self)
        self.sub_notebook.pack(fill="both", expand=True, padx=8, pady=2)
        self.sub_tabs = {}
        for key, title, editable in [
            ("input", "Epipe Input", True),
            ("vpls", "VPLS Output", False),
            ("delete", "Epipe Deletion", False),
            ("combined", "Combined Output", False),
        ]:
            frame = ttk.Frame(self.sub_notebook, padding=2)
            self.sub_notebook.add(frame, text=title)
            self.sub_tabs[key] = self._make_text(frame, editable)

        # Actions Bar
        actions = ttk.Frame(self, padding=(8, 4))
        actions.pack(fill="x")
        buttons = [
            ("Convert", self.convert),
            ("Load TXT/CFG", self.load_input),
            ("Save VPLS", lambda: self.save_text("vpls", "VPLS_Configuration.txt")),
            ("Save Epipe Deletion", lambda: self.save_text("delete", "Epipe_Deletion.txt")),
            ("Save Combined", lambda: self.save_text("combined", "VPLS_and_Epipe_Deletion.txt")),
            ("Copy Current Tab", self.copy_current),
        ]
        for label, command in buttons:
            ttk.Button(actions, text=label, command=command).pack(side="left", padx=2)
        ttk.Button(actions, text="Clear All", command=self.clear_all).pack(side="right", padx=2)

        # Status Bar
        self.status = tk.StringVar(value="Ready — paste or load Epipe config.")
        ttk.Label(self, textvariable=self.status, relief="sunken", anchor="w", padding=3).pack(fill="x", side="bottom")

    def _on_mousewheel(self, event):
        try:
            self.ports_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def _add_extra_port_row(self):
        next_idx = len(self.extra_ports) + 2
        new_row = AdditionalPortRow(
            self.ports_inner_frame,
            port_index=next_idx,
            remove_callback=self._remove_extra_port_row,
            default_lag="",
            default_cp="",
            default_qos="10"
        )
        self.extra_ports.append(new_row)
        needed_h = min(110, len(self.extra_ports) * 32)
        self.ports_canvas.configure(height=needed_h)

    def _remove_extra_port_row(self, row_obj):
        if row_obj in self.extra_ports:
            self.extra_ports.remove(row_obj)
        needed_h = min(110, max(40, len(self.extra_ports) * 32))
        self.ports_canvas.configure(height=needed_h)

    @staticmethod
    def _make_text(parent, editable):
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        widget = tk.Text(parent, wrap="none", font=("Consolas", 10), undo=editable,
                         bg="#17212b" if editable else "#111a22",
                         fg="#e7edf3" if editable else "#9ef0ad", insertbackground="white",
                         padx=6, pady=6)
        widget.grid(row=0, column=0, sticky="nsew")
        ybar = ttk.Scrollbar(parent, orient="vertical", command=widget.yview)
        ybar.grid(row=0, column=1, sticky="ns")
        xbar = ttk.Scrollbar(parent, orient="horizontal", command=widget.xview)
        xbar.grid(row=1, column=0, sticky="ew")
        widget.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        if not editable:
            widget.configure(state="disabled")
        return widget

    @staticmethod
    def _get(widget):
        return widget.get("1.0", "end-1c")

    @staticmethod
    def _set(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", text)
        widget.configure(state="disabled")

    @staticmethod
    def _indent(line):
        return len(line) - len(line.lstrip())

    @classmethod
    def _extract_block(cls, lines, start):
        base = cls._indent(lines[start])
        body = []
        i = start + 1
        while i < len(lines):
            if lines[i].strip().lower() == "exit" and cls._indent(lines[i]) == base:
                while body and not body[-1].strip():
                    body.pop()
                return body, i + 1
            body.append(lines[i])
            i += 1
        raise ValueError(f"Missing closing exit for: {lines[start].strip()}")

    def _parse_services(self, raw):
        lines = raw.splitlines()
        starts = [i for i, line in enumerate(lines) if self.SERVICE_RE.match(line)]
        if not starts:
            raise ValueError('No Epipe headers found. Expected: epipe 1213222 name "..." customer 1 create')
        services = []
        for n, start in enumerate(starts):
            end = starts[n + 1] if n + 1 < len(starts) else len(lines)
            services.append(self._parse_service(lines[start:end]))
        return services

    def _parse_service(self, lines):
        match = self.SERVICE_RE.match(lines[0])
        if not match:
            raise ValueError("Invalid Epipe header: " + lines[0])
        sid, raw_name, customer = match.groups()
        name = raw_name[1:-1] if raw_name.startswith('"') else raw_name
        service = {"id": sid, "name": name, "customer": customer,
                   "description": None, "mtu": None, "saps": [], "spokes": []}
        i = 1
        while i < len(lines):
            item = lines[i].strip()
            if not item:
                i += 1
                continue
            if item.startswith("description "):
                service["description"] = item
                i += 1
            elif item.startswith("service-mtu "):
                service["mtu"] = item
                i += 1
            elif re.match(r"^sap\s+", item, re.I):
                body, i = self._extract_block(lines, i)
                header = re.sub(r'\s+endpoint\s+"[^"]*"', "", item, flags=re.I)
                header = re.sub(r"\s+create\s*$", "", header, flags=re.I)
                service["saps"].append((header, body))
            elif re.match(r"^spoke-sdp\s+", item, re.I):
                body, i = self._extract_block(lines, i)
                m = re.match(r"^spoke-sdp\s+(\S+)", item, re.I)
                if m:
                    service["spokes"].append((m.group(1), body))
            else:
                i += 1
        if not service["saps"]:
            raise ValueError(f'Epipe {sid} has no SAP blocks.')
        return service

    def _detect_port1_and_vlan(self, service):
        sid = service["id"]
        ignored = [x.strip() for x in self.ignored_uplink_lags.get().split(",") if x.strip()]
        detected_port = None
        detected_vlan = None

        for sap_header, _ in service.get("saps", []):
            m = self.SAP_LAG_RE.search(sap_header)
            if m:
                lag = m.group(1)
                vlan = m.group(2)
                if lag not in ignored:
                    detected_port = lag
                    detected_vlan = vlan
                    break

        if not detected_vlan:
            for sap_header, _ in service.get("saps", []):
                m = self.SAP_LAG_RE.search(sap_header)
                if m:
                    detected_port = m.group(1)
                    detected_vlan = m.group(2)
                    break

        if not detected_vlan:
            lead = self.epipe_lead_digit.get().strip()
            if sid.startswith(lead) and len(sid) > len(lead) + 2:
                rem = sid[len(lead):]
                detected_port = rem[:2]
                detected_vlan = rem[2:]
            else:
                detected_vlan = sid[-4:]
                detected_port = "30"

        return detected_port, detected_vlan

    @staticmethod
    def _body_lines(body, indent=8):
        nonblank = [len(x) - len(x.lstrip()) for x in body if x.strip()]
        source_min = min(nonblank) if nonblank else 0
        pad = " " * indent
        result = []
        for line in body:
            if not line.strip():
                result.append("")
            else:
                stripped_line = line.lstrip() if source_min == 0 else line[source_min:]
                result.append(pad + stripped_line)
        return result

    def _convert_one(self, service):
        port1_lag, olt_vlan = self._detect_port1_and_vlan(service)
        vpls_lead = self.vpls_lead_digit.get().strip()
        vpls_id = f"{vpls_lead}{port1_lag}{olt_vlan}"

        vpls_name = f'{service["name"]}_{olt_vlan}'
        shg = f"vpls-{olt_vlan}"

        out = [f'/configure service vpls {vpls_id} name "{vpls_name}" customer {service["customer"]} create']
        if service["description"]:
            out.append(f'    description "{vpls_name}"')
        if service["mtu"]:
            out.append(f'    {service["mtu"]}')
        out += [f'    split-horizon-group "{shg}" create', '    exit', '    stp', '        shutdown', '    exit']

        for index, (header, body) in enumerate(service["saps"]):
            shg_text = f' split-horizon-group "{shg}"' if index > 0 else ""
            out.append(f'    {header}{shg_text} create')
            out.extend(self._body_lines(body, 8))
            out.append('    exit')

        for port_row in self.extra_ports:
            lag, cp, qos = port_row.get_values()
            if lag:
                cp_suffix = f".cp-{cp}" if cp else ""
                out.append(f'    sap lag-{lag}:{olt_vlan}{cp_suffix} split-horizon-group "{shg}" create')
                if qos:
                    out += ['        ingress', f'            qos {qos}', '        exit']
                out += ['        no shutdown', '    exit']

        if service.get("spokes"):
            spoke_id, spoke_body = service["spokes"][0]
            out.append(f'    spoke-sdp {spoke_id} create')
            out.extend(self._body_lines(spoke_body, 8))
            out.append('    exit')

        out += ['    no shutdown', '    info', '    exit all']
        return "\n".join(out)

    @staticmethod
    def _deletion(service):
        sid = service["id"]
        out = [
            f"/configure service epipe {sid}",
            "    shutdown"
        ]

        for spoke_id, _ in service.get("spokes", []):
            out.append(f"    spoke-sdp {spoke_id} shutdown")
            out.append(f"    no spoke-sdp {spoke_id}")

        for sap_header, _ in service.get("saps", []):
            sap_id = re.sub(r"^sap\s+", "", sap_header.strip(), flags=re.I)
            out.append(f"    sap {sap_id} shutdown")
            out.append(f"    no sap {sap_id}")

        out.extend([
            "    back",
            f"    no epipe {sid}",
            "exit all"
        ])
        return "\n".join(out)

    def convert(self):
        raw = self._get(self.sub_tabs["input"])
        if not raw.strip():
            messagebox.showwarning("Input required", "Paste or load Epipe configuration first.", parent=self)
            return
        try:
            services = self._parse_services(raw)
            vpls, deletes, combined, errors = [], [], [], []
            for service in services:
                try:
                    vpls_text = self._convert_one(service)
                    delete_text = self._deletion(service)
                    vpls.append(vpls_text)
                    deletes.append(delete_text)
                    combined.append(vpls_text + "\n\n" + delete_text)
                except Exception as exc:
                    errors.append(f'Epipe {service["id"]}: {exc}')
            self._set(self.sub_tabs["vpls"], "\n\n".join(vpls))
            self._set(self.sub_tabs["delete"], "\n\n".join(deletes))
            self._set(self.sub_tabs["combined"], "\n\n".join(combined))
            self.status.set(f"Parsed: {len(services)} | Converted: {len(vpls)} | Errors: {len(errors)}")
            if errors:
                messagebox.showwarning("Conversion warnings", "\n\n".join(errors), parent=self)
            elif vpls:
                self.sub_notebook.select(1)
                messagebox.showinfo("Done", f"Successfully converted {len(vpls)} Epipe service(s).", parent=self)
        except Exception as exc:
            messagebox.showerror("Conversion error", str(exc), parent=self)
            self.status.set("Conversion failed — check input and settings.")

    def load_input(self):
        path = filedialog.askopenfilename(parent=self, filetypes=[("Config/Text", "*.txt *.cfg *.log"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                data = f.read()
            self.sub_tabs["input"].delete("1.0", "end")
            self.sub_tabs["input"].insert("1.0", data)
            self.sub_notebook.select(0)
            self.status.set(f"Loaded: {path}")
        except OSError as exc:
            messagebox.showerror("File error", str(exc), parent=self)

    def save_text(self, key, default_name):
        data = self._get(self.sub_tabs[key])
        if not data.strip():
            messagebox.showwarning("Empty output", "Convert configuration first.", parent=self)
            return
        path = filedialog.asksaveasfilename(parent=self, initialfile=default_name, defaultextension=".txt",
                                            filetypes=[("Text files", "*.txt"), ("Config files", "*.cfg"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(data.rstrip() + "\n")
            self.status.set(f"Saved: {path}")
            messagebox.showinfo("Saved", f"Output saved:\n{path}", parent=self)
        except OSError as exc:
            messagebox.showerror("Save error", str(exc), parent=self)

    def copy_current(self):
        index = self.sub_notebook.index(self.sub_notebook.select())
        key = ["input", "vpls", "delete", "combined"][index]
        data = self._get(self.sub_tabs[key])
        if not data.strip():
            messagebox.showwarning("Empty tab", "Nothing to copy.", parent=self)
            return
        self.clipboard_clear()
        self.clipboard_append(data)
        self.update()
        self.status.set("Current tab copied to clipboard.")

    def clear_all(self):
        if not messagebox.askyesno("Clear all", "Clear input and all generated outputs?", parent=self):
            return
        self.sub_tabs["input"].delete("1.0", "end")
        for key in ("vpls", "delete", "combined"):
            self._set(self.sub_tabs[key], "")
        self.sub_notebook.select(0)
        self.status.set("Ready — all fields cleared.")
        
        # ==========================================
        # MAIN WORK AREA
        # ==========================================
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)
        # ---------- LEFT EXPLORER ----------
        self.left_frame = ttk.Frame(self.main_frame, width=250)
        self.left_frame.pack(side="left", fill="y", padx=(5,0), pady=5)
        ttk.Label(
        self.left_frame,
              text="🔍 Search Tool"
               ).pack(fill="x", padx=5, pady=(5,2))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(
                           self.left_frame,
                                       textvariable=self.search_var)
        self.search_entry.pack(fill="x", padx=5, pady=(0,5))
        self.tree = ttk.Treeview(self.left_frame,show="tree")
        self.tree.pack(fill="both",expand=True,padx=5,pady=5)
        # ---------- RIGHT WORKSPACE ----------
        self.right_frame = ttk.Frame(self.main_frame)
        self.right_frame.pack(side="right", fill="both", expand=True)
        self.nb = ttk.Notebook(self.right_frame)
        self.nb.pack(fill="both", expand=True)
        # Hide notebook tabs
        style = ttk.Style()
        style.layout("TNotebook.Tab", [])
        # Category nodes
        nokia_node = self.tree.insert("","end",text="📁 Nokia Services",open=True)
        monitor_node = self.tree.insert("","end",text="📁 Monitoring",open=True)
        migration_node = self.tree.insert("","end",text="📁 Migration",open=True)
        report_node = self.tree.insert("","end",text="📁 Reports",open=True)
        self.tool_map = {}
        

        self.nb.bind_all("<MouseWheel>", self.on_tab_scroll)
        self.root.bind("<Control-Right>", self.next_tab)
        self.root.bind("<Control-Left>", self.prev_tab)



        # O&M tabs
        self.t_epipe = EpipeGenTab(self.nb)
        self.nb.add(self.t_epipe, text="Epipe_Gen")
        item = self.tree.insert(nokia_node,"end",text="Epipe Generator")
        self.tool_map[item] = self.t_epipe
        

        self.t_vpls = VPLSGenTab(self.nb)
        self.nb.add(self.t_vpls, text="VPLS Generator")
        item = self.tree.insert(nokia_node,"end",text="VPLS Generator")
        self.tool_map[item] = self.t_vpls

        self.t_sapd = SAPDeleteTab(self.nb)
        self.nb.add(self.t_sapd, text="SR_SAP-Del")
        item = self.tree.insert(nokia_node,"end",text="SR_SAP-Del")
        self.tool_map[item] = self.t_sapd

        self.t_erem = EpipeRemoverTab(self.nb)
        self.nb.add(self.t_erem, text="EpipeRemoverTab")
        item = self.tree.insert(nokia_node,"end",text="EpipeRemoverTab")
        self.tool_map[item] = self.t_erem

        self.t_emod = EpipeModifyTab(self.nb)
        self.nb.add(self.t_emod, text="EpipeModify")
        item = self.tree.insert(nokia_node,"end",text="EpipeModify")
        self.tool_map[item] = self.t_emod

        self.t_vconv = VPLSConvertTab(self.nb)
        self.nb.add(self.t_vconv, text="VPLS-Deletion-IXR")
        item = self.tree.insert(nokia_node,"end",text="VPLS-Deletion-IXR")
        self.tool_map[item] = self.t_vconv

        # PFCP Session #########################################

        self.t_pfcp = PFCPMonitorTab(self.nb)
        self.nb.add(self.t_pfcp.frame, text="PFCP_Monitor")
        item = self.tree.insert(monitor_node,"end",text="PFCP_Monitor")
        self.tool_map[item] = self.t_pfcp.frame

        #"SDP Filter"#####################################
        self.t_sdp = SDPFilterTab(self.nb)
        self.nb.add(self.t_sdp, text="Spoke_SDP_Filter")
        item = self.tree.insert(nokia_node,"end",text="Spoke_SDP_Filter")
        self.tool_map[item] = self.t_sdp

        #Local-AS Finder######################################
        self.t_local_as = LocalASFinderTab(self.nb)
        self.nb.add(self.t_local_as, text="Local_AS_Number_Finder")
        item = self.tree.insert(nokia_node,"end",text="Local_AS_Number_Finder")
        self.tool_map[item] = self.t_local_as

        #"Ping Monitor"#######################################
        self.t_ping = PingMonitorTab(self.nb)
        self.nb.add(self.t_ping, text="Ping_Monitor")
        item = self.tree.insert(monitor_node,"end",text="Ping_Monitor")
        self.tool_map[item] = self.t_ping

        #Script Runner########################################
        self.t_script_runner = ScriptRunnerTab(self.nb)
        self.nb.add(self.t_script_runner, text="ScriptRunner") 
        item = self.tree.insert(nokia_node,"end",text="ScriptRunner") 
        self.tool_map[item] = self.t_script_runner

        self.t_ip_subnet = IPSubnetTab(self.nb)
        self.nb.add(self.t_ip_subnet, text="IPSubnet")
        item = self.tree.insert(nokia_node,"end",text="IPSubnet")
        self.tool_map[item] = self.t_ip_subnet

        self.t_CPMonitor = CPMonitorTab(self.nb)
        self.nb.add(self.t_CPMonitor, text="CPMonitorTab")
        item = self.tree.insert(monitor_node,"end",text="CPMonitorTab")
        self.tool_map[item] = self.t_CPMonitor
        ####################### vpls modify tab ###############
        self.t_lag_tab = LagReplaceTab(self.nb)
        self.nb.add(self.t_lag_tab, text="VPLs_Lag_Replace_Change")
        item = self.tree.insert(nokia_node,"end",text="VPLs_Lag_Replace_Change")
        self.tool_map[item] = self.t_lag_tab

        self.t_EPipe_ReplaceGen = EpipeReplaceGenerateTab(self.nb)
        self.nb.add(self.t_EPipe_ReplaceGen, text="Epipe_Port_Change_IXR2")
        item = self.tree.insert(nokia_node,"end",text="Epipe_Port_Change_IXR2")
        self.tool_map[item] = self.t_EPipe_ReplaceGen
        ###########ILL Genrater ###############################
        self.tab_interface = InterfaceScriptTab(self.nb)
        self.nb.add(self.tab_interface, text="ILL_Genrator")
        item = self.tree.insert(migration_node,"end",text="ILL_Genrator")
        self.tool_map[item] = self.tab_interface
        #############"VPRN Collector"##############
        self.t_vprn = VprnCollectorTab(self.nb)
        self.nb.add(self.t_vprn, text="Vprn_Collector")
        item = self.tree.insert(nokia_node,"end",text="Vprn_Collector")
        self.tool_map[item] = self.t_vprn
        
        self.t_outer_vlan = OuterVLANReplaceTab(self.nb)
        self.nb.add(self.t_outer_vlan, text="OLT_SR_Gen OuterVLANReplaceTab")
        item = self.tree.insert(nokia_node,"end",text="OLT_SR_Gen OuterVLANReplaceTab")
        self.tool_map[item] = self.t_outer_vlan
    

        self.vprn_tab = VPRNConfigTab(self.nb)
        self.nb.add(self.vprn_tab, text="VPRN_Config_Gen")
        item = self.tree.insert(migration_node,"end",text="VPRN_Config_Gen")
        self.tool_map[item] = self.vprn_tab
       
       ##############CP_Report_GEN_TAB##################### 
        self.cpreport_tab = CPReportGen(self.nb)
        self.nb.add(self.cpreport_tab, text="CPReportGen")
        item = self.tree.insert(report_node,"end",text="CPReportGen")
        self.tool_map[item] = self.cpreport_tab
        
        ########### Epipe Migration TAB ADD#############

        self.t_nokia7250 = NokiaRouter7250Tab(self.nb)
        self.nb.add(self.t_nokia7250, text="Epipe_For_Migration")
        item = self.tree.insert(migration_node,"end",text="Epipe_For_Migration")
        self.tool_map[item] = self.t_nokia7250
        # Search event
        self.search_var.trace_add("write",self.filter_tools)
        # Tree selection event
        self.tree.bind("<<TreeviewSelect>>",self.open_tool)
        # Open first tab by default
        if self.nb.tabs():self.nb.select(self.nb.tabs()[0])
        #OLT extractor 
        self.olt_tab = OLTExtractorTab(self.nb)
        self.nb.add(self.olt_tab, text="OLT_Extractor_SR")
        item = self.tree.insert(migration_node,"end",text="OLT_Extractor_SR")
        self.tool_map[item] = self.olt_tab
        #####"OLT_details_finder_on_IXR"###########
        self.service_parser_tab = NokiaServiceParserTab(self.nb)
        self.nb.add(self.service_parser_tab, text="OLT_details_finder_on_IXR")
        item = self.tree.insert(migration_node,"end",text="OLT_details_finder_on_IXR")
        self.tool_map[item] = self.service_parser_tab
        
        ##########################################################
        #VPLS Migration tab ###################
        self.vpls_config_tab = VPLSConfigTab1(self.nb)
        self.nb.add(self.vpls_config_tab,text="VPLS_Config_Generator")
        item = self.tree.insert(migration_node,"end",text="VPLS_Config_Generator")
        self.tool_map[item] = self.vpls_config_tab
        
        self.olt_tab = OLTVoiceExtractorTab(self.nb)
        self.nb.add(self.olt_tab, text="OLT_Voice_Extractor_SR")
        item = self.tree.insert(migration_node,"end",text="OLT_Voice_Extractor_SR")
        self.tool_map[item] = self.olt_tab
        
        ############### BBNLINBAND TAB ############
        
        self.tab_interface = BbnlinbandTab(self.nb)
        self.nb.add(self.tab_interface, text="BbnlinbandTab")
        item = self.tree.insert(migration_node,"end",text="BbnlinbandTab")
        self.tool_map[item] = self.tab_interface
        
    def filter_tools(self, *args):
        search = self.search_var.get().lower()
        for parent in self.tree.get_children(""):
            for child in self.tree.get_children(parent):
               text = self.tree.item(child, "text").lower()
               if search in text:  
                   try:
                       self.tree.reattach(child, parent, "end")
                   except:
                       pass
               else:
                    try:
                        self.tree.detach(child)
                    except:
                        pass
                                         
    def open_tool(self, event=None):

        item = self.tree.focus()

        if item not in self.tool_map:
            return
            
        widget = self.tool_map[item]
        
        for idx, tab in enumerate(self.nb.tabs()):
            
            try:
                if str(widget) == str(self.nb.nametowidget(tab)):
                    self.nb.select(idx)
                    break
            except:
                pass
            


    def toggle_mode(self):
        new_mode = not self.theme.dark
        self.theme.apply(new_mode)
        self.mode_btn.configure(
            text="🌙 Dark Mode" if not new_mode else "☀️ Light Mode"
        )
        self.settings["dark_mode"] = new_mode
        save_settings(self.settings)

    def export_all_txt(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile="all_tools_export.txt",
            filetypes=[("Text files", "*.txt")]
        )
        if not path:
            return
        sections = {
            "VPLS Generator": self.t_vpls.get_output_text(),
            "EpipeGen": self.t_epipe.get_output_text(),
            "SAPDelete": self.t_sapd.output.get("1.0", tk.END),
            "EpipeRemover": self.t_erem.get_output_text(),
            "EpipeModify": self.t_emod.output.get("1.0", tk.END),
            "VPLS-Convert-Delete": self.t_vconv.output.get("1.0", tk.END),
            "PFCPMonitor_Output": self.t_pfcp.output_box.get("1.0", tk.END),
        }
        try:
            with open(path, "w", encoding="utf-8") as f:
                for k, v in sections.items():
                    f.write(f"==== {k} ====\n")
                    f.write(v.strip() + "\n\n")
            messagebox.showinfo("Exported", f"All outputs saved to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

    def export_all_excel(self):
        sheets = {
            "VPLS_Generator": self.t_vpls.get_output_text(),
            "EpipeGen": self.t_epipe.get_output_text(),
            "SAPDelete": self.t_sapd.output.get("1.0", tk.END),
            "EpipeRemover": self.t_erem.get_output_text(),
            "EpipeModify": self.t_emod.output.get("1.0", tk.END),
            "VPLS-Convert-Delete": self.t_vconv.output.get("1.0", tk.END),
            "PFCPMonitor": self.t_pfcp.output_box.get("1.0", tk.END),
        }
        export_all_excel(self.root, sheets)

    def open_vprn_collector(self):
        """
        Opens the PyQt5 VPRN Collector in a separate process
        using the same script file with '--vprn' argument.
        """
        if not HAS_PYQT5:
            messagebox.showerror(
                "PyQt5 Missing",
                "PyQt5 is not installed.\nInstall with: pip install PyQt5",
                parent=self.root
            )
            return
        if pd is None:
            messagebox.showwarning(
                "pandas Missing",
                "pandas + openpyxl are required for Excel in VPRN Collector.\n"
                "Install with: pip install pandas openpyxl",
                parent=self.root
            )
        if ConnectHandler is None:
            messagebox.showwarning(
                "Netmiko Missing",
                "netmiko is required for SSH connections in VPRN Collector.\n"
                "Install with: pip install netmiko",
                parent=self.root
            )
        try:
            script_path = os.path.abspath(sys.argv[0])
            subprocess.Popen([sys.executable, script_path, "--vprn"])
        except Exception as e:
            messagebox.showerror(
                "Launch Error",
                f"Failed to launch VPRN Collector:\n{e}",
                parent=self.root
            )
# ===== Mouse Wheel Tab Switching =====
    def on_tab_scroll(self, event):
     return "break"


# ===== Keyboard Navigation =====
    def next_tab(self, event=None):
     current = self.nb.index(self.nb.select())
     total = len(self.nb.tabs())
     self.nb.select((current + 1) % total)


    def prev_tab(self, event=None):
     current = self.nb.index(self.nb.select())
     total = len(self.nb.tabs())
     self.nb.select((current - 1) % total)
# =====================================================
# ===================== ENTRY POINT ====================
# =====================================================
# Run Tkinter main app
root = tk.Tk()
root.geometry(APP_GEOMETRY)

app = AllInOneApp(root)

# ================= TRAY INTEGRATION =================

def on_close():
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)

# ===================================================

root.mainloop()
