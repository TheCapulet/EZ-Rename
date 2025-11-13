#!/usr/bin/env python3
# -*- coding: utf-8 -*-



import os
import re
import sys
import json
import time
import shutil
import threading
import pathlib
import unicodedata
import urllib.parse
import urllib.request
import subprocess
import importlib
from typing import Optional, Tuple, List, Dict, Callable, Set

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# ---------------- Optional deps ----------------
_HAS_REQUESTS = False
try:
    import requests
    _HAS_REQUESTS = True
except Exception:
    _HAS_REQUESTS = False

_HAS_MUTAGEN = False
try:
    from mutagen import File as MutagenFile
    from mutagen.mp4 import MP4
    _HAS_MUTAGEN = True
except Exception:
    MutagenFile = None
    MP4 = None
    _HAS_MUTAGEN = False

APP_TITLE = "EZ-Rename - TV Episode Renamer"
STYLE_NS = "App"
APP_UA = "tvrename-gui/3.2 (+https://api.tvmaze.com)"

VIDEO_EXTS = {'.mkv', '.mp4', '.m4v', '.mov', '.avi', '.ts', '.wmv'}
SUB_EXTS   = {'.srt'}

# ---------------- Regex ----------------
RX_SExx_STR = r'\bS\s*(?P<season>\d{1,2})\s*[-._ ]*E\s*(?P<ep>\d{1,2})\b'
RX_X_STR    = r'\b(?P<season>\d{1,2})\s*[x×]\s*(?P<ep>\d{1,2})\b'
RX_SExx = re.compile(RX_SExx_STR, re.I)
RX_X    = re.compile(RX_X_STR, re.I)

FILENAME_REGEXES = [
    re.compile(r'^(?P<show>.+?)\s*[-._ ]*' + RX_SExx_STR, re.I),
    re.compile(r'^(?P<show>.+?)\s*[-._ ]*' + RX_X_STR,   re.I),
]

NOISE_TOKENS = {
    '1080p','2160p','1440p','720p','480p','web','webrip','webdl','web-dl','hdrip','bdrip','brrip',
    'bluray','blu-ray','hdtv','dvdrip','dvdscr','remux','amzn','amazon','nf','netflix','dsnp','disney',
    'hmax','hbo','paramount','atvp','appletv','appletv+','itv','bbc','hulu','max',
    'extended','proper','repack','internal','x264','h264','x265','h265','hevc','xvid','av1','aac',
    'ddp5','ddp5.1','dd5.1','eac3','dts','truehd','atmos','galaxytv','eztv','rartv','rarbg',
    'ntb','tbs','sva','chs','chs.eng','sub','subs','dub','dual','multi','imax'
}

# user-configurable extra tokens (filled by App)
EXTRA_NOISE_TOKENS: Set[str] = set()

def all_noise_tokens() -> Set[str]:
    return NOISE_TOKENS | EXTRA_NOISE_TOKENS

# ---------------- Tooltip helper ----------------
class Tooltip:
    def __init__(self, widget, text: str, delay: int = 700):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._after_id = None
        self.tipwindow: Optional[tk.Toplevel] = None

        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")
        widget.bind("<ButtonPress>", self._on_leave, add="+")

    def _on_enter(self, _event=None):
        self._schedule()

    def _on_leave(self, _event=None):
        self._unschedule()
        self._hide()

    def _schedule(self):
        self._unschedule()
        self._after_id = self.widget.after(self.delay, self._show)

    def _unschedule(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show(self):
        if self.tipwindow or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 20
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        except Exception:
            return
        tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        frame = ttk.Frame(tw, padding=4, borderwidth=1, relief='solid')
        frame.pack(fill='both', expand=True)
        label = ttk.Label(frame, text=self.text, justify='left', wraplength=360)
        label.pack(fill='both', expand=True)
        self.tipwindow = tw

    def _hide(self):
        if self.tipwindow is not None:
            try:
                self.tipwindow.destroy()
            except Exception:
                pass
            self.tipwindow = None

def create_tooltip(widget, text: str):
    Tooltip(widget, text)
    return widget

# ---------------- HTTP helpers ----------------
class HTTPResult:
    def __init__(self, data: Optional[dict], status: int, error: Optional[str]):
        self.data = data
        self.status = status
        self.error = error

def http_get_json(url: str, timeout: int = 15) -> HTTPResult:
    if _HAS_REQUESTS:
        try:
            r = requests.get(url, headers={"User-Agent": APP_UA}, timeout=timeout)
            if r.status_code == 200:
                try:
                    return HTTPResult(r.json(), 200, None)
                except Exception as e:
                    return HTTPResult(None, r.status_code, f"Invalid JSON: {e}")
            return HTTPResult(None, r.status_code, None)
        except Exception as e:
            return HTTPResult(None, -1, str(e))
    try:
        req = urllib.request.Request(url, headers={"User-Agent": APP_UA})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            try:
                return HTTPResult(json.loads(raw), resp.getcode(), None)
            except Exception as e:
                return HTTPResult(None, resp.getcode(), f"Invalid JSON: {e}")
    except urllib.error.HTTPError as e:
        return HTTPResult(None, e.code, None)
    except Exception as e:
        return HTTPResult(None, -1, str(e))

URL_SEARCH       = "https://api.tvmaze.com/search/shows?q={q}"
URL_EP_BY_NUMBER = "https://api.tvmaze.com/shows/{id}/episodebynumber?season={s}&number={e}"

def tvmaze_search_show_candidates(show_guess: str) -> List[dict]:
    url = URL_SEARCH.format(q=urllib.parse.quote(show_guess))
    r = http_get_json(url)
    if r.status != 200 or r.data is None:
        return []
    return [item.get("show") or {} for item in r.data]

def tvmaze_episode_title(show_id: int, season: int, episode: int) -> Tuple[Optional[str], Optional[str]]:
    url = URL_EP_BY_NUMBER.format(id=show_id, s=season, e=episode)
    r = http_get_json(url)
    if r.status == 404:
        return None, "Not found in TVMaze"
    if r.status != 200 or r.data is None:
        return None, (r.error or f"TVMaze error HTTP {r.status}")
    return r.data.get("name"), None

# ---------------- Parse / Name helpers ----------------
def slug_to_title(s: str) -> str:
    s = re.sub(r'[._\-]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return ' '.join(tok if tok.isupper() else tok.title() for tok in s.split(' '))

def sanitize_show_guess(show: str) -> str:
    s = (show or "").lower()
    tokens = re.split(r'[._\-\s]+', s) if s else []
    noise = all_noise_tokens()
    cleaned = [t for t in tokens if t and t not in noise]
    out = []
    for t in cleaned:
        if re.match(r'^s\d{1,2}e\d{1,2}', t, re.I) or re.match(r'^\d{1,2}[x×]\d{1,2}', t, re.I):
            break
        out.append(t)
    if not out:
        out = cleaned or tokens
    return slug_to_title(' '.join(out).strip())

def parse_filename(name: str) -> Optional[Tuple[Optional[str],int,int,int]]:
    stem = pathlib.Path(name).stem

    for rx in FILENAME_REGEXES:
        m = rx.search(stem)
        if m:
            show = sanitize_show_guess(m.group('show'))
            return (show or None, int(m.group('season')), int(m.group('ep')), m.end())

    m = RX_SExx.search(stem) or RX_X.search(stem)
    if m:
        left = stem[:m.start()]
        show = sanitize_show_guess(re.sub(r'[-._ ]+$', '', left))
        return ((show or None), int(m.group('season')), int(m.group('ep')), m.end())

    return None

PARENS_BLOCK = re.compile(r'\s*[\(\[].*?[\)\]]')
MULTI_SEPS   = re.compile(r'\s*[-._]+\s*')

def extract_existing_title(stem: str, marker_end: int) -> str:
    tail = stem[marker_end:]
    tail = MULTI_SEPS.sub(' ', tail, count=1).strip(' -._')
    tail = PARENS_BLOCK.sub('', tail)
    tail = re.sub(r'\s+', ' ', tail).strip()
    parts = [p for p in re.split(r'[._ ]+', tail) if p]
    stop_tokens = all_noise_tokens() | {'mp4','mkv','m4v','mov','avi','wmv','ts'}
    clean_words = []
    for w in parts:
        clean_words.append(w)
        if w.lower() in stop_tokens:
            break
    title = ' '.join(clean_words).strip(' -._')
    return title if title else tail

def safe_filename(s: str) -> str:
    s = unicodedata.normalize('NFKC', s)
    s = s.replace(':', ' -')
    s = re.sub(r'[\\/<>|?*"„“”]', '', s).strip().rstrip(' .')
    return s

def plan_new_name(old_path: str, title: str, season: int, episode: int) -> str:
    p = pathlib.Path(old_path)
    new_stem = f"S{season:02d}E{episode:02d} - {title}"
    return str(p.with_name(safe_filename(new_stem) + p.suffix))

def iter_video_files(root: str, recursive: bool) -> List[str]:
    paths = []
    root = os.path.abspath(root)
    if recursive:
        for dp, _, fns in os.walk(root):
            for fn in fns:
                if pathlib.Path(fn).suffix.lower() in VIDEO_EXTS:
                    paths.append(os.path.join(dp, fn))
    else:
        for fn in os.listdir(root):
            full = os.path.join(root, fn)
            if os.path.isfile(full) and pathlib.Path(full).suffix.lower() in VIDEO_EXTS:
                paths.append(full)
    return sorted(paths)

def matching_subtitles(video_path: str) -> List[str]:
    p = pathlib.Path(video_path)
    return [str(p.with_suffix(ext)) for ext in SUB_EXTS if p.with_suffix(ext).exists()]

# ---------------- Metadata ----------------
def _which_mkvpropedit() -> Optional[str]:
    p = shutil.which("mkvpropedit") or shutil.which("mkvpropedit.exe")
    if p:
        return p
    if sys.platform.startswith("win"):
        candidate = r"C:\Program Files\MKVToolNix\mkvpropedit.exe"
        if os.path.exists(candidate):
            return candidate
    return None

def set_mkv_title(path: str, title: str) -> Tuple[bool, str]:
    mkvprop = _which_mkvpropedit()
    if not mkvprop:
        return False, "mkvpropedit not found"
    try:
        cmd = [mkvprop, path, "--edit", "info", "--set", f"title={title}"]
        proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if proc.returncode == 0:
            return True, "OK"
        return False, f"mkvpropedit error: {proc.stderr.strip() or proc.stdout.strip()}"
    except Exception as e:
        return False, f"mkvpropedit exec error: {e}"

def set_mp4_title(path: str, title: str) -> Tuple[bool, str]:
    if not _HAS_MUTAGEN or MP4 is None:
        return False, "mutagen not installed"
    try:
        mp = MP4(path)
        mp["\xa9nam"] = [title]
        mp.save()
        return True, "OK"
    except Exception as e:
        return False, f"mutagen error: {e}"

def set_avi_title_mutagen(path: str, title: str) -> Tuple[bool, str]:
    if not _HAS_MUTAGEN or MutagenFile is None:
        return False, "mutagen not installed"
    try:
        f = MutagenFile(path)
        if f is None:
            return False, "mutagen could not open AVI"
        f.tags["INAM"] = title
        f.save()
        return True, "OK"
    except Exception as e:
        return False, f"mutagen AVI error: {e}"

def windows_set_shell_title(path: str, title: str) -> Tuple[bool, str]:
    if not sys.platform.startswith("win"):
        return False, "not Windows"
    try:
        import ctypes
        from ctypes import wintypes

        class PROPERTYKEY(ctypes.Structure):
            _fields_ = [("fmtid", ctypes.c_byte * 16), ("pid", ctypes.c_ulong)]

        def _CLSIDFromString(s: str):
            CLSIDFromString = ctypes.windll.ole32.CLSIDFromString
            CLSIDFromString.argtypes = [wintypes.LPCWSTR, ctypes.POINTER(ctypes.c_byte * 16)]
            clsid = (ctypes.c_byte * 16)()
            hr = CLSIDFromString(s, ctypes.byref(clsid))
            if hr != 0:
                raise OSError("CLSIDFromString failed")
            return clsid

        fmtid = _CLSIDFromString("{F29F85E0-4FF9-1068-AB91-08002B27B3D9}")
        PKEY_Title = PROPERTYKEY(fmtid, 2)

        class PROPVARIANT(ctypes.Structure):
            _fields_ = [
                ("vt", ctypes.c_ushort),
                ("wReserved1", ctypes.c_ubyte),
                ("wReserved2", ctypes.c_ubyte),
                ("wReserved3", ctypes.c_ulong),
                ("pszVal", wintypes.LPWSTR),
                ("padding", ctypes.c_ulong)
            ]
        VT_LPWSTR = 31

        IPropertyStore = ctypes.c_void_p

        SHGetPropertyStoreFromParsingName = ctypes.windll.shell32.SHGetPropertyStoreFromParsingName
        SHGetPropertyStoreFromParsingName.argtypes = [
            wintypes.LPCWSTR, ctypes.c_void_p, ctypes.c_uint32, ctypes.POINTER(ctypes.c_byte * 16),
            ctypes.POINTER(IPropertyStore)
        ]
        SHGetPropertyStoreFromParsingName.restype = ctypes.c_long

        def _ipropstore_setvalue(store, pkey, pvar):
            vtbl = ctypes.cast(store, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
            func = ctypes.CFUNCTYPE(ctypes.c_long, IPropertyStore, ctypes.POINTER(PROPERTYKEY), ctypes.POINTER(PROPVARIANT))(vtbl[7])
            return func(store, ctypes.byref(pkey), ctypes.byref(pvar))

        def _ipropstore_commit(store):
            vtbl = ctypes.cast(store, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
            func = ctypes.CFUNCTYPE(ctypes.c_long, IPropertyStore)(vtbl[8])
            return func(store)

        IID_IPropertyStore = _CLSIDFromString("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}")

        store = IPropertyStore()
        hr = SHGetPropertyStoreFromParsingName(path, None, 0, ctypes.byref(IID_IPropertyStore), ctypes.byref(store))
        if hr != 0:
            return False, f"Shell property store open failed: HRESULT {hr:#x}"

        class PROPVARIANT_LPWSTR(PROPVARIANT):
            pass

        var = PROPVARIANT_LPWSTR()
        var.vt = 31  # VT_LPWSTR
        var.pszVal = title

        if _ipropstore_setvalue(store, PKEY_Title, var) != 0:
            return False, "Shell SetValue failed"
        if _ipropstore_commit(store) != 0:
            return False, "Shell Commit failed"
        return True, "OK"
    except Exception as e:
        return False, f"Windows shell title error: {e}"

def write_title_metadata_any(path: str, new_title: str) -> Tuple[bool, str]:
    ext = pathlib.Path(path).suffix.lower()
    if ext == ".mkv":
        ok, msg = set_mkv_title(path, new_title)
        if ok: return True, "OK"
        ok2, msg2 = windows_set_shell_title(path, new_title)
        return (ok2, msg if ok2 else f"{msg}; fallback: {msg2}")
    if ext in {".mp4", ".m4v", ".mov"}:
        ok, msg = set_mp4_title(path, new_title)
        if ok: return True, "OK"
        ok2, msg2 = windows_set_shell_title(path, new_title)
        return (ok2, msg if ok2 else f"{msg}; fallback: {msg2}")
    if ext == ".avi":
        ok, msg = set_avi_title_mutagen(path, new_title)
        if ok: return True, "OK"
        ok2, msg2 = windows_set_shell_title(path, new_title)
        return (ok2, msg if ok2 else f"{msg}; fallback: {msg2}")
    ok3, msg3 = windows_set_shell_title(path, new_title)
    return (ok3, "OK" if ok3 else f"metadata not supported for {ext}; fallback: {msg3}")

# ---------------- .nfo ----------------
def escape_xml(s: str) -> str:
    return (s.replace("&","&amp;")
             .replace("<","&lt;")
             .replace(">","&gt;")
             .replace('"',"&quot;")
             .replace("'","&apos;"))

def write_nfo_sidecar(video_path: str, show_name: Optional[str], season: int, episode: int, title: str) -> Tuple[bool, str, str]:
    nfo_path = str(pathlib.Path(video_path).with_suffix(".nfo"))
    try:
        show_xml = f"<showtitle>{escape_xml(show_name or '')}</showtitle>" if show_name else ""
        xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<episodedetails>
  <title>{escape_xml(title)}</title>
  <season>{season}</season>
  <episode>{episode}</episode>
  {show_xml}
</episodedetails>
"""
        with open(nfo_path, "w", encoding="utf-8") as f:
            f.write(xml)
        return True, "OK", nfo_path
    except Exception as e:
        return False, f".nfo write error: {e}", nfo_path

# ---------------- Show picker ----------------
class ShowPickerDialog(tk.Toplevel):
    def __init__(self, master, guess: str, cands: List[dict]):
        super().__init__(master)
        self.title(f"Select show for: {guess}")
        self.resizable(False, False)
        self.selected = None

        ttk.Label(self, text=f"Multiple matches for “{guess}”. Select the correct show:").grid(
            row=0, column=0, sticky="w", padx=10, pady=(10,6)
        )

        self.listbox = tk.Listbox(self, width=80, height=10, exportselection=False)
        self.listbox.grid(row=1, column=0, padx=10, sticky="nsew")
        self.cands = cands

        rows = []
        for sh in cands:
            name = sh.get("name") or "Unknown"
            premiered = (sh.get("premiered") or "")[:4] or "—"
            network = (sh.get("network") or {}).get("name") or (sh.get("webChannel") or {}).get("name") or ""
            rows.append(f"{name}  ({premiered})  {network}".strip())
        for r in rows:
            self.listbox.insert('end', r)
        self.listbox.selection_set(0)

        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, sticky="e", padx=10, pady=10)
        ttk.Button(btns, text="OK", command=self._ok).grid(row=0, column=0, padx=4)
        ttk.Button(btns, text="Cancel", command=self._cancel).grid(row=0, column=1, padx=4)

        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())
        self.grab_set()
        self.transient(master)
        self.listbox.focus_set()

    def _ok(self):
        idxs = self.listbox.curselection()
        if not idxs:
            self._cancel(); return
        self.selected = self.cands[idxs[0]]
        self.destroy()

    def _cancel(self):
        self.selected = None
        self.destroy()

# ---------------- Theming ----------------
def apply_theme(root: tk.Tk, listbox: tk.Listbox, dark: bool):
    style = ttk.Style(root)
    try:
        style.theme_use("clam" if dark else "default")
    except tk.TclError:
        style.theme_use("clam")

    if dark:
        BG  = "#101317"; FG = "#E8ECF1"; MUT = "#A5AAB3"
        BTN = "#1A1E24"; ENT = "#181C22"; SEL = "#2A73C5"
        TBK = "#0F1216"; TFG = "#E8ECF1"; TSB = "#1E232A"
        HPB = "#4DA3FF"; BDR = "#2A2F36"; HDR = "#1E232A"
    else:
        BG  = "#F3F4F7"; FG = "#16191D"; MUT = "#6B6F78"
        BTN = "#ECEFF4"; ENT = "#FFFFFF"; SEL = "#CDE6FF"
        TBK = "#FFFFFF"; TFG = "#16191D"; TSB = "#DBDEE3"
        HPB = "#3A6EA5"; BDR = "#D3D7DE"; HDR = "#ECEFF4"

    try:
        root.configure(bg=BG)
    except Exception:
        pass

    PAD = {"padding": (6, 2)}  # slightly tighter vertical padding

    style.configure(f"{STYLE_NS}.TFrame", background=BG)
    style.configure(f"{STYLE_NS}.TLabel", background=BG, foreground=FG, **PAD)

    style.configure(f"{STYLE_NS}.TCheckbutton", background=BG, foreground=FG, **PAD)
    style.map(
        f"{STYLE_NS}.TCheckbutton",
        background=[('active', BG), ('pressed', BG), ('selected', BG), ('disabled', BG)],
        foreground=[('disabled', MUT)]
    )

    style.configure(f"{STYLE_NS}.TButton", background=BTN, foreground=FG, relief='flat', bordercolor=BDR, **PAD)
    style.map(f"{STYLE_NS}.TButton",
              background=[('active', BTN if dark else "#E6EAF0"),
                          ('pressed', BTN if dark else "#E0E5EC")],
              foreground=[('disabled', MUT)])

    style.configure(f"{STYLE_NS}.TEntry", fieldbackground=ENT, foreground=FG, insertcolor=FG, bordercolor=BDR)
    style.configure(f"{STYLE_NS}.TSpinbox", fieldbackground=ENT, foreground=FG, arrowsize=12, bordercolor=BDR)

    style.configure(f"{STYLE_NS}.TPanedwindow", background=BG, bordercolor=BDR, relief='flat')

    # Scrollbars
    style.configure("Vertical.TScrollbar", troughcolor=TSB, background=TSB, bordercolor=BDR)
    style.configure("Horizontal.TScrollbar", troughcolor=TSB, background=TSB, bordercolor=BDR)

    # Treeview
    style.configure("Treeview",
                    background=TBK, foreground=TFG, fieldbackground=TBK,
                    bordercolor=BDR, lightcolor=BDR, darkcolor=BDR, rowheight=22)
    style.map("Treeview",
              background=[('selected', SEL)],
              foreground=[('selected', FG)])
    style.configure("Treeview.Heading", background=HDR, foreground=FG, bordercolor=BDR, relief='flat')
    style.map("Treeview.Heading",
              background=[('active', HDR), ('pressed', HDR)],
              relief=[('active', 'flat'), ('pressed', 'flat')])

    # Progressbar
    style.configure("TProgressbar", troughcolor=TSB, background=HPB, bordercolor=BDR)
    style.configure("Horizontal.TProgressbar", troughcolor=TSB, background=HPB, bordercolor=BDR)
    style.configure("Vertical.TProgressbar", troughcolor=TSB, background=HPB, bordercolor=BDR)

    # Listbox colors
    try:
        listbox.configure(bg=TBK, fg=TFG,
                          selectbackground=SEL, selectforeground=FG,
                          highlightthickness=1, highlightcolor=BDR, highlightbackground=BDR,
                          relief='solid')
    except Exception:
        pass

# ---------------- TSV ----------------
def save_tsv(path: str, rows: List[Tuple[str, ...]], headers: List[str]):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write('\t'.join(headers) + '\n')
        for r in rows:
            f.write('\t'.join(r) + '\n')

def load_backup_tsv(path: str) -> List[Tuple[str, str, str]]:
    """Load backup TSV: TYPE, OLD_PATH, NEW_PATH."""
    out: List[Tuple[str, str, str]] = []
    with open(path, 'r', encoding='utf-8') as f:
        first = True
        for line in f:
            line = line.rstrip('\n')
            if first:
                first = False
                continue  # skip header
            if not line.strip():
                continue
            parts = line.split('\t')
            if len(parts) < 3:
                continue
            typ, old_p, new_p = parts[0], parts[1], parts[2]
            out.append((typ, old_p, new_p))
    return out

# ---------------- Mutagen reload / deps ----------------
def _reload_mutagen_flag() -> bool:
    global _HAS_MUTAGEN, MutagenFile, MP4
    try:
        importlib.invalidate_caches()
        import importlib.util as iu
        spec = iu.find_spec("mutagen.mp4")
        if spec is None:
            _HAS_MUTAGEN = False
            return False
        importlib.import_module("mutagen.mp4")
        from mutagen import File as MF
        from mutagen.mp4 import MP4 as MP4T
        MutagenFile = MF
        MP4 = MP4T
        _HAS_MUTAGEN = True
        return True
    except Exception:
        _HAS_MUTAGEN = False
        return False

def detect_dep_status() -> Tuple[bool, bool, Optional[str]]:
    mut = _reload_mutagen_flag()
    mkv = bool(_which_mkvpropedit())
    what = []
    if mut: what.append("mutagen")
    if mkv: what.append("mkvpropedit")
    return mut, mkv, ", ".join(what) if what else "none"

def run_cmd_capture(cmd: List[str], shell: bool = False) -> Tuple[int, str]:
    try:
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, shell=shell)
        return p.returncode, p.stdout
    except Exception as e:
        return 1, f"[exec error] {e}"

def run_elevated_install_ps(cmdline: str) -> Tuple[int, str]:
    if not sys.platform.startswith("win"):
        return 1, "Elevation only supported on Windows."
    inner = cmdline.replace("'", "''")
    arglist = f"'-NoProfile','-ExecutionPolicy','Bypass','-Command','{inner}'"
    ps_cmd = [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
        "Start-Process", "powershell", "-Verb", "RunAs", "-Wait",
        "-ArgumentList", arglist
    ]
    try:
        p = subprocess.run(ps_cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        return (0 if p.returncode == 0 else p.returncode, p.stdout)
    except Exception as e:
        return 1, f"[elevated exec error] {e}"

def install_mutagen_steps() -> List[List[str]]:
    py = sys.executable or "python"
    return [[py, "-m", "pip", "install", "--user", "--upgrade", "mutagen"]]

def install_mkvtoolnix_steps_windows() -> List[List[str]]:
    steps = []
    if shutil.which("winget") or shutil.which("winget.exe"):
        steps += [
            ["winget", "source", "update"],
            ["winget", "search", "MoritzBunkus.MKVToolNix"],
            ["winget", "install", "--id", "MoritzBunkus.MKVToolNix", "-e",
             "--accept-package-agreements", "--accept-source-agreements"]
        ]
    if shutil.which("choco") or shutil.which("choco.exe"):
        steps.append(["choco", "install", "-y", "mkvtoolnix"])
    return steps

def install_mkvtoolnix_steps_macos() -> List[List[str]]:
    if shutil.which("brew"):
        return [["brew", "install", "mkvtoolnix"]]
    return []

def install_mkvtoolnix_steps_linux() -> List[List[str]]:
    if shutil.which("apt-get"):
        return [["sudo", "apt-get", "update"],
                ["sudo", "apt-get", "install", "-y", "mkvtoolnix"]]
    if shutil.which("dnf"):
        return [["sudo", "dnf", "install", "-y", "mkvtoolnix"]]
    if shutil.which("pacman"):
        return [["sudo", "pacman", "-S", "--noconfirm", "mkvtoolnix"]]
    return []

def is_permission_error(out: str) -> bool:
    s = (out or "").lower()
    return any(tok in s for tok in [
        "access is denied", "access denied", "administrator", "elevated", "permission denied",
        "is denied", "requires administrator"
    ])

class OutputDialog(tk.Toplevel):
    def __init__(self, master, title="Command output"):
        super().__init__(master)
        self.title(title)
        self.geometry("780x420")
        self.resizable(True, True)
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill='both', expand=True)
        self.txt = tk.Text(frm, wrap='word', height=20)
        ysb = ttk.Scrollbar(frm, orient='vertical', command=self.txt.yview)
        self.txt.configure(yscrollcommand=ysb.set)
        self.txt.pack(side='left', fill='both', expand=True)
        ysb.pack(side='right', fill='y')
        ttk.Button(self, text="Close", command=self.destroy).pack(pady=6)
        self.txt.insert('end', "")
        self.txt.configure(state='disabled')

    def append(self, s: str):
        self.txt.configure(state='normal')
        self.txt.insert('end', s + ("\n" if not s.endswith("\n") else ""))
        self.txt.see('end')
        self.txt.configure(state='disabled')
        self.update_idletasks()

# ---------------- Planner ----------------
class RenamePlanner:
    def __init__(
        self,
        folder: str,
        recursive: bool,
        rate_delay: float = 0.40,
        format_only: bool = False,
        choose_show_cb: Optional[Callable[[str, List[dict]], Optional[dict]]] = None,
        override_show: Optional[str] = None,
        prompt_missing_cb: Optional[Callable[[], Optional[str]]] = None,
        write_metadata: bool = True,
        write_meta_if_ok: bool = False,
        write_nfo: bool = False
    ):
        self.folder = os.path.abspath(folder)
        self.recursive = recursive
        self.rate_delay = max(0.0, rate_delay)
        self.format_only = format_only
        self.choose_show_cb = choose_show_cb
        self.write_metadata = write_metadata
        self.write_meta_if_ok = write_meta_if_ok
        self.override_show = (override_show or "").strip() or None
        self.prompt_missing_cb = prompt_missing_cb
        self.write_nfo = write_nfo
        self._prompt_cache: Optional[str] = None

        self.cache: Dict[str, Tuple[int,str]] = {}
        self.changes: List[Tuple[str, ...]] = []
        self.failures: List[Tuple[str,str]] = []
        self.stats: Dict[str,int] = {
            "videos_total": 0,
            "parsed_ok": 0,
            "parsed_fail": 0,
            "show_not_found": 0,
            "ep_not_found": 0,
            "already_correct": 0,
        }
        self._stop = False

    def stop(self): self._stop = True
    def _note(self, path: str, reason: str): self.failures.append((path, reason))

    def _choose_show(self, guess: str, cands: List[dict]) -> Optional[Tuple[int,str]]:
        if not self.choose_show_cb:
            if cands:
                sh = cands[0]
                return sh.get("id"), sh.get("name") or guess
            return None
        chosen = self.choose_show_cb(guess, cands)
        if not chosen: return None
        return chosen.get("id"), chosen.get("name") or guess

    def _resolve_show(self, show_guess: str) -> Optional[Tuple[int,str]]:
        key = (show_guess or "").strip()
        if key in self.cache:
            return self.cache[key]

        def _query(q: str) -> Optional[Tuple[int,str]]:
            q_key = (q or "").strip()
            if q_key in self.cache:
                return self.cache[q_key]
            cands = tvmaze_search_show_candidates(q_key)
            if not cands:
                return None
            ident = (cands[0].get("id"), cands[0].get("name")) if len(cands) == 1 else self._choose_show(q_key, cands)
            if ident:
                self.cache[q_key] = ident
            return ident

        ident = _query(key) if key else None
        if ident:
            return ident

        alt = self._get_show_guess_or_prompt(None)
        alt_key = (alt or "").strip()
        if alt_key and alt_key != key:
            ident = _query(alt_key)
            if ident:
                if key:
                    self.cache[key] = ident
                return ident

        return None

    def _get_show_guess_or_prompt(self, current_guess: Optional[str]) -> Optional[str]:
        if current_guess:
            return current_guess
        if self.override_show:
            return self.override_show
        if self._prompt_cache:
            return self._prompt_cache
        if self.prompt_missing_cb:
            entered = self.prompt_missing_cb()
            if entered:
                self._prompt_cache = entered.strip()
                return self._prompt_cache
        return None

    def scan(self, progress_cb=None, status_cb=None):
        files = iter_video_files(self.folder, self.recursive)
        total = len(files)
        self.stats["videos_total"] = total
        if status_cb: status_cb(f"Scanning {total} video files...")
        self.changes.clear(); self.failures.clear()

        for idx, path in enumerate(files, 1):
            if self._stop:
                if status_cb: status_cb("Scan canceled."); return
            base = os.path.basename(path)
            try:
                parsed = parse_filename(base)
                if not parsed:
                    self.stats["parsed_fail"] += 1
                    self._note(path, "No SxxEyy / 'Sxx Eyy' / 1xYY pattern")
                    if progress_cb: progress_cb(idx, total)
                    continue

                show_guess, season, episode, marker_end = parsed
                self.stats["parsed_ok"] += 1

                if self.format_only:
                    stem = pathlib.Path(base).stem
                    title = extract_existing_title(stem, marker_end) or f"Episode {episode}"
                    new_path = plan_new_name(path, title, season, episode)
                    if os.path.abspath(new_path) == os.path.abspath(path):
                        self.stats["already_correct"] += 1
                        stem_title = pathlib.Path(new_path).stem
                        self.changes.append(("metaonly", path, stem_title))
                    else:
                        self.changes.append(("video", path, new_path))
                        for sub in matching_subtitles(path):
                            sub_new = str(pathlib.Path(new_path).with_suffix(pathlib.Path(sub).suffix))
                            if os.path.abspath(sub_new) != os.path.abspath(sub):
                                self.changes.append(("subtitle", sub, sub_new))
                    if progress_cb: progress_cb(idx, total)
                    continue

                real_guess = self._get_show_guess_or_prompt(show_guess)
                if not real_guess:
                    self.stats["show_not_found"] += 1
                    self._note(path, "Show name missing")
                    if progress_cb: progress_cb(idx, total)
                    continue

                ident = self._resolve_show(real_guess)
                if not ident:
                    self.stats["show_not_found"] += 1
                    self._note(path, f"Show not found in TVMaze (guess='{real_guess}')")
                    if progress_cb: progress_cb(idx, total)
                    continue

                show_id, official = ident
                title, err = tvmaze_episode_title(show_id, season, episode)
                if err or not title:
                    self.stats["ep_not_found"] += 1
                    self._note(path, f"{err or 'Episode not found'} for '{official}' S{season:02d}E{episode:02d}")
                    if progress_cb: progress_cb(idx, total)
                    time.sleep(self.rate_delay); continue

                time.sleep(self.rate_delay)
                new_path = plan_new_name(path, title, season, episode)
                if os.path.abspath(new_path) == os.path.abspath(path):
                    self.stats["already_correct"] += 1
                    stem_title = pathlib.Path(new_path).stem
                    self.changes.append(("metaonly", path, stem_title))
                else:
                    self.changes.append(("video", path, new_path))
                    for sub in matching_subtitles(path):
                        sub_new = str(pathlib.Path(new_path).with_suffix(pathlib.Path(sub).suffix))
                        if os.path.abspath(sub_new) != os.path.abspath(sub):
                            self.changes.append(("subtitle", sub, sub_new))

            except Exception as e:
                self._note(path, f"Error: {e}")

            if progress_cb: progress_cb(idx, total)

        if status_cb:
            s = self.stats
            planned_v = sum(1 for t,_,_ in self.changes if t == "video")
            planned_s = sum(1 for t,_,_ in self.changes if t == "subtitle")
            planned_m = sum(1 for t,_,_ in self.changes if t == "metaonly")
            status_cb(
                f"Scan complete — Videos: {s['videos_total']}, Parsed: {s['parsed_ok']}, "
                f"Unmatched: {s['parsed_fail']}, Show not found: {s['show_not_found']}, "
                f"Episode not found: {s['ep_not_found']}, Already correct: {s['already_correct']}. "
                f"Planned: {planned_v} video, {planned_s} subtitles, {planned_m} metadata writes."
            )

    def apply(self, write_meta_if_ok: bool, write_nfo: bool,
              progress_cb=None, status_cb=None) -> List[Tuple[str, str, str, str]]:
        vids  = [(t,o,n) for (t,o,n) in self.changes if t == "video"]
        subs  = [(t,o,n) for (t,o,n) in self.changes if t == "subtitle"]
        metas = [(t,o,n) for (t,o,n) in self.changes if t == "metaonly"]
        ordered = vids + subs + metas

        results: List[Tuple[str,str,str,str]] = []
        total = len(ordered)
        if status_cb: status_cb(f"Renaming {total} items...")

        for i, (typ, old, new) in enumerate(ordered, 1):
            if self._stop:
                if status_cb: status_cb("Rename canceled."); break
            try:
                if typ == "metaonly":
                    stem_title = new
                    ok, msg = write_title_metadata_any(old, stem_title)
                    results.append(("meta", old, stem_title, "OK" if ok else f"ERR: {msg}"))
                    if progress_cb: progress_cb(i, total)
                    continue

                if typ == "video":
                    final_target = new
                    counter = 2
                    while os.path.exists(final_target):
                        p = pathlib.Path(new)
                        final_target = str(p.with_stem(f"{p.stem} ({counter})"))
                        counter += 1
                    os.rename(old, final_target)
                    results.append((typ, old, final_target, "OK"))

                    stem_title = pathlib.Path(final_target).stem
                    ok, msg = write_title_metadata_any(final_target, stem_title)
                    results.append(("meta", final_target, stem_title, "OK" if ok else f"ERR: {msg}"))

                    if write_nfo:
                        parsed = parse_filename(os.path.basename(final_target)) or (None, 0, 0, 0)
                        show_guess, season, episode, _ = parsed
                        pretty_title = stem_title.split(" - ", 1)[-1]
                        ok, msg, nfo = write_nfo_sidecar(
                            final_target, show_guess, season, episode, pretty_title
                        )
                        results.append(("nfo", final_target, nfo, "OK" if ok else f"ERR: {msg}"))

                elif typ == "subtitle":
                    final_target = new
                    counter = 2
                    while os.path.exists(final_target):
                        p = pathlib.Path(new)
                        final_target = str(p.with_stem(f"{p.stem} ({counter})"))
                        counter += 1
                    os.rename(old, final_target)
                    results.append((typ, old, final_target, "OK"))

            except Exception as e:
                results.append((typ, old, new, f"ERR: {e}"))
            if progress_cb:
                progress_cb(i, total)
            time.sleep(0.02)

        if status_cb: status_cb("Rename pass complete.")
        return results

# ---------------- App ----------------
OPTIONS_FILE = os.path.join(os.path.expanduser("~"), ".tv_renamer_options.json")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x900")
        self.minsize(980, 660)

        self.folder_var = tk.StringVar(value=os.getcwd())
        self.recursive_var = tk.BooleanVar(value=True)
        self.delay_var = tk.DoubleVar(value=0.40)
        self.format_only_var = tk.BooleanVar(value=False)
        self.dark_mode_var = tk.BooleanVar(value=True)  # default dark
        self.write_meta_if_ok_var = tk.BooleanVar(value=True)  # ON by default
        self.write_nfo_var = tk.BooleanVar(value=False)

        # track custom tokens as list for saving
        self.custom_noise_tokens: List[str] = []

        mut, mkv, _ = detect_dep_status()
        self.dep_status_var = tk.StringVar(
            value=f"mutagen={'OK' if mut else 'missing'} | mkvpropedit={'OK' if mkv else 'missing'}"
        )

        self.status_var = tk.StringVar(value="Ready.")
        self.summary_var = tk.StringVar(value="")
        self.planner: Optional[RenamePlanner] = None
        self.last_prompted_show: Optional[str] = None

        self.plan_rows: List[Tuple[str, str, str]] = []
        self.fail_rows: List[Tuple[str, str]] = []
        self.result_rows: List[Tuple[str, str, str, str]] = []

        self._build_ui()
        self._load_options_if_present()
        self.after(0, self._apply_theme_and_fix)

    def _build_ui(self):
        root_pane = ttk.Frame(self, style=f"{STYLE_NS}.TFrame", padding=(10,0,10,4))
        root_pane.pack(fill='x')

        left = ttk.Frame(root_pane, style=f"{STYLE_NS}.TFrame")
        left = ttk.Frame(root_pane, style=f"{STYLE_NS}.TFrame")
        left.grid(row=0, column=0, sticky='ew', padx=(0,10))
        root_pane.columnconfigure(0, weight=1)

        right = ttk.Frame(root_pane, style=f"{STYLE_NS}.TFrame")
        right.grid(row=0, column=1, sticky='n')

        ttk.Label(left, text="Folder:", style=f"{STYLE_NS}.TLabel").grid(row=0, column=0, sticky='w')
        self.ent_folder = ttk.Entry(left, textvariable=self.folder_var, style=f"{STYLE_NS}.TEntry")
        self.ent_folder.grid(row=0, column=1, sticky='ew', padx=6)
        left.columnconfigure(1, weight=2)
        browse_btn = ttk.Button(left, text="Browse…", command=self.on_browse, style=f"{STYLE_NS}.TButton")
        browse_btn.grid(row=0, column=2, padx=(4,0))
        create_tooltip(self.ent_folder, "Root folder containing the video files you want to scan and rename.")
        create_tooltip(browse_btn, "Pick the folder that contains your TV episodes.")

        btns = ttk.Frame(left, style=f"{STYLE_NS}.TFrame", padding=(0,2,0,0))
        btns.grid(row=1, column=0, columnspan=3, sticky='w')
        self.btn_scan = ttk.Button(btns, text="Scan", command=self.on_scan, style=f"{STYLE_NS}.TButton")
        self.btn_scan.grid(row=0, column=0, padx=(0,6))
        create_tooltip(self.btn_scan, "Analyze files and build a rename plan.\nNo files are changed during Scan.")

        self.btn_apply = ttk.Button(btns, text="Apply Renames", command=self.on_apply, state='disabled', style=f"{STYLE_NS}.TButton")
        self.btn_apply.grid(row=0, column=1, padx=6)
        create_tooltip(self.btn_apply, "Perform the planned renames and metadata writes shown below.")

        self.btn_save_backup = ttk.Button(btns, text="Backup Names", command=self.on_save_backup, state='disabled', style=f"{STYLE_NS}.TButton")
        self.btn_save_backup.grid(row=0, column=2, padx=6)
        create_tooltip(self.btn_save_backup, "Save a TSV file with original and new paths.\nUse it later to restore filenames.")

        self.btn_restore_backup = ttk.Button(btns, text="Restore From Backup", command=self.on_restore_backup, style=f"{STYLE_NS}.TButton")
        self.btn_restore_backup.grid(row=0, column=3, padx=6)
        create_tooltip(self.btn_restore_backup, "Load a previously saved backup TSV and rename files back to their original names.")

        self.btn_save_options = ttk.Button(btns, text="Save Options", command=self.on_save_options, style=f"{STYLE_NS}.TButton")
        self.btn_save_options.grid(row=0, column=4, padx=6)
        create_tooltip(self.btn_save_options, "Save current settings (folder, options, custom tokens) to a config file in your home directory.")

        self.btn_check_deps = ttk.Button(btns, text="Check Dependencies", command=self.on_install_deps, style=f"{STYLE_NS}.TButton")
        self.btn_check_deps.grid(row=0, column=5, padx=6)
        create_tooltip(self.btn_check_deps, "Detect or install optional tools (mutagen, mkvpropedit) used for writing embedded metadata.")

        ttk.Label(left, text="Rate delay (seconds):", style=f"{STYLE_NS}.TLabel").grid(row=2, column=0, sticky='w', pady=(2,0))
        self.spin_delay = ttk.Spinbox(left, from_=0.0, to=5.0, increment=0.05, textvariable=self.delay_var, width=6, style=f"{STYLE_NS}.TSpinbox")
        self.spin_delay.grid(row=2, column=1, sticky='w', pady=(2,0))
        create_tooltip(self.spin_delay, "Delay between TVMaze API requests.\nLower = faster, higher = less chance of rate limiting.")

        ttk.Label(right, text="Options", style=f"{STYLE_NS}.TLabel").grid(row=0, column=0, sticky='w')
        chk_dark = ttk.Checkbutton(right, text="Dark mode", variable=self.dark_mode_var, command=self.on_toggle_dark, style=f"{STYLE_NS}.TCheckbutton")
        chk_dark.grid(row=1, column=0, sticky='w')
        create_tooltip(chk_dark, "Toggle between dark and light themes for this app.")

        chk_recursive = ttk.Checkbutton(right, text="Recursive", variable=self.recursive_var, style=f"{STYLE_NS}.TCheckbutton")
        chk_recursive.grid(row=2, column=0, sticky='w')
        create_tooltip(chk_recursive, "If enabled, include video files in all subfolders of the selected folder.")

        chk_format = ttk.Checkbutton(right, text="Format-only", variable=self.format_only_var, style=f"{STYLE_NS}.TCheckbutton")
        chk_format.grid(row=3, column=0, sticky='w')
        create_tooltip(chk_format, "Do NOT call TVMaze.\nJust normalize filenames based on the existing SxxEyy pattern and title.")

        chk_meta_ok = ttk.Checkbutton(right, text="Write metadata even if correct", variable=self.write_meta_if_ok_var, style=f"{STYLE_NS}.TCheckbutton")
        chk_meta_ok.grid(row=4, column=0, sticky='w')
        create_tooltip(chk_meta_ok, "If enabled, embedded title metadata is rewritten even when the filename is already in the target format.")

        chk_nfo = ttk.Checkbutton(right, text="Write .nfo sidecar", variable=self.write_nfo_var, style=f"{STYLE_NS}.TCheckbutton")
        chk_nfo.grid(row=5, column=0, sticky='w')
        create_tooltip(chk_nfo, ".nfo sidecar: small XML file written next to the video.\nUsed by media managers like Kodi/Emby/Plex to read episode info.")

        btn_tokens = ttk.Button(right, text="Custom noise tokens…", command=self.on_edit_noise_tokens, style=f"{STYLE_NS}.TButton")
        btn_tokens.grid(row=6, column=0, sticky='w', pady=(2,0))
        create_tooltip(btn_tokens, "Add extra words to ignore when detecting show names.\nExample: group tags or release labels you always want stripped.")

        ttk.Separator(right, orient='horizontal').grid(row=7, column=0, sticky='ew', pady=4)
        dep_lbl = ttk.Label(right, textvariable=self.dep_status_var, style=f"{STYLE_NS}.TLabel")
        dep_lbl.grid(row=8, column=0, sticky='w')
        create_tooltip(dep_lbl, "Status of optional dependencies used for writing metadata.\nmutagen = MP4/AVI tags, mkvpropedit = MKV title tags.")

        frm_sum = ttk.Frame(self, style=f"{STYLE_NS}.TFrame", padding=(10,2,10,0))
        frm_sum.pack(fill='x')
        ttk.Label(frm_sum, textvariable=self.summary_var, style=f"{STYLE_NS}.TLabel").pack(anchor='w')

        paned = ttk.Panedwindow(self, orient='vertical', style=f"{STYLE_NS}.TPanedwindow")
        paned.pack(fill='both', expand=True, padx=10, pady=(0,6))

        frm_plan = ttk.Frame(paned, style=f"{STYLE_NS}.TFrame")
        paned.add(frm_plan, weight=3)
        ttk.Label(frm_plan, text="Planned Renames:", style=f"{STYLE_NS}.TLabel").pack(anchor='w')

        cols = ("type", "old", "arrow", "new")
        self.tree = ttk.Treeview(frm_plan, columns=cols, show='headings', height=10)
        self.tree.heading("type", text="Type")
        self.tree.heading("old", text="Current Filename")
        self.tree.heading("arrow", text="→")
        self.tree.heading("new", text="New Filename")
        self.tree.column("type", width=100, anchor='center')
        self.tree.column("old", width=520, anchor='w')
        self.tree.column("arrow", width=28, anchor='center')
        self.tree.column("new", width=520, anchor='w')
        ysb = ttk.Scrollbar(frm_plan, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=ysb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        ysb.pack(side='right', fill='y')

        frm_fail = ttk.Frame(paned, style=f"{STYLE_NS}.TFrame")
        paned.add(frm_fail, weight=1)
        ttk.Label(frm_fail, text="Skipped / Failed (with reasons):", style=f"{STYLE_NS}.TLabel").pack(anchor='w')
        self.lst_fail = tk.Listbox(frm_fail, height=8)
        ysb2 = ttk.Scrollbar(frm_fail, orient='vertical', command=self.lst_fail.yview)
        self.lst_fail.configure(yscrollcommand=ysb2.set)
        self.lst_fail.pack(side='left', fill='both', expand=True)
        ysb2.pack(side='right', fill='y')

        frm_bottom = ttk.Frame(self, style=f"{STYLE_NS}.TFrame", padding=(10,0,10,10))
        frm_bottom.pack(fill='x')
        self.progress = ttk.Progressbar(frm_bottom, mode='determinate')
        self.progress.pack(fill='x')
        self.lbl_status = ttk.Label(frm_bottom, textvariable=self.status_var, anchor='w', style=f"{STYLE_NS}.TLabel")
        self.lbl_status.pack(fill='x', pady=(4,0))
        create_tooltip(self.progress, "Progress for the current scan or rename operation.")

    def _apply_theme_and_fix(self):
        apply_theme(self, self.lst_fail, bool(self.dark_mode_var.get()))

    # ----- Options save/load -----
    def _options_dict(self) -> dict:
        return {
            "folder": self.folder_var.get(),
            "recursive": bool(self.recursive_var.get()),
            "delay": float(self.delay_var.get()),
            "format_only": bool(self.format_only_var.get()),
            "dark_mode": bool(self.dark_mode_var.get()),
            "write_meta_if_ok": bool(self.write_meta_if_ok_var.get()),
            "write_nfo": bool(self.write_nfo_var.get()),
            "custom_noise_tokens": sorted(list(EXTRA_NOISE_TOKENS)),
        }

    def _load_options_if_present(self):
        global EXTRA_NOISE_TOKENS
        try:
            if os.path.exists(OPTIONS_FILE):
                with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
                    d = json.load(f)
                self.folder_var.set(d.get("folder", self.folder_var.get()))
                self.recursive_var.set(bool(d.get("recursive", True)))
                self.delay_var.set(float(d.get("delay", 0.40)))
                self.format_only_var.set(bool(d.get("format_only", False)))
                self.dark_mode_var.set(bool(d.get("dark_mode", True)))
                self.write_meta_if_ok_var.set(bool(d.get("write_meta_if_ok", True)))
                self.write_nfo_var.set(bool(d.get("write_nfo", False)))

                custom = d.get("custom_noise_tokens", [])
                if isinstance(custom, list):
                    EXTRA_NOISE_TOKENS = {str(t).lower() for t in custom if str(t).strip()}
                    self.custom_noise_tokens = sorted(EXTRA_NOISE_TOKENS)
        except Exception:
            pass

    def on_save_options(self):
        try:
            with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
                json.dump(self._options_dict(), f, indent=2)
            messagebox.showinfo("Saved", f"Options saved to:\n{OPTIONS_FILE}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save options:\n{e}")

    # ----- UI plumbing -----
    def _lock_controls(self, running: bool):
        state = 'disabled' if running else 'normal'
        self.ent_folder.config(state=state)
        self.spin_delay.config(state=state)
        self.btn_scan.config(state='disabled' if running else 'normal')
        self.btn_apply.config(state='disabled')
        self.btn_save_backup.config(state='disabled' if running else ('normal' if self.plan_rows else 'disabled'))
        self.btn_restore_backup.config(state='disabled' if running else 'normal')
        self.btn_save_options.config(state='disabled' if running else 'normal')
        self.btn_check_deps.config(state='disabled' if running else 'normal')

    def _after_scan_controls(self):
        self.btn_apply.config(state='normal' if self.plan_rows else 'disabled')
        self.btn_save_backup.config(state='normal' if self.plan_rows else 'disabled')

    def _after_apply_controls(self):
        pass

    def on_toggle_dark(self):
        self._apply_theme_and_fix()

    def choose_show_blocking(self, guess: str, cands: List[dict]) -> Optional[dict]:
        result = {"selected": None}
        done = threading.Event()
        def _ask():
            dlg = ShowPickerDialog(self, guess, cands)
            self.wait_window(dlg)
            result["selected"] = dlg.selected
            done.set()
        self.after(0, _ask)
        done.wait()
        return result["selected"]

    def prompt_for_show_blocking(self) -> Optional[str]:
        result = {"text": None}
        done = threading.Event()
        def _ask():
            txt = simpledialog.askstring("Show name",
                                         "Enter the show title to use for TVMaze:",
                                         parent=self)
            t = (txt.strip() if txt else None)
            result["text"] = t
            if t:
                self.last_prompted_show = t
            done.set()
        self.after(0, _ask)
        done.wait()
        return result["text"]

    def on_browse(self):
        d = filedialog.askdirectory(initialdir=self.folder_var.get() or os.getcwd(), title="Select folder")
        if d:
            self.folder_var.set(d)

    def on_install_deps(self):
        dlg = OutputDialog(self, "Install / Check Dependencies")
        dlg.append("Checking current status…")
        mut, mkv, _ = detect_dep_status()
        dlg.append(f" - mutagen: {'present' if mut else 'missing'}")
        dlg.append(f" - mkvpropedit: {'present' if mkv else 'missing'}")

        want_mut = not mut
        want_mkv = not mkv

        if not want_mut and not want_mkv:
            dlg.append("\nAll dependencies already installed.")
            self._refresh_dep_status(); self.update_idletasks()
            return

        if want_mut:
            dlg.append("\nInstalling mutagen (pip)…")
            for cmd in install_mutagen_steps():
                dlg.append(f"$ {' '.join(cmd)}")
                rc, out = run_cmd_capture(cmd)
                dlg.append(out.strip())
                if rc != 0:
                    dlg.append(f"[ERROR] pip returned {rc}. Try running the above command manually.")
                    break
            self._refresh_dep_status(); self.update_idletasks()

        if want_mkv:
            dlg.append("\nInstalling MKVToolNix (mkvpropedit)…")
            steps: List[List[str]] = []
            if sys.platform.startswith("win"):
                steps = install_mkvtoolnix_steps_windows()
            elif sys.platform == "darwin":
                steps = install_mkvtoolnix_steps_macos()
                if not steps:
                    dlg.append("Homebrew not found. Install Homebrew or MKVToolNix manually.")
            else:
                steps = install_mkvtoolnix_steps_linux()
                if not steps:
                    dlg.append("No supported package manager (apt/dnf/pacman) detected. Install mkvtoolnix manually.")

            for cmd in steps:
                dlg.append(f"$ {' '.join(cmd)}")
                rc, out = run_cmd_capture(cmd)
                dlg.append(out.strip())
                if rc == 0:
                    continue
                if sys.platform.startswith("win") and is_permission_error(out):
                    if messagebox.askyesno("Elevation required", "Installer needs administrator rights. Run elevated now?"):
                        rc2, out2 = run_elevated_install_ps(" ".join(cmd))
                        dlg.append(f"[elevated rc={rc2}] {out2.strip() if out2 else ''}")
                    else:
                        dlg.append("[WARN] Skipped elevation at user request.")
                else:
                    dlg.append(f"[WARN] Command returned {rc}. Continuing…")

            self._refresh_dep_status(); self.update_idletasks()
            _, mkv_now, _ = detect_dep_status()
            dlg.append(f"mkvpropedit detection: {'FOUND' if mkv_now else 'MISSING'} (path: {_which_mkvpropedit() or 'n/a'})")

            if sys.platform.startswith("win") and not mkv_now:
                default_url = "https://mkvtoolnix.download/windows/releases/96.0/mkvtoolnix-64-bit-96.0-setup.exe"
                if messagebox.askyesno("Manual installer",
                                       "Package manager didn’t complete. Install MKVToolNix via direct installer URL (elevated)?"):
                    url = simpledialog.askstring(
                        "Installer URL",
                        "Paste the official MKVToolNix installer URL (.exe). It will be run elevated.\n"
                        "Leave the default to fetch a current 64-bit setup from the official site:",
                        parent=self,
                        initialvalue=default_url
                    )
                    if url:
                        url = url.strip()
                        dlg.append(f"\nDownloading installer from:\n{url}")
                        try:
                            import tempfile
                            tmpdir = tempfile.gettempdir()
                            fname = os.path.join(tmpdir, os.path.basename(urllib.parse.urlparse(url).path) or "mkvtoolnix_installer.exe")
                            urllib.request.urlretrieve(url, fname)
                            dlg.append(f"Saved to: {fname}")
                            rc3, out3 = run_elevated_install_ps(f'"{fname}"')
                            if out3: dlg.append(out3.strip())
                            dlg.append(f"[elevated] return code: {rc3}")
                        except Exception as e:
                            dlg.append(f"[ERROR] Manual install failed: {e}")

            self._refresh_dep_status(); self.update_idletasks()
            _, mkv_end, _ = detect_dep_status()
            dlg.append(f"Final mkvpropedit detection: {'FOUND' if mkv_end else 'MISSING'} (path: {_which_mkvpropedit() or 'n/a'})")

        _reload_mutagen_flag()
        self._refresh_dep_status()
        self.update_idletasks()
        mut2, mkv2, _ = detect_dep_status()
        dlg.append("\nRe-checking status…")
        dlg.append(f"Now: mutagen={'present' if mut2 else 'missing'}, mkvpropedit={'present' if mkv2 else 'missing'}")
        if not (mut2 and mkv2):
            dlg.append("\nSome dependencies are still missing. If you saw permission errors, re-run installs in an elevated shell.")
        dlg.append("\nDone.")

    def _refresh_dep_status(self):
        mut, mkv, _ = detect_dep_status()
        self.dep_status_var.set(f"mutagen={'OK' if mut else 'missing'} | mkvpropedit={'OK' if mkv else 'missing'}")

    # === Scan / Apply ===
    def on_scan(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Select a valid folder.")
            return

        self.plan_rows.clear()
        self.fail_rows.clear()
        for w in self.tree.get_children():
            self.tree.delete(w)
        self.lst_fail.delete(0, 'end')
        self.progress['value'] = 0
        self.status_var.set("Starting scan…")
        self.summary_var.set("")
        self._lock_controls(True)

        prompt_cb = self.prompt_for_show_blocking
        explicit = None
        ovr = explicit or (self.last_prompted_show)

        self.planner = RenamePlanner(
            folder=folder,
            recursive=bool(self.recursive_var.get()),
            rate_delay=float(self.delay_var.get()),
            format_only=bool(self.format_only_var.get()),
            choose_show_cb=self.choose_show_blocking,
            override_show=ovr,
            prompt_missing_cb=prompt_cb,
            write_metadata=True,
            write_meta_if_ok=bool(self.write_meta_if_ok_var.get()),
            write_nfo=bool(self.write_nfo_var.get())
        )
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def _scan_thread(self):
        def progress_cb(i, total):
            self.progress['maximum'] = max(1, total)
            self.progress['value'] = i
        def status_cb(msg):
            self.status_var.set(msg)
        try:
            self.planner.scan(progress_cb=progress_cb, status_cb=status_cb)
            self.plan_rows = []
            for item in self.planner.changes:
                if item[0] == "video":
                    _, o, n = item
                    self.plan_rows.append(("video", os.path.basename(o), os.path.basename(n)))
                elif item[0] == "subtitle":
                    _, o, n = item
                    self.plan_rows.append(("subtitle", os.path.basename(o), os.path.basename(n)))
                elif item[0] == "metaonly":
                    _, o, title = item
                    self.plan_rows.append(("metadata", os.path.basename(o), title))
            self.fail_rows = [(os.path.basename(p), reason) for (p,reason) in self.planner.failures]
        except Exception as e:
            self.status_var.set(f"Scan error: {e}")
        finally:
            self.after(0, self._scan_finished)

    def _scan_finished(self):
        self.tree.delete(*self.tree.get_children())
        for typ, old, new in self.plan_rows:
            arrow = "→" if typ != "metadata" else "✎"
            right = new if typ != "metadata" else f"write title = '{new}'"
            self.tree.insert('', 'end', values=(typ, old, arrow, right))
        self.lst_fail.delete(0, 'end')
        for path, reason in self.fail_rows:
            self.lst_fail.insert('end', f"{path} — {reason}")

        if self.planner:
            s = self.planner.stats
            planned_v = sum(1 for t,_,_ in self.plan_rows if t == "video")
            planned_s = sum(1 for t,_,_ in self.plan_rows if t == "subtitle")
            planned_m = sum(1 for t,_,_ in self.plan_rows if t == "metadata")
            self.summary_var.set(
                f"Videos: {s['videos_total']} | Parsed: {s['parsed_ok']} | "
                f"Unmatched: {s['parsed_fail']} | Show not found: {s['show_not_found']} | "
                f"Episode not found: {s['ep_not_found']} | Already correct: {s['already_correct']} | "
                f"Planned: {planned_v} video, {planned_s} subtitles, {planned_m} metadata writes."
            )

        self._lock_controls(False)
        self._after_scan_controls()
        self.status_var.set("Scan complete. Review plan and skipped items above.")

    def on_apply(self):
        if not self.planner or not self.planner.changes:
            messagebox.showinfo("Info", "Nothing to rename.")
            return

        self.result_rows.clear()
        self.progress['value'] = 0
        self.status_var.set("Starting rename…")
        self._lock_controls(True)
        threading.Thread(target=self._apply_thread, daemon=True).start()

    def _apply_thread(self):
        def progress_cb(i, total):
            self.progress['maximum'] = max(1, total)
            self.progress['value'] = i
        def status_cb(msg):
            self.status_var.set(msg)
        try:
            results = self.planner.apply(
                write_meta_if_ok=bool(self.write_meta_if_ok_var.get()),
                write_nfo=bool(self.write_nfo_var.get()),
                progress_cb=progress_cb, status_cb=status_cb
            )
            self.result_rows = []
            for typ, o, n, r in results:
                o_disp = os.path.basename(o)
                n_disp = os.path.basename(n) if (typ != "meta") else n
                self.result_rows.append((typ, o_disp, n_disp, r))
        except Exception as e:
            self.status_var.set(f"Rename error: {e}")
        finally:
            self.after(0, self._apply_finished)

    def _apply_finished(self):
        ok = sum(1 for *_ , r in self.result_rows if r == "OK")
        errs = [row for row in self.result_rows if row[3] != "OK"]
        self._lock_controls(False)
        self._after_apply_controls()
        msg = f"Rename complete: {ok} OK, {len(errs)} errors."
        self.status_var.set(msg)
        if errs:
            messagebox.showwarning("Completed with errors", msg)
        else:
            messagebox.showinfo("Done", msg)

    # === Backup / Restore ===
    def on_save_backup(self):
        if not self.planner or not self.planner.changes:
            messagebox.showinfo("Info", "No planned renames to back up.\nRun a scan first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save Name Backup TSV",
            defaultextension=".tsv",
            filetypes=[("TSV files","*.tsv"), ("All files","*.*")]
        )
        if not path:
            return
        try:
            rows: List[Tuple[str, str, str]] = []
            for typ, old, new in self.planner.changes:
                if typ in ("video", "subtitle"):
                    rows.append((typ, os.path.abspath(old), os.path.abspath(new)))
            if not rows:
                messagebox.showinfo("Info", "No video or subtitle renames to back up.")
                return
            save_tsv(path, rows, ["TYPE","OLD_PATH","NEW_PATH"])
            messagebox.showinfo("Saved", f"Backup saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save backup:\n{e}")

    def on_restore_backup(self):
        path = filedialog.askopenfilename(
            title="Open Name Backup TSV",
            filetypes=[("TSV files","*.tsv"), ("All files","*.*")]
        )
        if not path:
            return
        try:
            mapping = load_backup_tsv(path)
            if not mapping:
                messagebox.showwarning("Empty backup", "No entries found in backup.")
                return
            restored = 0
            skipped = 0
            for typ, old_path, new_path in mapping:
                if typ not in ("video", "subtitle"):
                    skipped += 1
                    continue
                old_path = os.path.abspath(old_path)
                new_path = os.path.abspath(new_path)
                if not os.path.exists(new_path):
                    skipped += 1
                    continue
                if os.path.exists(old_path):
                    skipped += 1
                    continue
                try:
                    os.rename(new_path, old_path)
                    restored += 1
                except Exception:
                    skipped += 1
            messagebox.showinfo(
                "Restore complete",
                f"Restore complete.\nRestored: {restored}\nSkipped: {skipped}"
            )
            self.status_var.set(f"Restore complete: {restored} restored, {skipped} skipped.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not restore from backup:\n{e}")

    # === Custom noise tokens ===
    def on_edit_noise_tokens(self):
        global EXTRA_NOISE_TOKENS
        existing = " ".join(sorted(EXTRA_NOISE_TOKENS))
        txt = simpledialog.askstring(
            "Custom noise tokens",
            "Enter extra noise tokens separated by spaces or commas.\n"
            "These are stripped from show names just like 1080p, WEB-DL, etc.",
            initialvalue=existing,
            parent=self
        )
        if txt is None:
            return
        tokens = set(t.strip().lower() for t in re.split(r"[,\s]+", txt) if t.strip())
        EXTRA_NOISE_TOKENS = tokens
        self.custom_noise_tokens = sorted(EXTRA_NOISE_TOKENS)
        messagebox.showinfo(
            "Custom tokens updated",
            f"{len(EXTRA_NOISE_TOKENS)} custom noise tokens set.\n"
            "Click Save Options if you want to persist them."
        )

if __name__ == "__main__":
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    app = App()
    app.mainloop()
