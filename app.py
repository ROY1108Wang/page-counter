# -*- coding: utf-8 -*-
"""
轻量文件页数统计助手 V1.0
支持: pdf, doc/docx/docm, xls/xlsx/xlsm, ppt/pptx/pptm
原理: 直接读取文档内部元数据(OOXML docProps / OLE 摘要信息 / PDF 页对象),
      不启动 Office, 拖入即统计, 秒出单文件页数 + 总页数, 可导出 CSV/Excel。
可选: 元数据缺失时调用本机 Office COM 读取(较慢, 需安装 MS Office 或 WPS)。
"""
import os
import re
import csv
import datetime
import traceback
import threading
import queue
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------- 可选依赖 ----------
try:
    import olefile  # 读取 doc/xls/ppt 的 OLE 摘要信息
except ImportError:
    olefile = None

try:
    from pypdf import PdfReader  # PDF 页数
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # 拖拽支持
    _ROOT_CLS = TkinterDnD.Tk
except ImportError:
    DND_FILES = None
    _ROOT_CLS = tk.Tk

APP_EXT = {".pdf", ".doc", ".docx", ".docm", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx", ".pptm"}
DOC_EXT = {".doc", ".docx", ".docm"}
XLS_EXT = {".xls", ".xlsx", ".xlsm"}
PPT_EXT = {".ppt", ".pptx", ".pptm"}
ZIP_EXT = {".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm"}

DOCX_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"


def is_doc_file(name):
    """是否为受支持且可统计的文档(排除 Word/PPT 打开时产生的 ~$ 锁定临时文件)"""
    b = os.path.basename(name)
    if b.startswith("~$"):
        return False
    return os.path.splitext(b)[1].lower() in APP_EXT

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.log")


def log_error(msg):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write("[%s] %s\n" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg))
    except Exception:
        pass


# ================= 核心统计逻辑 =================
def count_pdf(path):
    if PdfReader is not None:
        reader = PdfReader(path)
        n = len(reader.pages)
        if n:
            return n
    # 无 pypdf 时的正则兜底
    with open(path, "rb") as f:
        data = f.read()
    if data.startswith(b"%PDF"):
        n = len(re.findall(rb"/Type\s*/Page(?![s])", data))
        if n:
            return n
    return None


def _zip_app_prop(path, tag):
    """读取 OOXML docProps/app.xml 里的某个属性, 如 Pages / Slides"""
    with zipfile.ZipFile(path) as z:
        with z.open("docProps/app.xml") as f:
            tree = ET.parse(f)
    node = tree.getroot().find(DOCX_NS + tag)
    if node is not None and node.text and node.text.isdigit():
        return int(node.text)
    return None


def count_docx(path):
    return _zip_app_prop(path, "Pages")


def count_pptx(path):
    return _zip_app_prop(path, "Slides")


def count_xlsx(path):
    n = _zip_app_prop(path, "Pages")  # Excel 一般不写 Pages
    if n:
        return n
    # 兜底: 用工作表数量近似"页数"
    with zipfile.ZipFile(path) as z:
        with z.open("xl/workbook.xml") as f:
            tree = ET.parse(f)
    return sum(1 for _ in tree.getroot().iter(
        "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet"))


def count_ole(path):
    """doc/xls/ppt: 读取 OLE 摘要信息中的页数(属性 13, PageCount)"""
    if olefile is None:
        return None
    try:
        with olefile.OleFileIO(path) as ole:
            meta = ole.get_metadata()
            pc = getattr(meta, "page_count", None)
            if pc:
                return int(pc)
    except Exception:
        pass
    return None


def count_pages(path):
    """返回页数或 None(无数据)"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return count_pdf(path)
    if ext in DOC_EXT and ext != ".doc":
        return count_docx(path)
    if ext in PPT_EXT and ext != ".ppt":
        return count_pptx(path)
    if ext in XLS_EXT and ext != ".xls":
        return count_xlsx(path)
    if ext in APP_EXT:
        return count_ole(path)
    return None


# ================= Office COM 兜底(可选) =================
class ComCounter:
    """延迟启动 Word/Excel/PowerPoint COM, 用于元数据缺失的文件"""

    def __init__(self):
        self.apps = {}

    def _get(self, key, progids):
        if key in self.apps:
            return self.apps[key]
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        last_err = None
        for pid in progids:
            try:
                app = win32com.client.DispatchEx(pid)
                app.Visible = False
                app.DisplayAlerts = 0
                self.apps[key] = app
                return app
            except Exception as e:
                last_err = e
        raise last_err

    def count(self, path):
        ext = os.path.splitext(path)[1].lower()
        path = os.path.abspath(path)
        if ext in DOC_EXT:
            word = self._get("word", ["Word.Application", "KWPS.Application", "WPS.Application"])
            doc = self._open_any(word.Documents, path,
                                 {"FileName": path, "ReadOnly": True, "AddToRecentFiles": False},
                                 {"FileName": path})
            try:
                return doc.ComputeStatistics(2)  # wdStatisticPages
            finally:
                doc.Close(False)
        if ext in PPT_EXT:
            ppt = self._get("ppt", ["PowerPoint.Application", "KWPP.Application", "WPP.Application"])
            pres = self._open_any(ppt.Presentations, path,
                                  {"FileName": path, "ReadOnly": True, "WithWindow": False},
                                  {"FileName": path})
            try:
                return pres.Slides.Count
            finally:
                pres.Close()
        if ext in XLS_EXT:
            excel = self._get("excel", ["Excel.Application", "KET.Application", "ET.Application"])
            wb = self._open_any(excel.Workbooks, path,
                                {"FileName": path, "ReadOnly": True},
                                {"FileName": path})
            try:
                total = 0
                for ws in wb.Worksheets:
                    ws.Activate()
                    try:
                        total += int(excel.ExecuteExcel4Macro("GET.DOCUMENT(50)"))
                    except Exception:
                        total += 1
                return total
            finally:
                wb.Close(False)
        return None

    @staticmethod
    def _open_any(collection, path, kw1, kw2):
        """WPS 对 COM 参数支持不完整, 依次尝试多种打开方式"""
        try:
            return collection.Open(**kw1)
        except Exception:
            try:
                return collection.Open(**kw2)
            except Exception:
                return collection.Open(path)

    def quit(self):
        for app in self.apps.values():
            try:
                app.Quit()
            except Exception:
                pass
        self.apps.clear()
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


# ================= 主界面 =================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("轻量文件页数统计助手 V1.0 <拖入 pdf/doc/docx/xls/xlsx/ppt/pptx 等文件或文件夹>")
        self.root.geometry("1080x640")
        self.root.minsize(760, 480)

        self.rows = {}          # path -> dict(row)
        self.order = []         # 保持插入顺序
        self.ui_queue = queue.Queue()
        self.worker = None
        self.com = None

        self._build_ui()

        if DND_FILES is not None:
            root.drop_target_register(DND_FILES)
            root.dnd_bind("<<Drop>>", self.on_drop)
        else:
            self.set_status("提示: 未安装 tkinterdnd2, 拖拽不可用, 请使用\"添加文件/添加文件夹\"按钮")

        self.root.after(80, self._poll_queue)

    # ---------- 界面 ----------
    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            pass
        style.configure("Treeview", rowheight=24, font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))

        header = tk.Frame(self.root, bg="#8E1616", height=52)
        header.pack(fill="x")
        tk.Label(header, text="轻量文件页数统计助手", font=("Microsoft YaHei UI", 14, "bold"),
                 fg="white", bg="#8E1616").pack(side="left", padx=14)
        tk.Label(header, text="纯元数据解析 · 不启动 Office · 秒出结果 · 可导出表格",
                 font=("Microsoft YaHei UI", 9), fg="#FFD9D9", bg="#8E1616").pack(side="left", padx=8)

        # 汇总栏
        stats = tk.Frame(self.root)
        stats.pack(fill="x", padx=8, pady=4)
        self.lbl_files = tk.Label(stats, text="文件总数: 0", font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_done = tk.Label(stats, text="已处理: 0", font=("Microsoft YaHei UI", 10, "bold"))
        self.lbl_pages = tk.Label(stats, text="文件总页数: 0", font=("Microsoft YaHei UI", 10, "bold"), fg="#B00000")
        self.lbl_files.pack(side="left", padx=(4, 18))
        self.lbl_done.pack(side="left", padx=18)
        self.lbl_pages.pack(side="left", padx=18)
        self.status_var = tk.StringVar(value="就绪")
        tk.Label(stats, textvariable=self.status_var, fg="#666").pack(side="right", padx=6)

        # 计价栏: 两行四列, 标签右对齐、输入框等宽对齐
        price = tk.LabelFrame(self.root, text="计价", padx=10, pady=6)
        price.pack(fill="x", padx=8, pady=(0, 4))
        self.var_extra = tk.StringVar(value="0")         # 附加费(元)
        self.var_unit = tk.StringVar(value="0.05")       # 单价(元)
        self.var_total_pages = tk.StringVar(value="0")   # 总页
        self.var_duplex = tk.StringVar(value="0")        # 双面总页(张)
        self.var_money_single = tk.StringVar(value="0.00")  # 单面总价
        self.var_money_duplex = tk.StringVar(value="0.00")  # 双面总价
        self.var_threshold = tk.StringVar(value="")      # 满减, 如 10-5
        self.var_ship = tk.StringVar(value="")           # 满多少元才发货
        self.var_mode = tk.StringVar(value="双面")        # 计价方式: 双面/单面
        self.var_more = tk.StringVar(value="—")          # 还能加打

        def pfield(row, col, label, var, editable, width=10):
            tk.Label(price, text=label, anchor="e").grid(row=row, column=col * 2, padx=(14, 4), sticky="e")
            e = tk.Entry(price, textvariable=var, width=width, justify="center",
                         font=("Microsoft YaHei UI", 10, "bold"))
            if not editable:
                e.configure(state="readonly", readonlybackground="#F5F5F5",
                            disabledforeground="#B00000")
            e.grid(row=row, column=col * 2 + 1, pady=3, sticky="w")
            return e

        pfield(0, 0, "附加费(元):", self.var_extra, True)
        pfield(0, 1, "总页:", self.var_total_pages, False)
        pfield(0, 2, "双面总页(张):", self.var_duplex, False)
        pfield(0, 3, "单价(元):", self.var_unit, True)
        pfield(1, 0, "单面总价(元):", self.var_money_single, False)
        pfield(1, 1, "双面总价(元):", self.var_money_duplex, False)
        pfield(1, 2, "满减(如10-5):", self.var_threshold, True, 8)
        pfield(1, 3, "满多少发货(元):", self.var_ship, True, 8)
        # 第三行: 计价方式 + 还能加打(Label 完整显示不截断)
        tk.Label(price, text="计价方式:", anchor="e").grid(row=2, column=0, padx=(14, 4), sticky="e")
        self.cmb_mode = ttk.Combobox(price, textvariable=self.var_mode, values=("双面", "单面"),
                                     state="readonly", width=8, font=("Microsoft YaHei UI", 10))
        self.cmb_mode.grid(row=2, column=1, sticky="w", pady=3)
        tk.Label(price, text="还能加打:", anchor="e").grid(row=2, column=2, padx=(14, 4), sticky="e")
        tk.Label(price, textvariable=self.var_more, fg="#B00000", anchor="w",
                 font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=2, column=3, columnspan=5, sticky="w")
        for c in range(8):
            price.columnconfigure(c, weight=1 if c % 2 == 0 else 0)
        self.var_extra.trace_add("write", self._update_price)
        self.var_unit.trace_add("write", self._update_price)
        self.var_threshold.trace_add("write", self._update_price)
        self.var_ship.trace_add("write", self._update_price)
        self.var_mode.trace_add("write", self._update_price)

        body = tk.Frame(self.root)
        body.pack(fill="both", expand=True, padx=8)

        # 左侧列表
        left = tk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        cols = ("path", "status", "pages")
        self.tree = ttk.Treeview(left, columns=cols, show="headings")
        self.tree.heading("path", text="文件路径")
        self.tree.heading("status", text="状态")
        self.tree.heading("pages", text="页数")
        self.tree.column("path", width=560, anchor="w")
        self.tree.column("status", width=90, anchor="center", stretch=False)
        self.tree.column("pages", width=70, anchor="center", stretch=False)
        self.tree.tag_configure("pending", foreground="#888888")
        self.tree.tag_configure("done", foreground="#000000")
        self.tree.tag_configure("fail", foreground="#B00000")
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # 右侧操作区
        right = tk.LabelFrame(body, text="操作", padx=10, pady=8)
        right.pack(side="left", fill="y", padx=(8, 0))

        self.var_sub = tk.BooleanVar(value=True)
        self.var_com = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="包含子文件夹", variable=self.var_sub).pack(fill="x", pady=2)
        ttk.Checkbutton(right, text="缺页数时自动用 Office/WPS 补算(稍慢)",
                        variable=self.var_com).pack(fill="x", pady=(2, 8))

        def btn(text, cmd, top=True):
            b = ttk.Button(right, text=text, command=cmd, width=22)
            b.pack(fill="x", pady=3)
            return b

        btn("添加文件...", self.add_files)
        btn("添加文件夹...", self.add_folder)
        btn("从剪贴板导入", self.import_clipboard)
        ttk.Separator(right).pack(fill="x", pady=6)
        btn("导出 Excel", lambda: self.export(True))
        btn("导出 CSV", lambda: self.export(False))
        ttk.Separator(right).pack(fill="x", pady=6)
        btn("移除选中", self.remove_selected)
        btn("清空列表", self.clear_all)

    # ---------- 添加文件 ----------
    def on_drop(self, event):
        try:
            log_error("[拖拽] 收到: " + str(event.data)[:300])
            paths = self.root.tk.splitlist(event.data)
            self.add_paths([p.strip("{}") for p in paths])
        except Exception:
            log_error("拖拽处理异常:\n" + traceback.format_exc())
            messagebox.showerror(
                "拖拽出错",
                "拖拽处理异常，请改用\"添加文件/添加文件夹\"按钮，\n"
                "或在资源管理器复制文件后点\"从剪贴板导入\"。\n"
                "详情见 app.log")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=[("支持的文档", "*.pdf *.doc *.docx *.docm *.xls *.xlsx *.xlsm *.ppt *.pptx *.pptm"),
                       ("所有文件", "*.*")])
        self.add_paths(paths)

    def add_folder(self):
        d = filedialog.askdirectory(title="选择文件夹")
        if d:
            self.add_paths([d])

    def _clipboard_files(self):
        """读取资源管理器复制(Ctrl+C)的文件/文件夹路径(CF_HDROP), 失败则回退到文本行"""
        try:
            import ctypes
            user32 = ctypes.windll.user32
            shell32 = ctypes.windll.shell32
            CF_HDROP = 15
            if not user32.IsClipboardFormatAvailable(CF_HDROP) or not user32.OpenClipboard(0):
                raise OSError
            try:
                h = user32.GetClipboardData(CF_HDROP)
                if not h:
                    raise OSError
                count = shell32.DragQueryFileW(h, 0xFFFFFFFF, None, 0)
                paths, buf = [], ctypes.create_unicode_buffer(520)
                for i in range(count):
                    shell32.DragQueryFileW(h, i, buf, 520)
                    paths.append(buf.value)
                return paths
            finally:
                user32.CloseClipboard()
        except Exception:
            try:
                text = self.root.clipboard_get()
            except Exception:
                return []
            return [p.strip().strip('"') for p in text.splitlines() if p.strip()]

    def import_clipboard(self):
        paths = self._clipboard_files()
        log_error("[剪贴板] 收到 %d 项" % len(paths))
        if not paths:
            messagebox.showinfo("提示", "剪贴板里没有文件。\n请先在资源管理器选中文件/文件夹，按 Ctrl+C 复制后再点本按钮。")
            return
        self.add_paths(paths)

    def add_paths(self, paths):
        paths = [p for p in paths if p]
        log_error("[导入] 收到 %d 项: %s" % (
            len(paths), " | ".join(str(p)[:150] for p in paths[:8])))
        files = []
        for p in paths:
            p = os.path.normpath(p)
            if os.path.isdir(p):
                if self.var_sub.get():
                    for root, _dirs, names in os.walk(p):
                        files += [os.path.join(root, n) for n in names if is_doc_file(n)]
                else:
                    try:
                        files += [os.path.join(p, n) for n in os.listdir(p) if is_doc_file(n)]
                    except OSError:
                        pass
            elif os.path.isfile(p) and is_doc_file(p):
                files.append(p)
        added = 0
        for f in files:
            f = os.path.normpath(f)
            if f in self.rows:
                continue
            rec = {"path": f, "name": os.path.basename(f), "status": "待处理", "pages": None}
            self.rows[f] = rec
            self.order.append(f)
            self.tree.insert("", "end", iid=f, values=(f, "待处理", "—"), tags=("pending",))
            added += 1
        self.refresh_stats()
        if added:
            self.set_status("已添加 %d 个文件" % added)
            self.start_worker()
        elif files or paths:
            self.set_status("未发现可识别的文档（支持 pdf/doc/docx/xls/xlsx/ppt/pptx 等）")

    # ---------- 后台统计线程 ----------
    def start_worker(self):
        if self.worker and self.worker.is_alive():
            self._retry_when_done = True   # 统计中新文件已入列, 完成后自动接续
            return
        pending = [p for p in self.order if self.rows[p]["status"] == "待处理"]
        if not pending:
            return
        use_com = self.var_com.get()
        self.worker = threading.Thread(target=self._work, args=(pending, use_com), daemon=True)
        self.worker.start()

    def _work(self, pending, use_com):
        com = None
        try:
            if use_com:
                com = ComCounter()
            for path in pending:
                self.ui_queue.put(("status", path, "处理中"))
                pages, status = None, ""
                try:
                    pages = count_pages(path)
                    if pages is None and com is not None:
                        pages = com.count(path)
                    status = "完成" if pages is not None else "无页数数据"
                except PermissionError:
                    status = "无法读取"
                except zipfile.BadZipFile:
                    status = "无法读取"
                except Exception as e:
                    status = "失败"
                    log_error("统计失败 %s\n%s" % (path, traceback.format_exc()))
                    self.ui_queue.put(("error", path, str(e)))
                self.ui_queue.put(("result", path, status, pages))
        finally:
            if com is not None:
                com.quit()
            self.ui_queue.put(("done",))

    def _poll_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                kind = msg[0]
                if kind == "status":
                    _, path, status = msg
                    if path in self.rows:
                        self.rows[path]["status"] = status
                        self.tree.set(path, "status", status)
                elif kind == "result":
                    _, path, status, pages = msg
                    if path in self.rows:
                        rec = self.rows[path]
                        rec["status"] = status
                        rec["pages"] = pages
                        tag = "done" if pages is not None else "fail"
                        self.tree.item(path, values=(path, status, "—" if pages is None else pages), tags=(tag,))
                elif kind == "error":
                    self.set_status("错误: %s" % msg[2])
                elif kind == "done":
                    self.worker = None
                    self.refresh_stats()
                    self.set_status("统计完成")
                    if getattr(self, "_retry_when_done", False):
                        self._retry_when_done = False
                        self.start_worker()
        except queue.Empty:
            pass
        self.refresh_stats()
        self.root.after(80, self._poll_queue)

    # ---------- 其它操作 ----------
    def refresh_stats(self):
        total = len(self.rows)
        done = sum(1 for r in self.rows.values() if r["status"] not in ("待处理", "处理中"))
        pages = sum(r["pages"] or 0 for r in self.rows.values())
        self.lbl_files.config(text="文件总数: %d" % total)
        self.lbl_done.config(text="已处理: %d" % done)
        self.lbl_pages.config(text="文件总页数: %d" % pages)
        self._update_price()

    @staticmethod
    def _parse_threshold(text):
        """解析满减: '10-5' → (10, 5) 表示满10减5; '5' → (5, 0) 表示满5(无减免)"""
        s = (text or "").strip()
        if not s:
            return None
        m = re.match(r"^(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)$", s)
        if m:
            return float(m.group(1)), float(m.group(2))
        try:
            return float(s), 0.0
        except ValueError:
            return None

    def _update_price(self, *_):
        """计价: 双面总页 = ceil(总页/2) 张;
           单面总价 = 总页 x 单价 + 附加费; 双面总价 = 双面总页 x 单价 + 附加费;
           满减: 计算实付还差多少元、还能加打多少页"""
        try:
            fee = float(self.var_extra.get() or 0)
        except ValueError:
            fee = 0.0
        try:
            unit = float(self.var_unit.get() or 0)
        except ValueError:
            unit = 0.0
        pages = sum(r["pages"] or 0 for r in self.rows.values())
        duplex = (pages + 1) // 2
        self.var_total_pages.set(str(pages))
        self.var_duplex.set(str(duplex))
        single_total = pages * unit + fee
        duplex_total = duplex * unit + fee
        self.var_money_single.set("%.2f" % single_total)
        self.var_money_duplex.set("%.2f" % duplex_total)

        # 满减 + 满多少发货: 发货看"满减后实付"; 按所选方式(默认双面)换算还能加打多少(含面数)
        rule = self._parse_threshold(self.var_threshold.get())
        try:
            ship = float(self.var_ship.get() or 0)
        except ValueError:
            ship = 0.0
        if rule is None and ship <= 0:
            self.var_more.set("—")
            return
        is_duplex = self.var_mode.get() != "单面"
        base = duplex if is_duplex else pages
        count_word = "张" if is_duplex else "页"
        faces_per_unit = 2 if is_duplex else 1   # 1张双面纸 = 2面
        total_now = base * unit + fee
        if ship <= 0:
            paid = total_now - rule[1] if (rule and total_now >= rule[0]) else total_now
            self.var_more.set("实付 %.2f 元 (未填发货门槛)" % paid)
            return
        # 实付(若已享受满减)
        paid = total_now - rule[1] if (rule and total_now >= rule[0]) else total_now
        if paid >= ship:
            self.var_more.set("已满足发货条件 (实付 %.2f 元)" % paid)
            return
        if unit <= 0:
            self.var_more.set("还差 %.2f 元 (单价为0, 无法换算)" % (ship - paid))
            return
        if rule and total_now < rule[0]:
            # 还没到满减门槛: 目标总价 = max(满额, 发货门槛+减额), 保证用上券且实付达标
            target_total = max(rule[0], ship + rule[1])
            need = target_total - total_now
        else:
            need = ship - paid
        n = int(-(-need // unit))          # 向上取整: 还能加打的张/页数
        faces = n * faces_per_unit         # 折合面数(1张双面=2面)
        self.var_more.set("还能打印 %d %s\n折合 %d 面" % (n, count_word, faces))

    def set_status(self, text):
        self.status_var.set(text)

    def remove_selected(self):
        for iid in self.tree.selection():
            self.rows.pop(iid, None)
            if iid in self.order:
                self.order.remove(iid)
            self.tree.delete(iid)
        self.refresh_stats()

    def clear_all(self):
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        self.order.clear()
        self.refresh_stats()
        self.set_status("已清空")

    def export(self, as_excel):
        if not self.rows:
            messagebox.showinfo("提示", "列表为空, 请先添加文件")
            return
        if as_excel:
            try:
                import openpyxl  # noqa
            except ImportError:
                messagebox.showwarning("提示", "未安装 openpyxl, 将导出为 CSV。\n可运行: pip install openpyxl")
                as_excel = False
        ext = ".xlsx" if as_excel else ".csv"
        ft = [("Excel 工作簿", "*.xlsx")] if as_excel else [("CSV 文件", "*.csv")]
        out = filedialog.asksaveasfilename(title="导出统计结果", defaultextension=ext,
                                           filetypes=ft, initialfile="页数统计" + ext)
        if not out:
            return
        try:
            if as_excel:
                self._export_xlsx(out)
            else:
                self._export_csv(out)
            self.set_status("已导出: " + out)
            if messagebox.askyesno("完成", "已导出到:\n%s\n\n是否立即打开?" % out):
                os.startfile(out)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _data_rows(self):
        data = []
        for p in self.order:
            r = self.rows[p]
            data.append([r["path"], r["name"],
                         os.path.splitext(r["name"])[1].lower().lstrip(".").upper(),
                         r["status"], r["pages"] if r["pages"] is not None else ""])
        return data

    def _price_summary_rows(self):
        try:
            fee = float(self.var_extra.get() or 0)
        except ValueError:
            fee = 0.0
        try:
            unit = float(self.var_unit.get() or 0)
        except ValueError:
            unit = 0.0
        pages = sum(r["pages"] or 0 for r in self.rows.values())
        duplex = (pages + 1) // 2
        rows = [["总页", "", "", "", pages],
                ["双面总页(张)", "", "", "", duplex],
                ["附加费(元)", "", "", "", "%.2f" % fee],
                ["单价(元)", "", "", "", unit],
                ["单面总价(元)", "", "", "", "%.2f" % (pages * unit + fee)],
                ["双面总价(元)", "", "", "", "%.2f" % (duplex * unit + fee)]]
        rule = self._parse_threshold(self.var_threshold.get())
        if rule is not None:
            rows.append(["满减", "", "", "", "满%.2f减%.2f" % rule])
        try:
            ship = float(self.var_ship.get() or 0)
        except ValueError:
            ship = 0.0
        if ship > 0:
            rows.append(["满多少发货(元)", "", "", "", "%.2f" % ship])
        rows.append(["计价方式", "", "", "", self.var_mode.get()])
        rows.append(["还能加打", "", "", "", self.var_more.get()])
        return rows

    def _export_csv(self, out):
        with open(out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["文件路径", "文件名", "类型", "状态", "页数"])
            w.writerows(self._data_rows())
            w.writerow(["总计", "", "", "", sum(r["pages"] or 0 for r in self.rows.values())])
            w.writerows(self._price_summary_rows())

    def _export_xlsx(self, out):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = Workbook()
        ws = wb.active
        ws.title = "页数统计"
        ws.append(["文件路径", "文件名", "类型", "状态", "页数"])
        head_fill = PatternFill("solid", fgColor="8E1616")
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
            c.alignment = Alignment(horizontal="center")
        for row in self._data_rows():
            ws.append(row)
        ws.append(["总计", "", "", "", sum(r["pages"] or 0 for r in self.rows.values())])
        ws.cell(row=ws.max_row, column=5).font = Font(bold=True)
        for row in self._price_summary_rows():
            ws.append(row)
            ws.cell(row=ws.max_row, column=5).font = Font(bold=True)
        ws.column_dimensions["A"].width = 70
        ws.column_dimensions["B"].width = 32
        for col in ("C", "D", "E"):
            ws.column_dimensions[col].width = 12
        ws.freeze_panes = "A2"
        wb.save(out)


def main():
    try:
        import ctypes
        if ctypes.windll.shell32.IsUserAnAdmin():
            log_error("警告: 程序以管理员权限运行, Windows 会拦截资源管理器的拖拽, 请以普通权限启动")
    except Exception:
        pass
    root = _ROOT_CLS()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
