import os
import cv2
import webbrowser
import math
import re
import threading
import subprocess
import winreg
from datetime import datetime
from tkinter import messagebox, filedialog
from functools import partial
import customtkinter as ctk

# ---------- Windows 标题读取（原版稳定实现） ----------
try:
    import win32com.client
except ImportError:
    win32com = None

def clean_string(raw_str):
    return str(raw_str).replace('\n', '').replace('\t', '').strip()

def get_windows_title(file_path):
    """从 Windows 文件属性中读取“标题”（固定索引 21）"""
    if os.name != "nt" or win32com is None:
        return ""
    title_content = ""
    shell = None
    folder = None
    file_item = None
    try:
        norm_path = os.path.normpath(file_path)
        folder_path = os.path.dirname(norm_path)
        file_name = os.path.basename(norm_path)
        shell = win32com.client.DispatchEx("Shell.Application")
        folder = shell.NameSpace(folder_path)
        if folder is None:
            return ""
        file_item = folder.ParseName(file_name)
        if file_item is not None:
            try:
                title_content = clean_string(folder.GetDetailsOf(file_item, 21))
            except:
                pass
    except Exception:
        return ""
    finally:
        if file_item:
            try:
                del file_item
            except:
                pass
        if folder:
            try:
                del folder
            except:
                pass
        if shell:
            try:
                del shell
            except:
                pass
    return title_content

def extract_url(text):
    """从文本中提取 URL（支持 www. 和 .com 等）"""
    if not text:
        return None
    text = text.replace('\n', ' ').replace('\t', ' ').strip()
    m = re.search(r'https?://[a-zA-Z0-9_\-./&?=%#]+', text)
    if m:
        return m.group(0)
    m = re.search(r'(www\.[a-zA-Z0-9\-]+(?:\.[a-zA-Z0-9\-]+)*\.[a-zA-Z]{2,})', text)
    if m:
        return "https://" + m.group(0)
    m = re.search(r'([a-zA-Z0-9\-]+\.[a-zA-Z]{2,})', text)
    if m:
        return "https://" + m.group(0)
    if "http" in text:
        start = text.find("http")
        return text[start:].split()[0]
    return None

# ==================== 自然排序（数字从小到大） ====================
def natural_key(text):
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', text)]

# ==================== 配置 ====================
class Config:
    WINDOW_WIDTH = 1100
    WINDOW_HEIGHT = 780
    MIN_WIDTH = 900
    MIN_HEIGHT = 650
    VIDEO_EXT = (".mp4", ".mov", ".mkv", ".avi", ".flv", ".webm", ".wmv", ".mpg", ".mpeg")
    PAGE_SIZE = 7
    SEARCH_URL_TPL = "https://www.baidu.com/s?wd={filename}"

    BG_LIGHT = "#f2f2f7"
    BG_DARK = "#1c1c1e"
    CARD_LIGHT = "#ffffff"
    CARD_DARK = "#2c2c2e"
    SEPARATOR_LIGHT = "#e5e5ea"
    SEPARATOR_DARK = "#3a3a3c"
    PRIMARY_BLUE = "#007aff"
    PRIMARY_BLUE_HOVER = "#0055cc"
    TEXT_PRIMARY_LIGHT = "#1c1c1e"
    TEXT_PRIMARY_DARK = "#ffffff"
    TEXT_SECONDARY_LIGHT = "#3a3a3c"
    TEXT_SECONDARY_DARK = "#8e8e93"
    LINK_COLOR = "#007aff"
    VERSION_GRAY = "#8e8e93"

# ==================== 数据模型 ====================
class VideoItem:
    __slots__ = ("name", "path", "resolution", "duration_str", "size_str", "mtime", "title", "jump_url")
    def __init__(self, name, path, resolution, duration_str, size_str, mtime, title, jump_url):
        self.name = name
        self.path = path
        self.resolution = resolution
        self.duration_str = duration_str
        self.size_str = size_str
        self.mtime = mtime
        self.title = title
        self.jump_url = jump_url

class VideoParser:
    @staticmethod
    def get_windows_title(file_path):
        return get_windows_title(file_path)

    @staticmethod
    def extract_url(text):
        return extract_url(text)

    @classmethod
    def parse_video(cls, file_path):
        try:
            cap = cv2.VideoCapture(file_path)
            width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            fps = cap.get(cv2.CAP_PROP_FPS)
            total_frames = cap.get(cv2.CAP_PROP_FRAME_COUNT)
            duration = total_frames / fps if fps > 0 else 0
            cap.release()
        except:
            width = height = 0
            duration = 0

        h = int(duration // 3600)
        m = int((duration % 3600) // 60)
        s = int(duration % 60)
        duration_str = f"{h:02d}:{m:02d}:{s:02d}"

        size_bytes = os.path.getsize(file_path)
        units = ["B", "KB", "MB", "GB"]
        idx = int(math.log(size_bytes, 1024)) if size_bytes > 0 else 0
        size_str = f"{size_bytes / (1024 ** idx):.2f} {units[idx]}"

        mtime = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M")
        title = cls.get_windows_title(file_path)
        jump_url = cls.extract_url(title)
        resolution = f"{width}×{height}" if width else "未知分辨率"

        return VideoItem(
            name=os.path.basename(file_path),
            path=file_path,
            resolution=resolution,
            duration_str=duration_str,
            size_str=size_str,
            mtime=mtime,
            title=title,
            jump_url=jump_url
        )

class VideoScanner:
    def __init__(self, scan_subfolder=True):
        self.scan_subfolder = scan_subfolder

    def scan(self, root_dir):
        result = []
        if self.scan_subfolder:
            for dirpath, _, files in os.walk(root_dir):
                for f in files:
                    if f.lower().endswith(Config.VIDEO_EXT):
                        result.append(os.path.join(dirpath, f))
        else:
            for f in os.listdir(root_dir):
                fp = os.path.join(root_dir, f)
                if os.path.isfile(fp) and f.lower().endswith(Config.VIDEO_EXT):
                    result.append(fp)
        return result

# ==================== 主窗口 ====================
class GlassVideoFinder(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("视频检索器")
        self.geometry(f"{Config.WINDOW_WIDTH}x{Config.WINDOW_HEIGHT}")
        self.minsize(Config.MIN_WIDTH, Config.MIN_HEIGHT)
        self.center_window()

        self.video_items = []
        self.filtered_items = []
        self.current_page = 1
        self.total_pages = 1
        self.loading = False
        self.scanner = VideoScanner(scan_subfolder=True)

        # 新增：浏览器播放开关状态，默认 False（使用默认播放器）
        self.use_browser = False

        self.build_ui()
        self.switch_page("page1")

    def center_window(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - Config.WINDOW_WIDTH) // 2
        y = (self.winfo_screenheight() - Config.WINDOW_HEIGHT) // 2
        self.geometry(f"+{x}+{y}")

    def build_ui(self):
        self.root_container = ctk.CTkFrame(self, fg_color="transparent")
        self.root_container.pack(fill="both", expand=True)

        self.page1 = ctk.CTkFrame(self.root_container, fg_color="transparent")
        self.page2 = ctk.CTkFrame(self.root_container, fg_color="transparent")
        self.build_page1()
        self.build_page2()

    def build_page1(self):
        main = ctk.CTkFrame(self.page1, fg_color=(Config.BG_LIGHT, Config.BG_DARK), corner_radius=0)
        main.pack(fill="both", expand=True)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(4, 0))
        ctk.CTkLabel(header, text="视频检索器", font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK)).pack(anchor="w")

        toolbar = ctk.CTkFrame(main, fg_color="transparent")
        toolbar.pack(fill="x", padx=20, pady=(0, 2))
        self.btn_select = ctk.CTkButton(toolbar, text="选择文件夹", width=120, corner_radius=10,
                                        fg_color=Config.PRIMARY_BLUE, hover_color=Config.PRIMARY_BLUE_HOVER,
                                        command=self.select_folder)
        self.btn_select.pack(side="left")
        self.search_entry = ctk.CTkEntry(toolbar, placeholder_text="搜索文件名", width=180,
                                         corner_radius=10, border_width=0,
                                         fg_color=(Config.CARD_LIGHT, Config.CARD_DARK))
        self.search_entry.pack(side="right")
        self.search_entry.bind('<KeyRelease>', self.on_search)

        self.scroll_frame = ctk.CTkScrollableFrame(main, fg_color="transparent",
                                                   scrollbar_button_color=(Config.SEPARATOR_LIGHT, Config.SEPARATOR_DARK),
                                                   scrollbar_button_hover_color=(Config.PRIMARY_BLUE, Config.PRIMARY_BLUE))
        self.scroll_frame.pack(fill="both", expand=True, padx=20, pady=(0, 4))

        self.empty_label = ctk.CTkLabel(self.scroll_frame, text="📂 点击「选择文件夹」开始扫描",
                                        font=ctk.CTkFont(size=16),
                                        text_color=(Config.TEXT_SECONDARY_LIGHT, Config.TEXT_SECONDARY_DARK))
        self.empty_label.pack(pady=60)

        # ========== 底部导航栏（已居中，且高度一致） ==========
        nav = ctk.CTkFrame(main, fg_color="transparent")
        nav.pack(fill="x", padx=20, pady=(8, 14))

        # 左侧：设置按钮
        self.btn_settings = ctk.CTkButton(nav, text="设置", width=80, height=34, corner_radius=8,
                                          fg_color="transparent", text_color=(Config.PRIMARY_BLUE, Config.PRIMARY_BLUE),
                                          hover_color=(Config.SEPARATOR_LIGHT, Config.SEPARATOR_DARK),
                                          command=partial(self.switch_page, "page2"))
        self.btn_settings.pack(side="left", padx=(0, 10), anchor="n")

        # 右侧：透明占位框架（与设置按钮等宽等高，实现对称）
        spacer_right = ctk.CTkFrame(nav, fg_color="transparent", width=80, height=34)
        spacer_right.pack(side="right", padx=(10, 0), anchor="n")

        # 中间：分页按钮组（自动居中）
        center = ctk.CTkFrame(nav, fg_color="transparent")
        center.pack(expand=True, anchor="n")

        self.btn_home = ctk.CTkButton(center, text="⏮", width=48, height=34, corner_radius=8,
                                      fg_color="transparent", text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK),
                                      hover_color=(Config.SEPARATOR_LIGHT, Config.SEPARATOR_DARK),
                                      command=self.go_home)
        self.btn_home.pack(side="left", padx=4)

        self.btn_prev = ctk.CTkButton(center, text="◀", width=48, height=34, corner_radius=8,
                                      fg_color="transparent", text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK),
                                      hover_color=(Config.SEPARATOR_LIGHT, Config.SEPARATOR_DARK),
                                      command=self.go_prev)
        self.btn_prev.pack(side="left", padx=4)

        self.page_label = ctk.CTkLabel(center, text="1 / 1", font=ctk.CTkFont(size=14),
                                       text_color=(Config.TEXT_SECONDARY_LIGHT, Config.TEXT_SECONDARY_DARK))
        self.page_label.pack(side="left", padx=16)

        self.btn_next = ctk.CTkButton(center, text="▶", width=48, height=34, corner_radius=8,
                                      fg_color="transparent", text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK),
                                      hover_color=(Config.SEPARATOR_LIGHT, Config.SEPARATOR_DARK),
                                      command=self.go_next)
        self.btn_next.pack(side="left", padx=4)

        self.btn_last = ctk.CTkButton(center, text="⏭", width=48, height=34, corner_radius=8,
                                      fg_color="transparent", text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK),
                                      hover_color=(Config.SEPARATOR_LIGHT, Config.SEPARATOR_DARK),
                                      command=self.go_last)
        self.btn_last.pack(side="left", padx=4)

        self.load_label = ctk.CTkLabel(header, text="", font=ctk.CTkFont(size=14), text_color=Config.PRIMARY_BLUE)
        self.load_label.pack(side="right", padx=(0, 10))

    # ==================== 设置页面（新增浏览器播放开关） ====================
    def build_page2(self):
        container = ctk.CTkFrame(self.page2, fg_color=(Config.BG_LIGHT, Config.BG_DARK), corner_radius=0)
        container.pack(fill="both", expand=True)

        top_frame = ctk.CTkFrame(container, fg_color="transparent")
        top_frame.pack(fill="x", padx=20, pady=(30, 20))

        ctk.CTkLabel(top_frame, text="设置", font=ctk.CTkFont(size=28, weight="bold"),
                     text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK)).pack(side="left")

        ctk.CTkLabel(top_frame, text="VF V1", font=ctk.CTkFont(size=12),
                     text_color=Config.VERSION_GRAY).pack(side="right")

        # ---- 扫描子文件夹开关 ----
        card1 = ctk.CTkFrame(container, fg_color=(Config.CARD_LIGHT, Config.CARD_DARK), corner_radius=12)
        card1.pack(fill="x", padx=20, pady=10)

        sw_frame1 = ctk.CTkFrame(card1, fg_color="transparent")
        sw_frame1.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(sw_frame1, text="扫描子文件夹", font=ctk.CTkFont(size=15),
                     text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK)).pack(side="left")

        self.subfolder_var = ctk.BooleanVar(value=self.scanner.scan_subfolder)
        sw1 = ctk.CTkSwitch(sw_frame1, text="", variable=self.subfolder_var,
                           command=self.toggle_subfolder,
                           progress_color=Config.PRIMARY_BLUE, button_color=Config.PRIMARY_BLUE)
        sw1.pack(side="right")

        ctk.CTkLabel(container, text="开启：遍历所有子文件夹\n关闭：仅扫描所选文件夹根目录",
                     font=ctk.CTkFont(size=13),
                     text_color=(Config.TEXT_SECONDARY_LIGHT, Config.TEXT_SECONDARY_DARK),
                     justify="left").pack(anchor="w", padx=20, pady=(5, 10))

        # ---- 新增：使用浏览器播放开关 ----
        card2 = ctk.CTkFrame(container, fg_color=(Config.CARD_LIGHT, Config.CARD_DARK), corner_radius=12)
        card2.pack(fill="x", padx=20, pady=10)

        sw_frame2 = ctk.CTkFrame(card2, fg_color="transparent")
        sw_frame2.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(sw_frame2, text="使用浏览器播放", font=ctk.CTkFont(size=15),
                     text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK)).pack(side="left")

        self.browser_var = ctk.BooleanVar(value=self.use_browser)
        sw2 = ctk.CTkSwitch(sw_frame2, text="", variable=self.browser_var,
                           command=self.toggle_browser,
                           progress_color=Config.PRIMARY_BLUE, button_color=Config.PRIMARY_BLUE)
        sw2.pack(side="right")

        ctk.CTkLabel(container, text="开启：使用默认浏览器播放视频\n关闭：使用系统默认播放器播放",
                     font=ctk.CTkFont(size=13),
                     text_color=(Config.TEXT_SECONDARY_LIGHT, Config.TEXT_SECONDARY_DARK),
                     justify="left").pack(anchor="w", padx=20, pady=(5, 10))

        # ---- 导出链接和返回按钮 ----
        export_btn = ctk.CTkButton(container, text="导出链接 (log.txt)", width=160, corner_radius=10,
                                   fg_color=Config.PRIMARY_BLUE, hover_color=Config.PRIMARY_BLUE_HOVER,
                                   command=self.export_links)
        export_btn.pack(anchor="w", padx=20, pady=(0, 10))

        ctk.CTkButton(container, text="返回", width=100, corner_radius=10,
                      fg_color=Config.PRIMARY_BLUE, hover_color=Config.PRIMARY_BLUE_HOVER,
                      command=partial(self.switch_page, "page1")).pack(anchor="w", padx=20, pady=20)

    # ---------- 导出链接 ----------
    def export_links(self):
        if not self.video_items:
            messagebox.showinfo("提示", "请先扫描视频文件夹")
            return
        file_path = os.path.join(os.getcwd(), "log.txt")
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                for item in self.video_items:
                    link = item.jump_url if item.jump_url else "无链接"
                    f.write(f"{item.name}\t{link}\n")
            webbrowser.open(file_path)
            messagebox.showinfo("导出成功", f"已导出 {len(self.video_items)} 条记录到 {file_path}\n浏览器将打开该文件。")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def switch_page(self, page):
        self.page1.pack_forget()
        self.page2.pack_forget()
        if page == "page1":
            self.page1.pack(fill="both", expand=True)
        else:
            self.page2.pack(fill="both", expand=True)

    def toggle_subfolder(self):
        self.scanner.scan_subfolder = self.subfolder_var.get()

    # ---------- 新增：切换浏览器播放 ----------
    def toggle_browser(self):
        self.use_browser = self.browser_var.get()

    # ---------- 搜索 ----------
    def on_search(self, event=None):
        keyword = self.search_entry.get().strip()
        if not keyword:
            self.filtered_items = self.video_items.copy()
        else:
            keyword_lower = keyword.lower()
            self.filtered_items = [item for item in self.video_items if keyword_lower in item.name.lower()]
        self.current_page = 1
        self.total_pages = max(1, (len(self.filtered_items) + Config.PAGE_SIZE - 1) // Config.PAGE_SIZE)
        self.update_page_label()
        self.render_page()

    # ---------- 选择文件夹 ----------
    def select_folder(self):
        folder = filedialog.askdirectory(title="选择视频文件夹")
        if not folder:
            return
        self.btn_select.configure(text=os.path.basename(folder))
        self.clear_cards()
        self.video_items.clear()
        self.filtered_items.clear()
        self.loading = True
        self.load_animate(0)

        def scan_thread():
            paths = self.scanner.scan(folder)
            items = []
            for p in paths:
                try:
                    items.append(VideoParser.parse_video(p))
                except Exception as e:
                    print(f"解析失败 {p}: {e}")
            self.after(0, self.on_scan_done, items)

        threading.Thread(target=scan_thread, daemon=True).start()

    def load_animate(self, step):
        if not self.loading:
            self.load_label.configure(text="")
            return
        dots = [".", "..", "...", ""]
        self.load_label.configure(text="加载中" + dots[step % len(dots)])
        self.after(300, self.load_animate, step + 1)

    def on_scan_done(self, items):
        self.loading = False
        self.load_label.configure(text="")
        self.video_items = sorted(items, key=lambda x: natural_key(x.name))
        self.search_entry.delete(0, 'end')
        self.filtered_items = self.video_items.copy()
        self.total_pages = max(1, (len(self.filtered_items) + Config.PAGE_SIZE - 1) // Config.PAGE_SIZE)
        self.current_page = 1
        self.update_page_label()
        self.render_page()

    def clear_cards(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()

    # ---------- 渲染 ----------
    def render_page(self):
        self.clear_cards()
        if not self.filtered_items:
            self.empty_label.pack(pady=60)
            return
        start = (self.current_page - 1) * Config.PAGE_SIZE
        end = min(start + Config.PAGE_SIZE, len(self.filtered_items))
        for idx, item in enumerate(self.filtered_items[start:end]):
            card = ctk.CTkFrame(self.scroll_frame, fg_color=(Config.CARD_LIGHT, Config.CARD_DARK), corner_radius=10)
            card.pack(fill="x", pady=(4, 0))
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", padx=14, pady=10)

            top = ctk.CTkFrame(inner, fg_color="transparent")
            top.pack(fill="x")
            ctk.CTkLabel(top, text=f"🎬 {item.name}", font=ctk.CTkFont(size=15, weight="bold"),
                         text_color=(Config.TEXT_PRIMARY_LIGHT, Config.TEXT_PRIMARY_DARK)).pack(side="left")
            btn_frame = ctk.CTkFrame(top, fg_color="transparent")
            btn_frame.pack(side="right")
            ctk.CTkButton(btn_frame, text="本地", width=50, height=24, corner_radius=6,
                          fg_color=Config.PRIMARY_BLUE, hover_color=Config.PRIMARY_BLUE_HOVER,
                          command=partial(self.open_local, item.path)).pack(side="left", padx=2)
            ctk.CTkButton(btn_frame, text="网页", width=50, height=24, corner_radius=6,
                          fg_color=Config.PRIMARY_BLUE, hover_color=Config.PRIMARY_BLUE_HOVER,
                          command=partial(self.open_web, item)).pack(side="left", padx=2)

            bottom = ctk.CTkFrame(inner, fg_color="transparent")
            bottom.pack(fill="x", pady=(4, 0))
            meta = f"{item.resolution} · {item.duration_str} · {item.size_str} · {item.mtime}"
            ctk.CTkLabel(bottom, text=meta, font=ctk.CTkFont(size=12),
                         text_color=(Config.TEXT_SECONDARY_LIGHT, Config.TEXT_SECONDARY_DARK)).pack(side="left")

            if item.jump_url:
                right_text = f"🔗 {item.jump_url}"
                right_color = Config.LINK_COLOR
            else:
                right_text = "无链接"
                right_color = Config.VERSION_GRAY
            ctk.CTkLabel(bottom, text=right_text, font=ctk.CTkFont(size=12), text_color=right_color).pack(side="right")

            if idx < len(self.filtered_items[start:end]) - 1:
                sep = ctk.CTkFrame(card, fg_color=(Config.SEPARATOR_LIGHT, Config.SEPARATOR_DARK), height=1)
                sep.pack(fill="x", padx=14)

    # ---------- 打开本地文件（支持浏览器播放开关） ----------
    def open_local(self, path):
        if not os.path.exists(path):
            messagebox.showwarning("文件丢失", f"视频不存在：{path}")
            return
        if self.use_browser:
            # 尝试用默认浏览器打开
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'http\shell\open\command') as key:
                    cmd = winreg.QueryValue(key, None)
                if cmd.startswith('"'):
                    browser_path = cmd.split('"')[1]
                else:
                    browser_path = cmd.split()[0]
                subprocess.Popen([browser_path, path], shell=False)
            except Exception:
                # 失败则回退
                os.startfile(path)
        else:
            # 使用系统默认播放器
            os.startfile(path)

    def open_web(self, item):
        if item.jump_url:
            webbrowser.open(item.jump_url)
        else:
            webbrowser.open(Config.SEARCH_URL_TPL.format(filename=item.name.replace(" ", "+")))

    # ---------- 分页 ----------
    def update_page_label(self):
        self.page_label.configure(text=f"{self.current_page} / {self.total_pages}")

    def go_home(self):
        self.current_page = 1
        self.update_page_label()
        self.render_page()

    def go_last(self):
        self.current_page = self.total_pages
        self.update_page_label()
        self.render_page()

    def go_prev(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.update_page_label()
            self.render_page()

    def go_next(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.update_page_label()
            self.render_page()

if __name__ == "__main__":
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    app = GlassVideoFinder()
    app.mainloop()