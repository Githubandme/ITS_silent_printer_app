import asyncio
import websockets
import tkinter as tk
from tkinter import filedialog, StringVar, messagebox
from tkinter.scrolledtext import ScrolledText
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button, Label, Entry, Combobox
import subprocess
import os
import datetime
import threading
import urllib.request
import win32print
import sys
import configparser

# 替换为 SumatraPDF.exe 的路径，假设与脚本在同一目录
# 修改为始终从当前工作目录加载
SUMATRAPDF_PATH = "SumatraPDF.exe"
LISTEN_PORT = 1972
CACHE_DIR = "cache_pdfs"
LOG_FILE = "printer_log.txt"
SETTINGS_FILE = "printer_settings.ini" # 设置文件路径
ICON_FILE = "favicon.ico" # 图标文件名

# 保留的日志行数和缓存文件数量
MAX_LOG_LINES = 10
MAX_CACHED_FILES = 10

os.makedirs(CACHE_DIR, exist_ok=True)

class SilentPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ITS静默打印服务工具")
        self.style = Style(theme="flatly")
        
        # 隐藏窗口直到UI完全加载
        self.root.withdraw()

        self.selected_printer = StringVar()
        self.cache_dir = StringVar(value=CACHE_DIR)
        self.paper_width = StringVar()
        self.paper_height = StringVar()

        self.text_log = None 

        # 设置窗口图标
        self._set_window_icon()

        self.load_settings()

        self.setup_ui()
        # 初始化日志颜色标签
        self._setup_log_colors() 
        self.log("🖨️ 静默打印服务启动...", type="system")
        self.log(f"📡 监听地址：ws://127.0.0.1:{LISTEN_PORT}", type="system")
        # 修改日志信息，显示实际使用的SumatraPDF路径
        self.log(f"🗃️ SumatraPDF 路径: {self.get_sumatra_path()}", type="info")
        self.log(f"📄 当前纸张大小设置为: 宽度 {self.paper_width.get()}mm x 高度 {self.paper_height.get()}mm。", type="info")

        if not os.path.exists(self.get_sumatra_path()):
            self.log(f"❗ 错误：未找到 SumatraPDF.exe！请确保 {self.get_sumatra_path()} 存在。", type="error")
            self.log("   请从官方网站下载 SumatraPDF (便携版推荐)：https://www.sumatrapdfreader.org/download-free-pdf-viewer", type="info")
        else:
            threading.Thread(target=self.start_server, daemon=True).start()
        
        # UI加载完成后显示窗口
        self.root.deiconify()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _set_window_icon(self):
        """设置窗口图标"""
        # 查找图标文件在当前工作目录
        icon_path = os.path.join(os.getcwd(), ICON_FILE)
        
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
                # 设置任务栏图标
                self.root.wm_iconbitmap(icon_path)
            except tk.TclError as e:
                pass # 在初始化日志系统之前，这里无法使用self.log。所以暂时pass
        else:
            # 图标文件不存在时忽略，不报错
            pass


    def load_settings(self):
        """从配置文件加载打印设置"""
        config = configparser.ConfigParser()
        config.read(SETTINGS_FILE)

        if 'PrinterSettings' in config and 'printer_name' in config['PrinterSettings']:
            self.selected_printer.set(config['PrinterSettings']['printer_name'])
        else:
            self.selected_printer.set(win32print.GetDefaultPrinter())
        
        if 'PaperSettings' in config and 'width_mm' in config['PaperSettings']:
            self.paper_width.set(config['PaperSettings']['width_mm'])
        else:
            self.paper_width.set("100") # 默认宽度
        
        if 'PaperSettings' in config and 'height_mm' in config['PaperSettings']:
            self.paper_height.set(config['PaperSettings']['height_mm'])
        else:
            self.paper_height.set("150") # 默认高度

    def save_settings(self):
        """保存当前打印设置到配置文件"""
        config = configparser.ConfigParser()
        config['PrinterSettings'] = {'printer_name': self.selected_printer.get()}
        config['PaperSettings'] = {
            'width_mm': self.paper_width.get(),
            'height_mm': self.paper_height.get()
        }
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
        self.log("配置已保存。", type="info")

    def on_closing(self):
        """处理窗口关闭事件，保存设置并退出"""
        self.save_settings()
        self.root.destroy()
        sys.exit(0)

    def _setup_log_colors(self):
        """设置日志文本框的颜色标签"""
        if self.text_log:
            self.text_log.tag_config('info', foreground='blue')
            self.text_log.tag_config('success', foreground='green')
            self.text_log.tag_config('warning', foreground='orange')
            self.text_log.tag_config('error', foreground='red')
            self.text_log.tag_config('system', foreground='purple')

    def setup_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(pady=5)

        # 打印机选择
        Label(frame, text="选择打印机:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        printer_list = [p for p in win32print.EnumPrinters(2)]
        printer_names = [p[2] for p in printer_list]
        self.printer_combo = Combobox(frame, values=printer_names, textvariable=self.selected_printer, width=37)
        self.printer_combo.grid(row=0, column=1, padx=5, pady=2, sticky="ew")

        # 纸张宽度设置
        Label(frame, text="纸张宽度 (mm):").grid(row=1, column=0, padx=5, pady=2, sticky="w")
        Entry(frame, textvariable=self.paper_width, width=10).grid(row=1, column=1, padx=5, pady=2, sticky="w")

        # 纸张高度设置
        Label(frame, text="纸张高度 (mm):").grid(row=2, column=0, padx=5, pady=2, sticky="w")
        Entry(frame, textvariable=self.paper_height, width=10).grid(row=2, column=1, padx=5, pady=2, sticky="w")

        # 缓存目录设置
        Label(frame, text="缓存目录:").grid(row=3, column=0, padx=5, pady=2, sticky="w")
        Entry(frame, textvariable=self.cache_dir, width=40).grid(row=3, column=1, padx=5, pady=2, sticky="ew")
        Button(frame, text="浏览", command=self.choose_dir).grid(row=3, column=2, padx=5, pady=2)

        # 查看历史记录按钮
        Button(frame, text="查看历史记录", command=self.open_log).grid(row=4, column=1, pady=5)

        # 日志文本框
        self.text_log = ScrolledText(self.root, height=25, width=100)
        self.text_log.pack(padx=10, pady=10)


    def choose_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.cache_dir.set(path)
            os.makedirs(path, exist_ok=True)

    def open_log(self):
        # 确保日志文件存在，否则os.startfile会报错
        if os.path.exists(LOG_FILE):
            os.startfile(LOG_FILE)
        else:
            messagebox.showinfo("信息", "日志文件不存在。")

    def log(self, message: str, type: str = 'info'):
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        full_message = f"[{timestamp}] {message}\n"
        
        # 将日志添加到GUI文本框，并应用颜色标签
        if self.text_log:
            self.text_log.insert(tk.END, full_message, type)
            self.text_log.see(tk.END)
        else:
            print(full_message, end='')
        
        # 读取现有日志行
        log_lines = []
        if os.path.exists(LOG_FILE):
            try:
                with open(LOG_FILE, "r", encoding="utf-8") as f:
                    log_lines = f.readlines()
            except Exception as e:
                if self.text_log: self.text_log.insert(tk.END, f"⚠️ 读取日志文件失败: {e}\n", 'error')
                else: print(f"⚠️ 读取日志文件失败: {e}\n", end='')
                log_lines = []

        # 添加新日志行
        log_lines.append(full_message)
        
        # 只保留最后 MAX_LOG_LINES 行
        if len(log_lines) > MAX_LOG_LINES:
            log_lines = log_lines[-MAX_LOG_LINES:]
        
        # 写回日志文件
        with open(LOG_FILE, "w", encoding="utf-8") as f:
            f.writelines(log_lines)

    def start_server(self):
        asyncio.run(self.run_ws_server())

    async def run_ws_server(self):
        async def handler(websocket):
            self.log("🔌 连接成功", type="system")
            try:
                async for message in websocket:
                    self.log(f"📨 收到消息: {message}", type="info")
                    await self.handle_print_job(message)
            except websockets.exceptions.ConnectionClosed:
                self.log("❌ 客户端断开连接", type="system")

        async with websockets.serve(handler, "127.0.0.1", LISTEN_PORT):
            await asyncio.Future()

    def _cleanup_cache_files(self):
        """清理缓存目录，只保留最近的 MAX_CACHED_FILES 个文件"""
        files = []
        for f_name in os.listdir(self.cache_dir.get()):
            f_path = os.path.join(self.cache_dir.get(), f_name)
            if os.path.isfile(f_path) and f_name.lower().endswith('.pdf'):
                # 获取文件的修改时间
                files.append((f_path, os.path.getmtime(f_path)))
        
        # 按修改时间排序，最早的在前
        files.sort(key=lambda x: x[1])
        
        # 如果文件数量超过限制，则删除最旧的文件
        if len(files) > MAX_CACHED_FILES:
            for i in range(len(files) - MAX_CACHED_FILES):
                try:
                    os.remove(files[i][0])
                    self.log(f"🗑️ 已删除旧的缓存文件: {os.path.basename(files[i][0])}", type="warning")
                except Exception as e:
                    self.log(f"⚠️ 删除旧缓存文件失败 {os.path.basename(files[i][0])}: {e}", type="error")

    def get_sumatra_path(self):
        """获取SumatraPDF的正确路径，始终从当前工作目录加载"""
        # 始终从当前工作目录加载SumatraPDF.exe
        return os.path.join(os.getcwd(), "SumatraPDF.exe")

    async def handle_print_job(self, msg: str):
        if not os.path.exists(self.get_sumatra_path()):
            self.log("⚠️ 打印失败：SumatraPDF.exe 不存在，无法执行打印操作。", type="error")
            self.log("   请确保已将 SumatraPDF.exe 复制到程序同目录下。", type="info")
            return

        try:
            job_id, pdf_url = msg.split(",", 1)
            pdf_filename = os.path.join(self.cache_dir.get(), f"{job_id}.pdf")

            self.log(f"⬇️ 下载PDF: {pdf_url}", type="info")
            urllib.request.urlretrieve(pdf_url, pdf_filename)
            self.log(f"✅ 已保存到: {pdf_filename}", type="success")

            self.log("🖨️ 正在打印 (使用 SumatraPDF)...", type="info")

            try:
                width = float(self.paper_width.get())
                height = float(self.paper_height.get())
                if width <= 0 or height <= 0:
                    raise ValueError("宽度和高度必须大于0。")
            except ValueError as e:
                self.log(f"❗ 错误：无效的纸张尺寸输入 - {e}", type="error")
                messagebox.showerror("输入错误", f"纸张宽度和高度必须是有效的正数！\n{e}")
                return

            paper_setting = f"paperSize={width}x{height}mm" 

            print_command = [
                self.get_sumatra_path(),  # 使用新的方法获取路径
                "-print-to", self.selected_printer.get(),
                "-print-settings", paper_setting,
                "-silent",
                pdf_filename
            ]

            self.log(f"📋 执行命令: {' '.join(print_command)}", type="info")

            result = subprocess.run(
                print_command,
                check=False,
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if result.returncode != 0:
                error_output = result.stdout + result.stderr
                self.log(f"⚠️ 打印失败: 错误码 {result.returncode} - {error_output.strip()}", type="error")
                self.log(f"   请确保 SumatraPDF.exe 存在且路径正确，并且您的打印机 '{self.selected_printer.get()}' 可用。", type="info")
                self.log(f"   检查纸张大小设置 '{paper_setting}' 是否被打印机支持。", type="info")
                self.log(f"   有时需要管理员权限运行此程序才能正常使用命令行打印。", type="info")
            else:
                self.log("✅ 打印完成", type="success")
                # 打印成功后清理缓存文件
                self._cleanup_cache_files() 

        except Exception as e:
            self.log(f"⚠️ 打印失败: {e}", type="error")

if __name__ == "__main__":
    root = tk.Tk()
    app = SilentPrinterApp(root)
    root.mainloop()