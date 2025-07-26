'''
Description: Convert Word/PPT to PDF or Add Watermark to PDF
version: 1.6
Author: Chandler Lu
Date: 2025-07-23 20:15:41
LastEditTime: 2025-07-26 11:15:14
'''

# -*- coding: utf-8 -*-

from pathlib import Path
from tkinter import ttk, filedialog, messagebox, scrolledtext
from typing import List, Union, Optional
from platformdirs import user_data_dir
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import urllib.request
import webbrowser

TKINTERDND_AVAILABLE = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TKINTERDND_AVAILABLE = True
except ImportError:
    pass


def get_resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class SettingsManager:
    def __init__(self, logger=None):
        self.logger = logger
        self.config_dir = Path(user_data_dir("com.chandler.rabpdf"))
        self.config_file = self.config_dir / "settings.json"
        self._ensure_config_dir()

    def log(self, message):
        if self.logger:
            self.logger(message)
        else:
            print(message)

    def _ensure_config_dir(self):
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.log(f"无法创建配置目录：{e}")

    def load_settings(self) -> dict:
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    return settings
            except (json.JSONDecodeError, IOError) as e:
                self.log(f"加载配置失败：{e}")
        return {}

    def save_settings(self, settings: dict):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
        except IOError as e:
            self.log(f"保存配置失败：{e}")


class DependencyManager:
    def __init__(self, logger=None):
        self.logger = logger
        self.system = platform.system()
        self.machine = platform.machine()

    def log(self, message):
        if self.logger:
            self.logger(message)
        else:
            print(message)

    def install_python_package(self, package_name: str) -> bool:
        try:
            self.log(f"正在安装 {package_name}...")
            command = [sys.executable, "-m", "pip",
                       "install"] + package_name.split()
            subprocess.run(command, capture_output=True,
                           text=True, check=True, encoding='utf-8')
            self.log(f"{package_name} 安装成功")
            return True
        except subprocess.CalledProcessError as e:
            self.log(f"{package_name} 安装失败：{e.stderr}")
            return False

    def check_libreoffice(self) -> bool:
        if self.system == 'Darwin':
            return os.path.exists("/Applications/LibreOffice.app")
        elif self.system == "Windows":
            paths = [r"C:\Program Files\LibreOffice\program\soffice.exe",
                     r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
            return any(os.path.exists(path) for path in paths)
        else:
            try:
                subprocess.run(["libreoffice", "--version"],
                               capture_output=True, check=True)
                return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                return False

    def find_libreoffice_path(self) -> str:
        if self.system == 'Darwin':
            macos_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
            if os.path.exists(macos_path):
                return f'"{macos_path}"'
        elif self.system == "Windows":
            paths = [r"C:\Program Files\LibreOffice\program\soffice.exe",
                     r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
            for path in paths:
                if os.path.exists(path):
                    return f'"{path}"'
        return "libreoffice"

    def get_libreoffice_download_url(self) -> Optional[str]:
        # version = "25.2.5"
        # base_url = f"https://mirror-hk.koddos.net/tdf/libreoffice/stable/{version}/"

        if self.system == "Windows":
            if self.machine.endswith('64'):
                # return f"{base_url}win/x86_64/LibreOffice_{version}_Win_x86-64.msi"
                return f"https://pan.yeslu.cn/f/wEH3/LibreOffice_25.2.5_Win_x86-64.msi"
        elif self.system == "Darwin":
            if self.machine == "arm64":
                # return f"{base_url}mac/aarch64/LibreOffice_{version}_MacOS_aarch64.dmg"
                return f"https://pan.yeslu.cn/f/EnS3/LibreOffice_25.2.5_MacOS_aarch64.dmg"
            elif self.machine == "x86_64":
                # return f"{base_url}mac/x86_64/LibreOffice_{version}_MacOS_x86-64.dmg"
                return f"https://pan.yeslu.cn/f/qMh4/LibreOffice_25.2.5_MacOS_x86-64.dmg"

        self.log(
            f"当前系统 ({self.system} {self.machine}) 暂不支持自动安装 LibreOffice。")
        return None

    def download_file(self, url: str, dest_path: str) -> bool:
        try:
            self.log(f"正在下载文件...")

            last_percent = -1

            def progress_hook(block_num, block_size, total_size):
                nonlocal last_percent
                if total_size > 0:
                    percent = int((block_num * block_size * 100) / total_size)
                    if percent > last_percent and percent % 5 == 0:
                        self.log(f"   下载进度：{percent}%")
                        last_percent = percent
            urllib.request.urlretrieve(
                url, dest_path, reporthook=progress_hook)
            self.log(f"下载完成：{os.path.basename(dest_path)}")
            return True
        except Exception as e:
            self.log(f"下载失败：{e}")
            return False

    def install_libreoffice_windows(self, msi_path: str) -> bool:
        try:
            self.log("正在静默安装 LibreOffice... 这可能需要几分钟，请稍候。")
            cmd = ['msiexec', '/i', msi_path, '/qn']
            subprocess.run(cmd, check=True, capture_output=True, text=True)
            self.log("LibreOffice 安装成功！")
            return True
        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            self.log(f"静默安装失败：{e.stderr or e}")
            self.log("正在尝试打开安装程序，请手动完成安装。")
            os.startfile(msi_path)
            return False

    def install_libreoffice_macos(self, dmg_path: str) -> bool:
        mount_point = None
        try:
            self.log(f"正在挂载磁盘映像：{os.path.basename(dmg_path)}...")
            attach_result = subprocess.run(
                ['hdiutil', 'attach', dmg_path], check=True, capture_output=True, text=True)
            for line in attach_result.stdout.strip().split('\n'):
                if '/Volumes/' in line:
                    mount_point = line.split('\t')[-1].strip()
                    break

            if not mount_point or not os.path.exists(mount_point):
                raise RuntimeError("挂载 DMG 失败：找不到挂载点。")
            self.log(f"映像已挂载于：{mount_point}")

            app_file = next((f for f in os.listdir(mount_point)
                            if f.endswith('.app')), None)
            if not app_file:
                raise RuntimeError("在 DMG 中找不到 .app 文件。")

            source_app_path = os.path.join(mount_point, app_file)
            self.log(f"正在安装 {app_file} 到 /Applications 文件夹...")
            self.log("系统将提示您输入管理员密码以授权安装。")

            applescript = f'do shell script "cp -R \\"{source_app_path}\\" \\"/Applications/\\"" with administrator privileges'
            subprocess.run(['osascript', '-e', applescript], check=True)

            self.log("LibreOffice 安装成功！")
            return True
        except (subprocess.CalledProcessError, RuntimeError, FileNotFoundError) as e:
            self.log(f"macOS 安装失败：{e}")
            return False
        finally:
            if mount_point:
                self.log(f"正在卸载磁盘映像...")
                subprocess.run(
                    ['hdiutil', 'detach', mount_point], capture_output=True)


class WatermarkManager:
    def __init__(self, logger=None):
        self.logger = logger
        self.font_path = None
        self.temp_dir = None
        self.font_name = "Source Han Serif CN"
        self.font_registered = False
        if self._get_accessible_font_path():
            self._register_font()
        else:
            self.log("错误：无法访问或创建字体文件的可写副本。")

    def log(self, message):
        if self.logger:
            self.logger(message)
        else:
            print(message)

    def _get_accessible_font_path(self) -> Optional[str]:
        try:
            original_font_path = get_resource_path("SourceHanSerif.ttf")
            if not os.path.exists(original_font_path):
                self.log(f"警告：在资源中找不到字体文件 {original_font_path}。")
                return None

            base_temp_dir = tempfile.gettempdir()
            self.temp_dir = os.path.join(base_temp_dir, "com.chandler.rabpdf")
            os.makedirs(self.temp_dir, exist_ok=True)
            accessible_font_path = os.path.join(self.temp_dir, "SourceHanSerif.ttf")
            shutil.copy2(original_font_path, accessible_font_path)
            self.font_path = accessible_font_path
            return accessible_font_path
        except Exception as e:
            self.log(f"处理字体文件时出错：{e}")
            return None

    def _register_font(self):
        try:
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
        except ImportError:
            self.log("缺少 ReportLab 库，无法注册字体。")
            return

        try:
            pdfmetrics.registerFont(TTFont(self.font_name, self.font_path))
            self.font_registered = True
        except Exception as e:
            self.log(f"注册字体失败：{e}")

    def add_watermark(self, pdf_path, watermark_text, output_path, opacity, font_size, rotation):
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.colors import gray
            import PyPDF2
        except ImportError:
            self.log("缺少 PDF 库。")
            return False
        if not self.font_registered:
            self.log("字体未注册。")
            return False

        pdf_path, output_path, watermark_canvas_path = Path(
            pdf_path), Path(output_path), None
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                if not pdf_reader.pages:
                    self.log(f"PDF '{pdf_path.name}' 为空。")
                    return False
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
                    watermark_canvas_path = temp_file.name

                page_width, page_height = map(
                    float, pdf_reader.pages[0].mediabox[2:])

                c = canvas.Canvas(watermark_canvas_path,
                                  pagesize=(page_width, page_height))
                c.setFillColor(gray, alpha=opacity)
                c.setFont(self.font_name, font_size)
                c.translate(page_width / 2, page_height / 2)
                c.rotate(rotation)
                text_width = c.stringWidth(
                    watermark_text, self.font_name, font_size)

                for x in range(-int(page_width*2), int(page_width*2), int(text_width*1.5)):
                    for y in range(-int(page_height*2), int(page_height*2), int(font_size*5)):
                        c.drawString(x, y, watermark_text)
                c.save()

                with open(watermark_canvas_path, 'rb') as watermark_file:
                    watermark_page = PyPDF2.PdfReader(watermark_file).pages[0]
                    pdf_writer = PyPDF2.PdfWriter()
                    for page in pdf_reader.pages:
                        page.merge_page(watermark_page)
                        pdf_writer.add_page(page)
                    with open(output_path, 'wb') as output_file:
                        pdf_writer.write(output_file)

            self.log(f"水印添加成功：{output_path.name}")
            return True
        except Exception as e:
            self.log(f"添加水印时出错：{str(e)}")
            return False
        finally:
            if watermark_canvas_path and os.path.exists(watermark_canvas_path):
                os.unlink(watermark_canvas_path)


class OfficeToPDFConverter:
    def __init__(self, logger=None):
        self.logger = logger
        self.system = platform.system()
        self.dependency_manager = DependencyManager(logger)

    def log(self, message):
        if self.logger:
            self.logger(message)
        else:
            print(message)

    def find_libreoffice_path(self):
        return self.dependency_manager.find_libreoffice_path()

    def convert_with_libreoffice(self, input_path: Path, output_dir: Path) -> bool:
        soffice_cmd = self.dependency_manager.find_libreoffice_path()
        try:
            if self.system in ["Windows", "Darwin"] and soffice_cmd.startswith('"'):
                cmd_str = f'{soffice_cmd} --headless --convert-to pdf --outdir "{output_dir}" "{input_path}"'
            else:
                cmd_str = [soffice_cmd, "--headless", "--convert-to",
                           "pdf", "--outdir", str(output_dir), str(input_path)]

            self.log(f"正在使用 LibreOffice 转换 {input_path.name}...")
            result = subprocess.run(
                cmd_str, shell=True, capture_output=True, text=True, encoding='utf-8')
            output_file = output_dir / f"{input_path.stem}.pdf"
            if result.returncode == 0 and output_file.exists():
                self.log(f"转换成功：{output_file.name}")
                return True
            self.log(f"LibreOffice 转换失败：{result.stdout or result.stderr}")
            return False
        except Exception as e:
            self.log(f"LibreOffice 转换出错：{str(e)}")
            return False

    def convert_with_comtypes(self, input_path: Path, output_path: Path) -> bool:
        if self.system != "Windows":
            self.log("comtypes 仅支持 Windows")
            return False
        try:
            import comtypes.client
        except ImportError:
            self.log("comtypes 未安装")
            return False

        file_extension = input_path.suffix.lower()
        if file_extension in ['.doc', '.docx']:
            word, doc = None, None
            try:
                self.log(f"正在使用 Word 转换 {input_path.name}...")
                word = comtypes.client.CreateObject("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(str(input_path.resolve()))

                wdFormatPDF = 17
                doc.SaveAs(str(output_path.resolve()), FileFormat=wdFormatPDF)

                self.log(f"转换成功：{output_path.name}")
                return True
            except Exception as e:
                self.log(f"Word/comtypes 转换出错：{str(e)}")
                return False
            finally:
                if doc:
                    doc.Close(SaveChanges=0)
                if word:
                    word.Quit()
        else:
            powerpoint, presentation = None, None
            try:
                self.log(f"正在使用 PowerPoint 转换 {input_path.name}...")
                powerpoint = comtypes.client.CreateObject(
                    "Powerpoint.Application")
                powerpoint.Visible = 1
                presentation = powerpoint.Presentations.Open(
                    str(input_path.resolve()))
                presentation.SaveAs(str(output_path.resolve()), 32)
                self.log(f"转换成功：{output_path.name}")
                return True
            except Exception as e:
                self.log(f"PowerPoint/comtypes 转换出错：{str(e)}")
                return False
            finally:
                if presentation:
                    presentation.Close()
                if powerpoint:
                    powerpoint.Quit()


class OfficeToPDFGUI:
    def __init__(self):
        if TKINTERDND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("RabPDF")
        self.root.geometry("800x750")
        self.root.minsize(600, 500)

        self.log_message(
            f"欢迎使用 RabPDF，本软件可以帮助您完成 Word/PPT 转换为 PDF 格式并添加水印，或者直接为 PDF 添加水印。")

        # try:
        #     icon_path = get_resource_path("icon/icon.png")
        #     self.icon_image = tk.PhotoImage(file=icon_path)
        #     self.root.iconphoto(True, self.icon_image)
        # except tk.TclError:
        #     pass

        self.dependency_manager = DependencyManager(self.log_message)
        self.watermark_manager = WatermarkManager(self.log_message)
        self.converter = OfficeToPDFConverter(self.log_message)
        self.settings_manager = SettingsManager(self.log_message)

        settings = self.settings_manager.load_settings()

        self.input_files = []
        self.output_directory = tk.StringVar()
        self.conversion_method = tk.StringVar(value="auto")
        self.add_watermark = tk.BooleanVar(value=True)

        self.watermark_history = settings.get('watermark_history', [])
        last_text = settings.get(
            'last_watermark_text', "深势科技版权所有，请勿外传")
        opacity = settings.get('opacity', 0.2)
        size = settings.get('size', 25)
        rotation = settings.get('rotation', 30)

        self.watermark_text = tk.StringVar(value=last_text)
        self.watermark_opacity = tk.DoubleVar(value=opacity)
        self.watermark_size = tk.IntVar(value=size)
        self.watermark_rotation = tk.IntVar(value=rotation)

        self.setup_ui()
        self.check_dependencies()

        if TKINTERDND_AVAILABLE:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self._handle_drop)

    def _on_mousewheel(self, event):
        if platform.system() == 'Darwin':
            self.canvas.yview_scroll(-1 * event.delta, "units")
        elif platform.system() == 'Linux':
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _update_opacity_label(self, value):
        self.opacity_value_label.config(text=f"{int(float(value) * 100)}%")

    def _update_size_label(self, value):
        self.size_value_label.config(text=f"{int(float(value))}")

    def _update_rotation_label(self, value):
        self.rotation_value_label.config(text=f"{int(float(value))}°")

    def check_and_guide_macos_permissions(self):
        if platform.system() != 'Darwin':
            return True
        try:
            os.listdir(os.path.expanduser('~/Documents'))
            return True
        except PermissionError:
            messagebox.showwarning("警告", "可能缺少完全磁盘访问权限，确认后跳转设置页。")
        try:
            subprocess.run(
                ['open', 'x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles'], check=True)
        except Exception as e:
            self.log_message(
                "无法自动打开系统设置。请手动前往“系统设置”->“隐私与安全性”->“完全磁盘访问权限”进行设置。")
        return False

    def setup_ui(self):
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)

        self.canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            container, orient="vertical", command=self.canvas.yview)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>", self._on_mousewheel)
        self.root.bind_all("<Button-5>", self._on_mousewheel)

        self.scrollable_frame = ttk.Frame(self.canvas, padding="10")
        self.frame_id = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw")

        def _configure_frame(event): self.canvas.configure(
            scrollregion=self.canvas.bbox("all"))
        def _configure_canvas(event): self.canvas.itemconfigure(
            self.frame_id, width=event.width)

        self.scrollable_frame.bind("<Configure>", _configure_frame)
        self.canvas.bind("<Configure>", _configure_canvas)

        main_frame = self.scrollable_frame
        main_frame.columnconfigure(0, weight=1)
        row = 0

        title_label = ttk.Label(
            main_frame, text="Word/PPT/PDF 处理工作流", font=("Source Han Serif CN", 16, "bold"))
        title_label.grid(row=row, column=0, pady=(0, 20), sticky='n')
        row += 1

        files_frame = ttk.LabelFrame(
            main_frame, text="第一步：选择文件 (或拖拽文件到此窗口)", padding="10")
        files_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        row += 1
        files_frame.columnconfigure(0, weight=1)

        listbox_frame = ttk.Frame(files_frame)
        listbox_frame.grid(row=0, column=0, sticky='ew')
        listbox_frame.columnconfigure(0, weight=1)
        self.files_listbox = tk.Listbox(listbox_frame, height=5)
        self.files_listbox.grid(row=0, column=0, sticky="ew")
        files_scrollbar = ttk.Scrollbar(
            listbox_frame, orient="vertical", command=self.files_listbox.yview)
        files_scrollbar.grid(row=0, column=1, sticky="ns")
        self.files_listbox.config(yscrollcommand=files_scrollbar.set)

        btn_frame = ttk.Frame(files_frame)
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        ttk.Button(btn_frame, text="添加文件", command=self.add_files).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="添加文件夹", command=self.add_folder).pack(
            side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="清空列表",
                   command=self.clear_files).pack(side=tk.LEFT)

        output_frame = ttk.LabelFrame(
            main_frame, text="第二步：设置输出", padding="10")
        output_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        row += 1
        output_frame.columnconfigure(1, weight=1)
        ttk.Label(output_frame, text="输出目录：").grid(
            row=0, column=0, sticky=tk.W)
        ttk.Entry(output_frame, textvariable=self.output_directory).grid(
            row=0, column=1, sticky="ew", padx=5)
        ttk.Button(output_frame, text="浏览",
                   command=self.select_output_dir).grid(row=0, column=2)

        settings_frame = ttk.LabelFrame(
            main_frame, text="第三步：配置选项", padding="10")
        settings_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        row += 1
        settings_frame.columnconfigure(1, weight=1)

        ttk.Label(settings_frame, text="Office2PDF 转换方法：").grid(
            row=0, column=0, sticky=tk.W)
        method_combo = ttk.Combobox(settings_frame, textvariable=self.conversion_method,
                                    values=["auto", "comtypes", "libreoffice"], state="readonly")
        method_combo.grid(row=0, column=1, sticky=tk.W, padx=5)
        if platform.system() != "Windows":
            self.conversion_method.set("libreoffice")
            method_combo.config(
                values=["auto", "libreoffice"], state="disabled")

        ttk.Checkbutton(settings_frame, text="为输出的 PDF 添加水印", variable=self.add_watermark,
                        command=self.toggle_watermark_options).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))

        self.watermark_options_frame = ttk.Frame(settings_frame)
        self.watermark_options_frame.grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(5, 0), padx=20)

        self.watermark_options_frame.columnconfigure(1, weight=1)
        self.watermark_options_frame.columnconfigure(2, weight=0)

        ttk.Label(self.watermark_options_frame, text="水印文字：").grid(
            row=0, column=0, sticky=tk.W)
        self.watermark_text_combo = ttk.Combobox(
            self.watermark_options_frame, textvariable=self.watermark_text)
        self.watermark_text_combo['values'] = self.watermark_history
        self.watermark_text_combo.grid(
            row=0, column=1, columnspan=2, sticky="ew", pady=(0, 5))

        ttk.Label(self.watermark_options_frame, text="透明度：").grid(
            row=1, column=0, sticky=tk.W)
        ttk.Scale(self.watermark_options_frame, from_=0.05, to=1.0, variable=self.watermark_opacity,
                  orient=tk.HORIZONTAL, command=self._update_opacity_label).grid(row=1, column=1, sticky="ew", pady=(0, 5))
        self.opacity_value_label = ttk.Label(
            self.watermark_options_frame, width=5, anchor="w")
        self.opacity_value_label.grid(
            row=1, column=2, sticky=tk.W, padx=(5, 0))

        ttk.Label(self.watermark_options_frame, text="字体大小：").grid(
            row=2, column=0, sticky=tk.W)
        ttk.Scale(self.watermark_options_frame, from_=20, to=100, variable=self.watermark_size,
                  orient=tk.HORIZONTAL, command=self._update_size_label).grid(row=2, column=1, sticky="ew", pady=(0, 5))
        self.size_value_label = ttk.Label(
            self.watermark_options_frame, width=5, anchor="w")
        self.size_value_label.grid(row=2, column=2, sticky=tk.W, padx=(5, 0))

        ttk.Label(self.watermark_options_frame, text="旋转角度：").grid(
            row=3, column=0, sticky=tk.W)
        ttk.Scale(self.watermark_options_frame, from_=0, to=90, variable=self.watermark_rotation,
                  orient=tk.HORIZONTAL, command=self._update_rotation_label).grid(row=3, column=1, sticky="ew")
        self.rotation_value_label = ttk.Label(
            self.watermark_options_frame, width=5, anchor="w")
        self.rotation_value_label.grid(
            row=3, column=2, sticky=tk.W, padx=(5, 0))

        self._update_opacity_label(self.watermark_opacity.get())
        self._update_size_label(self.watermark_size.get())
        self._update_rotation_label(self.watermark_rotation.get())

        self.toggle_watermark_options()

        action_frame = ttk.LabelFrame(
            main_frame, text="第四步：开始处理", padding="10")
        action_frame.grid(row=row, column=0, sticky="ew", pady=(10, 10))
        row += 1
        self.convert_button = ttk.Button(
            action_frame, text="开始处理", command=self.start_processing_with_permission_check)
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        self.progress = ttk.Progressbar(
            action_frame, length=200, mode='indeterminate')
        self.progress.pack(side=tk.LEFT)

        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding="10")
        log_frame.grid(row=row, column=0, sticky="nsew")
        row += 1
        main_frame.rowconfigure(row - 1, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=8, wrap=tk.WORD, state='disabled')
        self.log_text.grid(row=0, column=0, sticky="nsew")
        ttk.Button(log_frame, text="清空日志", command=self.clear_log).grid(
            row=1, column=0, pady=(5, 0), sticky='e')

        deps_frame = ttk.LabelFrame(
            main_frame, text="依赖（在 Windows 下可选）", padding="10")
        deps_frame.grid(row=row, column=0, sticky="ew", pady=(10, 0))
        row += 1
        deps_frame.columnconfigure(1, weight=1)
        self.libreoffice_status = ttk.Label(
            deps_frame, text="LibreOffice (Word/PPT 转换): 检查中...")
        self.libreoffice_status.grid(row=0, column=0, sticky='w', pady=2)

        deps_btn_frame = ttk.Frame(deps_frame)
        deps_btn_frame.grid(row=0, column=1, sticky='e')
        self.install_lo_button = ttk.Button(
            deps_btn_frame, text="安装 LibreOffice", command=self.install_libreoffice)
        self.install_lo_button.pack(side=tk.LEFT, padx=(0, 5))
        self.download_lo_button = ttk.Button(
            deps_btn_frame, text="手动下载 LibreOffice", command=self.open_download_page)
        self.download_lo_button.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(deps_btn_frame, text="刷新状态",
                   command=self.check_dependencies).pack(side=tk.LEFT)

    def install_libreoffice(self):
        if self.dependency_manager.check_libreoffice():
            messagebox.showinfo("信息", "LibreOffice 已经安装。")
            return
        if messagebox.askyesno("确认", "将开始下载并安装 LibreOffice。\n此过程可能需要几分钟，并可能需要管理员权限。\n是否继续？"):
            self.install_lo_button.config(state="disabled")
            threading.Thread(
                target=self._install_libreoffice_thread, daemon=True).start()

    def open_download_page(self):
        webbrowser.open_new(
            "https://mirrors.cloud.tencent.com/libreoffice/libreoffice/stable/")

    def _install_libreoffice_thread(self):
        try:
            url = self.dependency_manager.get_libreoffice_download_url()
            if not url:
                return
            with tempfile.TemporaryDirectory() as temp_dir:
                file_ext = ".msi" if platform.system() == "Windows" else ".dmg"
                installer_path = os.path.join(
                    temp_dir, f"LibreOffice_installer{file_ext}")
                if not self.dependency_manager.download_file(url, installer_path):
                    return

                if platform.system() == "Windows":
                    self.dependency_manager.install_libreoffice_windows(
                        installer_path)
                elif platform.system() == "Darwin":
                    self.dependency_manager.install_libreoffice_macos(
                        installer_path)
        except Exception as e:
            self.log_message(f"安装过程中发生意外错误：{e}")
        finally:
            self.check_dependencies()

    def check_dependencies(self):
        def check():
            if self.dependency_manager.check_libreoffice():
                self.libreoffice_status.config(
                    text="LibreOffice (Word/PPT 转换): 已安装", foreground="green")
                if self.root.winfo_exists():
                    self.install_lo_button.config(state="disabled")
            else:
                self.libreoffice_status.config(
                    text="LibreOffice (Word/PPT 转换): 未安装", foreground="red")
                if self.root.winfo_exists():
                    self.install_lo_button.config(state="normal")
        threading.Thread(target=check, daemon=True).start()

    def _add_file_paths(self, file_paths: List[str]):
        if not self.output_directory.get() and file_paths:
            self.output_directory.set(str(Path(file_paths[0]).parent))
            self.log_message(f"输出目录已自动设置为：{self.output_directory.get()}")
        added_count = 0
        for file in file_paths:
            if not file.lower().endswith(('.pptx', '.pdf', '.ppt', '.docx', '.doc')):
                continue
            if file not in self.input_files:
                self.input_files.append(file)
                self.files_listbox.insert(tk.END, os.path.basename(file))
                added_count += 1
        if added_count > 0:
            self.log_message(f"添加了 {added_count} 个新文件。")

    def _handle_drop(self, event):
        try:
            self.log_message(f"检测到拖放操作...")
            self._add_file_paths(self.root.tk.splitlist(event.data))
        except Exception as e:
            self.log_message(f"处理拖放文件时出错：{e}")

    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择文件", filetypes=[("可处理文件", "*.pptx *.ppt *.pdf *.doc *.docx"), ("所有文件", "*.*")])
        if files:
            self._add_file_paths(list(files))

    def add_folder(self):
        folder = filedialog.askdirectory(title="选择文件夹")
        if folder:
            found = [str(f) for ext in ("*.pptx", "*.ppt", "*.pdf", "*.doc", "*.docx")
                     for f in Path(folder).glob(ext)]
            if found:
                self._add_file_paths(found)

    def clear_files(self): self.input_files.clear(
    ); self.files_listbox.delete(0, tk.END)

    def select_output_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.output_directory.set(d)

    def start_processing_with_permission_check(self):
        if platform.system() == 'Darwin':
            if not self.check_and_guide_macos_permissions():
                self.log_message("已启动权限设置流程，请手动将本程序添加至“完全磁盘访问权限”中！")
                self.log_message("添加后，请使用 Command+Q 结束本进程，并重新打开。")
                return
        self.start_processing()

    def start_processing(self):
        if not self.input_files:
            messagebox.showwarning("警告", "请先添加要处理的文件。")
            return
        if not self.output_directory.get():
            messagebox.showwarning("警告", "请选择输出目录。")
            return
        self.convert_button.config(state="disabled")
        self.progress.start()
        threading.Thread(target=self._perform_processing, daemon=True).start()

    def _save_current_settings(self):
        """Saves current watermark settings and updates history."""
        current_text = self.watermark_text.get().strip()
        if not current_text:
            return

        if current_text in self.watermark_history:
            self.watermark_history.remove(current_text)
        self.watermark_history.insert(0, current_text)
        self.watermark_history = self.watermark_history[:5]

        settings = {
            'last_watermark_text': current_text,
            'opacity': self.watermark_opacity.get(),
            'size': self.watermark_size.get(),
            'rotation': self.watermark_rotation.get(),
            'watermark_history': self.watermark_history,
        }
        self.settings_manager.save_settings(settings)

        if hasattr(self, 'watermark_text_combo'):
            self.watermark_text_combo['values'] = self.watermark_history

    def _perform_processing(self):
        try:
            should_save_settings = self.add_watermark.get()
            output_dir = Path(self.output_directory.get())
            output_dir.mkdir(parents=True, exist_ok=True)
            success_count, total_files = 0, len(self.input_files)
            for i, file_path_str in enumerate(self.input_files, 1):
                self.log_message(
                    f"\n--- [{i}/{total_files}] 处理：{os.path.basename(file_path_str)} ---")
                input_path, processed_pdf_path = Path(file_path_str), None
                if input_path.suffix.lower() in ['.pptx', '.ppt', '.docx', '.doc']:
                    method = self.conversion_method.get()
                    conversion_success = False
                    output_pdf_path = output_dir / f"{input_path.stem}.pdf"
                    use_comtypes = (method == "comtypes") or \
                        (method == "auto" and platform.system() == "Windows")

                    if use_comtypes:
                        conversion_success = self.converter.convert_with_comtypes(
                            input_path, output_pdf_path)
                        if not conversion_success and method == "auto":
                            self.log_message(
                                "PowerPoint (comtypes) 转换失败。正在自动尝试 LibreOffice...")
                            conversion_success = self.converter.convert_with_libreoffice(
                                input_path, output_dir)
                    else:
                        conversion_success = self.converter.convert_with_libreoffice(
                            input_path, output_dir)

                    if conversion_success:
                        processed_pdf_path = output_pdf_path
                    else:
                        self.log_message(f"文件 {input_path.name} 转换失败。")
                        continue
                elif input_path.suffix.lower() == '.pdf':
                    self.log_message("PDF 文件，跳过转换。")
                    processed_pdf_path = input_path

                if processed_pdf_path and processed_pdf_path.exists():
                    if self.add_watermark.get():
                        watermarked_output_path = output_dir / \
                            f"{processed_pdf_path.stem}_watermarked.pdf"
                        if self.watermark_manager.add_watermark(processed_pdf_path, self.watermark_text.get(), watermarked_output_path, self.watermark_opacity.get(), self.watermark_size.get(), self.watermark_rotation.get()):
                            success_count += 1
                        else:
                            self.log_message(
                                f"文件 {processed_pdf_path.name} 添加水印失败。")
                    else:
                        final_path = output_dir / processed_pdf_path.name
                        if not final_path.exists() or not final_path.samefile(processed_pdf_path):
                            shutil.copy2(processed_pdf_path, final_path)
                        self.log_message(f"文件处理完成（无水印）。")
                        success_count += 1
            self.log_message(f"\n全部处理完成！成功：{success_count}/{total_files}")
            if should_save_settings and success_count > 0:
                self.root.after(0, self._save_current_settings)
            # if success_count > 0 and self.root.winfo_exists():
            #     self.root.after(0, lambda: messagebox.askyesno(
            #         "完成", f"成功处理 {success_count} 个文件!\n是否打开输出文件夹？") and self._open_folder(output_dir))
        except Exception as e:
            self.log_message(f"处理过程中发生严重错误：{str(e)}")
        finally:
            if self.root.winfo_exists():
                self.root.after(0, lambda: (self.progress.stop(),
                                self.convert_button.config(state="normal")))

    def log_message(self, message):
        def _log():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, f"{message}\n")
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        if hasattr(self, 'root') and self.root.winfo_exists():
            self.root.after(0, _log)

    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def toggle_watermark_options(self):
        state = tk.NORMAL if self.add_watermark.get() else tk.DISABLED
        for child in self.watermark_options_frame.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass

    # def _open_folder(self, folder_path):
    #     try:
    #         if platform.system() == "Windows":
    #             os.startfile(folder_path)
    #         elif platform.system() == "Darwin":
    #             subprocess.run(["open", str(folder_path)])
    #         else:
    #             subprocess.run(["xdg-open", str(folder_path)])
    #     except Exception as e:
    #         self.log_message(f"无法自动打开文件夹：{str(e)}")

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.mainloop()

    def on_close(self):
        try:
            self._save_current_settings()
            temp_dir = getattr(self.watermark_manager, "temp_dir", None)
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

        except Exception as e:
            self.log_message(f"关闭程序时出错：{e}")
        finally:
            self.root.destroy()


def main():
    try:
        app = OfficeToPDFGUI()
        app.run()
    except Exception as e:
        print(f"程序启动失败：{str(e)}")
        try:
            messagebox.showerror("严重错误", f"程序启动失败：{str(e)}")
        except tk.TclError:
            pass


if __name__ == "__main__":
    main()
