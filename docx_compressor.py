import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import zipfile
import threading
import tempfile
import io
import platform
from pathlib import Path

try:
    from PIL import Image, ImageEnhance
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False


# ──────────────────────────────────────────────
#  跨平台字体检测
# ──────────────────────────────────────────────
def get_font_family():
    system = platform.system()
    if system == "Darwin":
        return "PingFang SC"
    elif system == "Windows":
        return "Microsoft YaHei UI"
    else:
        return "Noto Sans CJK SC"

FONT_FAMILY = get_font_family()


# ──────────────────────────────────────────────
#  公共工具函数
# ──────────────────────────────────────────────
def has_transparency(img):
    if img.mode == "RGBA":
        return img.split()[3].getextrema()[0] < 255
    if img.mode == "P":
        return "transparency" in img.info
    return False


# ──────────────────────────────────────────────
#  DOCX 核心压缩逻辑
# ──────────────────────────────────────────────
SUPPORTED_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}

def compress_image(img_path, quality, max_dim, max_bytes, force_jpeg, log_fn):
    ext = img_path.suffix.lower()
    if ext not in SUPPORTED_EXT:
        return 0, 0, False, img_path

    original_size = img_path.stat().st_size

    try:
        with Image.open(img_path) as img:
            img.load()
            is_transparent = has_transparency(img)

            if max_dim > 0:
                w, h = img.size
                if w > max_dim or h > max_dim:
                    ratio = min(max_dim / w, max_dim / h)
                    img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

            def to_rgb(image):
                if image.mode == 'RGBA':
                    bg = Image.new('RGB', image.size, (255, 255, 255))
                    bg.paste(image, mask=image.split()[3])
                    return bg
                if image.mode != 'RGB':
                    return image.convert('RGB')
                return image

            if is_transparent and not force_jpeg:
                save_fmt = 'PNG'
                log_fn(f"  🔒 {img_path.name}：含透明通道，保持 PNG 格式")
            elif is_transparent and force_jpeg:
                save_fmt = 'JPEG'
                log_fn(f"  ⚠️  {img_path.name}：含透明通道，强制转 JPEG（透明区域填白色）")
            elif ext in {'.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp'}:
                save_fmt = 'JPEG'
            else:
                save_fmt = 'PNG'

            cur_quality = quality
            cur_img = img
            converted_to_jpeg = False
            new_path = img_path

            while True:
                buf = io.BytesIO()
                if save_fmt == 'JPEG':
                    to_rgb(cur_img).save(buf, format='JPEG', quality=cur_quality, optimize=True)
                else:
                    cur_img.save(buf, format='PNG', optimize=True, compress_level=9)

                size_ok = buf.tell() <= max_bytes

                if size_ok or cur_quality <= 20:
                    if save_fmt == 'JPEG' and ext not in {'.jpg', '.jpeg'}:
                        new_path = img_path.with_suffix('.jpg')
                        new_path.write_bytes(buf.getvalue())
                        img_path.unlink()
                        converted_to_jpeg = True
                        log_fn(f"  🔄 {img_path.name} → {new_path.name}（已重命名）")
                    else:
                        img_path.write_bytes(buf.getvalue())
                    break

                if save_fmt == 'PNG' and (not is_transparent or force_jpeg):
                    save_fmt = 'JPEG'
                    cur_img = to_rgb(cur_img)
                    log_fn(f"  🔄 {img_path.name}：PNG 超限，转 JPEG 继续压缩")
                else:
                    cur_quality -= 5

    except Exception as exc:
        log_fn(f"  ⚠  跳过 {img_path.name}：{exc}")
        return original_size, original_size, False, img_path

    new_size = new_path.stat().st_size
    return original_size, new_size, converted_to_jpeg, new_path


def update_rels_for_renamed(extract_dir, old_name, new_name, log_fn):
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return
    content = rels_path.read_text(encoding='utf-8')
    if old_name in content:
        new_content = content.replace(old_name, new_name)
        rels_path.write_text(new_content, encoding='utf-8')
        log_fn(f"  📝 已更新 .rels 引用：{old_name} → {new_name}")


def run_compress_docx(input_path, quality, max_dim, max_mb,
                      force_jpeg, progress_fn, status_fn, log_fn):
    output_path = input_path.parent / (input_path.stem + "_compressed.docx")
    max_bytes = int(max_mb * 1024 * 1024)
    orig_total = input_path.stat().st_size

    log_fn(f"📂 文件：{input_path.name}")
    log_fn(f"📦 原始大小：{orig_total / 1024 / 1024:.2f} MB\n")

    with tempfile.TemporaryDirectory() as tmp:
        extract_dir = Path(tmp) / "doc"

        status_fn("正在解压文档…")
        log_fn("── Step 1 解压 ──────────────────────────")
        try:
            with zipfile.ZipFile(input_path, 'r') as z:
                z.extractall(extract_dir)
            log_fn("  ✅ 解压成功\n")
        except Exception as e:
            log_fn(f"  ❌ 解压失败：{e}")
            return

        media_dir = extract_dir / "word" / "media"
        if not media_dir.exists():
            log_fn("⚠  未找到 word/media 目录，文档中可能没有图片。")
            return

        all_files = [f for f in media_dir.iterdir() if f.is_file()]
        log_fn(f"── Step 2 共发现 {len(all_files)} 个媒体文件 ──────────")
        log_fn(f"\n── Step 3 开始逐张压缩 ──────────────────")

        total_orig = total_new = 0
        for idx, fpath in enumerate(all_files, 1):
            progress_fn((idx / len(all_files)) * 80)
            status_fn(f"处理 {idx}/{len(all_files)}: {fpath.name}")

            orig, new, renamed, new_path = compress_image(
                fpath, quality, max_dim, max_bytes, force_jpeg, log_fn)

            total_orig += orig
            total_new  += new

            if renamed:
                update_rels_for_renamed(
                    extract_dir, fpath.name, new_path.name, log_fn)

            if orig == 0:
                log_fn(f"  ➡  {fpath.name}（非图片格式，跳过）")
            elif orig == new:
                log_fn(f"  ➡  {fpath.name}：{orig / 1024:.0f} KB（无需压缩）")
            else:
                ratio = (1 - new / orig) * 100 if orig else 0
                log_fn(f"  ✅ {fpath.name}: "
                       f"{orig / 1024:.0f} KB → {new / 1024:.0f} KB "
                       f"(↓{ratio:.0f}%)")

        log_fn(f"\n  图片合计：{total_orig / 1024 / 1024:.2f} MB → "
               f"{total_new / 1024 / 1024:.2f} MB\n")

        status_fn("正在重新打包…")
        log_fn("── Step 4 重新打包为 .docx ──────────────")
        progress_fn(85)
        try:
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for f in extract_dir.rglob('*'):
                    if f.is_file():
                        zout.write(f, f.relative_to(extract_dir))
            log_fn("  ✅ 打包成功")
        except Exception as e:
            log_fn(f"  ❌ 打包失败：{e}")
            return

        progress_fn(100)
        final_size = output_path.stat().st_size
        saved_pct = (1 - final_size / orig_total) * 100 if orig_total else 0
        log_fn(f"\n{'=' * 52}\n  ✅  处理完成！\n"
               f"  文档大小：{orig_total / 1024 / 1024:.2f} MB  →  "
               f"{final_size / 1024 / 1024:.2f} MB  (节省 {saved_pct:.1f}%)\n"
               f"  输出文件：{output_path}\n{'=' * 52}")
        status_fn(f"✅  完成！节省 {saved_pct:.1f}% 空间")
        return str(output_path), orig_total, final_size


# ──────────────────────────────────────────────
#  PDF 核心压缩逻辑（支持扫描件极限强压模式）
# ──────────────────────────────────────────────
def compress_image_bytes_for_pdf(img_bytes, quality, max_dim, max_bytes, scan_mode):
    """
    针对 PDF 内嵌入图片的专用压缩。
    如果开启 scan_mode（扫描件模式）：强行将分辨率降至1280px以下，转高对比灰度，质量降低至 55。
    """
    try:
        with Image.open(io.BytesIO(img_bytes)) as img:
            img.load()

            # ── 1. 扫描件模式强行参数重置 ─────────────
            if scan_mode:
                target_dim = 1280 if max_dim == 0 or max_dim > 1280 else max_dim
                target_quality = 55
            else:
                target_dim = max_dim
                target_quality = quality

            is_gray = False
            # 扫描件模式或单色图转灰度
            if scan_mode or img.mode in ('1', 'L', 'P'):
                if not has_transparency(img):
                    img = img.convert('L')
                    # 扫描件增强对比度，使文字更清晰
                    if scan_mode:
                        enhancer = ImageEnhance.Contrast(img)
                        img = enhancer.enhance(1.3)
                    is_gray = True

            # ── 2. 强行降低分辨率（降采样） ───────────
            if target_dim > 0:
                w, h = img.size
                if w > target_dim or h > target_dim:
                    ratio = min(target_dim / w, target_dim / h)
                    img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

            # ── 3. 色彩格式转换 ────────────────────────
            if not is_gray:
                if img.mode == 'RGBA':
                    bg = Image.new('RGB', img.size, (255, 255, 255))
                    bg.paste(img, mask=img.split()[3])
                    img = bg
                elif img.mode != 'RGB':
                    img = img.convert('RGB')

            new_size = img.size

            # ── 4. 降质保存 ────────────────────────────
            cur_quality = target_quality
            while True:
                buf = io.BytesIO()
                img.save(buf, format='JPEG', quality=cur_quality, optimize=True)
                if buf.tell() <= max_bytes or cur_quality <= 20:
                    return buf.getvalue(), new_size, is_gray
                cur_quality -= 5

    except Exception:
        return img_bytes, None, False


def run_compress_pdf(input_path, quality, max_dim, max_mb,
                     force_jpeg, scan_mode, progress_fn, status_fn, log_fn):
    if not FITZ_AVAILABLE:
        log_fn("❌ 未安装 PyMuPDF，无法处理 PDF。")
        status_fn("❌ 缺少 PyMuPDF 库")
        return

    output_path = input_path.parent / (input_path.stem + "_compressed.pdf")
    max_bytes   = int(max_mb * 1024 * 1024)
    orig_total  = input_path.stat().st_size

    log_fn(f"📂 文件：{input_path.name}")
    log_fn(f"📦 原始大小：{orig_total / 1024 / 1024:.2f} MB")
    if scan_mode:
        log_fn("⚡ 模式：【合同/扫描件极限强压模式已开启】（目标尺寸<=1280px，质量=55，灰度强化）\n")
    else:
        log_fn("⚙️  模式：【普通压缩模式】\n")

    try:
        doc = fitz.open(str(input_path))
    except Exception as e:
        log_fn(f"❌ 打开 PDF 失败：{e}")
        return

    total_pages = len(doc)
    log_fn(f"── Step 1 共 {total_pages} 页，开始压缩图片 ──")

    total_orig = total_new = img_count = 0
    processed_xrefs = set()

    for page_num in range(total_pages):
        page     = doc[page_num]
        img_list = page.get_images(full=True)
        status_fn(f"处理第 {page_num + 1}/{total_pages} 页…")
        progress_fn((page_num / total_pages) * 85)

        for img_info in img_list:
            xref = img_info[0]

            if xref in processed_xrefs:
                continue
            processed_xrefs.add(xref)

            try:
                base_image = doc.extract_image(xref)
                img_bytes  = base_image["image"]
                orig_size  = len(img_bytes)

                if orig_size < 5 * 1024:
                    log_fn(f"  ➡  第{page_num+1}页 xref={xref}：{orig_size / 1024:.0f} KB（太小，跳过）")
                    continue

                new_bytes, new_wh, is_gray = compress_image_bytes_for_pdf(
                    img_bytes, quality, max_dim, max_bytes, scan_mode)

                if new_wh is None:
                    continue

                new_file_size = len(new_bytes)

                # 非扫描模式下，体积变大才拦截；扫描模式下强制替换
                if not scan_mode and new_file_size >= orig_size:
                    log_fn(f"  🛑 第{page_num+1}页 xref={xref}：原图已最优（{orig_size/1024:.0f}KB -> 压后{new_file_size/1024:.0f}KB），保留原图！")
                    continue

                total_orig += orig_size
                total_new  += new_file_size
                img_count  += 1

                doc.update_stream(xref, new_bytes, compress=False)
                doc.xref_set_key(xref, "Filter",     "/DCTDecode")
                doc.xref_set_key(xref, "ColorSpace", "/DeviceGray" if is_gray else "/DeviceRGB")
                doc.xref_set_key(xref, "Width",      str(new_wh[0]))
                doc.xref_set_key(xref, "Height",     str(new_wh[1]))
                doc.xref_set_key(xref, "BitsPerComponent", "8")

                ratio = (1 - new_file_size / orig_size) * 100
                log_fn(f"  ✅ 第{page_num+1}页 xref={xref}：{orig_size/1024:.0f} KB → {new_file_size/1024:.0f} KB (↓{ratio:.0f}%)")

            except Exception as e:
                log_fn(f"  ⚠  第{page_num+1}页 xref={xref} 失败：{e}")

    log_fn(f"\n  共成功处理了 {img_count} 张图片")

    status_fn("正在保存 PDF…")
    log_fn("── Step 2 保存并重构 PDF ──────────────")
    progress_fn(90)
    try:
        doc.save(str(output_path), garbage=4, deflate=True)
        doc.close()
    except Exception as e:
        log_fn(f"  ❌ 保存失败：{e}")
        doc.close()
        return

    progress_fn(100)
    final_size = output_path.stat().st_size

    # 全局防变大拦截（仅在非扫描模式触发）
    if not scan_mode and final_size >= orig_total:
        if output_path.exists():
            output_path.unlink()
        log_fn(f"\n{'=' * 52}\n  🛑 提示：原文件已处于极佳压缩状态，已自动取消保存以防文件变大。\n"
               f"  💡 提示：如果是合同/扫描件，请勾选「⚡ 开启合同/扫描件极限强压模式」重新尝试！\n{'=' * 52}")
        status_fn("🛑 原文档已最优，无需压缩")
        return str(input_path), orig_total, orig_total

    saved_pct = (1 - final_size / orig_total) * 100
    log_fn(f"\n{'=' * 52}\n  ✅  处理完成！\n"
           f"  文档大小：{orig_total/1024/1024:.2f} MB  →  "
           f"{final_size/1024/1024:.2f} MB  (节省 {saved_pct:.1f}%)\n"
           f"  输出文件：{output_path}\n{'=' * 52}")
    status_fn(f"✅  完成！节省 {saved_pct:.1f}% 空间")
    return str(output_path), orig_total, final_size


# ──────────────────────────────────────────────
#  可滚动主容器
# ──────────────────────────────────────────────
class ScrollableFrame(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.canvas = tk.Canvas(self, bg="#F5F5F5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.inner = tk.Frame(self.canvas, bg="#F5F5F5")
        self._win_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>",   self._on_mousewheel)
        self.canvas.bind_all("<Button-5>",   self._on_mousewheel)

    def _on_inner_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self._win_id, width=event.width)

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            delta = -1 if event.delta > 0 else 1
            self.canvas.yview_scroll(delta, "units")


# ──────────────────────────────────────────────
#  GUI 界面
# ──────────────────────────────────────────────
class App:
    BG     = "#F5F5F5"
    BLUE   = "#1976D2"
    GREEN  = "#388E3C"
    ORANGE = "#E65100"
    RED    = "#C62828"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.FONT       = (FONT_FAMILY, 10)
        self.FONT_BIG   = (FONT_FAMILY, 13, "bold")
        self.FONT_SMALL = (FONT_FAMILY, 8)
        self.FONT_MED   = (FONT_FAMILY, 9)
        self.FONT_BTN   = (FONT_FAMILY, 11, "bold")

        root.title("文档图片批量压缩工具（DOCX + PDF）")

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        win_w = min(820, sw - 40)
        win_h = min(750, sh - 80)
        x = (sw - win_w) // 2
        y = (sh - win_h) // 2
        root.geometry(f"{win_w}x{win_h}+{x}+{y}")

        root.resizable(True, True)
        root.minsize(640, 480)
        root.configure(bg=self.BG)

        if not PIL_AVAILABLE:
            messagebox.showerror(
                "缺少依赖",
                "请先在命令行执行：\n\n  pip install Pillow PyMuPDF\n\n重新运行程序。"
            )
            root.destroy()
            return

        self._build_ui()

    def _build_ui(self):
        # 顶部标题栏
        header = tk.Frame(self.root, bg=self.BLUE, height=50)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        tk.Label(
            header, text="📄  文档图片批量压缩工具（DOCX + PDF）",
            font=self.FONT_BIG, bg=self.BLUE, fg="white"
        ).pack(side=tk.LEFT, padx=16, pady=12)

        tk.Label(
            header,
            text=f"平台：{platform.system()}  |  PDF：{'✅' if FITZ_AVAILABLE else '❌ 需安装 PyMuPDF'}",
            font=self.FONT_SMALL, bg=self.BLUE, fg="#BBDEFB"
        ).pack(side=tk.RIGHT, padx=16)

        # 底部按钮栏
        btn_bar = tk.Frame(self.root, bg="#EEEEEE", pady=8, padx=18, relief=tk.RIDGE, bd=1)
        btn_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.start_btn = tk.Button(
            btn_bar, text="▶  开始压缩",
            font=self.FONT_BTN, bg=self.GREEN, fg="white",
            bd=0, padx=22, pady=6, cursor="hand2", command=self._start
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(
            btn_bar, text="清除日志", font=self.FONT,
            bd=0, padx=12, pady=6, cursor="hand2", command=self._clear_log
        ).pack(side=tk.LEFT)

        tk.Label(
            btn_bar, text="输出保存在原文件同目录（文件名加 _compressed）",
            font=self.FONT_SMALL, fg="gray", bg="#EEEEEE"
        ).pack(side=tk.RIGHT)

        # 内部滚动区域
        scroll_frame = ScrollableFrame(self.root)
        scroll_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        body = scroll_frame.inner
        pad  = {"padx": 18, "pady": 5}

        # 1. 选择文件
        box1 = tk.LabelFrame(body, text=" 选择文件（支持 .docx 和 .pdf）", font=self.FONT, bg=self.BG, padx=8, pady=8)
        box1.pack(fill=tk.X, **pad)

        self.file_var = tk.StringVar()
        tk.Entry(box1, textvariable=self.file_var, font=self.FONT).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(box1, text="浏览…", font=self.FONT, bg=self.BLUE, fg="white", bd=0, padx=10, pady=3, cursor="hand2", command=self._browse).pack(side=tk.LEFT, padx=(6, 0))

        # 2. 压缩参数
        box2 = tk.LabelFrame(body, text=" 压缩参数 ", font=self.FONT, bg=self.BG, padx=8, pady=8)
        box2.pack(fill=tk.X, **pad)

        # 🔥【关键新增】：合同/扫描件专用极限模式
        row_scan = tk.Frame(box2, bg="#FFF3E0", padx=8, pady=6, bd=1, relief=tk.SOLID)
        row_scan.pack(fill=tk.X, pady=(0, 8))
        
        self.scan_mode_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            row_scan,
            text="⚡ 开启「合同/扫描件」极限强压模式（专治 60KB 左右等死板小图，强制降分辨率与灰度，再省 50%+）",
            variable=self.scan_mode_var,
            font=(FONT_FAMILY, 9, "bold"), bg="#FFF3E0", fg=self.ORANGE,
            selectcolor="#FFF3E0", activebackground="#FFF3E0"
        ).pack(side=tk.LEFT)

        # 图片质量
        row1 = tk.Frame(box2, bg=self.BG)
        row1.pack(fill=tk.X, pady=3)
        tk.Label(row1, text="常规压缩质量：", font=self.FONT, bg=self.BG, width=12, anchor='w').pack(side=tk.LEFT)
        self.quality_var = tk.IntVar(value=85)
        tk.Scale(row1, from_=20, to=95, orient=tk.HORIZONTAL, variable=self.quality_var, length=220, bg=self.BG, highlightthickness=0).pack(side=tk.LEFT)
        tk.Label(row1, textvariable=self.quality_var, font=self.FONT, bg=self.BG, width=3).pack(side=tk.LEFT)
        tk.Label(row1, text="（推荐 85；开启极限强压模式时此项自动调整）", font=self.FONT_SMALL, fg="gray", bg=self.BG).pack(side=tk.LEFT)

        # 最大边长
        row2 = tk.Frame(box2, bg=self.BG)
        row2.pack(fill=tk.X, pady=3)
        tk.Label(row2, text="常规最大边长：", font=self.FONT, bg=self.BG, width=12, anchor='w').pack(side=tk.LEFT)
        self.dim_var = tk.IntVar(value=2480)
        for label, val in [("不限制", 0), ("1920 px", 1920), ("2480 px（A4 300dpi）", 2480), ("3508 px（高清）", 3508)]:
            tk.Radiobutton(row2, text=label, variable=self.dim_var, value=val, font=self.FONT_MED, bg=self.BG).pack(side=tk.LEFT, padx=4)

        # 单图上限
        row3 = tk.Frame(box2, bg=self.BG)
        row3.pack(fill=tk.X, pady=3)
        tk.Label(row3, text="单图上限：", font=self.FONT, bg=self.BG, width=12, anchor='w').pack(side=tk.LEFT)
        self.maxmb_var = tk.DoubleVar(value=1.0)
        for label, val in [("0.1 MB", 0.1), ("0.3 MB", 0.3), ("0.5 MB", 0.5), ("1 MB", 1.0), ("2 MB", 2.0)]:
            tk.Radiobutton(row3, text=label, variable=self.maxmb_var, value=val, font=self.FONT_MED, bg=self.BG).pack(side=tk.LEFT, padx=3)

        # 强制转 JPEG
        row4 = tk.Frame(box2, bg=self.BG)
        row4.pack(fill=tk.X, pady=(6, 2))
        self.force_jpeg_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            row4, text="强制将含透明通道的 PNG 也转换为 JPEG（透明变白底）",
            variable=self.force_jpeg_var, font=self.FONT_MED, bg=self.BG,
            fg=self.ORANGE, selectcolor=self.BG, activebackground=self.BG,
            command=self._on_force_jpeg_toggle
        ).pack(side=tk.LEFT)

        self.force_jpeg_warn = tk.Label(
            box2, text="⚠️  注意：透明区域会变白底，请确认文档中没有依赖透明背景的图章/Logo。",
            font=self.FONT_SMALL, fg="white", bg=self.ORANGE, wraplength=720, justify='left', padx=8, pady=4
        )

        # 进度与状态
        box3 = tk.LabelFrame(body, text=" 进度 ", font=self.FONT, bg=self.BG, padx=8, pady=6)
        box3.pack(fill=tk.X, **pad)

        self.progress = ttk.Progressbar(box3, mode='determinate')
        self.progress.pack(fill=tk.X)
        self.status_var = tk.StringVar(value="等待开始…")
        tk.Label(box3, textvariable=self.status_var, font=self.FONT, bg=self.BG, fg="gray", anchor='w').pack(fill=tk.X, pady=(4, 0))

        # 日志
        box4 = tk.LabelFrame(body, text=" 处理日志 ", font=self.FONT, bg=self.BG, padx=5, pady=5)
        box4.pack(fill=tk.BOTH, expand=True, **pad)

        self.log_box = scrolledtext.ScrolledText(
            box4, height=10,
            font=("Menlo" if platform.system() == "Darwin" else "Consolas", 9),
            state='disabled', bg="#1E1E1E", fg="#D4D4D4", insertbackground="white"
        )
        self.log_box.pack(fill=tk.BOTH, expand=True)

    def _on_force_jpeg_toggle(self):
        if self.force_jpeg_var.get():
            self.force_jpeg_warn.pack(fill=tk.X, pady=(2, 4))
        else:
            self.force_jpeg_warn.pack_forget()

    def _browse(self):
        path = filedialog.askopenfilename(
            title="选择文件",
            filetypes=[("支持的文件", "*.docx *.pdf"), ("Word 文档", "*.docx"), ("PDF 文件", "*.pdf")]
        )
        if path:
            self.file_var.set(path)

    def _log(self, msg):
        self.log_box.config(state='normal')
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state='disabled')

    def _clear_log(self):
        self.log_box.config(state='normal')
        self.log_box.delete("1.0", tk.END)
        self.log_box.config(state='disabled')

    def _set_progress(self, val):
        self.progress['value'] = val
        self.root.update_idletasks()

    def _set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def _start(self):
        path_str = self.file_var.get().strip()
        if not path_str:
            messagebox.showwarning("提示", "请先选择一个文件！")
            return

        path = Path(path_str)
        if not path.exists():
            messagebox.showerror("错误", "文件不存在，请重新选择！")
            return

        suffix = path.suffix.lower()
        if suffix not in {'.docx', '.pdf'}:
            messagebox.showwarning("格式错误", "仅支持 .docx 和 .pdf 文件。")
            return

        if suffix == '.pdf' and not FITZ_AVAILABLE:
            messagebox.showerror("缺少依赖", "处理 PDF 需要安装 PyMuPDF：pip install PyMuPDF")
            return

        self.start_btn.config(state='disabled', text="处理中…")
        self._clear_log()
        self.progress['value'] = 0

        def task():
            try:
                if suffix == '.docx':
                    result = run_compress_docx(
                        path, self.quality_var.get(),
                        self.dim_var.get(), self.maxmb_var.get(),
                        self.force_jpeg_var.get(),
                        self._set_progress, self._set_status, self._log
                    )
                else:
                    result = run_compress_pdf(
                        path, self.quality_var.get(),
                        self.dim_var.get(), self.maxmb_var.get(),
                        self.force_jpeg_var.get(),
                        self.scan_mode_var.get(),  # 传入扫描件模式勾选状态
                        self._set_progress, self._set_status, self._log
                    )
                if result:
                    out_path, orig, final = result
                    saved = (1 - final / orig) * 100 if orig else 0
                    self.root.after(0, lambda: messagebox.showinfo(
                        "🎉 压缩完成",
                        f"处理完成！\n\n原始大小：{orig / 1024 / 1024:.2f} MB\n"
                        f"压缩后：  {final / 1024 / 1024:.2f} MB\n"
                        f"节省：    {saved:.1f}%\n\n输出文件：\n{out_path}"
                    ))
            except Exception as exc:
                self._log(f"\n❌ 发生异常：{exc}")
                self._set_status("❌ 处理失败，请查看日志")
            finally:
                self.root.after(0, lambda: self.start_btn.config(state='normal', text="▶  开始压缩"))

        threading.Thread(target=task, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
