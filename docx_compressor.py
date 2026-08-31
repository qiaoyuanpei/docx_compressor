import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import zipfile
import threading
import tempfile
import io
import platform
from pathlib import Path

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


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
#  核心压缩逻辑
# ──────────────────────────────────────────────
SUPPORTED_EXT = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}


def has_transparency(img: "Image.Image") -> bool:
    """检测图片是否含有透明像素"""
    if img.mode == "RGBA":
        extrema = img.split()[3].getextrema()
        return extrema[0] < 255
    if img.mode == "P":
        return "transparency" in img.info
    return False


def compress_image(img_path: Path, quality: int, max_dim: int,
                   max_bytes: int, force_jpeg: bool, log_fn) -> tuple:
    """
    压缩单张图片
    force_jpeg=True  → 所有图片（含透明 PNG）强制转为 JPEG，透明区域填白色背景
    force_jpeg=False → 含透明通道的 PNG 保持 PNG 格式，不转 JPEG
    返回：(原始字节数, 新字节数, 是否重命名, 新路径)
    """
    ext = img_path.suffix.lower()
    if ext not in SUPPORTED_EXT:
        return 0, 0, False, img_path

    original_size = img_path.stat().st_size

    try:
        with Image.open(img_path) as img:
            img.load()

            # ── 1. 检测透明通道 ──────────────────────
            is_transparent = has_transparency(img)

            # ── 2. 等比缩放 ──────────────────────────
            if max_dim > 0:
                w, h = img.size
                if w > max_dim or h > max_dim:
                    ratio = min(max_dim / w, max_dim / h)
                    img = img.resize(
                        (int(w * ratio), int(h * ratio)),
                        Image.LANCZOS
                    )

            def to_rgb(image):
                """统一转为 RGB，RGBA 则以白底合并透明通道"""
                if image.mode == 'RGBA':
                    bg = Image.new('RGB', image.size, (255, 255, 255))
                    bg.paste(image, mask=image.split()[3])
                    return bg
                if image.mode != 'RGB':
                    return image.convert('RGB')
                return image

            # ── 3. 决定目标格式 ──────────────────────
            if is_transparent and not force_jpeg:
                # 有透明 + 未强制转换 → 保持 PNG
                save_fmt = 'PNG'
                log_fn(f"  🔒 {img_path.name}：含透明通道，保持 PNG 格式")
            elif is_transparent and force_jpeg:
                # 有透明 + 强制转换 → 转 JPEG，白底填充
                save_fmt = 'JPEG'
                log_fn(f"  ⚠️  {img_path.name}：含透明通道，强制转 JPEG（透明区域填白色）")
            elif ext in {'.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp'}:
                save_fmt = 'JPEG'
            else:
                save_fmt = 'PNG'  # 无透明 PNG 先尝试 PNG 压缩

            cur_quality = quality
            cur_img = img
            converted_to_jpeg = False
            new_path = img_path

            # ── 4. 压缩循环 ──────────────────────────
            while True:
                buf = io.BytesIO()

                if save_fmt == 'JPEG':
                    to_rgb(cur_img).save(
                        buf, format='JPEG',
                        quality=cur_quality, optimize=True
                    )
                else:
                    cur_img.save(
                        buf, format='PNG',
                        optimize=True, compress_level=9
                    )

                size_ok = buf.tell() <= max_bytes

                if size_ok or cur_quality <= 20:
                    # 写回文件
                    if (save_fmt == 'JPEG'
                            and ext not in {'.jpg', '.jpeg'}):
                        # 后缀名不同 → 重命名
                        new_path = img_path.with_suffix('.jpg')
                        new_path.write_bytes(buf.getvalue())
                        img_path.unlink()
                        converted_to_jpeg = True
                        log_fn(
                            f"  🔄 {img_path.name} → "
                            f"{new_path.name}（已重命名）"
                        )
                    else:
                        img_path.write_bytes(buf.getvalue())
                    break

                # PNG 超限且（无透明 或 强制转换）→ 转 JPEG
                if save_fmt == 'PNG' and (not is_transparent or force_jpeg):
                    save_fmt = 'JPEG'
                    cur_img = to_rgb(cur_img)
                    log_fn(
                        f"  🔄 {img_path.name}："
                        f"PNG 压缩后仍超限，转换为 JPEG 继续压缩"
                    )
                else:
                    cur_quality -= 5

    except Exception as exc:
        log_fn(f"  ⚠  跳过 {img_path.name}：{exc}")
        return original_size, original_size, False, img_path

    new_size = new_path.stat().st_size
    return original_size, new_size, converted_to_jpeg, new_path


def update_rels_for_renamed(extract_dir: Path,
                             old_name: str, new_name: str, log_fn):
    """
    PNG 改名为 JPG 后，同步更新 Word 内部 .rels 引用，防止图片丢失
    """
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return
    content = rels_path.read_text(encoding='utf-8')
    if old_name in content:
        new_content = content.replace(old_name, new_name)
        rels_path.write_text(new_content, encoding='utf-8')
        log_fn(f"  📝 已更新 .rels 引用：{old_name} → {new_name}")


def run_compress(input_path: Path, quality: int, max_dim: int,
                 max_mb: float, force_jpeg: bool,
                 progress_fn, status_fn, log_fn):
    """
    主流程：解压 docx → 压缩图片 → 重新打包为 docx
    """
    output_path = input_path.parent / (
        input_path.stem + "_compressed.docx"
    )
    max_bytes = int(max_mb * 1024 * 1024)
    orig_total = input_path.stat().st_size

    log_fn(f"📂 文件：{input_path.name}")
    log_fn(f"📦 原始大小：{orig_total / 1024 / 1024:.2f} MB")
    log_fn(f"⚙️  参数：质量={quality}，最大边="
           f"{'不限' if max_dim == 0 else str(max_dim) + 'px'}，"
           f"单图上限={max_mb}MB，"
           f"强制转JPEG={'是 ⚠️' if force_jpeg else '否 🔒'}\n")

    with tempfile.TemporaryDirectory() as tmp:
        extract_dir = Path(tmp) / "doc"

        # ── Step 1：解压 ──────────────────────────────
        status_fn("正在解压文档…")
        log_fn("── Step 1 解压 ──────────────────────────")
        try:
            with zipfile.ZipFile(input_path, 'r') as z:
                z.extractall(extract_dir)
            log_fn("  ✅ 解压成功\n")
        except Exception as e:
            log_fn(f"  ❌ 解压失败：{e}")
            return

        # ── Step 2：查找图片 ──────────────────────────
        media_dir = extract_dir / "word" / "media"
        if not media_dir.exists():
            log_fn("⚠  未找到 word/media 目录，文档中可能没有图片。")
            return

        all_files = [f for f in media_dir.iterdir() if f.is_file()]
        log_fn(f"── Step 2 共发现 {len(all_files)} 个媒体文件 ──────────")

        # ── Step 3：逐张压缩 ──────────────────────────
        log_fn(f"\n── Step 3 开始逐张压缩 ──────────────────")
        total_orig = total_new = 0

        for idx, fpath in enumerate(all_files, 1):
            progress_fn((idx / len(all_files)) * 80)
            status_fn(f"处理 {idx}/{len(all_files)}: {fpath.name}")

            orig, new, renamed, new_path = compress_image(
                fpath, quality, max_dim, max_bytes, force_jpeg, log_fn
            )
            total_orig += orig
            total_new += new

            if renamed:
                update_rels_for_renamed(
                    extract_dir, fpath.name, new_path.name, log_fn
                )

            if orig == 0:
                log_fn(f"  ➡  {fpath.name}（非图片格式，跳过）")
            elif orig == new:
                log_fn(
                    f"  ➡  {fpath.name}："
                    f"{orig / 1024:.0f} KB（无需压缩）"
                )
            else:
                ratio = (1 - new / orig) * 100 if orig else 0
                log_fn(
                    f"  ✅ {fpath.name}: "
                    f"{orig / 1024:.0f} KB → {new / 1024:.0f} KB "
                    f"(↓{ratio:.0f}%)"
                )

        log_fn(
            f"\n  图片合计：{total_orig / 1024 / 1024:.2f} MB → "
            f"{total_new / 1024 / 1024:.2f} MB\n"
        )

        # ── Step 4：重新打包 ──────────────────────────
        status_fn("正在重新打包…")
        log_fn("── Step 4 重新打包为 .docx ──────────────")
        progress_fn(85)
        try:
            with zipfile.ZipFile(
                output_path, 'w', zipfile.ZIP_DEFLATED
            ) as zout:
                for f in extract_dir.rglob('*'):
                    if f.is_file():
                        zout.write(f, f.relative_to(extract_dir))
            log_fn("  ✅ 打包成功")
        except Exception as e:
            log_fn(f"  ❌ 打包失败：{e}")
            return

        progress_fn(100)

        final_size = output_path.stat().st_size
        saved_pct = (
            (1 - final_size / orig_total) * 100 if orig_total else 0
        )

        summary = (
            f"\n{'=' * 52}\n"
            f"  ✅  处理完成！\n"
            f"  文档大小：{orig_total / 1024 / 1024:.2f} MB  →  "
            f"{final_size / 1024 / 1024:.2f} MB  (节省 {saved_pct:.1f}%)\n"
            f"  输出文件：{output_path}\n"
            f"{'=' * 52}"
        )
        log_fn(summary)
        status_fn(f"✅  完成！节省 {saved_pct:.1f}% 空间")
        return str(output_path), orig_total, final_size


# ──────────────────────────────────────────────
#  GUI 界面
# ──────────────────────────────────────────────
class App:
    BG      = "#F5F5F5"
    BLUE    = "#1976D2"
    GREEN   = "#388E3C"
    ORANGE  = "#E65100"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.FONT       = (FONT_FAMILY, 10)
        self.FONT_BIG   = (FONT_FAMILY, 13, "bold")
        self.FONT_SMALL = (FONT_FAMILY, 8)
        self.FONT_MED   = (FONT_FAMILY, 9)
        self.FONT_BTN   = (FONT_FAMILY, 11, "bold")

        root.title("DOCX 图片批量压缩工具")
        root.geometry("780x680")
        root.resizable(False, False)
        root.configure(bg=self.BG)

        if not PIL_AVAILABLE:
            messagebox.showerror(
                "缺少依赖",
                "请先在命令行执行：\n\n"
                "  pip install Pillow\n\n"
                "然后重新运行本程序。"
            )
            root.destroy()
            return

        self._build_ui()

    def _build_ui(self):

        # ── 顶部标题栏 ────────────────────────────────
        header = tk.Frame(self.root, bg=self.BLUE, height=50)
        header.pack(fill=tk.X)
        tk.Label(
            header,
            text="📄  DOCX 图片批量压缩工具",
            font=self.FONT_BIG,
            bg=self.BLUE, fg="white"
        ).pack(side=tk.LEFT, padx=16, pady=12)
        tk.Label(
            header,
            text=f"运行平台：{platform.system()}",
            font=self.FONT_SMALL,
            bg=self.BLUE, fg="#BBDEFB"
        ).pack(side=tk.RIGHT, padx=16)

        body = tk.Frame(self.root, bg=self.BG, padx=18, pady=10)
        body.pack(fill=tk.BOTH, expand=True)

        # ── 文件选择 ──────────────────────────────────
        box1 = tk.LabelFrame(
            body, text=" 选择 .docx 文件 ",
            font=self.FONT, bg=self.BG, padx=8, pady=8
        )
        box1.pack(fill=tk.X, pady=(0, 8))

        self.file_var = tk.StringVar()
        tk.Entry(
            box1, textvariable=self.file_var,
            font=self.FONT, width=64
        ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(
            box1, text="浏览…", font=self.FONT,
            bg=self.BLUE, fg="white", bd=0,
            padx=10, pady=3, cursor="hand2",
            command=self._browse
        ).pack(side=tk.LEFT, padx=(6, 0))

        # ── 压缩参数 ──────────────────────────────────
        box2 = tk.LabelFrame(
            body, text=" 压缩参数 ",
            font=self.FONT, bg=self.BG, padx=8, pady=8
        )
        box2.pack(fill=tk.X, pady=(0, 8))

        # 图片质量
        row1 = tk.Frame(box2, bg=self.BG)
        row1.pack(fill=tk.X, pady=3)
        tk.Label(
            row1, text="压缩质量：",
            font=self.FONT, bg=self.BG, width=12, anchor='w'
        ).pack(side=tk.LEFT)
        self.quality_var = tk.IntVar(value=85)
        tk.Scale(
            row1, from_=20, to=95, orient=tk.HORIZONTAL,
            variable=self.quality_var, length=240,
            bg=self.BG, highlightthickness=0
        ).pack(side=tk.LEFT)
        tk.Label(
            row1, textvariable=self.quality_var,
            font=self.FONT, bg=self.BG, width=3
        ).pack(side=tk.LEFT)
        tk.Label(
            row1,
            text="（推荐 85；值越小文件越小，清晰度越低）",
            font=self.FONT_SMALL, fg="gray", bg=self.BG
        ).pack(side=tk.LEFT)

        # 最大边长
        row2 = tk.Frame(box2, bg=self.BG)
        row2.pack(fill=tk.X, pady=3)
        tk.Label(
            row2, text="最大边长：",
            font=self.FONT, bg=self.BG, width=12, anchor='w'
        ).pack(side=tk.LEFT)
        self.dim_var = tk.IntVar(value=2480)
        for label, val in [
            ("不限制", 0),
            ("1920 px", 1920),
            ("2480 px（A4 300dpi）", 2480),
            ("3508 px（A4 高清）", 3508)
        ]:
            tk.Radiobutton(
                row2, text=label, variable=self.dim_var,
                value=val, font=self.FONT_MED, bg=self.BG
            ).pack(side=tk.LEFT, padx=4)

        # 单图大小上限（新增 0.3 和 0.4）
        row3 = tk.Frame(box2, bg=self.BG)
        row3.pack(fill=tk.X, pady=3)
        tk.Label(
            row3, text="单图上限：",
            font=self.FONT, bg=self.BG, width=12, anchor='w'
        ).pack(side=tk.LEFT)
        self.maxmb_var = tk.DoubleVar(value=1.0)
        for label, val in [
            ("0.3 MB", 0.3),
            ("0.4 MB", 0.4),
            ("0.5 MB", 0.5),
            ("1 MB",   1.0),
            ("2 MB",   2.0),
            ("5 MB",   5.0),
        ]:
            tk.Radiobutton(
                row3, text=label, variable=self.maxmb_var,
                value=val, font=self.FONT_MED, bg=self.BG
            ).pack(side=tk.LEFT, padx=3)
        tk.Label(
            row3,
            text="（超限自动降质直到达标）",
            font=self.FONT_SMALL, fg="gray", bg=self.BG
        ).pack(side=tk.LEFT, padx=4)

        # ── 新增：透明 PNG 强制转 JPEG 选项 ──────────
        row4 = tk.Frame(box2, bg=self.BG)
        row4.pack(fill=tk.X, pady=(6, 2))

        self.force_jpeg_var = tk.BooleanVar(value=False)
        force_cb = tk.Checkbutton(
            row4,
            text="强制将含透明通道的 PNG 也转换为 JPEG（透明区域将填充为白色背景）",
            variable=self.force_jpeg_var,
            font=self.FONT_MED,
            bg=self.BG,
            fg=self.ORANGE,
            selectcolor=self.BG,
            activebackground=self.BG,
            command=self._on_force_jpeg_toggle
        )
        force_cb.pack(side=tk.LEFT)

        # 警告提示文字（默认隐藏）
        self.force_jpeg_warn = tk.Label(
            box2,
            text="⚠️  注意：勾选后，Logo、图章等透明图片的透明区域会变成白色方块，"
                 "请确认你的文档中没有依赖透明背景的图片，或你不介意此效果。",
            font=self.FONT_SMALL,
            fg="white", bg=self.ORANGE,
            wraplength=710, justify='left',
            padx=8, pady=4
        )
        # 初始隐藏，勾选后才显示

        # ── 进度 ──────────────────────────────────────
        box3 = tk.LabelFrame(
            body, text=" 进度 ",
            font=self.FONT, bg=self.BG, padx=8, pady=6
        )
        box3.pack(fill=tk.X, pady=(0, 8))

        self.progress = ttk.Progressbar(
            box3, length=720, mode='determinate'
        )
        self.progress.pack(fill=tk.X)
        self.status_var = tk.StringVar(value="等待开始…")
        tk.Label(
            box3, textvariable=self.status_var,
            font=self.FONT, bg=self.BG, fg="gray", anchor='w'
        ).pack(fill=tk.X, pady=(4, 0))

        # ── 日志 ──────────────────────────────────────
        box4 = tk.LabelFrame(
            body, text=" 处理日志 ",
            font=self.FONT, bg=self.BG, padx=5, pady=5
        )
        box4.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        self.log_box = scrolledtext.ScrolledText(
            box4, height=10,
            font=(
                "Menlo" if platform.system() == "Darwin"
                else "Consolas", 9
            ),
            state='disabled',
            bg="#1E1E1E", fg="#D4D4D4",
            insertbackground="white"
        )
        self.log_box.pack(fill=tk.BOTH, expand=True)

        # ── 底部按钮 ──────────────────────────────────
        btn_row = tk.Frame(body, bg=self.BG)
        btn_row.pack(fill=tk.X)

        self.start_btn = tk.Button(
            btn_row, text="▶  开始压缩",
            font=self.FONT_BTN,
            bg=self.GREEN, fg="white",
            bd=0, padx=22, pady=6,
            cursor="hand2",
            command=self._start
        )
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(
            btn_row, text="清除日志",
            font=self.FONT, bd=0,
            padx=12, pady=6,
            cursor="hand2",
            command=self._clear_log
        ).pack(side=tk.LEFT)

        tk.Label(
            btn_row,
            text="输出文件保存在原文件同目录（文件名加 _compressed）",
            font=self.FONT_SMALL, fg="gray", bg=self.BG
        ).pack(side=tk.RIGHT)

    # ── 勾选「强制转 JPEG」时显示/隐藏警告 ──────────
    def _on_force_jpeg_toggle(self):
        if self.force_jpeg_var.get():
            self.force_jpeg_warn.pack(fill=tk.X, pady=(2, 4), padx=0)
        else:
            self.force_jpeg_warn.pack_forget()

    # ── 事件处理 ──────────────────────────────────────
    def _browse(self):
        path = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if path:
            self.file_var.set(path)

    def _log(self, msg: str):
        self.log_box.config(state='normal')
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state='disabled')

    def _clear_log(self):
        self.log_box.config(state='normal')
        self.log_box.delete("1.0", tk.END)
        self.log_box.config(state='disabled')

    def _set_progress(self, val: float):
        self.progress['value'] = val
        self.root.update_idletasks()

    def _set_status(self, msg: str):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def _start(self):
        path_str = self.file_var.get().strip()
        if not path_str:
            messagebox.showwarning("提示", "请先选择一个 .docx 文件！")
            return

        path = Path(path_str)
        if not path.exists():
            messagebox.showerror("错误", "文件不存在，请重新选择！")
            return
        if path.suffix.lower() != '.docx':
            messagebox.showwarning(
                "格式错误",
                "请选择 .docx 格式的文件。\n\n"
                "如果你的文件是 .doc 格式，\n"
                "请先用 Word 另存为 .docx 后再使用本工具。"
            )
            return

        # 勾选了强制转 JPEG 时，再次弹窗确认
        if self.force_jpeg_var.get():
            confirm = messagebox.askyesno(
                "⚠️  请确认",
                "你已勾选「强制将透明 PNG 转为 JPEG」。\n\n"
                "这会导致 Logo、图章等含透明背景的图片\n"
                "透明区域变成白色方块。\n\n"
                "确定继续吗？"
            )
            if not confirm:
                return

        self.start_btn.config(state='disabled', text="处理中…")
        self._clear_log()
        self.progress['value'] = 0

        def task():
            try:
                result = run_compress(
                    input_path=path,
                    quality=self.quality_var.get(),
                    max_dim=self.dim_var.get(),
                    max_mb=self.maxmb_var.get(),
                    force_jpeg=self.force_jpeg_var.get(),
                    progress_fn=self._set_progress,
                    status_fn=self._set_status,
                    log_fn=self._log,
                )
                if result:
                    out_path, orig, final = result
                    saved = (1 - final / orig) * 100 if orig else 0
                    self.root.after(0, lambda: messagebox.showinfo(
                        "🎉 压缩完成",
                        f"压缩完成！\n\n"
                        f"原始大小：{orig  / 1024 / 1024:.2f} MB\n"
                        f"压缩后：  {final / 1024 / 1024:.2f} MB\n"
                        f"节省：    {saved:.1f}%\n\n"
                        f"输出文件：\n{out_path}"
                    ))
            except Exception as exc:
                self._log(f"\n❌ 发生异常：{exc}")
                self._set_status("❌ 处理失败，请查看日志")
            finally:
                self.root.after(0, lambda: self.start_btn.config(
                    state='normal', text="▶  开始压缩"
                ))

        threading.Thread(target=task, daemon=True).start()


# ──────────────────────────────────────────────
#  程序入口
# ──────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
