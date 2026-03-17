import os
import sys
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime


class DocToDocxConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word文档批量转换工具 (.doc → .docx)")
        self.root.geometry("800x600")

        # 尝试设置窗口图标
        try:
            self.root.iconbitmap(default='converter.ico')
        except:
            pass

        # 初始化变量
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.doc_files = []
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="准备就绪")
        self.conversion_stats = {"total": 0, "success": 0, "failed": 0, "skipped": 0}

        # 设置默认路径为桌面
        desktop = Path.home() / "Desktop"
        self.input_folder.set(str(desktop))
        self.output_folder.set(str(desktop / "Converted_Docx"))

        # 检查是否安装了必要的库
        self.check_dependencies()

        self.create_widgets()

    def check_dependencies(self):
        """检查依赖库"""
        self.win32_available = False
        try:
            from win32com import client as win32
            import pythoncom
            self.win32_available = True
        except ImportError as e:
            print(f"缺少依赖库: {e}")

    def create_widgets(self):
        """创建UI组件"""
        # 标题
        title_label = ttk.Label(
            self.root,
            text="Word文档批量转换工具 (.doc → .docx)",
            font=("微软雅黑", 14, "bold")
        )
        title_label.pack(pady=20)

        # 依赖状态显示
        self.create_dependency_section()

        # 输入输出路径设置
        self.create_path_sections()

        # 文件列表区域
        self.create_file_list_section()

        # 选项区域
        self.create_options_section()

        # 进度条
        self.create_progress_section()

        # 状态栏
        self.create_status_bar()

        # 按钮区域
        self.create_button_section()

    def create_dependency_section(self):
        """创建依赖库状态显示"""
        frame = ttk.Frame(self.root)
        frame.pack(fill="x", padx=20, pady=5)

        if self.win32_available:
            status_text = "✓ 依赖库已安装，可以开始转换"
            color = "green"
        else:
            status_text = "✗ 需要安装pywin32库才能使用此工具"
            color = "red"

        status_label = ttk.Label(
            frame,
            text=status_text,
            font=("微软雅黑", 9),
            foreground=color
        )
        status_label.pack()

        if not self.win32_available:
            install_btn = ttk.Button(
                frame,
                text="查看安装说明",
                command=self.show_install_instructions,
                width=15
            )
            install_btn.pack(pady=5)

    def create_path_sections(self):
        """创建输入输出路径区域"""
        # 输入路径
        input_frame = ttk.LabelFrame(self.root, text="输入文件夹 (包含.doc文件)", padding=10)
        input_frame.pack(fill="x", padx=20, pady=10)

        input_path_frame = ttk.Frame(input_frame)
        input_path_frame.pack(fill="x")

        ttk.Label(input_path_frame, text="输入路径:").pack(side="left", padx=5)

        self.input_entry = ttk.Entry(input_path_frame, textvariable=self.input_folder, width=50)
        self.input_entry.pack(side="left", padx=5, fill="x", expand=True)

        input_browse_btn = ttk.Button(
            input_path_frame,
            text="浏览",
            command=lambda: self.browse_folder(self.input_folder),
            width=10
        )
        input_browse_btn.pack(side="left", padx=5)

        # 输出路径
        output_frame = ttk.LabelFrame(self.root, text="输出文件夹 (.docx文件保存位置)", padding=10)
        output_frame.pack(fill="x", padx=20, pady=10)

        output_path_frame = ttk.Frame(output_frame)
        output_path_frame.pack(fill="x")

        ttk.Label(output_path_frame, text="输出路径:").pack(side="left", padx=5)

        self.output_entry = ttk.Entry(output_path_frame, textvariable=self.output_folder, width=50)
        self.output_entry.pack(side="left", padx=5, fill="x", expand=True)

        output_browse_btn = ttk.Button(
            output_path_frame,
            text="浏览",
            command=lambda: self.browse_folder(self.output_folder),
            width=10
        )
        output_browse_btn.pack(side="left", padx=5)

    def create_file_list_section(self):
        """创建文件列表区域"""
        frame = ttk.LabelFrame(self.root, text="待转换文件列表", padding=10)
        frame.pack(fill="both", expand=True, padx=20, pady=10)

        # 创建树状列表
        columns = ("序号", "文件名", "文件大小", "状态")
        self.file_tree = ttk.Treeview(
            frame,
            columns=columns,
            show="headings",
            height=10
        )

        # 设置列标题和宽度
        col_widths = [50, 250, 100, 100]
        for idx, col in enumerate(columns):
            self.file_tree.heading(col, text=col)
            self.file_tree.column(col, width=col_widths[idx])

        # 添加滚动条
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        # 布局
        self.file_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def create_options_section(self):
        """创建选项区域"""
        frame = ttk.LabelFrame(self.root, text="转换选项", padding=10)
        frame.pack(fill="x", padx=20, pady=10)

        options_frame = ttk.Frame(frame)
        options_frame.pack(fill="x")

        # 是否保留原始文件结构
        self.preserve_structure = tk.BooleanVar(value=True)
        preserve_check = ttk.Checkbutton(
            options_frame,
            text="保留原始文件夹结构",
            variable=self.preserve_structure
        )
        preserve_check.pack(side="left", padx=10)

        # 是否覆盖已存在文件
        self.overwrite_existing = tk.BooleanVar(value=False)
        overwrite_check = ttk.Checkbutton(
            options_frame,
            text="覆盖已存在的文件",
            variable=self.overwrite_existing
        )
        overwrite_check.pack(side="left", padx=10)

        # 递归处理子文件夹
        self.include_subfolders = tk.BooleanVar(value=False)
        subfolder_check = ttk.Checkbutton(
            options_frame,
            text="包含子文件夹",
            variable=self.include_subfolders
        )
        subfolder_check.pack(side="left", padx=10)

        # 扫描按钮
        scan_btn = ttk.Button(
            frame,
            text="扫描.doc文件",
            command=self.scan_doc_files,
            width=15
        )
        scan_btn.pack(pady=5)

    def create_progress_section(self):
        """创建进度条区域"""
        frame = ttk.Frame(self.root)
        frame.pack(fill="x", padx=20, pady=5)

        self.progress_bar = ttk.Progressbar(
            frame,
            variable=self.progress_var,
            maximum=100,
            length=400,
            mode='determinate'
        )
        self.progress_bar.pack(fill="x", padx=5)

    def create_status_bar(self):
        """创建状态栏"""
        frame = ttk.Frame(self.root)
        frame.pack(fill="x", padx=20, pady=5)

        self.status_label = ttk.Label(
            frame,
            textvariable=self.status_var,
            font=("微软雅黑", 9),
            foreground="blue"
        )
        self.status_label.pack()

        # 统计信息
        self.stats_label = ttk.Label(
            frame,
            text="",
            font=("微软雅黑", 8),
            foreground="gray"
        )
        self.stats_label.pack()

    def create_button_section(self):
        """创建按钮区域"""
        frame = ttk.Frame(self.root)
        frame.pack(pady=20)

        # 开始转换按钮
        self.convert_btn = ttk.Button(
            frame,
            text="开始转换",
            command=self.start_conversion,
            width=20,
            state="normal" if self.win32_available else "disabled"
        )
        self.convert_btn.pack(pady=5)

        # 清空列表按钮
        clear_btn = ttk.Button(
            frame,
            text="清空列表",
            command=self.clear_file_list,
            width=20
        )
        clear_btn.pack(pady=5)

    def browse_folder(self, path_var):
        """浏览文件夹"""
        folder = filedialog.askdirectory(
            title="选择文件夹",
            initialdir=path_var.get()
        )
        if folder:
            path_var.set(folder)

    def show_install_instructions(self):
        """显示安装说明"""
        instructions = """
        需要安装 pywin32 库才能使用此工具。

        安装方法：

        方法1：使用 pip 安装
        pip install pywin32

        方法2：如果方法1失败，可以尝试
        python -m pip install pywin32

        方法3：从官网下载安装
        https://github.com/mhammond/pywin32

        注意：此工具仅适用于 Windows 系统。

        安装完成后请重启此程序。
        """

        messagebox.showinfo("安装说明", instructions)

    def get_file_size(self, filepath):
        """获取文件大小并格式化显示"""
        try:
            size = os.path.getsize(filepath)
            # 转换为KB或MB
            if size < 1024:
                return f"{size} B"
            elif size < 1024 * 1024:
                return f"{size / 1024:.1f} KB"
            else:
                return f"{size / (1024 * 1024):.1f} MB"
        except:
            return "未知"

    def scan_doc_files(self):
        """扫描.doc文件"""
        input_dir = self.input_folder.get()

        if not input_dir or not os.path.exists(input_dir):
            messagebox.showwarning("警告", "请输入有效的输入文件夹路径！")
            return

        # 清空现有列表
        self.doc_files = []
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # 扫描文件
        try:
            if self.include_subfolders.get():
                # 递归扫描
                pattern = "**/*.doc"
            else:
                # 仅扫描当前目录
                pattern = "*.doc"

            input_path = Path(input_dir)
            doc_files = list(input_path.glob(pattern))

            if not doc_files:
                messagebox.showinfo("提示", f"在 {input_dir} 中未找到 .doc 文件")
                return

            # 添加到列表
            for i, file_path in enumerate(doc_files):
                self.doc_files.append(str(file_path))
                file_size = self.get_file_size(file_path)
                self.file_tree.insert(
                    "",
                    "end",
                    values=(i + 1, file_path.name, file_size, "等待转换")
                )

            # 更新状态
            self.status_var.set(f"找到 {len(doc_files)} 个 .doc 文件")

        except Exception as e:
            messagebox.showerror("错误", f"扫描文件时出错:\n{str(e)}")

    def clear_file_list(self):
        """清空文件列表"""
        self.doc_files = []
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        self.status_var.set("文件列表已清空")
        self.stats_label.config(text="")
        self.conversion_stats = {"total": 0, "success": 0, "failed": 0, "skipped": 0}

    def update_file_status(self, file_index, status):
        """更新文件处理状态"""
        for item in self.file_tree.get_children():
            values = self.file_tree.item(item, 'values')
            if values and int(values[0]) == file_index:
                self.file_tree.set(item, '状态', status)
                # 根据状态设置颜色
                if status == "转换成功":
                    self.file_tree.tag_configure('success', foreground='green')
                    self.file_tree.item(item, tags=('success',))
                elif status == "转换失败":
                    self.file_tree.tag_configure('failed', foreground='red')
                    self.file_tree.item(item, tags=('failed',))
                elif status == "已跳过":
                    self.file_tree.tag_configure('skipped', foreground='orange')
                    self.file_tree.item(item, tags=('skipped',))
                break

    def update_progress(self, value, status):
        """更新进度条和状态"""
        self.progress_var.set(value)
        self.status_var.set(status)
        self.root.update_idletasks()

    def update_stats(self):
        """更新统计信息"""
        stats_text = f"总计: {self.conversion_stats['total']} | "
        stats_text += f"成功: {self.conversion_stats['success']} | "
        stats_text += f"失败: {self.conversion_stats['failed']} | "
        stats_text += f"跳过: {self.conversion_stats['skipped']}"
        self.stats_label.config(text=stats_text)

    def convert_doc_to_docx(self, doc_path, docx_path):
        """将单个.doc文件转换为.docx"""
        try:
            from win32com import client as win32
            import pythoncom

            # 初始化COM（重要：在每个线程中都需要初始化）
            pythoncom.CoInitialize()

            # 确保输出目录存在
            os.makedirs(os.path.dirname(docx_path), exist_ok=True)

            # 打开Word应用程序
            word = win32.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False

            try:
                # 打开.doc文件
                doc = word.Documents.Open(doc_path)

                # 保存为.docx格式（16表示docx格式）
                doc.SaveAs(docx_path, FileFormat=16)

                # 关闭文档
                doc.Close()

                return True, "转换成功"

            except Exception as e:
                return False, f"转换失败: {str(e)}"

            finally:
                # 关闭Word应用程序
                word.Quit()
                pythoncom.CoUninitialize()

        except Exception as e:
            return False, f"转换失败: {str(e)}"

    def conversion_thread(self):
        """转换线程"""
        try:
            if not self.doc_files:
                messagebox.showwarning("警告", "没有要转换的文件！请先扫描文件。")
                self.convert_btn.config(state="normal")
                return

            # 重置统计
            self.conversion_stats = {"total": len(self.doc_files), "success": 0, "failed": 0, "skipped": 0}
            self.update_stats()

            # 检查输出目录是否存在
            output_dir = self.output_folder.get()
            if not output_dir:
                messagebox.showwarning("警告", "请选择输出文件夹！")
                self.convert_btn.config(state="normal")
                return

            # 创建输出目录（如果不存在）
            os.makedirs(output_dir, exist_ok=True)

            # 开始转换
            for i, doc_path in enumerate(self.doc_files):
                # 更新进度
                progress = (i / len(self.doc_files)) * 100
                self.update_progress(
                    progress,
                    f"正在转换: {os.path.basename(doc_path)} ({i + 1}/{len(self.doc_files)})"
                )

                # 确定输出路径
                doc_path_obj = Path(doc_path)
                input_dir = Path(self.input_folder.get())

                if self.preserve_structure.get() and self.include_subfolders.get():
                    # 保持原始文件夹结构
                    rel_path = doc_path_obj.relative_to(input_dir)
                    docx_path = Path(output_dir) / rel_path.with_suffix('.docx')
                else:
                    # 所有文件放在同一目录
                    docx_path = Path(output_dir) / doc_path_obj.with_suffix('.docx').name

                # 检查文件是否已存在
                if docx_path.exists() and not self.overwrite_existing.get():
                    self.update_file_status(i + 1, "已跳过")
                    self.conversion_stats["skipped"] += 1
                    self.update_stats()
                    continue

                # 更新文件状态
                self.update_file_status(i + 1, "转换中...")

                # 执行转换
                success, message = self.convert_doc_to_docx(str(doc_path), str(docx_path))

                if success:
                    self.update_file_status(i + 1, "转换成功")
                    self.conversion_stats["success"] += 1
                else:
                    self.update_file_status(i + 1, "转换失败")
                    self.conversion_stats["failed"] += 1
                    # 记录错误日志
                    with open(Path(output_dir) / "conversion_errors.log", "a", encoding="utf-8") as f:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        f.write(f"[{timestamp}] {doc_path} -> {message}\n")

                self.update_stats()

            # 完成
            self.update_progress(100, "转换完成！")

            # 显示结果
            result_message = f"""
转换完成！

总计: {self.conversion_stats['total']} 个文件
成功: {self.conversion_stats['success']} 个
失败: {self.conversion_stats['failed']} 个
跳过: {self.conversion_stats['skipped']} 个

输出目录: {output_dir}
            """

            if self.conversion_stats['failed'] > 0:
                result_message += "\n\n注意：有文件转换失败，请查看 conversion_errors.log 文件了解详情。"

            messagebox.showinfo("转换完成", result_message)

            # 启用按钮
            self.convert_btn.config(state="normal")

            # 询问是否打开输出文件夹
            if messagebox.askyesno("完成", "是否打开输出文件夹？"):
                try:
                    os.startfile(output_dir)
                except:
                    # 对于非Windows系统
                    import subprocess
                    try:
                        subprocess.run(['open', output_dir])
                    except:
                        pass

        except Exception as e:
            messagebox.showerror("错误", f"转换过程中出错:\n{str(e)}\n\n{traceback.format_exc()}")
            self.update_progress(0, "转换失败")
            self.convert_btn.config(state="normal")

    def start_conversion(self):
        """开始转换"""
        if not self.win32_available:
            messagebox.showerror("错误", "需要安装 pywin32 库才能使用此工具！")
            return

        if not self.doc_files:
            messagebox.showwarning("警告", "没有要转换的文件！请先扫描文件。")
            return

        # 确认开始转换
        confirm = messagebox.askyesno("确认", f"确定要开始转换 {len(self.doc_files)} 个文件吗？")
        if not confirm:
            return

        # 禁用按钮防止重复点击
        self.convert_btn.config(state="disabled")

        # 重置进度
        self.update_progress(0, "开始转换...")

        # 在新线程中执行转换
        thread = threading.Thread(target=self.conversion_thread)
        thread.daemon = True
        thread.start()


def main():
    root = tk.Tk()
    app = DocToDocxConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()