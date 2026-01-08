"""
Excel账单合并工具 - 主程序
支持批量导入多个Excel账单，根据预设配置自动提取和汇总数据
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
from pathlib import Path
from config_manager import ConfigManager
from excel_processor import ExcelProcessor
from config_editor import ConfigEditor


class MergeBillApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel账单合并工具")
        self.root.geometry("700x500")
        self.root.minsize(600, 400)
        
        # 初始化配置管理器
        self.config_manager = ConfigManager()
        self.excel_processor = ExcelProcessor()
        
        # 存储拖入的文件
        self.file_list = []
        
        self.setup_ui()
        
        # 尝试支持拖拽（如果可用）
        self.setup_drag_drop()
        
    def setup_drag_drop(self):
        """尝试设置拖拽功能（可选）"""
        try:
            from tkinterdnd2 import DND_FILES
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        except:
            # 如果tkinterdnd2不可用，仅使用按钮方式
            pass
        
    def setup_ui(self):
        """设置用户界面"""
        # 顶部工具栏
        toolbar = ttk.Frame(self.root)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        # 预设选择
        ttk.Label(toolbar, text="选择预设：").pack(side=tk.LEFT, padx=5)
        
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(
            toolbar, 
            textvariable=self.preset_var, 
            state='readonly',
            width=30
        )
        self.preset_combo.pack(side=tk.LEFT, padx=5)
        self.update_preset_list()
        
        # 管理预设按钮
        ttk.Button(
            toolbar, 
            text="管理预设", 
            command=self.open_config_editor
        ).pack(side=tk.LEFT, padx=5)
        
        # 中间文件区域
        file_frame = ttk.LabelFrame(self.root, text="文件列表", padding=10)
        file_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 文件列表框和滚动条
        list_container = ttk.Frame(file_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            list_container,
            yscrollcommand=scrollbar.set,
            selectmode=tk.EXTENDED,
            font=("微软雅黑", 9)
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # 用于拖拽的引用
        self.drop_frame = file_frame
        
        # 文件统计标签
        self.file_count_label = ttk.Label(
            file_frame,
            text="已添加 0 个文件",
            foreground="blue"
        )
        self.file_count_label.pack(pady=5)
        
        # 底部按钮区
        button_frame = ttk.Frame(self.root)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            button_frame,
            text="添加文件",
            command=self.browse_files,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="添加文件夹",
            command=self.browse_folder,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="移除选中",
            command=self.remove_selected_files,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="清空列表",
            command=self.clear_files,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        # 右侧开始按钮
        start_button = ttk.Button(
            button_frame,
            text="开始合并",
            command=self.start_merge,
            width=20
        )
        start_button.pack(side=tk.RIGHT, padx=5)
        
        # 设置按钮样式（绿色突出显示）
        style = ttk.Style()
        try:
            style.configure('Start.TButton', background='#4CAF50', foreground='green')
            start_button.configure(style='Start.TButton')
        except:
            pass
        
    def update_preset_list(self):
        """更新预设列表"""
        presets = self.config_manager.get_preset_names()
        self.preset_combo['values'] = presets
        if presets:
            self.preset_combo.current(0)
    
    def on_drop(self, event):
        """处理文件拖拽事件"""
        files = self.parse_drop_files(event.data)
        self.add_files(files)
    
    def parse_drop_files(self, data):
        """解析拖拽的文件路径"""
        # 处理不同格式的文件路径
        files = []
        data = data.strip()
        
        # Windows路径可能用{}包裹
        if data.startswith('{') and data.endswith('}'):
            data = data[1:-1]
        
        # 分割多个文件
        import re
        # 匹配被{}包裹的路径或空格分隔的路径
        pattern = r'\{([^}]+)\}|([^\s]+)'
        matches = re.findall(pattern, data)
        
        for match in matches:
            file_path = match[0] if match[0] else match[1]
            if file_path and os.path.isfile(file_path):
                # 只接受Excel文件
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    files.append(file_path)
        
        return files
    
    def add_files(self, files):
        """添加文件到列表"""
        added_count = 0
        for file_path in files:
            if file_path not in self.file_list:
                self.file_list.append(file_path)
                self.file_listbox.insert(tk.END, os.path.basename(file_path))
                added_count += 1
        
        if added_count > 0:
            self.update_file_count()
            
    def browse_files(self):
        """浏览选择文件"""
        files = filedialog.askopenfilenames(
            title="选择Excel文件（可多选）",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if files:
            self.add_files(list(files))
    
    def browse_folder(self):
        """浏览选择文件夹，自动添加所有Excel文件"""
        folder = filedialog.askdirectory(title="选择包含Excel文件的文件夹")
        if folder:
            excel_files = []
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith(('.xlsx', '.xls')):
                        excel_files.append(os.path.join(root, file))
            
            if excel_files:
                self.add_files(excel_files)
                messagebox.showinfo("提示", f"从文件夹中找到 {len(excel_files)} 个Excel文件")
            else:
                messagebox.showwarning("提示", "文件夹中没有找到Excel文件")
    
    def remove_selected_files(self):
        """移除选中的文件"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("提示", "请先选择要移除的文件")
            return
        
        # 从后往前删除，避免索引变化
        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            del self.file_list[index]
        
        self.update_file_count()
    
    def clear_files(self):
        """清空文件列表"""
        self.file_list.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_file_count()
    
    def update_file_count(self):
        """更新文件计数显示"""
        count = len(self.file_list)
        self.file_count_label.config(text=f"已添加 {count} 个文件")
    
    def start_merge(self):
        """开始合并操作"""
        # 检查是否有文件
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加要合并的Excel文件！")
            return
        
        # 检查是否选择预设
        preset_name = self.preset_var.get()
        if not preset_name:
            messagebox.showwarning("提示", "请先选择一个预设配置！")
            return
        
        # 获取预设配置
        preset = self.config_manager.get_preset(preset_name)
        if not preset or not preset.get('mappings'):
            messagebox.showwarning("提示", "所选预设没有配置映射项目！")
            return
        
        # 选择输出文件
        output_file = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        
        if not output_file:
            return
        
        try:
            # 显示处理进度
            progress_window = tk.Toplevel(self.root)
            progress_window.title("处理中...")
            progress_window.geometry("350x120")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # 居中显示
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (progress_window.winfo_width() // 2)
            y = (progress_window.winfo_screenheight() // 2) - (progress_window.winfo_height() // 2)
            progress_window.geometry(f"+{x}+{y}")
            
            ttk.Label(
                progress_window, 
                text=f"正在处理 {len(self.file_list)} 个文件，请稍候...",
                font=("微软雅黑", 10)
            ).pack(pady=20)
            
            progress_bar = ttk.Progressbar(
                progress_window, 
                mode='indeterminate',
                length=300
            )
            progress_bar.pack(fill=tk.X, padx=20, pady=10)
            progress_bar.start()
            
            self.root.update()
            
            # 执行合并
            result = self.excel_processor.merge_bills(
                self.file_list,
                preset['mappings'],
                output_file,
                preset.get('settlement_search_column', 'D'),
                preset.get('settlement_search_keyword', '折后总计')
            )
            
            progress_window.destroy()
            
            # 显示结果
            if result['success']:
                messagebox.showinfo(
                    "成功",
                    f"合并完成！\n\n"
                    f"✓ 成功处理: {result['success_count']} 个文件\n"
                    f"✗ 失败: {result['error_count']} 个文件\n\n"
                    f"结果已保存到:\n{output_file}"
                )
                # 清空列表
                self.clear_files()
            else:
                messagebox.showerror("错误", f"合并失败：{result['message']}")
                
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            messagebox.showerror("错误", f"处理过程中出现错误：{str(e)}")
    
    def open_config_editor(self):
        """打开配置编辑器"""
        editor = ConfigEditor(self.root, self.config_manager)
        self.root.wait_window(editor.window)
        # 刷新预设列表
        self.update_preset_list()


def main():
    # 尝试使用TkinterDnD，如果不可用则使用标准Tk
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except:
        root = tk.Tk()
    
    app = MergeBillApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
