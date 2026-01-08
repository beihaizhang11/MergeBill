"""
配置编辑器 - 用于管理预设和映射配置的界面
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os


class ConfigEditor:
    def __init__(self, parent, config_manager):
        self.parent = parent
        self.config_manager = config_manager
        
        # 创建新窗口
        self.window = tk.Toplevel(parent)
        self.window.title("预设配置管理")
        self.window.geometry("900x600")
        self.window.minsize(800, 500)
        
        # 模态窗口
        self.window.transient(parent)
        self.window.grab_set()
        
        self.current_preset = None
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.PanedWindow(self.window, orient=tk.HORIZONTAL)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左侧：预设列表
        left_frame = ttk.Frame(main_frame)
        main_frame.add(left_frame, weight=1)
        
        ttk.Label(left_frame, text="预设列表", font=("微软雅黑", 10, "bold")).pack(
            anchor=tk.W, padx=5, pady=5
        )
        
        # 预设列表框
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.preset_listbox = tk.Listbox(
            list_frame, 
            yscrollcommand=scrollbar.set,
            font=("微软雅黑", 9)
        )
        self.preset_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.preset_listbox.yview)
        
        self.preset_listbox.bind('<<ListboxSelect>>', self.on_preset_select)
        
        # 预设操作按钮
        preset_btn_frame = ttk.Frame(left_frame)
        preset_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(
            preset_btn_frame,
            text="新建",
            command=self.new_preset
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            preset_btn_frame,
            text="复制",
            command=self.duplicate_preset
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            preset_btn_frame,
            text="重命名",
            command=self.rename_preset
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            preset_btn_frame,
            text="删除",
            command=self.delete_preset
        ).pack(side=tk.LEFT, padx=2)
        
        # 右侧：映射配置
        right_frame = ttk.Frame(main_frame)
        main_frame.add(right_frame, weight=2)
        
        # 预设信息
        info_frame = ttk.LabelFrame(right_frame, text="预设信息", padding=10)
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(info_frame, text="预设名称：").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.preset_name_label = ttk.Label(info_frame, text="", font=("微软雅黑", 9, "bold"))
        self.preset_name_label.grid(row=0, column=1, sticky=tk.W, pady=2, padx=5, columnspan=2)
        
        ttk.Label(info_frame, text="说明：").grid(row=1, column=0, sticky=tk.W+tk.N, pady=2)
        self.desc_text = tk.Text(info_frame, height=2, width=40, font=("微软雅黑", 9))
        self.desc_text.grid(row=1, column=1, sticky=tk.W+tk.E, pady=2, padx=5, columnspan=2)
        
        # 结算金额配置
        ttk.Label(info_frame, text="结算金额搜索列：").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.search_column_entry = ttk.Entry(info_frame, width=10, font=("微软雅黑", 9))
        self.search_column_entry.grid(row=2, column=1, sticky=tk.W, pady=2, padx=5)
        ttk.Label(info_frame, text="如: D", foreground="gray").grid(row=2, column=2, sticky=tk.W, pady=2)
        
        ttk.Label(info_frame, text="结算金额关键词：").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.search_keyword_entry = ttk.Entry(info_frame, width=20, font=("微软雅黑", 9))
        self.search_keyword_entry.grid(row=3, column=1, sticky=tk.W, pady=2, padx=5)
        ttk.Label(info_frame, text="如: 折后总计", foreground="gray").grid(row=3, column=2, sticky=tk.W, pady=2)
        
        ttk.Button(
            info_frame,
            text="保存配置",
            command=self.save_preset_info
        ).grid(row=4, column=1, pady=10, sticky=tk.W)
        
        info_frame.columnconfigure(1, weight=1)
        
        # 映射列表
        mapping_frame = ttk.LabelFrame(right_frame, text="映射配置", padding=10)
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 工具栏
        toolbar = ttk.Frame(mapping_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Button(
            toolbar,
            text="添加映射",
            command=self.add_mapping
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="编辑映射",
            command=self.edit_mapping
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="删除映射",
            command=self.delete_mapping
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="上移",
            command=self.move_mapping_up
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="下移",
            command=self.move_mapping_down
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            toolbar,
            text="预览Excel",
            command=self.preview_excel
        ).pack(side=tk.RIGHT, padx=2)
        
        # Treeview显示映射
        tree_frame = ttk.Frame(mapping_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.mapping_tree = ttk.Treeview(
            tree_frame,
            columns=("name", "cell", "description"),
            show="headings",
            yscrollcommand=tree_scroll.set
        )
        self.mapping_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.mapping_tree.yview)
        
        self.mapping_tree.heading("name", text="项目名称")
        self.mapping_tree.heading("cell", text="单元格")
        self.mapping_tree.heading("description", text="说明")
        
        self.mapping_tree.column("name", width=150)
        self.mapping_tree.column("cell", width=80)
        self.mapping_tree.column("description", width=200)
        
        # 双击编辑
        self.mapping_tree.bind('<Double-Button-1>', lambda e: self.edit_mapping())
        
        # 底部按钮
        bottom_frame = ttk.Frame(self.window)
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            bottom_frame,
            text="关闭",
            command=self.window.destroy
        ).pack(side=tk.RIGHT, padx=5)
        
        # 加载预设列表
        self.refresh_preset_list()
    
    def refresh_preset_list(self):
        """刷新预设列表"""
        self.preset_listbox.delete(0, tk.END)
        presets = self.config_manager.get_preset_names()
        for preset in presets:
            self.preset_listbox.insert(tk.END, preset)
        
        if presets:
            self.preset_listbox.selection_set(0)
            self.on_preset_select(None)
    
    def on_preset_select(self, event):
        """预设选择事件"""
        selection = self.preset_listbox.curselection()
        if not selection:
            return
        
        preset_name = self.preset_listbox.get(selection[0])
        self.load_preset(preset_name)
    
    def load_preset(self, preset_name):
        """加载预设配置"""
        preset = self.config_manager.get_preset(preset_name)
        if not preset:
            return
        
        self.current_preset = preset_name
        self.preset_name_label.config(text=preset_name)
        
        # 加载说明
        self.desc_text.delete("1.0", tk.END)
        self.desc_text.insert("1.0", preset.get("description", ""))
        
        # 加载结算金额配置
        self.search_column_entry.delete(0, tk.END)
        self.search_column_entry.insert(0, preset.get("settlement_search_column", "D"))
        
        self.search_keyword_entry.delete(0, tk.END)
        self.search_keyword_entry.insert(0, preset.get("settlement_search_keyword", "折后总计"))
        
        # 加载映射
        self.mapping_tree.delete(*self.mapping_tree.get_children())
        for mapping in preset.get("mappings", []):
            self.mapping_tree.insert("", tk.END, values=(
                mapping.get("name", ""),
                mapping.get("cell", ""),
                mapping.get("description", "")
            ))
    
    def save_preset_info(self):
        """保存预设信息（说明和结算金额配置）"""
        if not self.current_preset:
            return
        
        description = self.desc_text.get("1.0", tk.END).strip()
        search_column = self.search_column_entry.get().strip().upper()
        search_keyword = self.search_keyword_entry.get().strip()
        
        # 验证搜索列格式
        if search_column and not search_column.isalpha():
            messagebox.showwarning("提示", "搜索列必须是字母（如 A, B, C, D...）")
            return
        
        if not search_keyword:
            messagebox.showwarning("提示", "结算金额关键词不能为空！")
            return
        
        self.config_manager.update_preset(
            self.current_preset,
            description=description,
            settlement_search_column=search_column if search_column else "D",
            settlement_search_keyword=search_keyword
        )
        messagebox.showinfo("成功", "配置已保存！")
    
    def new_preset(self):
        """新建预设"""
        dialog = PresetNameDialog(self.window, "新建预设")
        self.window.wait_window(dialog.window)
        
        if dialog.result:
            name = dialog.result
            if name in self.config_manager.get_preset_names():
                messagebox.showerror("错误", "预设名称已存在！")
                return
            
            self.config_manager.add_preset(name)
            self.refresh_preset_list()
            # 选中新建的预设
            presets = self.config_manager.get_preset_names()
            idx = presets.index(name)
            self.preset_listbox.selection_clear(0, tk.END)
            self.preset_listbox.selection_set(idx)
            self.on_preset_select(None)
    
    def duplicate_preset(self):
        """复制预设"""
        if not self.current_preset:
            messagebox.showwarning("提示", "请先选择要复制的预设！")
            return
        
        dialog = PresetNameDialog(self.window, "复制预设", f"{self.current_preset}_副本")
        self.window.wait_window(dialog.window)
        
        if dialog.result:
            new_name = dialog.result
            if new_name in self.config_manager.get_preset_names():
                messagebox.showerror("错误", "预设名称已存在！")
                return
            
            self.config_manager.duplicate_preset(self.current_preset, new_name)
            self.refresh_preset_list()
    
    def rename_preset(self):
        """重命名预设"""
        if not self.current_preset:
            messagebox.showwarning("提示", "请先选择要重命名的预设！")
            return
        
        dialog = PresetNameDialog(self.window, "重命名预设", self.current_preset)
        self.window.wait_window(dialog.window)
        
        if dialog.result:
            new_name = dialog.result
            if new_name != self.current_preset:
                if new_name in self.config_manager.get_preset_names():
                    messagebox.showerror("错误", "预设名称已存在！")
                    return
                
                self.config_manager.rename_preset(self.current_preset, new_name)
                self.current_preset = new_name
                self.refresh_preset_list()
    
    def delete_preset(self):
        """删除预设"""
        if not self.current_preset:
            messagebox.showwarning("提示", "请先选择要删除的预设！")
            return
        
        if messagebox.askyesno("确认", f"确定要删除预设 '{self.current_preset}' 吗？"):
            self.config_manager.delete_preset(self.current_preset)
            self.current_preset = None
            self.refresh_preset_list()
    
    def add_mapping(self):
        """添加映射"""
        if not self.current_preset:
            messagebox.showwarning("提示", "请先选择一个预设！")
            return
        
        dialog = MappingDialog(self.window, "添加映射")
        self.window.wait_window(dialog.window)
        
        if dialog.result:
            preset = self.config_manager.get_preset(self.current_preset)
            mappings = preset.get("mappings", [])
            mappings.append(dialog.result)
            self.config_manager.update_preset(self.current_preset, mappings=mappings)
            self.load_preset(self.current_preset)
    
    def edit_mapping(self):
        """编辑映射"""
        if not self.current_preset:
            return
        
        selection = self.mapping_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要编辑的映射！")
            return
        
        item = selection[0]
        values = self.mapping_tree.item(item, "values")
        
        mapping = {
            "name": values[0],
            "cell": values[1],
            "description": values[2]
        }
        
        dialog = MappingDialog(self.window, "编辑映射", mapping)
        self.window.wait_window(dialog.window)
        
        if dialog.result:
            preset = self.config_manager.get_preset(self.current_preset)
            mappings = preset.get("mappings", [])
            idx = self.mapping_tree.index(item)
            mappings[idx] = dialog.result
            self.config_manager.update_preset(self.current_preset, mappings=mappings)
            self.load_preset(self.current_preset)
    
    def delete_mapping(self):
        """删除映射"""
        if not self.current_preset:
            return
        
        selection = self.mapping_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的映射！")
            return
        
        if messagebox.askyesno("确认", "确定要删除选中的映射吗？"):
            item = selection[0]
            idx = self.mapping_tree.index(item)
            
            preset = self.config_manager.get_preset(self.current_preset)
            mappings = preset.get("mappings", [])
            del mappings[idx]
            self.config_manager.update_preset(self.current_preset, mappings=mappings)
            self.load_preset(self.current_preset)
    
    def move_mapping_up(self):
        """上移映射"""
        if not self.current_preset:
            return
        
        selection = self.mapping_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        idx = self.mapping_tree.index(item)
        
        if idx == 0:
            return
        
        preset = self.config_manager.get_preset(self.current_preset)
        mappings = preset.get("mappings", [])
        mappings[idx], mappings[idx - 1] = mappings[idx - 1], mappings[idx]
        self.config_manager.update_preset(self.current_preset, mappings=mappings)
        self.load_preset(self.current_preset)
        
        # 重新选中
        items = self.mapping_tree.get_children()
        self.mapping_tree.selection_set(items[idx - 1])
    
    def move_mapping_down(self):
        """下移映射"""
        if not self.current_preset:
            return
        
        selection = self.mapping_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        idx = self.mapping_tree.index(item)
        
        preset = self.config_manager.get_preset(self.current_preset)
        mappings = preset.get("mappings", [])
        
        if idx >= len(mappings) - 1:
            return
        
        mappings[idx], mappings[idx + 1] = mappings[idx + 1], mappings[idx]
        self.config_manager.update_preset(self.current_preset, mappings=mappings)
        self.load_preset(self.current_preset)
        
        # 重新选中
        items = self.mapping_tree.get_children()
        self.mapping_tree.selection_set(items[idx + 1])
    
    def preview_excel(self):
        """预览Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择要预览的Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        
        if file_path:
            from excel_processor import ExcelProcessor
            processor = ExcelProcessor()
            preview_data = processor.preview_file(file_path, max_rows=15, max_cols=10)
            
            if preview_data:
                PreviewWindow(self.window, file_path, preview_data)


class PresetNameDialog:
    """预设名称输入对话框"""
    def __init__(self, parent, title, default_value=""):
        self.result = None
        
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("350x120")
        self.window.transient(parent)
        self.window.grab_set()
        
        # 居中显示
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        
        # 输入框
        ttk.Label(self.window, text="预设名称：").pack(pady=(20, 5))
        self.name_entry = ttk.Entry(self.window, width=30)
        self.name_entry.pack(pady=5)
        self.name_entry.insert(0, default_value)
        self.name_entry.select_range(0, tk.END)
        self.name_entry.focus()
        
        # 按钮
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="确定", command=self.ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.window.destroy).pack(side=tk.LEFT, padx=5)
        
        # 绑定回车键
        self.name_entry.bind('<Return>', lambda e: self.ok())
        self.window.bind('<Escape>', lambda e: self.window.destroy())
    
    def ok(self):
        """确认"""
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showwarning("提示", "预设名称不能为空！")
            return
        
        self.result = name
        self.window.destroy()


class MappingDialog:
    """映射配置对话框"""
    def __init__(self, parent, title, mapping=None):
        self.result = None
        
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("400x220")
        self.window.transient(parent)
        self.window.grab_set()
        
        # 居中显示
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
        
        # 输入表单
        form_frame = ttk.Frame(self.window, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(form_frame, text="项目名称：").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(form_frame, width=30)
        self.name_entry.grid(row=0, column=1, sticky=tk.W+tk.E, pady=5)
        
        ttk.Label(form_frame, text="单元格位置：").grid(row=1, column=0, sticky=tk.W, pady=5)
        cell_frame = ttk.Frame(form_frame)
        cell_frame.grid(row=1, column=1, sticky=tk.W+tk.E, pady=5)
        self.cell_entry = ttk.Entry(cell_frame, width=15)
        self.cell_entry.pack(side=tk.LEFT)
        ttk.Label(cell_frame, text="例如: A1, B2, C10", foreground="gray").pack(side=tk.LEFT, padx=5)
        
        ttk.Label(form_frame, text="说明：").grid(row=2, column=0, sticky=tk.W+tk.N, pady=5)
        self.desc_text = tk.Text(form_frame, height=3, width=30)
        self.desc_text.grid(row=2, column=1, sticky=tk.W+tk.E, pady=5)
        
        form_frame.columnconfigure(1, weight=1)
        
        # 按钮
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="确定", command=self.ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.window.destroy).pack(side=tk.LEFT, padx=5)
        
        # 加载现有数据
        if mapping:
            self.name_entry.insert(0, mapping.get("name", ""))
            self.cell_entry.insert(0, mapping.get("cell", ""))
            self.desc_text.insert("1.0", mapping.get("description", ""))
        
        self.name_entry.focus()
        
        # 绑定回车键
        self.window.bind('<Return>', lambda e: self.ok())
        self.window.bind('<Escape>', lambda e: self.window.destroy())
    
    def ok(self):
        """确认"""
        name = self.name_entry.get().strip()
        cell = self.cell_entry.get().strip().upper()
        description = self.desc_text.get("1.0", tk.END).strip()
        
        if not name:
            messagebox.showwarning("提示", "项目名称不能为空！")
            return
        
        if not cell:
            messagebox.showwarning("提示", "单元格位置不能为空！")
            return
        
        # 验证单元格格式
        from config_manager import ConfigManager
        cm = ConfigManager()
        if not cm.validate_cell_reference(cell):
            messagebox.showwarning("提示", "单元格格式不正确！请使用如 A1, B2 的格式。")
            return
        
        self.result = {
            "name": name,
            "cell": cell,
            "description": description
        }
        self.window.destroy()


class PreviewWindow:
    """Excel预览窗口"""
    def __init__(self, parent, file_path, data):
        self.window = tk.Toplevel(parent)
        self.window.title(f"预览: {os.path.basename(file_path)}")
        self.window.geometry("800x500")
        self.window.transient(parent)
        
        # 说明
        ttk.Label(
            self.window,
            text="点击单元格可以查看其位置信息（用于配置映射）",
            foreground="blue"
        ).pack(pady=5)
        
        # 创建表格
        frame = ttk.Frame(self.window)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 滚动条
        v_scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        h_scroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
        
        # Canvas和Frame用于显示网格
        canvas = tk.Canvas(frame, yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        v_scroll.config(command=canvas.yview)
        h_scroll.config(command=canvas.xview)
        
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 创建网格
        grid_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=grid_frame, anchor=tk.NW)
        
        # 显示数据
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                from openpyxl.utils import get_column_letter
                cell_ref = f"{get_column_letter(col_idx + 1)}{row_idx + 1}"
                
                cell_label = tk.Label(
                    grid_frame,
                    text=str(cell_value) if cell_value else "",
                    borderwidth=1,
                    relief=tk.SOLID,
                    width=12,
                    height=2,
                    anchor=tk.W,
                    padx=5
                )
                cell_label.grid(row=row_idx, column=col_idx, sticky=tk.W+tk.E)
                
                # 绑定点击事件
                cell_label.bind(
                    '<Button-1>',
                    lambda e, ref=cell_ref, val=cell_value: self.on_cell_click(ref, val)
                )
        
        # 更新滚动区域
        grid_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
        
        # 底部信息栏
        self.info_label = ttk.Label(
            self.window,
            text="点击单元格查看信息",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.info_label.pack(side=tk.BOTTOM, fill=tk.X)
    
    def on_cell_click(self, cell_ref, value):
        """单元格点击事件"""
        self.info_label.config(
            text=f"单元格: {cell_ref}  |  值: {value if value else '(空)'}"
        )

