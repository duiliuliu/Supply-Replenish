# 加单商品分配系统 v2.0
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
from allocation_core import allocate_add_order, generate_result_dataframe, load_config, DEFAULT_CONFIG

class AllocationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("加单商品分配系统 v2.0")
        self.root.geometry("1000x850")
        self.root.configure(bg="#F8F9FA")
        
        self.config = load_config()
        self.file_path = None
        self.result_df = None
        self.reason_df = None
        
        self.setup_styles()
        self.create_widgets()
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('Title.TLabel', font=('Microsoft YaHei', 20, 'bold'), background='#F8F9FA', foreground='#1A1A2E')
        style.configure('Section.TLabelframe', font=('Microsoft YaHei', 11, 'bold'), background='#FFFFFF')
        style.configure('Section.TLabelframe.Label', font=('Microsoft YaHei', 11, 'bold'), foreground='#2D3436', background='#FFFFFF')
        style.configure('Config.TLabel', font=('Microsoft YaHei', 10), background='#FFFFFF', foreground='#2D3436')
        style.configure('Status.TLabel', font=('Microsoft YaHei', 10), background='#F8F9FA', foreground='#636E72')
        
        style.configure('Primary.TButton', font=('Microsoft YaHei', 11, 'bold'), padding=(20, 10))
        style.configure('Secondary.TButton', font=('Microsoft YaHei', 10), padding=(10, 5))
        
        style.configure('Treeview', font=('Microsoft YaHei', 9), rowheight=28)
        style.configure('Treeview.Heading', font=('Microsoft YaHei', 10, 'bold'))
    
    def create_widgets(self):
        main_frame = tk.Frame(self.root, bg='#F8F9FA')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        title_frame = tk.Frame(main_frame, bg='#F8F9FA')
        title_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(title_frame, text="加单商品分配系统", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        version_label = tk.Label(title_frame, text="v2.0", font=('Microsoft YaHei', 10), bg='#F8F9FA', fg='#636E72')
        version_label.pack(side=tk.LEFT, padx=(10, 0), pady=(8, 0))
        
        self.create_config_section(main_frame)
        self.create_logic_section(main_frame)
        self.create_file_section(main_frame)
        self.create_button_section(main_frame)
        self.create_status_section(main_frame)
        self.create_result_section(main_frame)
    
    def create_config_section(self, parent):
        config_frame = tk.LabelFrame(parent, text=" 参数配置 ", font=('Microsoft YaHei', 11, 'bold'), bg='#FFFFFF', fg='#2D3436', padx=15, pady=10)
        config_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.config_expanded = False
        
        header_frame = tk.Frame(config_frame, bg='#FFFFFF')
        header_frame.pack(fill=tk.X)
        
        self.config_toggle = tk.Label(header_frame, text="▶", font=('Microsoft YaHei', 12), bg='#FFFFFF', fg='#0984E3', cursor='hand2')
        self.config_toggle.pack(side=tk.LEFT)
        self.config_toggle.bind('<Button-1>', self.toggle_config)
        
        config_title = tk.Label(header_frame, text="点击展开/折叠参数配置", font=('Microsoft YaHei', 10), bg='#FFFFFF', fg='#636E72', cursor='hand2')
        config_title.pack(side=tk.LEFT, padx=(5, 0))
        config_title.bind('<Button-1>', self.toggle_config)
        
        self.config_content = tk.Frame(config_frame, bg='#FFFFFF')
        
        levels = ['SA', 'A', 'B', 'C', 'D', 'OL']
        
        coverage_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        coverage_frame.pack(fill=tk.X, pady=(10, 5))
        
        tk.Label(coverage_frame, text="覆盖周期（天）:", font=('Microsoft YaHei', 10, 'bold'), bg='#FFFFFF', fg='#2D3436').pack(side=tk.LEFT)
        
        self.coverage_entries = {}
        for i, level in enumerate(levels):
            frame = tk.Frame(coverage_frame, bg='#FFFFFF')
            frame.pack(side=tk.LEFT, padx=(15 if i == 0 else 10, 0))
            
            tk.Label(frame, text=level, font=('Microsoft YaHei', 9), bg='#FFFFFF', fg='#636E72').pack(side=tk.LEFT)
            
            entry = tk.Entry(frame, width=5, font=('Microsoft YaHei', 10), justify='center')
            entry.pack(side=tk.LEFT, padx=(3, 0))
            entry.insert(0, str(self.config.get('allocation_config', {}).get('coverage_days', {}).get(level, 14)))
            self.coverage_entries[level] = entry
        
        weight_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        weight_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(weight_frame, text="等级权重:", font=('Microsoft YaHei', 10, 'bold'), bg='#FFFFFF', fg='#2D3436').pack(side=tk.LEFT)
        
        self.weight_entries = {}
        for i, level in enumerate(levels):
            frame = tk.Frame(weight_frame, bg='#FFFFFF')
            frame.pack(side=tk.LEFT, padx=(15 if i == 0 else 10, 0))
            
            tk.Label(frame, text=level, font=('Microsoft YaHei', 9), bg='#FFFFFF', fg='#636E72').pack(side=tk.LEFT)
            
            entry = tk.Entry(frame, width=5, font=('Microsoft YaHei', 10), justify='center')
            entry.pack(side=tk.LEFT, padx=(3, 0))
            entry.insert(0, str(self.config.get('allocation_config', {}).get('level_weights', {}).get(level, 1.0)))
            self.weight_entries[level] = entry
        
        safety_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        safety_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(safety_frame, text="安全系数:", font=('Microsoft YaHei', 10, 'bold'), bg='#FFFFFF', fg='#2D3436').pack(side=tk.LEFT)
        
        self.safety_entries = {}
        for i, level in enumerate(levels):
            frame = tk.Frame(safety_frame, bg='#FFFFFF')
            frame.pack(side=tk.LEFT, padx=(15 if i == 0 else 10, 0))
            
            tk.Label(frame, text=level, font=('Microsoft YaHei', 9), bg='#FFFFFF', fg='#636E72').pack(side=tk.LEFT)
            
            entry = tk.Entry(frame, width=5, font=('Microsoft YaHei', 10), justify='center')
            entry.pack(side=tk.LEFT, padx=(3, 0))
            entry.insert(0, str(self.config.get('allocation_config', {}).get('safety_factors', {}).get(level, 0.3)))
            self.safety_entries[level] = entry
        
        btn_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        reset_btn = tk.Button(btn_frame, text="恢复默认值", font=('Microsoft YaHei', 9), bg='#DFE6E9', fg='#2D3436', relief=tk.FLAT, padx=15, pady=5, cursor='hand2', command=self.reset_config)
        reset_btn.pack(side=tk.LEFT)
        
        save_btn = tk.Button(btn_frame, text="保存配置", font=('Microsoft YaHei', 9), bg='#0984E3', fg='white', relief=tk.FLAT, padx=15, pady=5, cursor='hand2', command=self.save_config)
        save_btn.pack(side=tk.LEFT, padx=(10, 0))
    
    def toggle_config(self, event=None):
        if self.config_expanded:
            self.config_content.pack_forget()
            self.config_toggle.config(text="▶")
        else:
            self.config_content.pack(fill=tk.X, pady=(5, 0))
            self.config_toggle.config(text="▼")
        self.config_expanded = not self.config_expanded
    
    def reset_config(self):
        default_config = DEFAULT_CONFIG['allocation_config']
        for level in ['SA', 'A', 'B', 'C', 'D', 'OL']:
            self.coverage_entries[level].delete(0, tk.END)
            self.coverage_entries[level].insert(0, str(default_config['coverage_days'].get(level, 14)))
            self.weight_entries[level].delete(0, tk.END)
            self.weight_entries[level].insert(0, str(default_config['level_weights'].get(level, 1.0)))
            self.safety_entries[level].delete(0, tk.END)
            self.safety_entries[level].insert(0, str(default_config['safety_factors'].get(level, 0.3)))
    
    def save_config(self):
        try:
            config = {
                "version": "2.0",
                "updated_at": "2026-04-28",
                "allocation_config": {
                    "coverage_days": {},
                    "level_weights": {},
                    "safety_factors": {},
                    "max_remaining_per_store": 10
                }
            }
            
            for level in ['SA', 'A', 'B', 'C', 'D', 'OL']:
                config['allocation_config']['coverage_days'][level] = int(self.coverage_entries[level].get())
                config['allocation_config']['level_weights'][level] = float(self.weight_entries[level].get())
                config['allocation_config']['safety_factors'][level] = float(self.safety_entries[level].get())
            
            config_path = os.path.join(os.path.dirname(__file__), 'allocation_config.json')
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.config = config
            messagebox.showinfo("成功", "配置已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def create_logic_section(self, parent):
        logic_frame = tk.LabelFrame(parent, text=" 分配逻辑说明 ", font=('Microsoft YaHei', 11, 'bold'), bg='#FFFFFF', fg='#2D3436', padx=15, pady=10)
        logic_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.logic_expanded = False
        
        header_frame = tk.Frame(logic_frame, bg='#FFFFFF')
        header_frame.pack(fill=tk.X)
        
        self.logic_toggle = tk.Label(header_frame, text="▶", font=('Microsoft YaHei', 12), bg='#FFFFFF', fg='#0984E3', cursor='hand2')
        self.logic_toggle.pack(side=tk.LEFT)
        self.logic_toggle.bind('<Button-1>', self.toggle_logic)
        
        logic_title = tk.Label(header_frame, text="点击展开/折叠分配逻辑说明", font=('Microsoft YaHei', 10), bg='#FFFFFF', fg='#636E72', cursor='hand2')
        logic_title.pack(side=tk.LEFT, padx=(5, 0))
        logic_title.bind('<Button-1>', self.toggle_logic)
        
        self.logic_content = tk.Frame(logic_frame, bg='#FFFFFF')
        
        stages = [
            ("阶段1", "断码修复", "#E3F2FD", "#1976D2", "确保卖场达到最低库存要求"),
            ("阶段2", "销量匹配", "#E8F5E9", "#388E3C", "根据销量和覆盖周期计算目标库存"),
            ("阶段3", "销尽率优先", "#FFF8E1", "#F57C00", "按综合得分(销尽率×等级权重)分配"),
            ("阶段4", "剩余分配", "#FCE4EC", "#C2185B", "按等级优先级分配剩余库存")
        ]
        
        for stage_id, name, bg_color, border_color, desc in stages:
            stage_frame = tk.Frame(self.logic_content, bg=bg_color, relief=tk.SOLID, borderwidth=1)
            stage_frame.pack(fill=tk.X, pady=3)
            
            stage_label = tk.Label(stage_frame, text=stage_id, font=('Microsoft YaHei', 10, 'bold'), bg=border_color, fg='white', width=6)
            stage_label.pack(side=tk.LEFT, padx=(0, 10))
            
            content_frame = tk.Frame(stage_frame, bg=bg_color)
            content_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=5)
            
            tk.Label(content_frame, text=name, font=('Microsoft YaHei', 10, 'bold'), bg=bg_color, fg='#2D3436').pack(anchor=tk.W)
            tk.Label(content_frame, text=desc, font=('Microsoft YaHei', 9), bg=bg_color, fg='#636E72').pack(anchor=tk.W)
    
    def toggle_logic(self, event=None):
        if self.logic_expanded:
            self.logic_content.pack_forget()
            self.logic_toggle.config(text="▶")
        else:
            self.logic_content.pack(fill=tk.X, pady=(5, 0))
            self.logic_toggle.config(text="▼")
        self.logic_expanded = not self.logic_expanded
    
    def create_file_section(self, parent):
        file_frame = tk.LabelFrame(parent, text=" 文件选择 ", font=('Microsoft YaHei', 11, 'bold'), bg='#FFFFFF', fg='#2D3436', padx=15, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        input_frame = tk.Frame(file_frame, bg='#FFFFFF')
        input_frame.pack(fill=tk.X)
        
        self.file_entry = tk.Entry(input_frame, font=('Microsoft YaHei', 10), bg='#F8F9FA', relief=tk.SOLID, borderwidth=1)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_btn = tk.Button(input_frame, text="浏览", font=('Microsoft YaHei', 10, 'bold'), bg='#0984E3', fg='white', relief=tk.FLAT, padx=20, pady=5, cursor='hand2', command=self.browse_file)
        browse_btn.pack(side=tk.RIGHT)
    
    def create_button_section(self, parent):
        btn_frame = tk.Frame(parent, bg='#F8F9FA')
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.run_btn = tk.Button(btn_frame, text="执行加单分配", font=('Microsoft YaHei', 12, 'bold'), bg='#BDC3C7', fg='#7F8C8D', relief=tk.FLAT, padx=30, pady=10, cursor='arrow', state=tk.DISABLED)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        self.save_btn = tk.Button(btn_frame, text="保存结果", font=('Microsoft YaHei', 12, 'bold'), bg='#BDC3C7', fg='#7F8C8D', relief=tk.FLAT, padx=30, pady=10, cursor='arrow', state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT)
    
    def create_status_section(self, parent):
        status_frame = tk.Frame(parent, bg='#F8F9FA')
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = tk.Label(status_frame, text="等待文件选择...", font=('Microsoft YaHei', 10), bg='#F8F9FA', fg='#636E72')
        self.status_label.pack(side=tk.LEFT)
    
    def create_result_section(self, parent):
        result_frame = tk.LabelFrame(parent, text=" 分配结果预览（前20行） ", font=('Microsoft YaHei', 11, 'bold'), bg='#FFFFFF', fg='#2D3436', padx=10, pady=10)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        tree_container = tk.Frame(result_frame, bg='#FFFFFF')
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(tree_container, show='headings')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
    
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        
        if self.file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.file_path)
            self.enable_run_button()
            self.status_label.config(text="文件已选择，点击执行开始分配", fg='#0984E3')
    
    def enable_run_button(self):
        self.run_btn.config(state=tk.NORMAL, bg='#00B894', fg='white', cursor='hand2', command=self.run_allocation)
    
    def enable_save_button(self):
        self.save_btn.config(state=tk.NORMAL, bg='#6C5CE7', fg='white', cursor='hand2', command=self.save_result)
    
    def run_allocation(self):
        self.status_label.config(text="正在读取Excel文件...", fg='#F39C12')
        self.root.update()
        
        try:
            df_inventory = pd.read_excel(self.file_path, sheet_name='库存')
            df_sales = pd.read_excel(self.file_path, sheet_name='销售')
            df_store_level = pd.read_excel(self.file_path, sheet_name='卖场等级')
            df_add_order = pd.read_excel(self.file_path, sheet_name='加单数量')
            
            self.status_label.config(text="正在执行分配逻辑...", fg='#F39C12')
            self.root.update()
            
            allocation_result, allocation_reasons, stores_sorted, skus = allocate_add_order(
                df_inventory, df_sales, df_store_level, df_add_order, self.config
            )
            
            self.result_df, self.reason_df = generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus)
            
            self.show_preview()
            
            self.enable_save_button()
            self.status_label.config(text="分配完成！可以保存结果", fg='#00B894')
            
            messagebox.showinfo("成功", "加单分配完成！")
            
        except Exception as e:
            self.status_label.config(text="执行失败: " + str(e), fg='#E74C3C')
            messagebox.showerror("错误", "执行失败:\n" + str(e))
    
    def show_preview(self):
        for col in self.tree.get_children():
            self.tree.delete(col)
        
        self.tree['columns'] = list(self.result_df.columns)
        
        for col in self.result_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')
        
        for idx, row in self.result_df.head(20).iterrows():
            values = list(row)
            self.tree.insert('', 'end', values=values)
    
    def save_result(self):
        save_path = filedialog.asksaveasfilename(
            title="保存分配结果",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if save_path:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                self.result_df.to_excel(writer, sheet_name='分配数量', index=False)
                self.reason_df.to_excel(writer, sheet_name='分配原因', index=False)
            
            messagebox.showinfo("成功", f"结果已保存到:\n{save_path}")

def main():
    root = tk.Tk()
    app = AllocationApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
