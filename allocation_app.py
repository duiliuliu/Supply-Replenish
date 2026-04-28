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
        self.root.title("加单商品分配系统")
        self.root.geometry("1050x880")
        self.root.configure(bg="#FFFFFF")
        self.root.minsize(900, 700)
        
        self.config = load_config()
        self.file_path = None
        self.result_df = None
        self.reason_df = None
        
        self.setup_styles()
        self.create_widgets()
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('Title.TLabel', font=('SF Pro Display', 22, 'bold'), background='#FFFFFF', foreground='#1A1A1A')
        style.configure('Section.TLabelframe', font=('SF Pro Display', 13, 'bold'), background='#FFFFFF', borderwidth=0)
        style.configure('Section.TLabelframe.Label', font=('SF Pro Display', 13, 'bold'), foreground='#1A1A1A', background='#FFFFFF')
        style.configure('Config.TLabel', font=('SF Pro Display', 12), background='#FFFFFF', foreground='#6B6B6B')
        style.configure('Status.TLabel', font=('SF Pro Display', 12), background='#FFFFFF', foreground='#6B6B6B')
        
        style.configure('Primary.TButton', font=('SF Pro Display', 12, 'bold'), padding=(24, 12))
        style.configure('Secondary.TButton', font=('SF Pro Display', 11), padding=(12, 6))
        
        style.configure('Treeview', font=('SF Pro Display', 11), rowheight=28, background='#FFFFFF', foreground='#1A1A1A')
        style.configure('Treeview.Heading', font=('SF Pro Display', 11, 'bold'), background='#F7F7F7', foreground='#1A1A1A')
        style.map('Treeview', background=[('selected', '#1A1A1A')], foreground=[('selected', '#FFFFFF')])
    
    def create_widgets(self):
        main_frame = tk.Frame(self.root, bg='#FFFFFF')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)
        
        title_frame = tk.Frame(main_frame, bg='#FFFFFF')
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="加单商品分配系统", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        version_label = tk.Label(title_frame, text="v2.0", font=('SF Pro Display', 12), bg='#FFFFFF', fg='#9B9B9B')
        version_label.pack(side=tk.LEFT, padx=(12, 0), pady=(8, 0))
        
        self.create_config_section(main_frame)
        self.create_logic_section(main_frame)
        self.create_file_section(main_frame)
        self.create_button_section(main_frame)
        self.create_status_section(main_frame)
        self.create_result_section(main_frame)
    
    def create_card_frame(self, parent):
        card = tk.Frame(parent, bg='#FFFFFF', relief=tk.SOLID, borderwidth=1)
        card.config(borderwidth=1, relief=tk.SOLID, bg='#FFFFFF', highlightbackground='#E5E5E5', highlightcolor='#E5E5E5', highlightthickness=1)
        return card
    
    def create_config_section(self, parent):
        config_frame = tk.LabelFrame(parent, text="", font=('SF Pro Display', 13, 'bold'), bg='#FFFFFF', fg='#1A1A1A', padx=0, pady=0, bd=0)
        config_frame.pack(fill=tk.X, pady=(0, 16))
        
        config_card = self.create_card_frame(config_frame)
        config_card.pack(fill=tk.X)
        
        self.apply_card_style(config_card)
        
        self.config_expanded = False
        
        header_frame = tk.Frame(config_card, bg='#FFFFFF')
        header_frame.pack(fill=tk.X, pady=16, padx=16)
        header_frame.bind('<Button-1>', self.toggle_config)
        header_frame.config(cursor='hand2')
        
        self.config_toggle = tk.Label(header_frame, text="▶", font=('SF Pro Display', 12), bg='#FFFFFF', fg='#1A1A1A')
        self.config_toggle.pack(side=tk.RIGHT)
        
        config_title = tk.Label(header_frame, text="参数配置", font=('SF Pro Display', 13, 'bold'), bg='#FFFFFF', fg='#1A1A1A')
        config_title.pack(side=tk.LEFT)
        
        self.config_content = tk.Frame(config_card, bg='#FFFFFF')
        
        levels = ['SA', 'A', 'B', 'C', 'D', 'OL']
        
        coverage_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        coverage_frame.pack(fill=tk.X, pady=(0, 16))
        
        tk.Label(coverage_frame, text="覆盖周期（天）", font=('SF Pro Display', 12, 'bold'), bg='#FFFFFF', fg='#1A1A1A').pack(side=tk.LEFT, anchor=tk.N)
        
        self.coverage_entries = {}
        for i, level in enumerate(levels):
            frame = tk.Frame(coverage_frame, bg='#FFFFFF')
            frame.pack(side=tk.LEFT, padx=(20 if i == 0 else 16, 0))
            
            tk.Label(frame, text=level, font=('SF Pro Display', 11), bg='#FFFFFF', fg='#6B6B6B').pack(side=tk.TOP, pady=(0, 4))
            
            entry = tk.Entry(frame, width=8, font=('SF Pro Display', 12), justify='center', bg='#F7F7F7', relief=tk.SOLID, borderwidth=1)
            entry.pack(side=tk.TOP)
            entry.insert(0, str(self.config.get('allocation_config', {}).get('coverage_days', {}).get(level, 14)))
            self.coverage_entries[level] = entry
        
        weight_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        weight_frame.pack(fill=tk.X, pady=(0, 16))
        
        tk.Label(weight_frame, text="等级权重", font=('SF Pro Display', 12, 'bold'), bg='#FFFFFF', fg='#1A1A1A').pack(side=tk.LEFT, anchor=tk.N)
        
        self.weight_entries = {}
        for i, level in enumerate(levels):
            frame = tk.Frame(weight_frame, bg='#FFFFFF')
            frame.pack(side=tk.LEFT, padx=(20 if i == 0 else 16, 0))
            
            tk.Label(frame, text=level, font=('SF Pro Display', 11), bg='#FFFFFF', fg='#6B6B6B').pack(side=tk.TOP, pady=(0, 4))
            
            entry = tk.Entry(frame, width=8, font=('SF Pro Display', 12), justify='center', bg='#F7F7F7', relief=tk.SOLID, borderwidth=1)
            entry.pack(side=tk.TOP)
            entry.insert(0, str(self.config.get('allocation_config', {}).get('level_weights', {}).get(level, 1.0)))
            self.weight_entries[level] = entry
        
        safety_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        safety_frame.pack(fill=tk.X, pady=(0, 16))
        
        tk.Label(safety_frame, text="安全系数", font=('SF Pro Display', 12, 'bold'), bg='#FFFFFF', fg='#1A1A1A').pack(side=tk.LEFT, anchor=tk.N)
        
        self.safety_entries = {}
        for i, level in enumerate(levels):
            frame = tk.Frame(safety_frame, bg='#FFFFFF')
            frame.pack(side=tk.LEFT, padx=(20 if i == 0 else 16, 0))
            
            tk.Label(frame, text=level, font=('SF Pro Display', 11), bg='#FFFFFF', fg='#6B6B6B').pack(side=tk.TOP, pady=(0, 4))
            
            entry = tk.Entry(frame, width=8, font=('SF Pro Display', 12), justify='center', bg='#F7F7F7', relief=tk.SOLID, borderwidth=1)
            entry.pack(side=tk.TOP)
            entry.insert(0, str(self.config.get('allocation_config', {}).get('safety_factors', {}).get(level, 0.3)))
            self.safety_entries[level] = entry
        
        btn_frame = tk.Frame(self.config_content, bg='#FFFFFF')
        btn_frame.pack(fill=tk.X, pady=(0, 8))
        
        reset_btn = self.create_button(btn_frame, "恢复默认值", style='secondary')
        reset_btn.pack(side=tk.RIGHT)
        reset_btn.config(command=self.reset_config)
        
        save_btn = self.create_button(btn_frame, "保存配置", style='primary')
        save_btn.pack(side=tk.RIGHT, padx=(12, 0))
        save_btn.config(command=self.save_config)
    
    def apply_card_style(self, frame):
        try:
            frame.config(highlightbackground='#E5E5E5', highlightcolor='#E5E5E5', highlightthickness=1)
        except:
            pass
    
    def create_button(self, parent, text, style='primary'):
        if style == 'primary':
            btn = tk.Button(parent, text=text, font=('SF Pro Display', 11, 'bold'), 
                           bg='#1A1A1A', fg='white', relief=tk.FLAT, padx=20, pady=8, cursor='hand2')
            btn.bind('<Enter>', lambda e, b=btn: b.config(bg='#333333'))
            btn.bind('<Leave>', lambda e, b=btn: b.config(bg='#1A1A1A'))
            btn.bind('<ButtonPress>', lambda e, b=btn: b.config(bg='#444444'))
            btn.bind('<ButtonRelease>', lambda e, b=btn: b.config(bg='#333333'))
        else:
            btn = tk.Button(parent, text=text, font=('SF Pro Display', 11), 
                           bg='#F7F7F7', fg='#6B6B6B', relief=tk.SOLID, borderwidth=1, padx=20, pady=8, cursor='hand2')
            btn.bind('<Enter>', lambda e, b=btn: b.config(bg='#E5E5E5'))
            btn.bind('<Leave>', lambda e, b=btn: b.config(bg='#F7F7F7'))
            btn.bind('<ButtonPress>', lambda e, b=btn: b.config(bg='#D5D5D5'))
            btn.bind('<ButtonRelease>', lambda e, b=btn: b.config(bg='#E5E5E5'))
        return btn
    
    def toggle_config(self, event=None):
        if self.config_expanded:
            self.config_content.pack_forget()
            self.config_toggle.config(text="▶")
        else:
            self.config_content.pack(fill=tk.X, padx=16, pady=(0, 16))
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
        messagebox.showinfo("成功", "已恢复默认配置")
    
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
        logic_frame = tk.LabelFrame(parent, text="", font=('SF Pro Display', 13, 'bold'), bg='#FFFFFF', fg='#1A1A1A', padx=0, pady=0, bd=0)
        logic_frame.pack(fill=tk.X, pady=(0, 16))
        
        logic_card = self.create_card_frame(logic_frame)
        logic_card.pack(fill=tk.X)
        
        self.apply_card_style(logic_card)
        
        self.logic_expanded = True
        
        header_frame = tk.Frame(logic_card, bg='#FFFFFF')
        header_frame.pack(fill=tk.X, pady=16, padx=16)
        header_frame.bind('<Button-1>', self.toggle_logic)
        header_frame.config(cursor='hand2')
        
        self.logic_toggle = tk.Label(header_frame, text="▼", font=('SF Pro Display', 12), bg='#FFFFFF', fg='#1A1A1A')
        self.logic_toggle.pack(side=tk.RIGHT)
        
        logic_title = tk.Label(header_frame, text="分配逻辑说明", font=('SF Pro Display', 13, 'bold'), bg='#FFFFFF', fg='#1A1A1A')
        logic_title.pack(side=tk.LEFT)
        
        self.logic_content = tk.Frame(logic_card, bg='#FFFFFF')
        self.logic_content.pack(fill=tk.X, padx=16, pady=(0, 16))
        
        stages = [
            ("阶段1", "断码修复", "确保卖场达到最低库存要求"),
            ("阶段2", "销量匹配", "根据销量和覆盖周期计算目标库存"),
            ("阶段3", "销尽率优先", "按综合得分(销尽率×等级权重)分配"),
            ("阶段4", "剩余分配", "按等级优先级分配剩余库存")
        ]
        
        for idx, (stage_id, name, desc) in enumerate(stages):
            stage_frame = tk.Frame(self.logic_content, bg='#FFFFFF', relief=tk.SOLID, borderwidth=1)
            stage_frame.pack(fill=tk.X, pady=(0 if idx == 0 else 8, 0))
            
            stage_label = tk.Label(stage_frame, text=stage_id, font=('SF Pro Display', 11, 'bold'), bg='#1A1A1A', fg='white', width=8, padx=0, pady=10)
            stage_label.pack(side=tk.LEFT)
            
            content_frame = tk.Frame(stage_frame, bg='#FFFFFF')
            content_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=10, padx=12)
            
            name_label = tk.Label(content_frame, text=name, font=('SF Pro Display', 11, 'bold'), bg='#FFFFFF', fg='#1A1A1A')
            name_label.pack(anchor=tk.W)
            
            desc_label = tk.Label(content_frame, text=desc, font=('SF Pro Display', 11), bg='#FFFFFF', fg='#6B6B6B')
            desc_label.pack(anchor=tk.W, pady=(4, 0))
    
    def toggle_logic(self, event=None):
        if self.logic_expanded:
            self.logic_content.pack_forget()
            self.logic_toggle.config(text="▶")
        else:
            self.logic_content.pack(fill=tk.X, padx=16, pady=(0, 16))
            self.logic_toggle.config(text="▼")
        self.logic_expanded = not self.logic_expanded
    
    def create_file_section(self, parent):
        file_frame = tk.LabelFrame(parent, text="", font=('SF Pro Display', 13, 'bold'), bg='#FFFFFF', fg='#1A1A1A', padx=0, pady=0, bd=0)
        file_frame.pack(fill=tk.X, pady=(0, 16))
        
        file_card = self.create_card_frame(file_frame)
        file_card.pack(fill=tk.X)
        
        self.apply_card_style(file_card)
        
        input_frame = tk.Frame(file_card, bg='#FFFFFF')
        input_frame.pack(fill=tk.X, pady=16, padx=16)
        
        self.file_entry = tk.Entry(input_frame, font=('SF Pro Display', 12), bg='#F7F7F7', relief=tk.SOLID, borderwidth=1, disabledbackground='#F7F7F7')
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))
        
        browse_btn = self.create_button(input_frame, "浏览文件", style='primary')
        browse_btn.pack(side=tk.RIGHT)
        browse_btn.config(command=self.browse_file)
    
    def create_button_section(self, parent):
        btn_frame = tk.Frame(parent, bg='#FFFFFF')
        btn_frame.pack(fill=tk.X, pady=(0, 16))
        
        self.run_btn = tk.Button(btn_frame, text="执行加单分配", font=('SF Pro Display', 13, 'bold'), 
                                bg='#E5E5E5', fg='#9B9B9B', relief=tk.FLAT, padx=32, pady=12, 
                                cursor='arrow', state=tk.DISABLED)
        self.run_btn.pack(side=tk.LEFT, padx=(0, 16))
        
        self.save_btn = tk.Button(btn_frame, text="保存结果", font=('SF Pro Display', 13, 'bold'), 
                                 bg='#E5E5E5', fg='#9B9B9B', relief=tk.FLAT, padx=32, pady=12, 
                                 cursor='arrow', state=tk.DISABLED)
        self.save_btn.pack(side=tk.LEFT)
    
    def create_status_section(self, parent):
        status_frame = tk.Frame(parent, bg='#FFFFFF')
        status_frame.pack(fill=tk.X, pady=(0, 16))
        
        self.status_label = tk.Label(status_frame, text="等待文件选择...", font=('SF Pro Display', 12), bg='#FFFFFF', fg='#9B9B9B')
        self.status_label.pack(side=tk.LEFT)
        
        self.status_icon = tk.Label(status_frame, text="○", font=('SF Pro Display', 12), bg='#FFFFFF', fg='#9B9B9B')
        self.status_icon.pack(side=tk.LEFT, padx=(8, 0))
    
    def create_result_section(self, parent):
        result_frame = tk.LabelFrame(parent, text="", font=('SF Pro Display', 13, 'bold'), bg='#FFFFFF', fg='#1A1A1A', padx=0, pady=0, bd=0)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        result_card = self.create_card_frame(result_frame)
        result_card.pack(fill=tk.BOTH, expand=True)
        
        self.apply_card_style(result_card)
        
        header_frame = tk.Frame(result_card, bg='#FFFFFF')
        header_frame.pack(fill=tk.X, pady=16, padx=16)
        
        result_title = tk.Label(header_frame, text="分配结果预览", font=('SF Pro Display', 13, 'bold'), bg='#FFFFFF', fg='#1A1A1A')
        result_title.pack(side=tk.LEFT)
        
        result_subtitle = tk.Label(header_frame, text="(前20行)", font=('SF Pro Display', 12), bg='#FFFFFF', fg='#9B9B9B')
        result_subtitle.pack(side=tk.LEFT, padx=(8, 0))
        
        tree_container = tk.Frame(result_card, bg='#FFFFFF')
        tree_container.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))
        
        self.tree = ttk.Treeview(tree_container, show='headings')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.bind('<Motion>', lambda e: self.root.config(cursor='arrow'))
    
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        
        if self.file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.file_path)
            self.enable_run_button()
            self.update_status("文件已选择，点击执行开始分配", 'success')
    
    def update_status(self, text, status='info'):
        colors = {
            'info': ('#6B6B6B', '○'),
            'success': ('#4CAF50', '✓'),
            'warning': ('#FF9800', '●'),
            'error': ('#F44336', '✗')
        }
        fg_color, icon = colors.get(status, colors['info'])
        self.status_label.config(text=text, fg=fg_color)
        self.status_icon.config(text=icon, fg=fg_color)
    
    def enable_run_button(self):
        self.run_btn.config(state=tk.NORMAL, bg='#1A1A1A', fg='white', cursor='hand2', command=self.run_allocation)
        self.run_btn.bind('<Enter>', lambda e: self.run_btn.config(bg='#333333'))
        self.run_btn.bind('<Leave>', lambda e: self.run_btn.config(bg='#1A1A1A'))
        self.run_btn.bind('<ButtonPress>', lambda e: self.run_btn.config(bg='#444444'))
        self.run_btn.bind('<ButtonRelease>', lambda e: self.run_btn.config(bg='#333333'))
    
    def enable_save_button(self):
        self.save_btn.config(state=tk.NORMAL, bg='#1A1A1A', fg='white', cursor='hand2', command=self.save_result)
        self.save_btn.bind('<Enter>', lambda e: self.save_btn.config(bg='#333333'))
        self.save_btn.bind('<Leave>', lambda e: self.save_btn.config(bg='#1A1A1A'))
        self.save_btn.bind('<ButtonPress>', lambda e: self.save_btn.config(bg='#444444'))
        self.save_btn.bind('<ButtonRelease>', lambda e: self.save_btn.config(bg='#333333'))
    
    def run_allocation(self):
        self.update_status("正在读取Excel文件...", 'warning')
        self.root.update()
        
        try:
            df_inventory = pd.read_excel(self.file_path, sheet_name='库存')
            df_sales = pd.read_excel(self.file_path, sheet_name='销售')
            df_store_level = pd.read_excel(self.file_path, sheet_name='卖场等级')
            df_add_order = pd.read_excel(self.file_path, sheet_name='加单数量')
            
            self.update_status("正在执行分配逻辑...", 'warning')
            self.root.update()
            
            allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
                df_inventory, df_sales, df_store_level, df_add_order, self.config
            )
            
            self.result_df, self.reason_df = generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus, store_level_map)
            
            self.show_preview()
            
            self.enable_save_button()
            self.update_status("分配完成！可以保存结果", 'success')
            
            messagebox.showinfo("成功", "加单分配完成！")
            
        except Exception as e:
            self.update_status("执行失败: " + str(e), 'error')
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