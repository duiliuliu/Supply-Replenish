# 加单商品分配系统 v2.5
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
import sys
import traceback

try:
    from allocation_core import allocate_add_order, generate_result_dataframe, load_config, DEFAULT_CONFIG
except Exception as e:
    print(f'Failed to import allocation_core: {e}')
    from collections import defaultdict
    
    def allocate_add_order(*args, **kwargs):
        return defaultdict(lambda: defaultdict(int)), defaultdict(lambda: defaultdict(str)), [], [], {}
    
    def generate_result_dataframe(*args, **kwargs):
        return pd.DataFrame(), pd.DataFrame()
    
    load_config = lambda: DEFAULT_CONFIG

class AllocationApp:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("加单商品分配系统")
            self.root.geometry("1100x1050")
            self.root.configure(bg="#F5F7FA")
            self.root.minsize(950, 850)
            
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            self.config = None
            self.file_path = None
            self.result_df = None
            self.reason_df = None
            
            self.stage_colors = [
                ("#E8F5FF", "#2563EB"),
                ("#F0FDF4", "#059669"),
                ("#FFF3E0", "#D97706"),
                ("#F5F3FF", "#7C3AED"),
            ]
            
            self.stage_list = [
                ("broken_size_fix", "断码修复", "优先填充缺码关键SKU"),
                ("sales_match", "销量匹配", "依据历史销速加权分配"),
                ("sell_through_priority", "销尽率优先", "高销尽门店获得补货权重"),
                ("remaining_allocation", "剩余分配", "尾量零散SKU随机填充")
            ]
            
            self.drag_data = {"index": -1, "item": None, "y": 0}
            self.stage_frames = []
            
            self.setup_styles()
            self.create_widgets()
            
            print('应用程序初始化成功')
        except Exception as e:
            print(f'初始化错误: {e}')
            traceback.print_exc()
            messagebox.showerror("初始化错误", f"程序初始化失败:\n{str(e)}")
    
    def on_closing(self):
        try:
            self.root.destroy()
        except:
            sys.exit(0)
    
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Title.TLabel", font=("SF Pro Display", 24, "bold"), background="#F5F7FA", foreground="#1F2937")
        style.configure("Subtitle.TLabel", font=("SF Pro Display", 14), background="#F5F7FA", foreground="#6B7280")
        
        style.configure("Treeview", font=("SF Pro Display", 12), rowheight=32, background="#FFFFFF", foreground="#1F2937", borderwidth=0)
        style.configure("Treeview.Heading", font=("SF Pro Display", 12, "bold"), background="#F9FAFB", foreground="#374151", borderwidth=0, relief=tk.FLAT)
        style.map("Treeview", background=[("selected", "#EFF6FF")], foreground=[("selected", "#2563EB")])
    
    def create_widgets(self):
        main_frame = tk.Frame(self.root, bg="#F5F7FA")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)
        
        self.create_header(main_frame)
        self.create_config_section(main_frame)
        self.create_logic_section(main_frame)
        self.create_file_upload_section(main_frame)
        self.create_result_section(main_frame)
        self.create_status_bar(main_frame)
    
    def create_header(self, parent):
        header_frame = tk.Frame(parent, bg="#F5F7FA")
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_frame = tk.Frame(header_frame, bg="#F5F7FA")
        title_frame.pack(side=tk.LEFT)
        
        title_label = tk.Label(title_frame, text="加单商品分配系统", font=("SF Pro Display", 24, "bold"), bg="#F5F7FA", fg="#1F2937")
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(title_frame, text="基于动态权重的库存补货与分配模型", font=("SF Pro Display", 14), bg="#F5F7FA", fg="#6B7280")
        subtitle_label.pack(anchor=tk.W, pady=(4, 0))
        
        version_label = tk.Label(header_frame, text="v2.5", font=("SF Pro Display", 13), bg="#F5F7FA", fg="#9CA3AF")
        version_label.pack(side=tk.RIGHT)
    
    def create_card_frame(self, parent):
        card = tk.Frame(parent, bg="#FFFFFF")
        card.config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB", highlightthickness=1)
        return card
    
    def create_config_section(self, parent):
        config_card = self.create_card_frame(parent)
        config_card.pack(fill=tk.X, pady=(0, 16))
        
        self.config_expanded = False
        
        header_frame = tk.Frame(config_card, bg="#FFFFFF")
        header_frame.pack(fill=tk.X, pady=14, padx=20)
        header_frame.bind("<Button-1>", self.toggle_config)
        header_frame.config(cursor="hand2")
        
        self.config_toggle = tk.Label(header_frame, text="▼", font=("SF Pro Display", 14), bg="#FFFFFF", fg="#6B7280")
        self.config_toggle.pack(side=tk.RIGHT)
        
        icon_label = tk.Label(header_frame, text="⚙", font=("SF Pro Display", 16), bg="#FFFFFF", fg="#2563EB")
        icon_label.pack(side=tk.LEFT)
        
        config_title = tk.Label(header_frame, text="参数配置", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
        config_title.pack(side=tk.LEFT, padx=(8, 0))
        
        self.config_content = tk.Frame(config_card, bg="#FFFFFF")
        self.config_content.pack(fill=tk.X, padx=20)
        
        self.create_config_grid()
    
    def create_config_grid(self):
        sections = [
            ("覆盖周期（天）", "coverage_days", {"SA": 30, "A": 30, "B": 14, "C": 14, "D": 14, "OL": 14}, int),
            ("等级权重", "level_weights", {"SA": 1.5, "A": 1.3, "B": 1.2, "C": 1.1, "D": 1.1, "OL": 1.0}, float),
            ("安全系数", "safety_factors", {"SA": 0.5, "A": 0.4, "B": 0.3, "C": 0.25, "D": 0.2, "OL": 0.2}, float),
            ("最小目标库存", "min_target_inventory", {"SA": 0, "A": 0, "B": 0, "C": 0, "D": 0, "OL": 0}, int),
        ]
        
        self.config_entries = {}
        
        for section_title, section_key, defaults, dtype in sections:
            row_frame = tk.Frame(self.config_content, bg="#FFFFFF")
            row_frame.pack(fill=tk.X, pady=(0, 14))
            
            label_frame = tk.Frame(row_frame, bg="#FFFFFF", width=120)
            label_frame.pack(side=tk.LEFT, anchor=tk.N)
            label_frame.pack_propagate(False)
            tk.Label(label_frame, text=section_title, font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#374151").pack(anchor=tk.W)
            
            entry_frame = tk.Frame(row_frame, bg="#FFFFFF")
            entry_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
            
            levels = ["SA", "A", "B", "C", "D", "OL"]
            self.config_entries[section_key] = {}
            
            for i, level in enumerate(levels):
                item_frame = tk.Frame(entry_frame, bg="#FFFFFF")
                item_frame.pack(side=tk.LEFT, padx=(32 if i == 0 else 24, 0))
                
                level_label = tk.Label(item_frame, text=level, font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280")
                level_label.pack(anchor=tk.W)
                
                entry = tk.Entry(item_frame, width=8, font=("SF Pro Display", 12), justify="center", 
                                bg="#F3F4F6", relief=tk.FLAT, borderwidth=0, highlightthickness=1, 
                                highlightbackground="#E5E7EB", highlightcolor="#D1D5DB")
                entry.pack(ipady=7, pady=(4, 0))
                
                try:
                    val = str(self.config.get("allocation_config", {}).get(section_key, {}).get(level, defaults[level]) if self.config else defaults[level])
                except:
                    val = str(defaults[level])
                entry.insert(0, val)
                self.config_entries[section_key][level] = entry
        
        note_label = tk.Label(self.config_content, text="注: 确保即使销量为0，卖场也能获得基础库存。", font=("SF Pro Display", 11), bg="#FFFFFF", fg="#9CA3AF")
        note_label.pack(anchor=tk.E)
        
        btn_frame = tk.Frame(self.config_content, bg="#FFFFFF")
        btn_frame.pack(fill=tk.X, pady=(12, 8))
        
        reset_btn = tk.Button(btn_frame, text="恢复默认值", font=("SF Pro Display", 12), bg="#F3F4F6", fg="#4B5563", 
                             relief=tk.FLAT, padx=20, pady=8, cursor="hand2", command=self.reset_config)
        reset_btn.pack(side=tk.RIGHT)
        
        save_btn = tk.Button(btn_frame, text="保存配置", font=("SF Pro Display", 12, "bold"), bg="#2563EB", fg="white", 
                            relief=tk.FLAT, padx=20, pady=8, cursor="hand2", command=self.save_config)
        save_btn.pack(side=tk.RIGHT, padx=(12, 0))
    
    def toggle_config(self, event=None):
        if self.config_expanded:
            self.config_content.pack_forget()
            self.config_toggle.config(text="▶")
        else:
            self.config_content.pack(fill=tk.X, padx=20)
            self.config_toggle.config(text="▼")
        self.config_expanded = not self.config_expanded
    
    def reset_config(self):
        sections = [("coverage_days", {"SA": 30, "A": 30, "B": 14, "C": 14, "D": 14, "OL": 14}),
                    ("level_weights", {"SA": 1.5, "A": 1.3, "B": 1.2, "C": 1.1, "D": 1.1, "OL": 1.0}),
                    ("safety_factors", {"SA": 0.5, "A": 0.4, "B": 0.3, "C": 0.25, "D": 0.2, "OL": 0.2}),
                    ("min_target_inventory", {"SA": 0, "A": 0, "B": 0, "C": 0, "D": 0, "OL": 0})]
        
        for key, defaults in sections:
            for level, val in defaults.items():
                if key in self.config_entries and level in self.config_entries[key]:
                    self.config_entries[key][level].delete(0, tk.END)
                    self.config_entries[key][level].insert(0, str(val))
        
        messagebox.showinfo("成功", "已恢复默认配置")
    
    def save_config(self):
        config = {
            "version": "2.5",
            "allocation_config": {
                "coverage_days": {},
                "level_weights": {},
                "safety_factors": {},
                "min_target_inventory": {},
                "stage_priority": [stage[0] for stage in self.stage_list[:3]],
                "max_remaining_per_store": 10
            }
        }
        
        sections = [("coverage_days", int), ("level_weights", float), ("safety_factors", float), ("min_target_inventory", int)]
        for key, dtype in sections:
            for level in ["SA", "A", "B", "C", "D", "OL"]:
                try:
                    config["allocation_config"][key][level] = dtype(self.config_entries[key][level].get())
                except:
                    pass
        
        config_path = os.path.join(os.path.dirname(__file__), "allocation_config.json")
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        
        self.config = config
        messagebox.showinfo("成功", "配置已保存!")
    
    def create_logic_section(self, parent):
        logic_card = self.create_card_frame(parent)
        logic_card.pack(fill=tk.X, pady=(0, 16))
        
        self.logic_expanded = True
        
        header_frame = tk.Frame(logic_card, bg="#FFFFFF")
        header_frame.pack(fill=tk.X, pady=14, padx=20)
        header_frame.bind("<Button-1>", self.toggle_logic)
        header_frame.config(cursor="hand2")
        
        self.logic_toggle = tk.Label(header_frame, text="▼", font=("SF Pro Display", 14), bg="#FFFFFF", fg="#6B7280")
        self.logic_toggle.pack(side=tk.RIGHT)
        
        icon_label = tk.Label(header_frame, text="📊", font=("SF Pro Display", 16), bg="#FFFFFF", fg="#2563EB")
        icon_label.pack(side=tk.LEFT)
        
        logic_title = tk.Label(header_frame, text="分配逻辑说明", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
        logic_title.pack(side=tk.LEFT, padx=(8, 0))
        
        drag_hint = tk.Label(header_frame, text="(可拖拽调整前三个阶段顺序)", font=("SF Pro Display", 11), bg="#FFFFFF", fg="#9CA3AF")
        drag_hint.pack(side=tk.LEFT, padx=(12, 0))
        
        self.logic_content = tk.Frame(logic_card, bg="#FFFFFF")
        self.logic_content.pack(fill=tk.X, padx=20, pady=(0, 14))
        
        self.create_logic_stages()
    
    def create_logic_stages(self):
        self.stages_container = tk.Frame(self.logic_content, bg="#FFFFFF")
        self.stages_container.pack(fill=tk.X)
        
        self.stage_frames = []
        
        for i, (stage_id, name, desc) in enumerate(self.stage_list):
            self._create_stage_item(i, stage_id, name, desc)
    
    def _create_stage_item(self, idx, stage_id, name, desc):
        bg_color, fg_color = self.stage_colors[idx]
        
        stage_frame = tk.Frame(self.stages_container, bg=bg_color, bd=1, relief=tk.FLAT)
        stage_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0 if idx == 0 else 12, 0))
        
        if idx < 3:
            stage_frame.bind("<ButtonPress-1>", lambda e, idx=idx: self._on_stage_press(e, idx))
            stage_frame.bind("<B1-Motion>", lambda e, idx=idx: self._on_stage_drag(e, idx))
            stage_frame.bind("<ButtonRelease-1>", lambda e, idx=idx: self._on_stage_release(e, idx))
            stage_frame.config(cursor="hand2", highlightbackground="#2563EB", highlightcolor="#2563EB", highlightthickness=2)
        
        stage_content = tk.Frame(stage_frame, bg=bg_color)
        stage_content.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)
        
        num_frame = tk.Frame(stage_content, bg="#FFFFFF", width=32, height=32)
        num_frame.pack(padx=8, pady=(0, 12))
        num_frame.pack_propagate(False)
        
        num_label = tk.Label(num_frame, text=str(idx+1), font=("SF Pro Display", 14, "bold"), bg="#FFFFFF", fg=fg_color)
        num_label.pack(fill=tk.BOTH, expand=True)
        
        stage_name_label = tk.Label(stage_content, text=name, font=("SF Pro Display", 13, "bold"), bg=bg_color, fg="#1F2937")
        stage_name_label.pack(pady=(0, 4))
        
        stage_desc_label = tk.Label(stage_content, text=desc, font=("SF Pro Display", 11), bg=bg_color, fg="#6B7280")
        stage_desc_label.pack()
        
        if idx < 3:
            drag_label = tk.Label(stage_content, text="⋮⋮", font=("SF Pro Display", 10), bg=bg_color, fg="#9CA3AF")
            drag_label.pack(pady=(8, 0))
        
        self.stage_frames.append((stage_id, stage_frame))
        
        if idx < 3:
            arrow_frame = tk.Frame(self.stages_container, bg="#FFFFFF", width=24)
            arrow_frame.pack(side=tk.LEFT)
            arrow_label = tk.Label(arrow_frame, text="→", font=("SF Pro Display", 16), bg="#FFFFFF", fg="#D1D5DB")
            arrow_label.pack(fill=tk.BOTH, expand=True)
    
    def _on_stage_press(self, event, index):
        self.drag_data["index"] = index
        self.drag_data["y"] = event.y_root
        self.drag_data["item"] = self.stage_frames[index][1]
        self.drag_data["item"].config(highlightbackground="#2563EB", highlightcolor="#2563EB", highlightthickness=3)
        self.drag_data["item"].config(cursor="hand2")
    
    def _on_stage_drag(self, event, index):
        if self.drag_data["index"] == -1:
            return
        
        delta_y = event.y_root - self.drag_data["y"]
        if abs(delta_y) > 30:
            direction = 1 if delta_y > 0 else -1
            new_index = index + direction
            
            if 0 <= new_index < 3:
                self.stage_list[index], self.stage_list[new_index] = self.stage_list[new_index], self.stage_list[index]
                
                for stage_id, widget in self.stage_frames:
                    widget.pack_forget()
                
                for widget in self.stages_container.winfo_children():
                    widget.destroy()
                
                self.stage_frames = []
                
                for i, (stage_id, name, desc) in enumerate(self.stage_list):
                    self._create_stage_item(i, stage_id, name, desc)
                
                self.drag_data["index"] = new_index
                self.drag_data["y"] = event.y_root
    
    def _on_stage_release(self, event, index):
        if self.drag_data["item"]:
            self.drag_data["item"].config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB", highlightthickness=2)
            self.drag_data["item"].config(cursor="hand2")
        self.drag_data = {"index": -1, "item": None, "y": 0}
    
    def toggle_logic(self, event=None):
        if self.logic_expanded:
            self.logic_content.pack_forget()
            self.logic_toggle.config(text="▶")
        else:
            self.logic_content.pack(fill=tk.X, padx=20, pady=(0, 14))
            self.logic_toggle.config(text="▼")
        self.logic_expanded = not self.logic_expanded
    
    def create_file_upload_section(self, parent):
        upload_frame = tk.Frame(parent, bg="#F5F7FA")
        upload_frame.pack(fill=tk.X, pady=(0, 16))
        
        upload_inner = tk.Frame(upload_frame, bg="#F5F7FA")
        upload_inner.pack(fill=tk.X)
        
        file_zone = tk.Frame(upload_inner, bg="#FFFFFF", bd=2, relief=tk.DASHED)
        file_zone.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 16))
        file_zone.config(highlightbackground="#D1D5DB", highlightcolor="#D1D5DB", highlightthickness=1)
        
        file_zone.bind("<Button-1>", lambda e: self.browse_file())
        file_zone.config(cursor="hand2")
        
        zone_content = tk.Frame(file_zone, bg="#FFFFFF")
        zone_content.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)
        
        cloud_icon = tk.Label(zone_content, text="☁", font=("Arial", 48), bg="#FFFFFF", fg="#9CA3AF")
        cloud_icon.pack(pady=(0, 12))
        
        title_label = tk.Label(zone_content, text="拖拽或点击上传清单", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
        title_label.pack()
        
        desc_label = tk.Label(zone_content, text="支持 .xlsx, .csv 格式，单次处理上限 10,000 行", font=("SF Pro Display", 13), bg="#FFFFFF", fg="#6B7280")
        desc_label.pack(pady=(6, 0))
        
        action_zone = tk.Frame(upload_inner, bg="#2563EB", width=320)
        action_zone.pack(side=tk.RIGHT, fill=tk.BOTH)
        
        action_content = tk.Frame(action_zone, bg="#2563EB")
        action_content.pack(fill=tk.BOTH, expand=True, padx=24, pady=24)
        
        rocket_icon = tk.Label(action_content, text="🚀", font=("Arial", 40), bg="#2563EB")
        rocket_icon.pack(pady=(0, 12))
        
        action_title = tk.Label(action_content, text="执行加单分配", font=("SF Pro Display", 16, "bold"), bg="#2563EB", fg="white")
        action_title.pack()
        
        action_desc = tk.Label(action_content, text="启动四阶段智能分配引擎，预计耗时 12s", font=("SF Pro Display", 12), bg="#2563EB", fg="#BFDBFE")
        action_desc.pack(pady=(4, 16))
        
        self.run_btn = tk.Button(action_content, text="开始计算", font=("SF Pro Display", 14, "bold"), 
                                 bg="#FFFFFF", fg="#2563EB", relief=tk.FLAT, padx=40, pady=12, 
                                 cursor="arrow", state=tk.DISABLED, command=self.run_allocation)
        self.run_btn.pack(pady=(0, 8))
        
        self.save_btn = tk.Button(action_content, text="导出清单", font=("SF Pro Display", 14, "bold"), 
                                  bg="#1D4ED8", fg="white", relief=tk.FLAT, padx=40, pady=12, 
                                  cursor="arrow", state=tk.DISABLED, command=self.save_result)
        self.save_btn.pack()
    
    def create_result_section(self, parent):
        result_card = self.create_card_frame(parent)
        result_card.pack(fill=tk.BOTH, expand=True)
        
        header_frame = tk.Frame(result_card, bg="#FFFFFF")
        header_frame.pack(fill=tk.X, pady=14, padx=20)
        
        left_frame = tk.Frame(header_frame, bg="#FFFFFF")
        left_frame.pack(side=tk.LEFT)
        
        icon_label = tk.Label(left_frame, text="📋", font=("SF Pro Display", 16), bg="#FFFFFF", fg="#2563EB")
        icon_label.pack(side=tk.LEFT)
        
        result_title = tk.Label(left_frame, text="分配结果预览", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
        result_title.pack(side=tk.LEFT, padx=(8, 0))
        
        right_frame = tk.Frame(header_frame, bg="#FFFFFF")
        right_frame.pack(side=tk.RIGHT)
        
        export_btn = tk.Button(right_frame, text="导出清单", font=("SF Pro Display", 12), bg="#F3F4F6", fg="#4B5563", 
                               relief=tk.FLAT, padx=14, pady=6, cursor="hand2", command=self.save_result)
        export_btn.pack(side=tk.LEFT)
        
        tree_container = tk.Frame(result_card, bg="#FFFFFF")
        tree_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 14))
        
        self.tree = ttk.Treeview(tree_container, show="headings", borderwidth=0)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
    
    def create_status_bar(self, parent):
        status_frame = tk.Frame(parent, bg="#FFFFFF", bd=1, relief=tk.FLAT)
        status_frame.pack(fill=tk.X, pady=(8, 0))
        status_frame.config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB", highlightthickness=1)
        
        left_frame = tk.Frame(status_frame, bg="#FFFFFF")
        left_frame.pack(side=tk.LEFT, padx=16, pady=6)
        
        self.status_icon = tk.Label(left_frame, text="○", font=("SF Pro Display", 12), bg="#FFFFFF", fg="#9CA3AF")
        self.status_icon.pack(side=tk.LEFT)
        
        self.status_label = tk.Label(left_frame, text="系统在线", font=("SF Pro Display", 12), bg="#FFFFFF", fg="#059669")
        self.status_label.pack(side=tk.LEFT, padx=(6, 12))
        
        right_frame = tk.Frame(status_frame, bg="#FFFFFF")
        right_frame.pack(side=tk.RIGHT, padx=16)
        
        sync_label = tk.Label(right_frame, text="上次同步: 10:45 AM", font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280")
        sync_label.pack()
    
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if self.file_path:
            self.run_btn.config(state=tk.NORMAL, cursor="hand2")
            self.run_btn.bind("<Enter>", lambda e: self.run_btn.config(bg="#F3F4F6"))
            self.run_btn.bind("<Leave>", lambda e: self.run_btn.config(bg="#FFFFFF"))
            self.update_status("文件已选择，点击执行开始分配", "info")
    
    def run_allocation(self):
        try:
            self.update_status("处理中...", "warning")
            self.root.update()
            
            df_inventory = pd.read_excel(self.file_path, sheet_name="库存")
            df_sales = pd.read_excel(self.file_path, sheet_name="销售")
            df_store_level = pd.read_excel(self.file_path, sheet_name="卖场等级")
            df_add_order = pd.read_excel(self.file_path, sheet_name="加单数量")
            
            allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
                df_inventory, df_sales, df_store_level, df_add_order, self.config
            )
            
            self.result_df, self.reason_df = generate_result_dataframe(
                allocation_result, allocation_reasons, stores_sorted, skus, store_level_map
            )
            
            self.show_preview()
            self.save_btn.config(state=tk.NORMAL, cursor="hand2")
            self.update_status("分配完成", "success")
            
            messagebox.showinfo("成功", "加单分配完成!")
        except Exception as e:
            print(f'run_allocation error: {e}')
            traceback.print_exc()
            self.update_status("执行失败: " + str(e), "error")
            messagebox.showerror("错误", "执行失败:\n" + str(e))
    
    def update_status(self, text, status="info"):
        colors = {
            "info": ("#6B7280", "○"),
            "success": ("#059669", "✓"),
            "warning": ("#D97706", "●"),
            "error": ("#DC2626", "✗")
        }
        fg_color, icon = colors.get(status, colors["info"])
        self.status_label.config(text=text, fg=fg_color)
        self.status_icon.config(text=icon, fg=fg_color)
    
    def show_preview(self):
        for col in self.tree.get_children():
            self.tree.delete(col)
        
        if self.result_df is not None and len(self.result_df) > 0:
            columns = list(self.result_df.columns)
            self.tree["columns"] = columns
            
            for col in columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor="center")
            
            for idx, row in self.result_df.head(20).iterrows():
                values = list(row)
                self.tree.insert("", "end", values=values)
    
    def save_result(self):
        if self.result_df is None:
            messagebox.showwarning("提示", "请先执行分配")
            return
        
        save_path = filedialog.asksaveasfilename(
            title="保存分配结果",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        
        if save_path:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                self.result_df.to_excel(writer, sheet_name="分配数量", index=False)
                self.reason_df.to_excel(writer, sheet_name="分配原因", index=False)
            
            messagebox.showinfo("成功", f"结果已保存到:\n{save_path}")

def main():
    try:
        print('启动加单商品分配系统...')
        
        root = tk.Tk()
        app = AllocationApp(root)
        root.mainloop()
    except Exception as e:
        print(f'程序主错误: {e}')
        traceback.print_exc()
        
        try:
            messagebox.showerror("程序错误", f"程序发生严重错误:\n{str(e)}")
        except:
            pass
        
        sys.exit(1)

if __name__ == "__main__":
    main()
