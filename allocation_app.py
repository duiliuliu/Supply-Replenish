#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import sys
import traceback
import json
from allocation_core import allocate_add_order, generate_result_dataframe, DEFAULT_CONFIG, load_config, VERSION

class AllocationApp:
    def __init__(self):
        try:
            self.root = tk.Tk()
            
            try:
                self.config = load_config()
            except:
                self.config = DEFAULT_CONFIG
            
            self.version = VERSION
            
            self.root.title(f"加单商品分配系统 v{self.version}")
            self.root.geometry("900x850")
            self.root.minsize(850, 750)
            self.root.configure(bg="#F5F7FA")
            
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            self.file_path = None
            self.result_df = None
            self.reason_df = None
            
            self.stage_colors = [
                ("#E8F5FF", "#2563EB"),
                ("#F0FDF4", "#059669"),
                ("#FFF3E0", "#D97706"),
                ("#F5F3FF", "#7C3AED"),
            ]
            
            stage_names = {
                "broken_size_fix": ("broken_size_fix", "断码修复", "SA/A级核心尺码至少2件，非核心尺码至少1件；其他等级核心尺码至少1件"),
                "sales_match": ("sales_match", "销量匹配", "目标库存 = 平均日需求 × 覆盖周期 + 安全库存"),
                "sell_through_priority": ("sell_through_priority", "销尽率优先", "综合得分 = 销尽率 × 等级权重，降序分配"),
            }
            
            default_stage_list = [
                ("broken_size_fix", "断码修复", "SA/A级核心尺码至少2件，非核心尺码至少1件；其他等级核心尺码至少1件"),
                ("sales_match", "销量匹配", "目标库存 = 平均日需求 × 覆盖周期 + 安全库存"),
                ("sell_through_priority", "销尽率优先", "综合得分 = 销尽率 × 等级权重，降序分配"),
            ]
            
            remaining_stage = ("remaining_allocation", "剩余分配", "按等级优先级分配：SA → A → B → C → D → OL，单卖场上限10件")
            
            config_priority = self.config.get("allocation_config", {}).get("stage_priority", [])
            self.stage_list = []
            for stage_id in config_priority:
                if stage_id in stage_names:
                    self.stage_list.append(stage_names[stage_id])
            
            if len(self.stage_list) != 3:
                self.stage_list = default_stage_list.copy()
            
            self.stage_list.append(remaining_stage)
            
            self.stage_frames = []
            self.stage_vars = []
            
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
        
        style.configure("Treeview", font=("SF Pro Display", 12), rowheight=32, background="#FFFFFF", foreground="#1F2937")
        style.configure("Treeview.Heading", font=("SF Pro Display", 12, "bold"), background="#F9FAFB", foreground="#374151")
        style.map("Treeview", background=[("selected", "#EFF6FF")], foreground=[("selected", "#2563EB")])
    
    def create_widgets(self):
        main_container = tk.Frame(self.root, bg="#F5F7FA")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(main_container, bg="#F5F7FA", highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#F5F7FA")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=24, pady=20)
        scrollbar.pack(side="right", fill="y")
        
        self.create_header(scrollable_frame)
        self.create_config_section(scrollable_frame)
        self.create_logic_section(scrollable_frame)
        self.create_file_upload_section(scrollable_frame)
        self.create_result_section(scrollable_frame)
        self.create_status_bar(scrollable_frame)
    
    def create_header(self, parent):
        header_frame = tk.Frame(parent, bg="#F5F7FA")
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_frame = tk.Frame(header_frame, bg="#F5F7FA")
        title_frame.pack(side=tk.LEFT)
        
        title_label = tk.Label(title_frame, text="加单商品分配系统", font=("SF Pro Display", 24, "bold"), bg="#F5F7FA", fg="#1F2937")
        title_label.pack(anchor=tk.W)
        
        subtitle_label = tk.Label(title_frame, text="基于动态权重的库存补货与分配模型", font=("SF Pro Display", 14), bg="#F5F7FA", fg="#6B7280")
        subtitle_label.pack(anchor=tk.W, pady=(4, 0))
        
        # 打赏链接
        donate_label = tk.Label(title_frame, text="❤️ 请作者喝杯咖啡", font=("SF Pro Display", 11), bg="#F5F7FA", fg="#EF4444", cursor="hand2")
        donate_label.pack(anchor=tk.W, pady=(2, 0))
        donate_label.bind("<Button-1>", self.open_donate)
        
        version_label = tk.Label(header_frame, text=f"v{self.version}", font=("SF Pro Display", 13), bg="#F5F7FA", fg="#9CA3AF")
        version_label.pack(side=tk.RIGHT)
    
    def create_card_frame(self, parent):
        card = tk.Frame(parent, bg="#FFFFFF")
        card.config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB", highlightthickness=1)
        return card
    
    def create_config_section(self, parent):
        config_card = self.create_card_frame(parent)
        config_card.pack(fill=tk.X, pady=(0, 16))
        
        self.config_expanded = True
        
        header_frame = tk.Frame(config_card, bg="#FFFFFF")
        header_frame.pack(fill=tk.X, pady=14, padx=20)
        header_frame.bind("<Button-1>", self.toggle_config)
        header_frame.config(cursor="hand2")
        
        self.config_toggle = tk.Label(header_frame, text="▼", font=("SF Pro Display", 14), bg="#FFFFFF", fg="#6B7280")
        self.config_toggle.pack(side=tk.RIGHT)
        self.config_toggle.bind("<Button-1>", self.toggle_config)
        self.config_toggle.config(cursor="hand2")
        
        icon_label = tk.Label(header_frame, text="⚙", font=("SF Pro Display", 16), bg="#FFFFFF", fg="#2563EB")
        icon_label.pack(side=tk.LEFT)
        icon_label.bind("<Button-1>", self.toggle_config)
        icon_label.config(cursor="hand2")
        
        config_title = tk.Label(header_frame, text="参数配置", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
        config_title.pack(side=tk.LEFT, padx=(8, 0))
        config_title.bind("<Button-1>", self.toggle_config)
        config_title.config(cursor="hand2")
        
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
            
            tk.Label(row_frame, text=section_title, font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#1F2937", width=14, anchor="w").pack(side=tk.LEFT, padx=(0, 20))
            
            entry_frame = tk.Frame(row_frame, bg="#FFFFFF")
            entry_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
            
            levels = ["SA", "A", "B", "C", "D", "OL"]
            self.config_entries[section_key] = {}
            
            for i, level in enumerate(levels):
                item_frame = tk.Frame(entry_frame, bg="#FFFFFF")
                item_frame.pack(side=tk.LEFT, padx=(0 if i == 0 else 16, 0))
                
                level_label = tk.Label(item_frame, text=level, font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280")
                level_label.pack(anchor=tk.W)
                
                entry = tk.Entry(item_frame, width=8, font=("SF Pro Display", 12), justify="center", 
                                bg="#E0E7FF", fg="#1F2937", relief=tk.FLAT, borderwidth=0, highlightthickness=1, 
                                highlightbackground="#C7D2FE", highlightcolor="#93C5FD")
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
        
        # 延迟显示消息框，避免 Mac Tkinter 崩溃
        self.root.after(100, lambda: messagebox.showinfo("成功", "已恢复默认配置"))
    
    def save_config(self):
        from datetime import datetime
        config = {
            "updated_at": datetime.now().strftime("%Y-%m-%d"),
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
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.config = config
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
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
        self.logic_toggle.bind("<Button-1>", self.toggle_logic)
        self.logic_toggle.config(cursor="hand2")
        
        icon_label = tk.Label(header_frame, text="📊", font=("SF Pro Display", 16), bg="#FFFFFF", fg="#2563EB")
        icon_label.pack(side=tk.LEFT)
        icon_label.bind("<Button-1>", self.toggle_logic)
        icon_label.config(cursor="hand2")
        
        logic_title = tk.Label(header_frame, text="分配逻辑说明", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
        logic_title.pack(side=tk.LEFT, padx=(8, 0))
        logic_title.bind("<Button-1>", self.toggle_logic)
        logic_title.config(cursor="hand2")
        
        self.logic_content = tk.Frame(logic_card, bg="#FFFFFF")
        self.logic_content.pack(fill=tk.X, padx=20, pady=(0, 14))
        
        self.create_logic_stages()
    
    def create_logic_stages(self):
        order_frame = tk.Frame(self.logic_content, bg="#FFFFFF")
        order_frame.pack(fill=tk.X, pady=(0, 16))
        
        order_label = tk.Label(order_frame, text="阶段顺序调整", font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#374151")
        order_label.pack(side=tk.LEFT)
        
        self.stage_vars = []
        stage_options = ["断码修复", "销量匹配", "销尽率优先"]
        for i in range(3):
            stage_container = tk.Frame(order_frame, bg="#FFFFFF")
            stage_container.pack(side=tk.LEFT, padx=(16, 0))
            
            tk.Label(stage_container, text=f"第{i+1}阶段:", font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280").pack(side=tk.LEFT)
            
            var = tk.StringVar()
            var.set(self.stage_list[i][1])
            self.stage_vars.append(var)
            
            cb = ttk.Combobox(stage_container, textvariable=var, values=stage_options, state="readonly", width=12, font=("SF Pro Display", 12))
            cb.pack(side=tk.LEFT, padx=(4, 0))
        
        apply_btn = tk.Button(order_frame, text="应用顺序", font=("SF Pro Display", 12, "bold"), bg="#2563EB", fg="white", 
                             relief=tk.FLAT, padx=16, pady=6, cursor="hand2", command=self.apply_stage_order)
        apply_btn.pack(side=tk.LEFT, padx=(24, 0))
        
        reset_stage_btn = tk.Button(order_frame, text="恢复默认", font=("SF Pro Display", 12), bg="#F3F4F6", fg="#4B5563", 
                             relief=tk.FLAT, padx=16, pady=6, cursor="hand2", command=self.reset_stage_order)
        reset_stage_btn.pack(side=tk.LEFT, padx=(8, 0))
        
        self.stages_container = tk.Frame(self.logic_content, bg="#FFFFFF")
        self.stages_container.pack(fill=tk.X, pady=(12, 0))
        
        self.stage_frames = []
        
        for i, (stage_id, name, desc) in enumerate(self.stage_list):
            self._create_stage_item(i, stage_id, name, desc)
    
    def _create_stage_item(self, idx, stage_id, name, desc):
        try:
            if idx < len(self.stage_colors):
                bg_color, fg_color = self.stage_colors[idx]
            else:
                bg_color, fg_color = self.stage_colors[-1] if self.stage_colors else ("#F5F5F5", "#666666")
            
            stage_frame = tk.Frame(self.stages_container, bg=bg_color, width=180)
            stage_frame.config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB", highlightthickness=1)
            stage_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0 if idx == 0 else 8, 0))
            stage_frame.pack_propagate(False)
            
            stage_content = tk.Frame(stage_frame, bg=bg_color)
            stage_content.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
            
            num_frame = tk.Frame(stage_content, bg="#FFFFFF", width=28, height=28)
            num_frame.pack(padx=4, pady=(0, 8))
            num_frame.pack_propagate(False)
            
            num_label = tk.Label(num_frame, text=str(idx+1), font=("SF Pro Display", 12, "bold"), bg="#FFFFFF", fg=fg_color)
            num_label.pack(fill=tk.BOTH, expand=True)
            
            stage_name_label = tk.Label(stage_content, text=name, font=("SF Pro Display", 12, "bold"), bg=bg_color, fg="#1F2937")
            stage_name_label.pack(pady=(0, 4))
            
            wrapped_desc = self._wrap_text(desc, 7)
            stage_desc_label = tk.Label(stage_content, text=wrapped_desc, font=("SF Pro Display", 10), bg=bg_color, fg="#6B7280", justify=tk.LEFT, wraplength=160)
            stage_desc_label.pack(fill=tk.X)
            
            self.stage_frames.append((stage_id, stage_frame))
            
            if idx < 3:
                arrow_frame = tk.Frame(self.stages_container, bg="#FFFFFF", width=20)
                arrow_frame.pack(side=tk.LEFT)
                arrow_frame.pack_propagate(False)
                arrow_label = tk.Label(arrow_frame, text="→", font=("SF Pro Display", 14), bg="#FFFFFF", fg="#D1D5DB")
                arrow_label.pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            print(f"_create_stage_item error for idx {idx}: {e}")
            import traceback
            traceback.print_exc()
    
    def _wrap_text(self, text, chars_per_line=7):
        result = []
        i = 0
        while i < len(text):
            result.append(text[i:i+chars_per_line])
            i += chars_per_line
        return "\n".join(result)
    
    def apply_stage_order(self):
        try:
            selected_names = [var.get() for var in self.stage_vars]
            
            # 验证选择完整性
            if len(set(selected_names)) != 3:
                messagebox.showwarning("提示", "请确保每个阶段只选择一次！")
                return
            
            # 验证所有选择都有效
            all_valid_names = ["断码修复", "销量匹配", "销尽率优先"]
            for name in selected_names:
                if name not in all_valid_names:
                    messagebox.showerror("错误", f"无效的阶段名称: {name}")
                    return
            
            # 构建完整的阶段映射（包括所有可能的阶段名）
            all_stage_map = {
                "断码修复": ("broken_size_fix", "断码修复", "SA/A级核心尺码至少2件，非核心尺码至少1件；其他等级核心尺码至少1件"),
                "销量匹配": ("sales_match", "销量匹配", "目标库存 = 平均日需求 × 覆盖周期 + 安全库存"),
                "销尽率优先": ("sell_through_priority", "销尽率优先", "综合得分 = 销尽率 × 等级权重，降序分配"),
            }
            
            # 安全构建新的阶段列表
            new_stage_list = []
            for name in selected_names:
                if name in all_stage_map:
                    new_stage_list.append(all_stage_map[name])
            
            # 确保我们有3个阶段
            if len(new_stage_list) != 3:
                messagebox.showerror("错误", "阶段构建失败，请重试")
                return
            
            new_stage_list.append(("remaining_allocation", "剩余分配", "按等级优先级分配：SA → A → B → C → D → OL，单卖场上限10件"))
            self.stage_list = new_stage_list
            
            # 更新配置
            if "allocation_config" not in self.config:
                self.config["allocation_config"] = {}
            self.config["allocation_config"]["stage_priority"] = [stage[0] for stage in self.stage_list[:3]]
            
            # 安全更新阶段显示
            if hasattr(self, 'stages_container') and self.stages_container:
                try:
                    for widget in self.stages_container.winfo_children():
                        try:
                            widget.destroy()
                        except:
                            pass
                except Exception as e:
                    print(f"清理阶段容器失败: {e}")
            
            self.stage_frames = []
            
            # 确保 stages_container 存在
            if not hasattr(self, 'stages_container') or not self.stages_container:
                print("警告: stages_container 不存在，尝试重建")
                if hasattr(self, 'logic_content') and self.logic_content:
                    self.stages_container = tk.Frame(self.logic_content, bg="#FFFFFF")
                    self.stages_container.pack(fill=tk.X, pady=(12, 0))
            
            # 只在 stages_container 存在时创建阶段项
            if hasattr(self, 'stages_container') and self.stages_container:
                for i, (stage_id, name, desc) in enumerate(self.stage_list):
                    try:
                        self._create_stage_item(i, stage_id, name, desc)
                    except Exception as e:
                        print(f"创建阶段 {i} 失败: {e}")
            
            # 安全更新下拉框显示为新的顺序
            if hasattr(self, 'stage_vars'):
                for i, var in enumerate(self.stage_vars):
                    if i < len(self.stage_list):
                        var.set(self.stage_list[i][1])
            
            # 延迟显示消息框，避免 Mac Tkinter 崩溃
            self.root.after(100, lambda: messagebox.showinfo("成功", "阶段顺序已更新！"))
        except Exception as e:
            print(f"apply_stage_order error: {e}")
            traceback.print_exc()
            # 延迟显示消息框，避免 Mac Tkinter 崩溃
            self.root.after(100, lambda err=str(e): messagebox.showerror("错误", f"应用顺序失败:\n{err}"))
    
    def reset_stage_order(self):
        try:
            default_stage_list = [
                ("broken_size_fix", "断码修复", "SA/A级核心尺码至少2件，非核心尺码至少1件；其他等级核心尺码至少1件"),
                ("sales_match", "销量匹配", "目标库存 = 平均日需求 × 覆盖周期 + 安全库存"),
                ("sell_through_priority", "销尽率优先", "综合得分 = 销尽率 × 等级权重，降序分配"),
                ("remaining_allocation", "剩余分配", "按等级优先级分配：SA → A → B → C → D → OL，单卖场上限10件")
            ]
            self.stage_list = default_stage_list.copy()
            
            # 更新配置
            if "allocation_config" not in self.config:
                self.config["allocation_config"] = {}
            self.config["allocation_config"]["stage_priority"] = ["broken_size_fix", "sales_match", "sell_through_priority"]
            
            # 安全更新阶段显示
            if hasattr(self, 'stages_container') and self.stages_container:
                try:
                    for widget in self.stages_container.winfo_children():
                        try:
                            widget.destroy()
                        except:
                            pass
                except Exception as e:
                    print(f"清理阶段容器失败: {e}")
            
            self.stage_frames = []
            
            # 确保 stages_container 存在
            if not hasattr(self, 'stages_container') or not self.stages_container:
                print("警告: stages_container 不存在，尝试重建")
                if hasattr(self, 'logic_content') and self.logic_content:
                    self.stages_container = tk.Frame(self.logic_content, bg="#FFFFFF")
                    self.stages_container.pack(fill=tk.X, pady=(12, 0))
            
            # 只在 stages_container 存在时创建阶段项
            if hasattr(self, 'stages_container') and self.stages_container:
                for i, (stage_id, name, desc) in enumerate(self.stage_list):
                    try:
                        self._create_stage_item(i, stage_id, name, desc)
                    except Exception as e:
                        print(f"创建阶段 {i} 失败: {e}")
            
            # 安全更新下拉框显示为新的顺序
            if hasattr(self, 'stage_vars'):
                for i, var in enumerate(self.stage_vars):
                    if i < len(self.stage_list):
                        var.set(self.stage_list[i][1])
            
            # 延迟显示消息框，避免 Mac Tkinter 崩溃
            self.root.after(100, lambda: messagebox.showinfo("成功", "已恢复默认阶段顺序！"))
        except Exception as e:
            print(f"reset_stage_order error: {e}")
            traceback.print_exc()
            # 延迟显示消息框，避免 Mac Tkinter 崩溃
            self.root.after(100, lambda err=str(e): messagebox.showerror("错误", f"恢复默认失败:\n{err}"))
    
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
        
        file_zone = tk.Frame(upload_inner, bg="#FFFFFF")
        file_zone.config(highlightbackground="#D1D5DB", highlightcolor="#D1D5DB", highlightthickness=1)
        file_zone.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 16))
        
        file_zone.bind("<Button-1>", lambda e: self.browse_file())
        file_zone.config(cursor="hand2")
        
        zone_content = tk.Frame(file_zone, bg="#FFFFFF")
        zone_content.pack(fill=tk.BOTH, expand=True, padx=60, pady=50)
        zone_content.bind("<Button-1>", lambda e: self.browse_file())
        zone_content.config(cursor="hand2")
        
        cloud_icon = tk.Label(zone_content, text="☁", font=("Arial", 56), bg="#FFFFFF", fg="#9CA3AF")
        cloud_icon.pack(pady=(0, 16))
        cloud_icon.bind("<Button-1>", lambda e: self.browse_file())
        cloud_icon.config(cursor="hand2")
        
        title_label = tk.Label(zone_content, text="拖拽或点击上传清单", font=("SF Pro Display", 16, "bold"), bg="#FFFFFF", fg="#1F2937")
        title_label.pack()
        title_label.bind("<Button-1>", lambda e: self.browse_file())
        title_label.config(cursor="hand2")
        
        desc_label = tk.Label(zone_content, text="支持 .xlsx, .csv 格式，单次处理上限 10,000 行", font=("SF Pro Display", 14), bg="#FFFFFF", fg="#6B7280")
        desc_label.pack(pady=(8, 0))
        desc_label.bind("<Button-1>", lambda e: self.browse_file())
        desc_label.config(cursor="hand2")
        
        self.file_name_label = tk.Label(zone_content, text="", font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#2563EB")
        self.file_name_label.pack(pady=(12, 0))
        self.file_name_label.bind("<Button-1>", lambda e: self.browse_file())
        self.file_name_label.config(cursor="hand2")
        
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
        result_card.pack(fill=tk.X, pady=(0, 16))
        
        header_frame = tk.Frame(result_card, bg="#FFFFFF")
        header_frame.pack(fill=tk.X, pady=14, padx=20)
        
        left_frame = tk.Frame(header_frame, bg="#FFFFFF")
        left_frame.pack(side=tk.LEFT)
        
        icon_label = tk.Label(left_frame, text="📋", font=("SF Pro Display", 16), bg="#FFFFFF", fg="#2563EB")
        icon_label.pack(side=tk.LEFT)
        
        result_title = tk.Label(left_frame, text="分配结果预览（前20行）", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
        result_title.pack(side=tk.LEFT, padx=(8, 0))
        
        right_frame = tk.Frame(header_frame, bg="#FFFFFF")
        right_frame.pack(side=tk.RIGHT)
        
        export_btn = tk.Button(right_frame, text="导出清单", font=("SF Pro Display", 12), bg="#F3F4F6", fg="#4B5563", 
                               relief=tk.FLAT, padx=14, pady=6, cursor="hand2", command=self.save_result)
        export_btn.pack(side=tk.LEFT)
        
        tree_container = tk.Frame(result_card, bg="#FFFFFF")
        tree_container.pack(fill=tk.X, padx=20, pady=(0, 14))
        
        self.tree = ttk.Treeview(tree_container, show="headings", height=20)
        self.tree.pack(fill=tk.X)
    
    def create_status_bar(self, parent):
        status_frame = tk.Frame(parent, bg="#FFFFFF")
        status_frame.pack(fill=tk.X)
        status_frame.config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB", highlightthickness=1)
        
        self.status_label = tk.Label(status_frame, text="  ○ 等待操作", font=("SF Pro Display", 12), bg="#FFFFFF", fg="#9CA3AF", anchor="w")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=16, pady=12)
        
        sync_label = tk.Label(status_frame, text="上次同步: 2026-04-28 10:45", font=("SF Pro Display", 11), bg="#FFFFFF", fg="#9CA3AF")
        sync_label.pack(side=tk.RIGHT, padx=16, pady=12)
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            file_name = os.path.basename(file_path)
            file_dir = os.path.dirname(file_path)
            if len(file_name) > 40:
                display_name = file_name[:20] + "..." + file_name[-15:]
            else:
                display_name = file_name
            self.file_name_label.config(text=f"📄 {display_name}")
            self.update_status(f"✓ 文件已选择: {file_name}", "#2563EB")
            self.run_btn.config(state=tk.NORMAL, cursor="hand2")
    
    def update_status(self, text, color="#6B7280"):
        self.status_label.config(text=f"  {text}", fg=color)
    
    def run_allocation(self):
        if not self.file_path:
            messagebox.showwarning("提示", "请先选择Excel文件")
            return
        
        try:
            self.update_status("● 处理中...", "#D97706")
            self.root.update()
            
            df_inventory = pd.read_excel(self.file_path, sheet_name="库存")
            df_sales = pd.read_excel(self.file_path, sheet_name="销售")
            df_store_level = pd.read_excel(self.file_path, sheet_name="卖场等级")
            df_add_order = pd.read_excel(self.file_path, sheet_name="加单数量")
            
            config = self.config if self.config else {}
            if "allocation_config" not in config:
                config["allocation_config"] = {}
            config["allocation_config"]["stage_priority"] = [stage[0] for stage in self.stage_list[:3]]
            
            allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
                df_inventory, df_sales, df_store_level, df_add_order, config
            )
            
            self.result_df, self.reason_df = generate_result_dataframe(
                allocation_result, allocation_reasons, stores_sorted, skus, store_level_map,
                stage_priority=[stage[0] for stage in self.stage_list[:3]]
            )
            
            self.display_result()
            
            self.update_status("✓ 分配完成", "#059669")
            self.save_btn.config(state=tk.NORMAL, cursor="hand2")
            
            # 延迟显示消息框，避免 Mac Tkinter 崩溃
            self.root.after(100, lambda: messagebox.showinfo("成功", "加单分配完成"))
        except Exception as e:
            print(f"分配错误: {e}")
            traceback.print_exc()
            self.update_status(f"✗ 执行失败: {str(e)}", "#DC2626")
            # 延迟显示消息框，避免 Mac Tkinter 崩溃
            self.root.after(100, lambda err=str(e): messagebox.showerror("错误", f"执行分配失败:\n{err}"))
    
    def display_result(self):
        if self.result_df is None or len(self.result_df) == 0:
            return
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        columns = list(self.result_df.columns)
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        for _, row in self.result_df.head(20).iterrows():
            self.tree.insert("", tk.END, values=list(row))
    
    def save_result(self):
        if self.result_df is None:
            messagebox.showwarning("提示", "请先执行分配")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存分配结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            try:
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    self.result_df.to_excel(writer, sheet_name="分配数量", index=False)
                    self.reason_df.to_excel(writer, sheet_name="分配原因", index=False)
                messagebox.showinfo("成功", f"结果已保存到:\n{file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存结果失败:\n{str(e)}")
    
    def open_donate(self, event=None):
        import webbrowser
        webbrowser.open("https://duiliuliu.github.io/sponsor-page/")

if __name__ == "__main__":
    try:
        app = AllocationApp()
        app.root.mainloop()
    except Exception as e:
        print(f"启动错误: {e}")
        traceback.print_exc()
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("启动错误", f"程序启动失败:\n{str(e)}")
            root.destroy()
        except:
            pass
