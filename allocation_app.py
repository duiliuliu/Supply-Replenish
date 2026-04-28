# 加单商品分配系统 v2.3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
import sys
import traceback

# 尝试导入核心逻辑
try:
    from allocation_core import allocate_add_order, generate_result_dataframe, load_config, DEFAULT_CONFIG
except Exception as e:
    print(f'Failed to import allocation_core: {e}')
    import traceback
    traceback.print_exc()
    # 如果导入失败，创建默认函数
    def allocate_add_order(*args, **kwargs):
        from collections import defaultdict
        return defaultdict(lambda: defaultdict(int)), defaultdict(lambda: defaultdict(str)), [], [], {}
    
    def generate_result_dataframe(*args, **kwargs):
        return pd.DataFrame(), pd.DataFrame()
    
    load_config = lambda: DEFAULT_CONFIG

class AllocationApp:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("加单商品分配系统")
            self.root.geometry("1100x980")
            self.root.configure(bg="#F5F7FA")
            self.root.minsize(950, 820)
            
            # 设置关闭窗口的处理
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
                ("broken_size_fix", "断码修复"),
                ("sales_match", "销量匹配"),
                ("sell_through_priority", "销尽率优先")
            ]
            
            # 拖拽相关变量
            self.drag_data = {"index": -1, "item": None, "y": 0}
            
            self.setup_styles()
            self.create_widgets()
            
            print('应用程序初始化成功')
        except Exception as e:
            print(f'初始化错误: {e}')
            import traceback
            traceback.print_exc()
            messagebox.showerror("初始化错误", f"程序初始化失败:\n{str(e)}\n\n请查看控制台输出获取详细信息")
    
    def on_closing(self):
        try:
            print('关闭窗口')
            self.root.destroy()
        except Exception as e:
            print(f'关闭错误: {e}')
            sys.exit(0)
    
    def setup_styles(self):
        try:
            style = ttk.Style()
            style.theme_use("clam")
            
            style.configure("Title.TLabel", font=("SF Pro Display", 24, "bold"), background="#F5F7FA", foreground="#1F2937")
            style.configure("Section.TLabelframe", font=("SF Pro Display", 15, "bold"), background="#F5F7FA", borderwidth=0)
            style.configure("Section.TLabelframe.Label", font=("SF Pro Display", 15, "bold"), foreground="#1F2937", background="#F5F7FA")
            style.configure("Config.TLabel", font=("SF Pro Display", 13), background="#F5F7FA", foreground="#4B5563")
            style.configure("Status.TLabel", font=("SF Pro Display", 13), background="#F5F7FA", foreground="#4B5563")
            
            style.configure("Primary.TButton", font=("SF Pro Display", 13, "bold"), padding=(28, 14))
            style.configure("Secondary.TButton", font=("SF Pro Display", 12), padding=(16, 8))
            
            style.configure("Treeview", font=("SF Pro Display", 12), rowheight=32, background="#FFFFFF", foreground="#1F2937", borderwidth=0)
            style.configure("Treeview.Heading", font=("SF Pro Display", 12, "bold"), background="#F9FAFB", foreground="#374151", borderwidth=0, relief=tk.FLAT)
            style.map("Treeview", background=[("selected", "#EFF6FF")], foreground=[("selected", "#2563EB")])
        except Exception as e:
            print(f'setup_styles error: {e}')
    
    def create_widgets(self):
        try:
            main_frame = tk.Frame(self.root, bg="#F5F7FA")
            main_frame.pack(fill=tk.BOTH, expand=True, padx=32, pady=28)
            
            self.create_header(main_frame)
            self.create_config_section(main_frame)
            self.create_file_section(main_frame)
            self.create_logic_section(main_frame)
            self.create_button_section(main_frame)
            self.create_status_section(main_frame)
            self.create_result_section(main_frame)
        except Exception as e:
            print(f'create_widgets error: {e}')
            import traceback
            traceback.print_exc()
    
    def create_header(self, parent):
        try:
            header_frame = tk.Frame(parent, bg="#F5F7FA")
            header_frame.pack(fill=tk.X, pady=(0, 28))
            
            title_label = tk.Label(header_frame, text="加单商品分配系统", font=("SF Pro Display", 24, "bold"), bg="#F5F7FA", fg="#1F2937")
            title_label.pack(side=tk.LEFT)
            
            subtitle_label = tk.Label(header_frame, text="基于智能算法的库存分配工具", font=("SF Pro Display", 16), bg="#F5F7FA", fg="#6B7280")
            subtitle_label.pack(side=tk.LEFT, padx=(20, 0), pady=(8, 0))
            
            version_label = tk.Label(header_frame, text="v2.3", font=("SF Pro Display", 13), bg="#F5F7FA", fg="#9CA3AF")
            version_label.pack(side=tk.RIGHT, pady=(8, 0))
        except Exception as e:
            print(f'create_header error: {e}')
    
    def create_card_frame(self, parent, with_border=True):
        try:
            card = tk.Frame(parent, bg="#FFFFFF")
            if with_border:
                card.config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB", highlightthickness=1)
            return card
        except Exception as e:
            print(f'create_card_frame error: {e}')
            return tk.Frame(parent, bg="#FFFFFF")
    
    def create_config_section(self, parent):
        try:
            config_frame = tk.Frame(parent, bg="#F5F7FA")
            config_frame.pack(fill=tk.X, pady=(0, 20))
            
            config_card = self.create_card_frame(config_frame)
            config_card.pack(fill=tk.X)
            
            self.config_expanded = False
            
            header_frame = tk.Frame(config_card, bg="#FFFFFF")
            header_frame.pack(fill=tk.X, pady=16, padx=24)
            header_frame.bind("<Button-1>", self.toggle_config)
            header_frame.config(cursor="hand2")
            
            self.config_toggle = tk.Label(header_frame, text="▶", font=("SF Pro Display", 13), bg="#FFFFFF", fg="#6B7280")
            self.config_toggle.pack(side=tk.RIGHT)
            
            config_title = tk.Label(header_frame, text="参数配置", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
            config_title.pack(side=tk.LEFT)
            
            self.config_content = tk.Frame(config_card, bg="#FFFFFF")
            
            levels = ["SA", "A", "B", "C", "D", "OL"]
            
            coverage_frame = tk.Frame(self.config_content, bg="#FFFFFF")
            coverage_frame.pack(fill=tk.X, pady=(0, 18))
            
            tk.Label(coverage_frame, text="覆盖周期（天）", font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#374151").pack(side=tk.LEFT, anchor=tk.N)
            
            self.coverage_entries = {}
            for i, level in enumerate(levels):
                frame = tk.Frame(coverage_frame, bg="#FFFFFF")
                frame.pack(side=tk.LEFT, padx=(32 if i == 0 else 20, 0))
                
                tk.Label(frame, text=level, font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280").pack(side=tk.TOP, pady=(0, 6))
                
                entry = tk.Entry(frame, width=9, font=("SF Pro Display", 12), justify="center", bg="#F9FAFB", relief=tk.FLAT, borderwidth=0, highlightthickness=1, highlightbackground="#E5E7EB", highlightcolor="#D1D5DB")
                entry.pack(side=tk.TOP, ipady=8)
                try:
                    val = str(self.config.get("allocation_config", {}).get("coverage_days", {}).get(level, 14) if self.config else 14)
                except:
                    val = "14"
                entry.insert(0, val)
                self.coverage_entries[level] = entry
            
            weight_frame = tk.Frame(self.config_content, bg="#FFFFFF")
            weight_frame.pack(fill=tk.X, pady=(0, 18))
            
            tk.Label(weight_frame, text="等级权重", font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#374151").pack(side=tk.LEFT, anchor=tk.N)
            
            self.weight_entries = {}
            for i, level in enumerate(levels):
                frame = tk.Frame(weight_frame, bg="#FFFFFF")
                frame.pack(side=tk.LEFT, padx=(32 if i == 0 else 20, 0))
                
                tk.Label(frame, text=level, font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280").pack(side=tk.TOP, pady=(0, 6))
                
                entry = tk.Entry(frame, width=9, font=("SF Pro Display", 12), justify="center", bg="#F9FAFB", relief=tk.FLAT, borderwidth=0, highlightthickness=1, highlightbackground="#E5E7EB", highlightcolor="#D1D5DB")
                entry.pack(side=tk.TOP, ipady=8)
                try:
                    val = str(self.config.get("allocation_config", {}).get("level_weights", {}).get(level, 1.0) if self.config else 1.0)
                except:
                    val = "1.0"
                entry.insert(0, val)
                self.weight_entries[level] = entry
            
            safety_frame = tk.Frame(self.config_content, bg="#FFFFFF")
            safety_frame.pack(fill=tk.X, pady=(0, 18))
            
            tk.Label(safety_frame, text="安全系数", font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#374151").pack(side=tk.LEFT, anchor=tk.N)
            
            self.safety_entries = {}
            for i, level in enumerate(levels):
                frame = tk.Frame(safety_frame, bg="#FFFFFF")
                frame.pack(side=tk.LEFT, padx=(32 if i == 0 else 20, 0))
                
                tk.Label(frame, text=level, font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280").pack(side=tk.TOP, pady=(0, 6))
                
                entry = tk.Entry(frame, width=9, font=("SF Pro Display", 12), justify="center", bg="#F9FAFB", relief=tk.FLAT, borderwidth=0, highlightthickness=1, highlightbackground="#E5E7EB", highlightcolor="#D1D5DB")
                entry.pack(side=tk.TOP, ipady=8)
                try:
                    val = str(self.config.get("allocation_config", {}).get("safety_factors", {}).get(level, 0.3) if self.config else 0.3)
                except:
                    val = "0.3"
                entry.insert(0, val)
                self.safety_entries[level] = entry
            
            min_target_frame = tk.Frame(self.config_content, bg="#FFFFFF")
            min_target_frame.pack(fill=tk.X, pady=(0, 18))
            
            tk.Label(min_target_frame, text="最小目标库存", font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#374151").pack(side=tk.LEFT, anchor=tk.N)
            
            self.min_target_entries = {}
            for i, level in enumerate(levels):
                frame = tk.Frame(min_target_frame, bg="#FFFFFF")
                frame.pack(side=tk.LEFT, padx=(32 if i == 0 else 20, 0))
                
                tk.Label(frame, text=level, font=("SF Pro Display", 12), bg="#FFFFFF", fg="#6B7280").pack(side=tk.TOP, pady=(0, 6))
                
                entry = tk.Entry(frame, width=9, font=("SF Pro Display", 12), justify="center", bg="#F9FAFB", relief=tk.FLAT, borderwidth=0, highlightthickness=1, highlightbackground="#E5E7EB", highlightcolor="#D1D5DB")
                entry.pack(side=tk.TOP, ipady=8)
                try:
                    val = str(self.config.get("allocation_config", {}).get("min_target_inventory", {}).get(level, 0) if self.config else 0)
                except:
                    val = "0"
                entry.insert(0, val)
                self.min_target_entries[level] = entry
            
            stage_priority_frame = tk.Frame(self.config_content, bg="#FFFFFF")
            stage_priority_frame.pack(fill=tk.X, pady=(0, 20))
            
            tk.Label(stage_priority_frame, text="阶段优先级", font=("SF Pro Display", 13, "bold"), bg="#FFFFFF", fg="#374151").pack(anchor=tk.W, pady=(0, 10))
            
            stage_display_frame = tk.Frame(stage_priority_frame, bg="#FFFFFF")
            stage_display_frame.pack(fill=tk.X)
            
            self.stage_labels = []
            self.stage_display_frame = stage_display_frame
            try:
                current_priority = self.config.get("allocation_config", {}).get("stage_priority", ["broken_size_fix", "sales_match", "sell_through_priority"]) if self.config else ["broken_size_fix", "sales_match", "sell_through_priority"]
            except:
                current_priority = ["broken_size_fix", "sales_match", "sell_through_priority"]
            
            for i, stage_id in enumerate(current_priority):
                self._create_stage_item(i, stage_id)
            
            btn_frame = tk.Frame(self.config_content, bg="#FFFFFF")
            btn_frame.pack(fill=tk.X, pady=(0, 10))
            
            reset_btn = self.create_button(btn_frame, "恢复默认值", style="secondary")
            reset_btn.pack(side=tk.RIGHT)
            reset_btn.config(command=self.reset_config)
            
            save_btn = self.create_button(btn_frame, "保存配置", style="primary")
            save_btn.pack(side=tk.RIGHT, padx=(14, 0))
            save_btn.config(command=self.save_config)
        except Exception as e:
            print(f'create_config_section error: {e}')
            import traceback
            traceback.print_exc()
    
    def _create_stage_item(self, idx, stage_id):
        try:
            bg_color, fg_color = self.stage_colors[idx]
            stage_name = [name for id, name in self.stage_list if id == stage_id][0]
            
            stage_item = tk.Frame(self.stage_display_frame, bg=bg_color, bd=0, highlightthickness=2, highlightbackground="#E5E7EB")
            stage_item.pack(fill=tk.X, pady=(0, 8))
            
            # 设置拖拽处理
            stage_item.bind("<ButtonPress-1>", lambda e, idx=idx: self._on_stage_press(e, idx))
            stage_item.bind("<B1-Motion>", lambda e, idx=idx: self._on_stage_drag(e, idx))
            stage_item.bind("<ButtonRelease-1>", lambda e, idx=idx: self._on_stage_release(e, idx))
            stage_item.config(cursor="grab")
            
            tk.Label(stage_item, text=f"{idx+1}", font=("SF Pro Display", 12, "bold"), bg=fg_color, fg="white", width=4, padx=6, pady=8).pack(side=tk.LEFT)
            
            stage_info = tk.Frame(stage_item, bg=bg_color)
            stage_info.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=14, pady=8)
            tk.Label(stage_info, text=stage_name, font=("SF Pro Display", 12, "bold"), bg=bg_color, fg=fg_color).pack(anchor=tk.W)
            tk.Label(stage_info, text="拖拽调整顺序", font=("SF Pro Display", 10), bg=bg_color, fg="#9CA3AF").pack(anchor=tk.W, pady=(2, 0))
            
            btn_container = tk.Frame(stage_item, bg=bg_color)
            btn_container.pack(side=tk.RIGHT, padx=10)
            
            if idx > 0:
                up_btn = tk.Button(btn_container, text="↑", font=("SF Pro Display", 11, "bold"), bg="#FFFFFF", fg=fg_color, relief=tk.FLAT, padx=10, pady=6, cursor="hand2", activebackground="#E5E7EB")
                up_btn.pack(side=tk.LEFT, padx=(0, 5))
                up_btn.config(command=lambda idx=idx: self.move_stage(idx, -1))
            
            if idx < len(self.stage_list) - 1:
                down_btn = tk.Button(btn_container, text="↓", font=("SF Pro Display", 11, "bold"), bg="#FFFFFF", fg=fg_color, relief=tk.FLAT, padx=10, pady=6, cursor="hand2", activebackground="#E5E7EB")
                down_btn.pack(side=tk.LEFT)
                down_btn.config(command=lambda idx=idx: self.move_stage(idx, 1))
            
            self.stage_labels.append((stage_id, stage_item))
        except Exception as e:
            print(f'_create_stage_item error: {e}')
    
    def _on_stage_press(self, event, index):
        try:
            self.drag_data["index"] = index
            self.drag_data["y"] = event.y
            self.drag_data["item"] = self.stage_labels[index][1]
            self.drag_data["item"].config(highlightbackground="#2563EB", highlightcolor="#2563EB")
            self.drag_data["item"].config(cursor="grabbing")
        except Exception as e:
            print(f'_on_stage_press error: {e}')
    
    def _on_stage_drag(self, event, index):
        try:
            if self.drag_data["index"] == -1:
                return
            
            # 计算鼠标移动方向
            delta_y = event.y - self.drag_data["y"]
            if abs(delta_y) > 20:
                direction = 1 if delta_y > 0 else -1
                new_index = index + direction
                
                if 0 <= new_index < len(self.stage_labels):
                    self.move_stage(index, direction)
                    self.drag_data["index"] = new_index
                    self.drag_data["y"] = event.y
        except Exception as e:
            print(f'_on_stage_drag error: {e}')
    
    def _on_stage_release(self, event, index):
        try:
            if self.drag_data["item"]:
                self.drag_data["item"].config(highlightbackground="#E5E7EB", highlightcolor="#E5E7EB")
                self.drag_data["item"].config(cursor="grab")
            self.drag_data = {"index": -1, "item": None, "y": 0}
        except Exception as e:
            print(f'_on_stage_release error: {e}')
    
    def create_button(self, parent, text, style="primary"):
        try:
            if style == "primary":
                btn = tk.Button(parent, text=text, font=("SF Pro Display", 12, "bold"), bg="#2563EB", fg="white", relief=tk.FLAT, padx=24, pady=10, cursor="hand2", activebackground="#1D4ED8", activeforeground="white")
                btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#1E40AF"))
                btn.bind("<Leave>", lambda e, b=btn: b.config(bg="#2563EB"))
                btn.bind("<ButtonPress>", lambda e, b=btn: b.config(bg="#1D4ED8"))
                btn.bind("<ButtonRelease>", lambda e, b=btn: b.config(bg="#1E40AF"))
            else:
                btn = tk.Button(parent, text=text, font=("SF Pro Display", 12), bg="#F3F4F6", fg="#4B5563", relief=tk.FLAT, padx=24, pady=10, cursor="hand2", activebackground="#CBD5E1", activeforeground="#374151")
                btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#E5E7EB"))
                btn.bind("<Leave>", lambda e, b=btn: b.config(bg="#F3F4F6"))
                btn.bind("<ButtonPress>", lambda e, b=btn: b.config(bg="#D1D5DB"))
                btn.bind("<ButtonRelease>", lambda e, b=btn: b.config(bg="#E5E7EB"))
            return btn
        except Exception as e:
            print(f'create_button error: {e}')
            return tk.Button(parent, text=text)
    
    def move_stage(self, idx, direction):
        try:
            current_priority = [stage[0] for stage in self.stage_labels]
            new_idx = idx + direction
            
            if 0 <= new_idx < len(current_priority):
                current_priority[new_idx], current_priority[idx] = current_priority[idx], current_priority[new_idx]
                
                for stage_id, widget in self.stage_labels:
                    widget.pack_forget()
                self.stage_labels = []
                
                for i, stage_id in enumerate(current_priority):
                    self._create_stage_item(i, stage_id)
        except Exception as e:
            print(f'move_stage error: {e}')
    
    def toggle_config(self, event=None):
        try:
            if self.config_expanded:
                self.config_content.pack_forget()
                self.config_toggle.config(text="▶")
            else:
                self.config_content.pack(fill=tk.X, padx=24, pady=(0, 16))
                self.config_toggle.config(text="▼")
            self.config_expanded = not self.config_expanded
        except Exception as e:
            print(f'toggle_config error: {e}')
    
    def reset_config(self):
        try:
            default_config = DEFAULT_CONFIG["allocation_config"]
            levels = ["SA", "A", "B", "C", "D", "OL"]
            for level in levels:
                if level in self.coverage_entries:
                    self.coverage_entries[level].delete(0, tk.END)
                    self.coverage_entries[level].insert(0, str(default_config["coverage_days"].get(level, 14)))
                if level in self.weight_entries:
                    self.weight_entries[level].delete(0, tk.END)
                    self.weight_entries[level].insert(0, str(default_config["level_weights"].get(level, 1.0)))
                if level in self.safety_entries:
                    self.safety_entries[level].delete(0, tk.END)
                    self.safety_entries[level].insert(0, str(default_config["safety_factors"].get(level, 0.3)))
                if level in self.min_target_entries:
                    self.min_target_entries[level].delete(0, tk.END)
                    self.min_target_entries[level].insert(0, str(default_config["min_target_inventory"].get(level, 0)))
            
            default_priority = default_config.get("stage_priority", ["broken_size_fix", "sales_match", "sell_through_priority"])
            for stage_id, widget in self.stage_labels:
                widget.pack_forget()
            self.stage_labels = []
            
            for i, stage_id in enumerate(default_priority):
                self._create_stage_item(i, stage_id)
            
            messagebox.showinfo("成功", "已恢复默认配置")
        except Exception as e:
            print(f'reset_config error: {e}')
            messagebox.showerror("错误", f"恢复默认配置失败: {str(e)}")
    
    def save_config(self):
        try:
            config = {
                "version": "2.3",
                "updated_at": "2026-04-28",
                "allocation_config": {
                    "coverage_days": {},
                    "level_weights": {},
                    "safety_factors": {},
                    "min_target_inventory": {},
                    "stage_priority": [stage[0] for stage in self.stage_labels],
                    "max_remaining_per_store": 10
                }
            }
            
            for level in ["SA", "A", "B", "C", "D", "OL"]:
                try:
                    config["allocation_config"]["coverage_days"][level] = int(self.coverage_entries[level].get())
                except:
                    config["allocation_config"]["coverage_days"][level] = 14
                try:
                    config["allocation_config"]["level_weights"][level] = float(self.weight_entries[level].get())
                except:
                    config["allocation_config"]["level_weights"][level] = 1.0
                try:
                    config["allocation_config"]["safety_factors"][level] = float(self.safety_entries[level].get())
                except:
                    config["allocation_config"]["safety_factors"][level] = 0.3
                try:
                    config["allocation_config"]["min_target_inventory"][level] = int(self.min_target_entries[level].get())
                except:
                    config["allocation_config"]["min_target_inventory"][level] = 0
            
            config_path = os.path.join(os.path.dirname(__file__), "allocation_config.json")
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.config = config
            messagebox.showinfo("成功", "配置已保存!")
        except Exception as e:
            print(f'save_config error: {e}')
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
    
    def create_logic_section(self, parent):
        try:
            logic_frame = tk.Frame(parent, bg="#F5F7FA")
            logic_frame.pack(fill=tk.X, pady=(0, 20))
            
            logic_card = self.create_card_frame(logic_frame)
            logic_card.pack(fill=tk.X)
            
            self.logic_expanded = True
            
            header_frame = tk.Frame(logic_card, bg="#FFFFFF")
            header_frame.pack(fill=tk.X, pady=16, padx=24)
            header_frame.bind("<Button-1>", self.toggle_logic)
            header_frame.config(cursor="hand2")
            
            self.logic_toggle = tk.Label(header_frame, text="▼", font=("SF Pro Display", 13), bg="#FFFFFF", fg="#6B7280")
            self.logic_toggle.pack(side=tk.RIGHT)
            
            logic_title = tk.Label(header_frame, text="分配逻辑说明", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
            logic_title.pack(side=tk.LEFT)
            
            self.logic_content = tk.Frame(logic_card, bg="#FFFFFF")
            self.logic_content.pack(fill=tk.X, padx=24, pady=(0, 16))
            
            stages = [
                ("阶段1", "断码修复", "SA/A级：核心尺码(160,165)≥2件，非核心≥1件；其他等级：核心尺码≥1件"),
                ("阶段2", "销量匹配", "目标库存=平均日需求×覆盖周期+安全库存，且不低于最小目标库存"),
                ("阶段3", "销尽率优先", "综合得分=销尽率×等级权重，所有等级参与，按得分降序分配"),
                ("阶段4", "剩余分配", "按等级优先级SA→A→B→C→D→OL分配，单卖场上限10件")
            ]
            
            for i, (stage_id, name, desc) in enumerate(stages):
                bg_color, fg_color = self.stage_colors[i]
                
                stage_frame = tk.Frame(self.logic_content, bg=bg_color)
                stage_frame.pack(fill=tk.X, pady=(0 if i == 0 else 12, 0))
                
                stage_label = tk.Label(stage_frame, text=stage_id, font=("SF Pro Display", 12, "bold"), bg=fg_color, fg="white", width=9, padx=6, pady=10)
                stage_label.pack(side=tk.LEFT)
                
                content_frame = tk.Frame(stage_frame, bg=bg_color)
                content_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, pady=10, padx=16)
                
                tk.Label(content_frame, text=name, font=("SF Pro Display", 12, "bold"), bg=bg_color, fg=fg_color).pack(anchor=tk.W)
                
                desc_label = tk.Label(content_frame, text=desc, font=("SF Pro Display", 11), bg=bg_color, fg="#4B5563")
                desc_label.pack(anchor=tk.W, pady=(4, 0))
        except Exception as e:
            print(f'create_logic_section error: {e}')
    
    def toggle_logic(self, event=None):
        try:
            if self.logic_expanded:
                self.logic_content.pack_forget()
                self.logic_toggle.config(text="▶")
            else:
                self.logic_content.pack(fill=tk.X, padx=24, pady=(0, 16))
                self.logic_toggle.config(text="▼")
            self.logic_expanded = not self.logic_expanded
        except Exception as e:
            print(f'toggle_logic error: {e}')
    
    def create_file_section(self, parent):
        try:
            file_frame = tk.Frame(parent, bg="#F5F7FA")
            file_frame.pack(fill=tk.X, pady=(0, 20))
            
            file_card = self.create_card_frame(file_frame)
            file_card.pack(fill=tk.X)
            
            inner_frame = tk.Frame(file_card, bg="#FFFFFF")
            inner_frame.pack(fill=tk.X, pady=16, padx=24)
            
            tk.Label(inner_frame, text="文件选择", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937").pack(anchor=tk.W, pady=(0, 12))
            
            input_row = tk.Frame(inner_frame, bg="#FFFFFF")
            input_row.pack(fill=tk.X)
            
            self.file_entry = tk.Entry(input_row, font=("SF Pro Display", 13), bg="#F9FAFB", relief=tk.FLAT, borderwidth=0, highlightthickness=1, highlightbackground="#E5E7EB", highlightcolor="#D1D5DB", disabledbackground="#F3F4F6")
            self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 14), ipady=10)
            
            browse_btn = self.create_button(input_row, "浏览文件", style="primary")
            browse_btn.pack(side=tk.RIGHT)
            browse_btn.config(command=self.browse_file)
        except Exception as e:
            print(f'create_file_section error: {e}')
    
    def create_button_section(self, parent):
        try:
            btn_frame = tk.Frame(parent, bg="#F5F7FA")
            btn_frame.pack(fill=tk.X, pady=(0, 20))
            
            self.run_btn = tk.Button(btn_frame, text="执行加单分配", font=("SF Pro Display", 14, "bold"), bg="#D1D5DB", fg="#9CA3AF", relief=tk.FLAT, padx=40, pady=14, cursor="arrow", state=tk.DISABLED, activebackground="#D1D5DB", activeforeground="#9CA3AF")
            self.run_btn.pack(side=tk.LEFT, padx=(0, 18))
            
            self.save_btn = tk.Button(btn_frame, text="保存结果", font=("SF Pro Display", 14, "bold"), bg="#D1D5DB", fg="#9CA3AF", relief=tk.FLAT, padx=40, pady=14, cursor="arrow", state=tk.DISABLED, activebackground="#D1D5DB", activeforeground="#9CA3AF")
            self.save_btn.pack(side=tk.LEFT)
        except Exception as e:
            print(f'create_button_section error: {e}')
    
    def create_status_section(self, parent):
        try:
            status_frame = tk.Frame(parent, bg="#F5F7FA")
            status_frame.pack(fill=tk.X, pady=(0, 20))
            
            self.status_icon = tk.Label(status_frame, text="○", font=("SF Pro Display", 14), bg="#F5F7FA", fg="#9CA3AF")
            self.status_icon.pack(side=tk.LEFT)
            
            self.status_label = tk.Label(status_frame, text="等待文件选择...", font=("SF Pro Display", 13), bg="#F5F7FA", fg="#6B7280")
            self.status_label.pack(side=tk.LEFT, padx=(10, 0))
        except Exception as e:
            print(f'create_status_section error: {e}')
    
    def create_result_section(self, parent):
        try:
            result_frame = tk.Frame(parent, bg="#F5F7FA")
            result_frame.pack(fill=tk.BOTH, expand=True)
            
            result_card = self.create_card_frame(result_frame)
            result_card.pack(fill=tk.BOTH, expand=True)
            
            header_frame = tk.Frame(result_card, bg="#FFFFFF")
            header_frame.pack(fill=tk.X, pady=16, padx=24)
            
            result_title = tk.Label(header_frame, text="分配结果预览", font=("SF Pro Display", 15, "bold"), bg="#FFFFFF", fg="#1F2937")
            result_title.pack(side=tk.LEFT)
            
            result_subtitle = tk.Label(header_frame, text="(前20行)", font=("SF Pro Display", 13), bg="#FFFFFF", fg="#9CA3AF")
            result_subtitle.pack(side=tk.LEFT, padx=(10, 0))
            
            tree_container = tk.Frame(result_card, bg="#FFFFFF")
            tree_container.pack(fill=tk.BOTH, expand=True, padx=24, pady=(0, 16))
            
            self.tree = ttk.Treeview(tree_container, show="headings", borderwidth=0)
            self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.tree.configure(yscrollcommand=scrollbar.set)
        except Exception as e:
            print(f'create_result_section error: {e}')
    
    def browse_file(self):
        try:
            self.file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
            )
            
            if self.file_path:
                self.file_entry.delete(0, tk.END)
                self.file_entry.insert(0, self.file_path)
                self.enable_run_button()
                self.update_status("文件已选择，点击执行开始分配", "success")
        except Exception as e:
            print(f'browse_file error: {e}')
            self.update_status("文件选择失败: " + str(e), "error")
            messagebox.showerror("错误", "文件选择失败:\n" + str(e))
    
    def update_status(self, text, status="info"):
        try:
            colors = {
                "info": ("#6B7280", "○"),
                "success": ("#059669", "✓"),
                "warning": ("#D97706", "●"),
                "error": ("#DC2626", "✗")
            }
            fg_color, icon = colors.get(status, colors["info"])
            self.status_label.config(text=text, fg=fg_color)
            self.status_icon.config(text=icon, fg=fg_color)
        except Exception as e:
            print(f'update_status error: {e}')
    
    def enable_run_button(self):
        try:
            self.run_btn.config(state=tk.NORMAL, bg="#2563EB", fg="white", cursor="hand2", command=self.run_allocation, activebackground="#1D4ED8", activeforeground="white")
            self.run_btn.bind("<Enter>", lambda e: self.run_btn.config(bg="#1E40AF"))
            self.run_btn.bind("<Leave>", lambda e: self.run_btn.config(bg="#2563EB"))
            self.run_btn.bind("<ButtonPress>", lambda e: self.run_btn.config(bg="#1D4ED8"))
            self.run_btn.bind("<ButtonRelease>", lambda e: self.run_btn.config(bg="#1E40AF"))
        except Exception as e:
            print(f'enable_run_button error: {e}')
    
    def enable_save_button(self):
        try:
            self.save_btn.config(state=tk.NORMAL, bg="#2563EB", fg="white", cursor="hand2", command=self.save_result, activebackground="#1D4ED8", activeforeground="white")
            self.save_btn.bind("<Enter>", lambda e: self.save_btn.config(bg="#1E40AF"))
            self.save_btn.bind("<Leave>", lambda e: self.save_btn.config(bg="#2563EB"))
            self.save_btn.bind("<ButtonPress>", lambda e: self.save_btn.config(bg="#1D4ED8"))
            self.save_btn.bind("<ButtonRelease>", lambda e: self.save_btn.config(bg="#1E40AF"))
        except Exception as e:
            print(f'enable_save_button error: {e}')
    
    def run_allocation(self):
        try:
            self.update_status("正在读取Excel文件...", "warning")
            self.root.update()
            
            df_inventory = pd.read_excel(self.file_path, sheet_name="库存")
            df_sales = pd.read_excel(self.file_path, sheet_name="销售")
            df_store_level = pd.read_excel(self.file_path, sheet_name="卖场等级")
            df_add_order = pd.read_excel(self.file_path, sheet_name="加单数量")
            
            self.update_status("正在执行分配逻辑...", "warning")
            self.root.update()
            
            allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
                df_inventory, df_sales, df_store_level, df_add_order, self.config
            )
            
            self.result_df, self.reason_df = generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus, store_level_map)
            
            self.show_preview()
            
            self.enable_save_button()
            self.update_status("分配完成! 可以保存结果", "success")
            
            messagebox.showinfo("成功", "加单分配完成!")
        except Exception as e:
            print(f'run_allocation error: {e}')
            import traceback
            traceback.print_exc()
            self.update_status("执行失败: " + str(e), "error")
            messagebox.showerror("错误", "执行失败:\n" + str(e))
    
    def show_preview(self):
        try:
            for col in self.tree.get_children():
                self.tree.delete(col)
            
            self.tree["columns"] = list(self.result_df.columns)
            
            for col in self.result_df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=110, anchor="center")
            
            for idx, row in self.result_df.head(20).iterrows():
                values = list(row)
                self.tree.insert("", "end", values=values)
        except Exception as e:
            print(f'show_preview error: {e}')
    
    def save_result(self):
        try:
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
        except Exception as e:
            print(f'save_result error: {e}')
            messagebox.showerror("错误", f"保存失败: {str(e)}")

def main():
    try:
        print('启动加单商品分配系统...')
        
        # 先尝试加载配置，确保有配置可用
        try:
            from allocation_core import load_config
            config = load_config()
            print('配置加载成功')
        except Exception as e:
            print(f'配置加载失败: {e}')
        
        root = tk.Tk()
        app = AllocationApp(root)
        print('主循环开始')
        root.mainloop()
        print('主循环结束')
    except Exception as e:
        print(f'程序主错误: {e}')
        import traceback
        traceback.print_exc()
        
        # 尝试显示错误消息框
        try:
            messagebox.showerror("程序错误", f"程序发生严重错误:\n{str(e)}\n\n详细信息已输出到控制台")
        except:
            pass
        
        sys.exit(1)

if __name__ == "__main__":
    main()
