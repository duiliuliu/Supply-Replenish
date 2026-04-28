# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from allocation_core import allocate_add_order, generate_result_dataframe

class AllocationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("加单商品分配系统 v3")
        self.root.geometry("950x900")
        
        self.file_path = None
        self.result_df = None
        self.reason_df = None
        
        self.create_widgets()
    
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="加单商品分配系统", font=("Arial", 18, "bold"))
        title_label.pack(pady=10)
        
        # 分配逻辑说明区 - 卡片式设计
        logic_section = tk.LabelFrame(self.root, text="📊 分配逻辑说明", font=("Arial", 12, "bold"), 
                                      bg="#FAFAFA", fg="#333", padx=10, pady=10)
        logic_section.pack(pady=10, padx=20, fill=tk.X)
        
        # 创建4个阶段卡片
        stages = [
            {
                "priority": "P1",
                "name": "断码修复",
                "color": "#E3F2FD",
                "border_color": "#1976D2",
                "description": "确保卖场达到最低库存要求",
                "details": "SA/A级卖场：核心尺码(160/165)至少2件，非核心尺码至少1件\n其他等级：核心尺码至少1件"
            },
            {
                "priority": "P2",
                "name": "销量匹配", 
                "color": "#E8F5E9",
                "border_color": "#388E3C",
                "description": "根据30天销量分配商品",
                "details": "确保库存数量与销量相匹配，保持合理库存水平"
            },
            {
                "priority": "P3",
                "name": "销尽率优先",
                "color": "#FFF8E1",
                "border_color": "#F57C00",
                "description": "按销尽率降序分配（仅限B/C/D/OL级）",
                "details": "销尽率 = 销量 / (销量+库存)，优先分配给销尽率高的卖场"
            },
            {
                "priority": "P4",
                "name": "剩余分配",
                "color": "#FCE4EC",
                "border_color": "#C2185B",
                "description": "分配剩余商品给所有卖场",
                "details": "确保每个卖场单个尺码库存不超过15件"
            }
        ]
        
        # 使用Frame来布局卡片
        cards_frame = tk.Frame(logic_section, bg="#FAFAFA")
        cards_frame.pack(fill=tk.X)
        
        for stage in stages:
            card = tk.Frame(cards_frame, bg=stage["color"], relief=tk.SOLID, 
                           borderwidth=2, highlightbackground=stage["border_color"], 
                           highlightthickness=1)
            card.pack(fill=tk.X, pady=5, padx=5)
            
            # 优先级标签
            priority_label = tk.Label(card, text=stage["priority"], 
                                     font=("Arial", 14, "bold"), bg=stage["border_color"], 
                                     fg="white", width=4, padx=5)
            priority_label.pack(side=tk.LEFT)
            
            # 内容区域
            content_frame = tk.Frame(card, bg=stage["color"])
            content_frame.pack(side=tk.LEFT, padx=10, fill=tk.X)
            
            # 名称和描述
            name_label = tk.Label(content_frame, text=stage["name"], 
                                  font=("Arial", 11, "bold"), bg=stage["color"], fg="#333")
            name_label.pack(anchor=tk.W)
            
            desc_label = tk.Label(content_frame, text=stage["description"], 
                                  font=("Arial", 10), bg=stage["color"], fg="#666")
            desc_label.pack(anchor=tk.W, pady=(2, 0))
            
            # 详细说明
            details_label = tk.Label(content_frame, text=stage["details"], 
                                     font=("Arial", 9), bg=stage["color"], fg="#888",
                                     justify=tk.LEFT, wraplength=700)
            details_label.pack(anchor=tk.W, pady=(2, 5))
        
        # 提示信息
        hint_frame = tk.Frame(logic_section, bg="#E3F2FD")
        hint_frame.pack(fill=tk.X, pady=(10, 0), padx=5)
        hint_label = tk.Label(hint_frame, text="💡 提示：分配按优先级顺序执行，如有剩余数量会继续下一阶段，可能有多个分配原因叠加",
                              font=("Arial", 9, "italic"), bg="#E3F2FD", fg="#1976D2")
        hint_label.pack(padx=10, pady=5)
        
        # 文件选择区
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, fill=tk.X, padx=20)
        
        tk.Label(file_frame, text="Excel文件:", font=("Arial", 11)).pack(side=tk.LEFT)
        self.file_entry = tk.Entry(file_frame, width=50, font=("Arial", 10))
        self.file_entry.pack(side=tk.LEFT, padx=5)
        browse_btn = tk.Button(file_frame, text="浏览", command=self.browse_file, 
                               font=("Arial", 10), bg="#616161", fg="white")
        browse_btn.pack(side=tk.LEFT)
        
        # 按钮区
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=15)
        
        self.run_btn = tk.Button(btn_frame, text="执行加单分配", command=self.run_allocation, 
                               font=("Arial", 12, "bold"), bg="#1B5E20", fg="white", width=18, 
                               relief=tk.RAISED, borderwidth=3, disabledforeground="#BDBDBD")
        self.run_btn.pack(side=tk.LEFT, padx=10)
        self.run_btn.config(state=tk.DISABLED, bg="#BDBDBD", fg="#757575")
        
        self.save_btn = tk.Button(btn_frame, text="保存结果", command=self.save_result, 
                                 font=("Arial", 12, "bold"), bg="#0D47A1", fg="white", width=12,
                                 relief=tk.RAISED, borderwidth=3, disabledforeground="#BDBDBD")
        self.save_btn.pack(side=tk.LEFT, padx=10)
        self.save_btn.config(state=tk.DISABLED, bg="#BDBDBD", fg="#757575")
        
        # 进度区
        self.progress_frame = tk.Frame(self.root)
        self.progress_frame.pack(pady=5)
        self.status_label = tk.Label(self.progress_frame, text="等待文件选择...", fg="#616161", font=("Arial", 10))
        self.status_label.pack()
        
        # 结果预览
        preview_frame = tk.Frame(self.root)
        preview_frame.pack(pady=15, fill=tk.BOTH, expand=True, padx=20)
        
        tk.Label(preview_frame, text="分配结果预览（前20行）", font=("Arial", 11, "bold")).pack(anchor=tk.W)
        
        self.tree = ttk.Treeview(preview_frame, show='headings')
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 说明标签
        info_label = tk.Label(self.root, 
                            text="说明：支持跨平台运行，分配完成后可保存结果到新Excel文件",
                            font=("Arial", 9), fg="#616161")
        info_label.pack(side=tk.BOTTOM, pady=10)
    
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        
        if self.file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.file_path)
            self.run_btn.config(state=tk.NORMAL, bg="#1B5E20", fg="white")
            self.status_label.config(text="文件已选择，点击执行开始分配", fg="#1976D2")
    
    def run_allocation(self):
        self.status_label.config(text="正在读取Excel文件...", fg="orange")
        self.root.update()
        
        try:
            # 读取数据
            df_inventory = pd.read_excel(self.file_path, sheet_name='库存')
            df_sales = pd.read_excel(self.file_path, sheet_name='销售')
            df_store_level = pd.read_excel(self.file_path, sheet_name='卖场等级')
            df_add_order = pd.read_excel(self.file_path, sheet_name='加单数量')
            
            self.status_label.config(text="正在执行分配逻辑...", fg="orange")
            self.root.update()
            
            # 执行分配
            allocation_result, allocation_reasons, stores_sorted, skus = allocate_add_order(
                df_inventory, df_sales, df_store_level, df_add_order
            )
            
            self.result_df, self.reason_df = generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus)
            
            self.show_preview()
            
            self.save_btn.config(state=tk.NORMAL, bg="#0D47A1", fg="white")
            self.status_label.config(text="分配完成！可以保存结果", fg="#2E7D32")
            
            messagebox.showinfo("成功", "加单分配完成！")
            
        except Exception as e:
            self.status_label.config(text="执行失败: " + str(e), fg="red")
            messagebox.showerror("错误", "执行失败:\n" + str(e))
    
    def show_preview(self):
        # 清空现有列
        for col in self.tree.get_children():
            self.tree.delete(col)
        
        # 设置列
        columns = list(self.result_df.columns)
        self.tree.config(columns=columns)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        # 添加数据（前20行）
        preview_data = self.result_df.head(20)
        for _, row in preview_data.iterrows():
            self.tree.insert('', 'end', values=list(row))
    
    def save_result(self):
        if self.result_df is None:
            return
        
        save_path = filedialog.asksaveasfilename(
            title="保存结果",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        
        if save_path:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    self.result_df.to_excel(writer, sheet_name='分配数量', index=False)
                    self.reason_df.to_excel(writer, sheet_name='分配原因', index=False)
                messagebox.showinfo("保存成功", f"结果已保存到:\n{save_path}")
            except Exception as e:
                messagebox.showerror("保存失败", str(e))

def main():
    root = tk.Tk()
    app = AllocationApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
