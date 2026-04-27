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
        self.root.geometry("800x600")
        
        self.file_path = None
        self.result_df = None
        
        self.create_widgets()
    
    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="加单商品分配系统", font=("Arial", 18, "bold"))
        title_label.pack(pady=20)
        
        # 文件选择区
        file_frame = tk.Frame(self.root)
        file_frame.pack(pady=10, fill=tk.X, padx=20)
        
        tk.Label(file_frame, text="Excel文件:").pack(side=tk.LEFT)
        self.file_entry = tk.Entry(file_frame, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        browse_btn = tk.Button(file_frame, text="浏览", command=self.browse_file)
        browse_btn.pack(side=tk.LEFT)
        
        # 按钮区
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=20)
        
        self.run_btn = tk.Button(btn_frame, text="执行加单分配", command=self.run_allocation, 
                               font=("Arial", 12), bg="#2E7D32", fg="#F5F5F5", width=18)
        self.run_btn.pack(side=tk.LEFT, padx=10)
        self.run_btn.config(state=tk.DISABLED)
        
        self.save_btn = tk.Button(btn_frame, text="保存结果", command=self.save_result, 
                                 font=("Arial", 12), bg="#1565C0", fg="#F5F5F5", width=12)
        self.save_btn.pack(side=tk.LEFT, padx=10)
        self.save_btn.config(state=tk.DISABLED)
        
        # 进度区
        self.progress_frame = tk.Frame(self.root)
        self.progress_frame.pack(pady=10)
        self.status_label = tk.Label(self.progress_frame, text="等待文件选择...", fg="gray")
        self.status_label.pack()
        
        # 结果预览
        preview_frame = tk.Frame(self.root)
        preview_frame.pack(pady=20, fill=tk.BOTH, expand=True, padx=20)
        
        tk.Label(preview_frame, text="分配结果预览（前20行）:").pack(anchor=tk.W)
        
        self.tree = ttk.Treeview(preview_frame, show='headings')
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 说明标签
        info_label = tk.Label(self.root, 
                            text="说明：支持跨平台运行，分配完成后可保存结果到新Excel文件",
                            font=("Arial", 9), fg="gray")
        info_label.pack(side=tk.BOTTOM, pady=10)
    
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        
        if self.file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.file_path)
            self.run_btn.config(state=tk.NORMAL)
            self.status_label.config(text="文件已选择，点击执行开始分配", fg="blue")
    
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
            
            self.save_btn.config(state=tk.NORMAL)
            self.status_label.config(text="分配完成！可以保存结果", fg="green")
            
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
