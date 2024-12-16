import tkinter as tk
import tkinter.filedialog
import ttkbootstrap as ttk
from PIL import ImageTk
import threading
import pandas as pd
import os

from utils import (get_icon_img, get_like_img, get_excel_files)

version = "v1.0.0"


class mergeExcel():
    def __init__(self, master):
        self.master = master
        self.master.title("merge excel")
        self.master.resizable(False, False)
        self.master.geometry("450x350")
        self.master.iconphoto(False, ImageTk.PhotoImage(data=get_icon_img()))

        self.table_dir = tk.StringVar()
        self.table_dir.set("")
        self.merge_way = tk.IntVar()
        self.merge_way.set(0)
        self.header_row = tk.IntVar()
        self.header_row.set(1)
        self.process_text = tk.StringVar()
        self.process_text.set("0%")
        self.out_merge_file = None
        self.merge_result_text = tk.StringVar()
        self.merge_result_text.set("待处理")
        self.running = False
        self.merge_thread = None

        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(pady=5, fill=tk.BOTH, expand=True)

        # 合并tab
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="合并")

        # 选择路径
        select_dir_frame = ttk.Frame(self.tab1)
        select_dir_frame.pack(pady=(20, 0))
        ttk.Label(select_dir_frame, text="路径:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(select_dir_frame, width=25, textvariable=self.table_dir).pack(side=tk.LEFT, padx=5)
        ttk.Button(select_dir_frame, text="选择", command=self.select_dir, style="success").pack(side=tk.LEFT, padx=5)

        # 合并方式
        merge_way_frame = ttk.Frame(self.tab1)
        merge_way_frame.pack(pady=(15, 0))
        ttk.Radiobutton(merge_way_frame, text="按sheet合并", variable=self.merge_way, value=1,
                        command=self.switch_header_row_state, style="success").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(merge_way_frame, text="按行合并", variable=self.merge_way, value=0,
                        command=self.switch_header_row_state, style="success").pack(side=tk.LEFT, padx=10)
        header_row_state = tk.DISABLED if self.merge_way.get() == 1 else tk.NORMAL
        self.header_row_label = tk.Label(merge_way_frame, text="表头行数:", state=header_row_state)
        self.header_row_label.pack(side=tk.LEFT, padx=5)
        self.header_row_spinbox = ttk.Spinbox(merge_way_frame, from_=0, to=99,
                                              textvariable=self.header_row, width=2, state=header_row_state)
        self.header_row_spinbox.pack(side=tk.LEFT, padx=5)

        # 开始合并
        start_merge_frame = ttk.Frame(self.tab1)
        start_merge_frame.pack(pady=(20, 0))
        ttk.Button(start_merge_frame, text="开始合并", command=self.start, style="success").pack(side=tk.LEFT)

        # 进度条
        progress_bar_frame = ttk.Frame(self.tab1)
        progress_bar_frame.pack(pady=(15, 0))
        ttk.Label(progress_bar_frame, text="处理进度:").pack(side=tk.LEFT, padx=10)
        self.progress_bar = ttk.Progressbar(progress_bar_frame, orient=tk.HORIZONTAL,
                                            length=280, mode='determinate', style="success-striped")
        self.progress_bar.pack(side=tk.LEFT)
        self.process_text_label = tk.Label(progress_bar_frame, textvariable=self.process_text, width=5)
        self.process_text_label.pack(side=tk.LEFT, padx=5)

        # 结果提示
        result_frame = ttk.Frame(self.tab1)
        result_frame.pack(pady=(10, 0))
        ttk.Label(result_frame, text="处理结果:").grid(row=0, column=0, padx=(10,0))
        ttk.Label(result_frame, text="").grid(row=1, column=0)
        self.result_label = ttk.Label(result_frame, textvariable=self.merge_result_text, width=32, wraplength=320)
        self.result_label.grid(row=0, column=1, padx=(0,10), rowspan=2, sticky=tk.NW)

        # 赞赏tab
        self.tab4 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab4, text="赞赏")
        self.like_img = ImageTk.PhotoImage(data=get_like_img())
        ttk.Label(self.tab4, image=self.like_img).pack(pady=(5, 0))
        tk.Label(self.tab4, text=version, font=("Arial", 8), state=tk.DISABLED).pack(side=tk.BOTTOM, pady=5)

    def select_dir(self):
        # 选择文件夹并显示特定后缀的文件
        self.progress_bar.config(value=0)
        self.process_text.set("0%")
        self.merge_result_text.set("待处理")
        file_path = tk.filedialog.askdirectory()
        self.table_dir.set(file_path)

    def switch_header_row_state(self):
        header_row_state = tk.DISABLED if self.merge_way.get() == 1 else tk.NORMAL
        self.header_row_label.config(state=header_row_state)
        self.header_row_spinbox.config(state=header_row_state)

    def merge_excel(self):
        excel_files = get_excel_files(self.table_dir.get())

        if len(excel_files) == 0:
            self.running = False
            self.merge_result_text.set("未处理，所选文件夹中未找到excel文件")
            return

        df_list = []
        sheet_name_list = []
        excel_header = None
        if self.header_row.get() == 1:
            excel_header = 0
        elif self.header_row.get() > 1:
            excel_header = list(range(0, self.header_row.get()))
        for i, excel_file in enumerate(excel_files):
            process_value = i * 95 // len(excel_files)
            self.progress_bar.config(value=process_value)
            self.process_text.set(f"{process_value}%")
            df = pd.read_excel(excel_file, sheet_name=0, header=excel_header)
            sheet_name = os.path.splitext(os.path.basename(excel_file))[0]
            # 若按行合并，则检查表头是否一致
            if i > 0 and self.merge_way.get() == 0:
                if not df_list[0].columns.equals(df.columns):
                    self.running = False
                    self.process_text.set("0%")
                    self.progress_bar.config(value=0)
                    self.merge_result_text.set(f"失败，{sheet_name_list[0]} 与 {sheet_name} 表头不一致")
                    return
            sheet_name_list.append(sheet_name)
            df_list.append(df)

        self.out_merge_file = self.table_dir.get() + "_merged.xlsx"
        if self.merge_way.get() == 0:
            # 按行合并
            df_merge = pd.concat(df_list, ignore_index=True)
            if self.header_row.get() > 1:
                df_merge.to_excel(self.out_merge_file)
            elif self.header_row.get() == 1:
                df_merge.to_excel(self.out_merge_file, index=False)
        else:
            # 按sheet合并
            with pd.ExcelWriter(self.out_merge_file) as writer:
                for df, sheet_name in zip(df_list, sheet_name_list):
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        self.progress_bar.config(value=100)
        self.process_text.set("100%")
        self.running = False
        self.merge_result_text.set("完成，输出文件 " + os.path.basename(self.out_merge_file))

    def start(self):
        if not self.running:
            self.running = True
            self.merge_result_text.set("处理中，请稍候...")
            self.merge_thread = threading.Thread(target=self.merge_excel)
            self.merge_thread.start()

    def on_closing(self):
        if self.running:
            self.running = False
            self.merge_thread.join()
        self.master.destroy()


if __name__ == '__main__':
    root = ttk.Window()
    style = ttk.Style()
    style.theme_use('litera')
    app = mergeExcel(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
