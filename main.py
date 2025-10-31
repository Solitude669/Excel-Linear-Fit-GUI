import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.metrics import r2_score
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
import sys

plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Excel 线性拟合小工具')
        self.geometry('900x700')  # 增大默认窗口
        self.minsize(900, 700)    # 防止用户缩太小

        # --- 顶部文件选择 ---
        frm_top = ttk.Frame(self)
        frm_top.pack(pady=10, fill=tk.X, padx=20)
        ttk.Label(frm_top, text='Excel 文件：').pack(side=tk.LEFT)
        self.var_path = tk.StringVar()
        ttk.Entry(frm_top, textvariable=self.var_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(frm_top, text='浏览…', command=self.select_file).pack(side=tk.LEFT)

        # --- 参数显示 ---
        frm_info = ttk.LabelFrame(self, text='拟合结果')
        frm_info.pack(fill=tk.X, padx=20, pady=5)
        self.var_eq = tk.StringVar(value='方程：未计算')
        self.var_r2 = tk.StringVar(value='R²：未计算')
        ttk.Label(frm_info, textvariable=self.var_eq).pack(anchor=tk.W)
        ttk.Label(frm_info, textvariable=self.var_r2).pack(anchor=tk.W)

        # --- 图形嵌入（关键：允许扩展） ---
        self.fig = plt.Figure(figsize=(7, 5))
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, self)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # --- 底部按钮 ---
        frm_btn = ttk.Frame(self)
        frm_btn.pack(pady=10)
        ttk.Button(frm_btn, text='开始拟合', command=self.fit).pack(side=tk.LEFT, padx=5)
        ttk.Button(frm_btn, text='保存图片', command=self.save_pic).pack(side=tk.LEFT, padx=5)
        ttk.Button(frm_btn, text='退出', command=self.quit).pack(side=tk.LEFT, padx=5)

    # ---------------- 功能函数 ----------------
    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx *.xls')])
        if path:
            self.var_path.set(path)

    def fit(self):
        path = self.var_path.get()
        if not os.path.isfile(path):
            messagebox.showerror('错误', '请先选择有效的 Excel 文件！')
            return
        try:
            # ★ 关键：把第一行当表头读，如果失败再退回无表头
            df = pd.read_excel(path, header=None)          # 先无表头读一次
            first = df.iloc[0].astype(str).str.strip()     # 取第一行
            # 如果 A1、B1 都是合法非数字文本，就当表头；否则整表重新无表头读
            if (first.notna() & ~first.str.match(r'^-?\d+(\.\d+)?$')).all():
                df = pd.read_excel(path, header=0)         # 第二行开始当数据
                x_name, y_name = first[0], first[1]
            else:
                df = pd.read_excel(path, header=None)      # 整表都是数据
                x_name, y_name = 'X轴', 'Y轴'
            x = df.iloc[:, 0].values
            y = df.iloc[:, 1].values
        except Exception as e:
            messagebox.showerror('读取失败', str(e))
            return

        # 拟合
        k, b = np.polyfit(x, y, 1)
        y_fit = k * x + b
        r2 = r2_score(y, y_fit)

        # 更新文字
        self.var_eq.set(f'方程：y = {k:.4f}x + {b:.4f}')
        self.var_r2.set(f'R² = {r2:.4f}')

        # 画图
        self.ax.clear()
        self.ax.scatter(x, y, label='原始数据')
        self.ax.plot(x, y_fit, color='C0', label='拟合直线')
        self.ax.set_title('线性拟合结果')
        self.ax.set_xlabel(x_name)   # 用动态名称
        self.ax.set_ylabel(y_name)
        self.ax.legend()
        self.fig.tight_layout()
        self.canvas.draw()

    def save_pic(self):
        save_path = filedialog.asksaveasfilename(defaultextension='.png',
                                                 filetypes=[('PNG', '*.png'), ('PDF', '*.pdf')])
        if save_path:
            self.fig.savefig(save_path)
            messagebox.showinfo('完成', f'图片已保存至\n{save_path}')


# ---------- 打包用入口 ----------
def resource_path(relative):
    """PyInstaller 后资源文件路径"""
    base = getattr(sys, '_MEIPASS', os.path.abspath('.'))
    return os.path.join(base, relative)


if __name__ == '__main__':
    App().mainloop()