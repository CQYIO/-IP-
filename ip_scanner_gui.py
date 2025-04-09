import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import socket
import concurrent.futures
from datetime import datetime
from openpyxl import Workbook

def ping_ip(ip):
    try:
        result = subprocess.run(['ping', '-n', '1', '-w', '300', ip],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if result.returncode == 0:
            try:
                hostname = socket.gethostbyaddr(ip)[0]
            except socket.herror:
                hostname = '未知主机'
            return (ip, hostname, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    except Exception:
        pass
    return None

def scan_network(base_ip, result_callback):
    live_hosts = []

    def run_scan():
        with concurrent.futures.ThreadPoolExecutor(max_workers=100) as executor:
            futures = [executor.submit(ping_ip, f"{base_ip}{i}") for i in range(1, 255)]
            for future in concurrent.futures.as_completed(futures):
                result = future.result()
                if result:
                    live_hosts.append(result)
                    result_callback(result)

    run_scan()

    return live_hosts

def export_to_excel(data):
    if not data:
        messagebox.showwarning("没有数据", "当前没有扫描结果可导出。")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
    if file_path:
        wb = Workbook()
        ws = wb.active
        ws.append(['IP地址', '主机名', '响应时间'])
        for row in data:
            ws.append(row)
        wb.save(file_path)
        messagebox.showinfo("导出成功", f"已成功导出到：\n{file_path}")

# 创建 GUI 界面
class IPScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("单位内网IP扫描工具")
        self.data = []

        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="请输入网段（如 172.16.5.）:").grid(row=0, column=0, sticky=tk.W)
        self.ip_entry = ttk.Entry(frame, width=20)
        self.ip_entry.insert(0, "172.16.5.")
        self.ip_entry.grid(row=0, column=1, sticky=tk.W)

        self.scan_button = ttk.Button(frame, text="开始扫描", command=self.start_scan)
        self.scan_button.grid(row=0, column=2, padx=5)

        self.export_button = ttk.Button(frame, text="导出Excel", command=lambda: export_to_excel(self.data))
        self.export_button.grid(row=0, column=3, padx=5)

        self.tree = ttk.Treeview(frame, columns=('IP地址', '主机名', '响应时间'), show='headings')
        for col in ('IP地址', '主机名', '响应时间'):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        self.tree.grid(row=1, column=0, columnspan=4, pady=10, sticky='nsew')

        frame.rowconfigure(1, weight=1)
        frame.columnconfigure(1, weight=1)

    def start_scan(self):
        base_ip = self.ip_entry.get()
        if not base_ip:
            messagebox.showerror("输入错误", "请输入有效的网段。")
            return

        self.data.clear()
        for row in self.tree.get_children():
            self.tree.delete(row)

        self.scan_button.config(state='disabled')
        self.root.after(100, lambda: self.run_scan_thread(base_ip))

    def run_scan_thread(self, base_ip):
        def on_result(result):
            self.data.append(result)
            self.tree.insert('', tk.END, values=result)

        scan_network(base_ip, on_result)
        self.scan_button.config(state='normal')

# 运行程序
if __name__ == "__main__":
    root = tk.Tk()
    app = IPScannerApp(root)
    root.mainloop()
