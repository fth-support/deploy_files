import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import openpyxl
import os
import subprocess
import threading

# ค่าเริ่มต้นของ Path เครื่องมือ PsExec
PSEXEC_PATH = r"C:\Support\PSTools\PsExec.exe"

class DeployApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hybrid Mass Deployment Tool")
        self.root.geometry("1000x700")
        
        self.source_path = tk.StringVar()
        self.exe_path = tk.StringVar()
        
        self.setup_ui()

    def setup_ui(self):
        # ================= ส่วนที่ 1.1: Browse Files =================
        frame1 = tk.LabelFrame(self.root, text="1. Source & Execute Settings", padx=10, pady=10)
        frame1.pack(fill="x", padx=10, pady=5)

        tk.Label(frame1, text="Source (File/Folder):").grid(row=0, column=0, sticky="w")
        tk.Entry(frame1, textvariable=self.source_path, width=60).grid(row=0, column=1, padx=5)
        tk.Button(frame1, text="Browse Folder", command=lambda: self.source_path.set(filedialog.askdirectory())).grid(row=0, column=2, padx=2)
        tk.Button(frame1, text="Browse File", command=lambda: self.source_path.set(filedialog.askopenfilename())).grid(row=0, column=3, padx=2)

        tk.Label(frame1, text="EXE to Run (Optional):").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(frame1, textvariable=self.exe_path, width=60).grid(row=1, column=1, padx=5)
        tk.Button(frame1, text="Browse EXE", command=lambda: self.exe_path.set(filedialog.askopenfilename(filetypes=[("Executable", "*.exe")]))).grid(row=1, column=2, columnspan=2, sticky="ew", padx=2)

        # ================= ส่วนที่ 1.2: Excel Template =================
        frame2 = tk.LabelFrame(self.root, text="2. Import / Export Data", padx=10, pady=10)
        frame2.pack(fill="x", padx=10, pady=5)

        tk.Button(frame2, text="Export Template (Excel)", command=self.export_template).pack(side="left", padx=5)
        tk.Button(frame2, text="Import Data (Excel)", command=self.import_excel).pack(side="left", padx=5)
        tk.Button(frame2, text="Start Deployment", bg="green", fg="white", command=self.start_deployment).pack(side="right", padx=5)

        # ================= ส่วนที่ 1.3 & 1.4: Data Grid & Status =================
        frame3 = tk.LabelFrame(self.root, text="3. Deployment Targets & Status", padx=10, pady=10)
        frame3.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("IP", "Username", "Password", "Run_EXE", "Custom_Source", "Ping", "Copy", "Run")
        self.tree = ttk.Treeview(frame3, columns=columns, show="headings", selectmode="browse")
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")
        self.tree.column("Custom_Source", width=200)
        self.tree.pack(fill="both", expand=True)

        # ================= ส่วนแก้ไขข้อมูลเฉพาะ Row =================
        edit_frame = tk.Frame(frame3)
        edit_frame.pack(fill="x", pady=5)
        tk.Button(edit_frame, text="Edit Selected Row (Override)", command=self.edit_selected_row).pack(side="left")
        tk.Button(edit_frame, text="Clear Selected", command=self.delete_selected_row).pack(side="left", padx=5)

        # ================= ส่วนที่ 1.5: Summary =================
        self.summary_var = tk.StringVar()
        self.summary_var.set("Total: 0 | 🟢 Success: 0 | 🔴 Fail: 0")
        tk.Label(self.root, textvariable=self.summary_var, font=("Arial", 12, "bold")).pack(pady=10)

    # --- ฟังก์ชันจัดการ Excel ---
    def export_template(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="DeployTemplate.xlsx")
        if filepath:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["IP", "Username", "Password", "Run_EXE (Yes/No)", "Custom_Source_Path (Optional)"])
            ws.append(["192.168.1.50", "Administrator", "P@ssw0rd", "Yes", ""])
            wb.save(filepath)
            messagebox.showinfo("Success", "Template exported successfully!")

    def import_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            try:
                # เคลียร์ข้อมูลเก่าในตารางก่อน
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                # ใช้คำสั่งที่ถูกต้องในการโหลดไฟล์ Excel
                wb = openpyxl.load_workbook(filepath)
                ws = wb.active
                
                # อ่านข้อมูลตั้งแต่แถวที่ 2 เป็นต้นไป (ข้าม Header)
                for row in ws.iter_rows(min_row=2, values_only=True):
                    # เช็คว่ามีค่า IP (row[0]) หรือไม่ ถ้ามีถึงจะนำเข้า
                    if row[0] is not None and str(row[0]).strip() != "":
                        ip = row[0]
                        user = row[1] if row[1] else ""
                        pw = row[2] if row[2] else ""
                        run_opt = row[3] if row[3] else "No"
                        custom_src = row[4] if row[4] else ""
                        
                        # ⚪ = Pending (รอการทำงาน)
                        self.tree.insert("", "end", values=(ip, user, pw, run_opt, custom_src, "⚪", "⚪", "⚪"))
                
                self.update_summary()
                
            except Exception as e:
                # ถ้ามี Error เช่น ไฟล์พัง หรือถูกเปิดค้างไว้ใน Excel ให้แจ้งเตือน
                messagebox.showerror("Import Error", f"เกิดข้อผิดพลาดในการโหลดไฟล์ Excel:\n{str(e)}")

    # --- ฟังก์ชันแก้ไข Row ---
    def edit_selected_row(self):
        selected = self.tree.selection()
        if not selected: return
        item = self.tree.item(selected[0])['values']
        
        # สร้าง Popup แก้ไขด่วน
        top = tk.Toplevel(self.root)
        top.title("Edit Override")
        
        labels = ["IP", "Username", "Password", "Run_EXE (Yes/No)", "Custom_Source"]
        entries = []
        for i, text in enumerate(labels):
            tk.Label(top, text=text).grid(row=i, column=0, padx=5, pady=5)
            e = tk.Entry(top, width=40)
            e.insert(0, str(item[i]) if item[i] else "")
            e.grid(row=i, column=1, padx=5, pady=5)
            entries.append(e)
            
        def save_edit():
            new_vals = [e.get() for e in entries] + item[5:] # เก็บค่า Status เดิมไว้
            self.tree.item(selected[0], values=new_vals)
            top.destroy()
            self.update_summary()
            
        tk.Button(top, text="Save", command=save_edit).grid(row=5, column=1, sticky="e", pady=10)

    def delete_selected_row(self):
        selected = self.tree.selection()
        if selected:
            self.tree.delete(selected[0])
            self.update_summary()

    # --- ฟังก์ชัน Deployment หลัก ---
    def update_summary(self):
        total = len(self.tree.get_children())
        success = 0
        fail = 0
        for child in self.tree.get_children():
            vals = self.tree.item(child)['values']
            if "🔴" in vals: fail += 1
            elif "🟢" in vals[6]: success += 1 # นับว่า Copy สำเร็จถือว่าผ่านขั้นต้น
        self.summary_var.set(f"Total: {total} | 🟢 Success: {success} | 🔴 Fail: {fail}")

    def update_status(self, item_id, col_index, status_icon):
        vals = list(self.tree.item(item_id)['values'])
        vals[col_index] = status_icon
        self.tree.item(item_id, values=vals)
        self.root.update()

    def start_deployment(self):
        # เช็คก่อนว่ามีข้อมูลในตารางไหม
        if len(self.tree.get_children()) == 0:
            messagebox.showwarning("Warning", "กรุณา Import ข้อมูล Excel ก่อนเริ่ม Deployment")
            return
            
        threading.Thread(target=self._run_deployment_thread, daemon=True).start()

    def _run_deployment_thread(self):
        items = self.tree.get_children()
        for item_id in items:
            vals = self.tree.item(item_id)['values']
            ip, user, pw, run_opt, custom_src = vals[0], vals[1], vals[2], vals[3], vals[4]
            
            # 1. Ping Check
            self.update_status(item_id, 5, "🟡") 
            response = subprocess.call(['ping', '-n', '1', '-w', '1000', ip], stdout=subprocess.DEVNULL)
            if response != 0:
                self.update_status(item_id, 5, "🔴")
                self.update_status(item_id, 6, "🔴")
                self.update_status(item_id, 7, "🔴")
                continue
            self.update_status(item_id, 5, "🟢")

            # 2. Copy Files (ใช้ net use นำทางก่อน)
            self.update_status(item_id, 6, "🟡")
            src = custom_src if custom_src else self.source_path.get()
            if not src:
                self.update_status(item_id, 6, "⚪") # Skip
            else:
                subprocess.call(f'net use \\\\{ip}\\IPC$ /user:{user} {pw}', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                
                dest = f"\\\\{ip}\\C$\\TempDeploy"
                if os.path.isdir(src):
                    cmd_copy = f'robocopy "{src}" "{dest}" /E /njh /njs /nc /ns /np'
                else:
                    cmd_copy = f'echo F | xcopy /Y /F "{src}" "{dest}\\{os.path.basename(src)}"'
                
                res_copy = subprocess.call(cmd_copy, shell=True, stdout=subprocess.DEVNULL)
                subprocess.call(f'net use \\\\{ip}\\IPC$ /delete /y', shell=True, stdout=subprocess.DEVNULL)

                if res_copy in [0, 1, 2, 3]: 
                    self.update_status(item_id, 6, "🟢")
                else:
                    self.update_status(item_id, 6, "🔴")

            # 3. Run PsExec
            if str(run_opt).strip().lower() == 'yes':
                self.update_status(item_id, 7, "🟡")
                exe_target = self.exe_path.get()
                if exe_target:
                    exe_name = os.path.basename(exe_target)
                    remote_exe_path = f"C:\\TempDeploy\\{exe_name}"
                    
                    psexec_cmd = f'"{PSEXEC_PATH}" \\\\{ip} -u {user} -p {pw} -accepteula -s -d "{remote_exe_path}"'
                    res_psexec = subprocess.call(psexec_cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    
                    if res_psexec == 0:
                        self.update_status(item_id, 7, "🟢")
                    else:
                        self.update_status(item_id, 7, "🔴")
                else:
                    self.update_status(item_id, 7, "🔴") 
            else:
                self.update_status(item_id, 7, "⚪") 

        self.update_summary()
        messagebox.showinfo("Done", "Deployment Task Completed!")

if __name__ == "__main__":
    root = tk.Tk()
    app = DeployApp(root)
    root.mainloop()
