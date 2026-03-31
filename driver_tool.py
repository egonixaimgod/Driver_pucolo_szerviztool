import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import ctypes
import sys
import os
import winreg
import re
from datetime import datetime

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def resource_path(relative_path):
    """ Visszaadja az erőforrás abszolút útvonalát, kezeli a PyInstaller _MEIPASS mappáját """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class DriverCleanerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Third-Party Driver Kezelő és WU Letiltó")
        self.geometry("900x600")
        
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except:
            pass
        
        self.create_widgets()
        self.refresh_drivers()

    def create_widgets(self):
        # Top Frame - Windows Update
        wu_frame = ttk.LabelFrame(self, text="Windows Update Driver Frissítések Beállításai", padding=10)
        wu_frame.pack(fill=tk.X, padx=10, pady=5)

        self.wu_status_lbl = ttk.Label(wu_frame, text="Állapot: Ismeretlen", font=("Arial", 10, "bold"))
        self.wu_status_lbl.pack(side=tk.LEFT, padx=10)

        disable_wu_btn = ttk.Button(wu_frame, text="WU Driver Letöltés LETILTÁSA", command=self.disable_wu_drivers)
        disable_wu_btn.pack(side=tk.LEFT, padx=5)

        enable_wu_btn = ttk.Button(wu_frame, text="WU Driver Letöltés ENGEDÉLYEZÉSE", command=self.enable_wu_drivers)
        enable_wu_btn.pack(side=tk.LEFT, padx=5)

        self.check_wu_status()

        # Middle Frame - Backup
        backup_frame = ttk.LabelFrame(self, text="Biztonsági Mentés (Driver Export és Visszaállítási Pont)", padding=10)
        backup_frame.pack(fill=tk.X, padx=10, pady=5)

        rp_btn = ttk.Button(backup_frame, text="Új Rendszer-visszaállítási Pont Készítése", command=self.create_restore_point)
        rp_btn.pack(side=tk.LEFT, padx=5)

        export_btn = ttk.Button(backup_frame, text="Összes Third-Party Driver Lementése (Exportálás)", command=self.backup_drivers)
        export_btn.pack(side=tk.LEFT, padx=5)

        # Bottom Frame - Drivers
        drv_frame = ttk.LabelFrame(self, text="Telepített Harmadik Fél (Third-party) Driverek", padding=10)
        drv_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("published", "original", "provider", "class", "version")
        self.tree = ttk.Treeview(drv_frame, columns=columns, show="headings", selectmode="extended")
        
        self.tree.heading("published", text="Közzétett Név (oem.inf)")
        self.tree.heading("original", text="Eredeti Név")
        self.tree.heading("provider", text="Gyártó")
        self.tree.heading("class", text="Eszközosztály")
        self.tree.heading("version", text="Verzió/Dátum")

        self.tree.column("published", width=120)
        self.tree.column("original", width=150)
        self.tree.column("provider", width=200)
        self.tree.column("class", width=150)
        self.tree.column("version", width=200)

        scrollbar = ttk.Scrollbar(drv_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons under tree
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        refresh_btn = ttk.Button(btn_frame, text="Lista Frissítése", command=self.refresh_drivers)
        refresh_btn.pack(side=tk.LEFT, padx=5)

        select_all_btn = ttk.Button(btn_frame, text="Összes Kijelölése", command=self.select_all_drivers)
        select_all_btn.pack(side=tk.LEFT, padx=5)

        delete_btn = ttk.Button(btn_frame, text="Kiválasztott Driver(ek) TÖRLÉSE", command=self.delete_selected_drivers)
        delete_btn.pack(side=tk.RIGHT, padx=5)

    def get_third_party_drivers(self):
        try:
            # Hide console window
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            result = subprocess.run(['pnputil', '/enum-drivers'], capture_output=True, text=True, startupinfo=startupinfo)
            output = result.stdout
            
            drivers = []
            current_driver = {}
            
            for line in output.splitlines():
                line = line.strip()
                if not line:
                    if current_driver:
                        drivers.append(current_driver)
                        current_driver = {}
                    continue
                
                parts = line.split(":", 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    val = parts[1].strip()
                    
                    if "Published Name" in key or "Közzétett név" in key:
                        current_driver["published"] = val
                    elif "Original Name" in key or "Eredeti név" in key:
                        current_driver["original"] = val
                    elif "Provider Name" in key or "Szolgáltató neve" in key:
                        current_driver["provider"] = val
                    elif "Class Name" in key or "Osztály neve" in key:
                        current_driver["class"] = val
                    elif "Driver Version" in key or "Illesztőprogram verziója" in key:
                        current_driver["version"] = val

            if current_driver:
                drivers.append(current_driver)
                
            return drivers
        except Exception as e:
            messagebox.showerror("Hiba", f"Nem sikerült lekérdezni a drivereket:\n{str(e)}")
            return []

    def refresh_drivers(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        drivers = self.get_third_party_drivers()
        for d in drivers:
            if "published" in d:
                self.tree.insert("", tk.END, values=(
                    d.get("published", ""),
                    d.get("original", ""),
                    d.get("provider", ""),
                    d.get("class", ""),
                    d.get("version", "")
                ))

    def select_all_drivers(self):
        for item in self.tree.get_children():
            self.tree.selection_add(item)

    def delete_selected_drivers(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Figyelmeztetés", "Kérlek, válassz ki legalább egy drivert a törléshez!")
            return
            
        if not messagebox.askyesno("Megerősítés", f"Biztosan törölni szeretnéd a kiválasztott {len(selected)} drivert és az eszközökről is eltávolítod?"):
            return
            
        success_count = 0
        fail_count = 0
        
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        
        for item in selected:
            published_name = self.tree.item(item, "values")[0]
            try:
                # /uninstall /force removes it from driver store AND attempts uninstall from active devices
                res = subprocess.run(['pnputil', '/delete-driver', published_name, '/uninstall', '/force'], 
                                   capture_output=True, text=True, startupinfo=startupinfo)
                if res.returncode == 0 or "Deleted" in res.stdout or "törölve" in res.stdout:
                    success_count += 1
                else:
                    fail_count += 1
                    print(f"Hiba a {published_name} törlésekor: {res.stdout}")
            except Exception as e:
                fail_count += 1
                print(f"Kivétel a {published_name} törlésekor: {e}")
                
        messagebox.showinfo("Eredmény", f"Sikeresen törölve: {success_count}\nNem sikerült: {fail_count}")
        self.refresh_drivers()

    def check_wu_status(self):
        try:
            policy_disabled = False
            search_disabled = False
            
            # 1. Policy key (Windows 10/11)
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate", 0, winreg.KEY_READ) as key:
                    val, _ = winreg.QueryValueEx(key, "ExcludeWUDriversInQualityUpdate")
                    if val == 1: policy_disabled = True
            except FileNotFoundError:
                pass
                
            # 2. DriverSearching key (Eszköztelepítési beállítások)
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching", 0, winreg.KEY_READ) as key:
                    val, _ = winreg.QueryValueEx(key, "SearchOrderConfig")
                    if val == 0: search_disabled = True
            except FileNotFoundError:
                pass
                
            if policy_disabled and search_disabled:
                self.wu_status_lbl.config(text="Állapot: Teljesen LETILTVA (Házirend és Eszközbeállítás is)", foreground="red")
            elif policy_disabled:
                self.wu_status_lbl.config(text="Állapot: Házirend által LETILTVA (Képen: bekapcsolva)", foreground="red")
            elif search_disabled:
                self.wu_status_lbl.config(text="Állapot: Eszközbeállításokban LETILTVA", foreground="red")
            else:
                self.wu_status_lbl.config(text="Állapot: Driver frissítés ENGEDÉLYEZVE", foreground="green")
        except Exception as e:
            self.wu_status_lbl.config(text="Állapot: Ismeretlen", foreground="black")

    def disable_wu_drivers(self):
        try:
            # Policy - ExcludeWUDriversInQualityUpdate (Windows 10/11)
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
            with winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, "ExcludeWUDriversInQualityUpdate", 0, winreg.REG_DWORD, 1)
            
            # SearchOrderConfig
            key_path2 = r"SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching"
            with winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, key_path2, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, "SearchOrderConfig", 0, winreg.REG_DWORD, 0)
                
            messagebox.showinfo("Siker", "Windows Update driver telepítés sikeresen LETILTVA.")
            self.check_wu_status()
        except PermissionError:
            messagebox.showerror("Hiba", "Nincs jogosultság a Registry írásához. Futtasd Rendszergazdaként!")
        except Exception as e:
            messagebox.showerror("Hiba", f"Hiba történt:\n{str(e)}")

    def enable_wu_drivers(self):
        try:
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_WRITE) as key:
                    winreg.DeleteValue(key, "ExcludeWUDriversInQualityUpdate")
            except FileNotFoundError:
                pass

            key_path2 = r"SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching"
            with winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, key_path2, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, "SearchOrderConfig", 0, winreg.REG_DWORD, 1)

            messagebox.showinfo("Siker", "Windows Update driver telepítés sikeresen VISSZAÁLLÍTVA.")
            self.check_wu_status()
        except PermissionError:
            messagebox.showerror("Hiba", "Nincs jogosultság a Registry írásához. Futtasd Rendszergazdaként!")
        except Exception as e:
            messagebox.showerror("Hiba", f"Hiba történt:\n{str(e)}")

    def create_restore_point(self):
        desc = f"Driver_Cleaner_Backup_{datetime.now().strftime('%Y%md_%H%M%S')}"
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            # PowerShell Command to create a restore point
            cmd = f'powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "Checkpoint-Computer -Description \'{desc}\' -RestorePointType \'MODIFY_SETTINGS\'"'
            
            messagebox.showinfo("Folyamatban", "Rendszer-visszaállítási pont létrehozása elindult...\nEz eltarthat egy percig, kérlek várj!")
            res = subprocess.run(cmd, shell=True, startupinfo=startupinfo, capture_output=True, text=True)
            
            if res.returncode == 0:
                messagebox.showinfo("Siker", f"A '{desc}' nevű visszaállítási pont sikeresen létrejött!")
            else:
                messagebox.showerror("Hiba", f"Nem sikerült létrehozni a visszaállítási pontot. Biztosan engedélyezve van a Rendszervédelem a Windowsban?\n\nKimenet: {res.stderr}")
        except Exception as e:
            messagebox.showerror("Hiba", f"Kivétel történt:\n{str(e)}")

    def backup_drivers(self):
        dest_dir = filedialog.askdirectory(title="Válassz egy mappát a driverek kimentéséhez")
        if not dest_dir:
            return
            
        backup_folder = os.path.join(dest_dir, f"Driver_Backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(backup_folder, exist_ok=True)
        
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            
            messagebox.showinfo("Folyamatban", f"Driverek exportálása a(z) {backup_folder} mappába...\nKérlek várj, ez hosszú időt is igénybe vehet.")
            
            res = subprocess.run(['dism', '/online', '/export-driver', f'/destination:{backup_folder}'], 
                                 capture_output=True, text=True, startupinfo=startupinfo)
                                 
            if res.returncode == 0 or "successful" in res.stdout or "sikeres" in res.stdout:
                msg = f"A harmadik fél (third-party) driverek sikeresen lementve ide:\n{backup_folder}\n\nHa baj van, Sergei Strelec WinPE-ben a dism++ szoftverrel vagy parancssorból visszarakhatod őket!"
                messagebox.showinfo("Sikeres Export", msg)
            else:
                messagebox.showerror("Hiba", f"A dism hibaüzenettel tért vissza:\n{res.stdout}\n{res.stderr}")
        except Exception as e:
            messagebox.showerror("Kivétel", f"Váratlan hiba az exportálás során:\n{str(e)}")


if __name__ == "__main__":
    if not is_admin():
        # Felemeljük a jogosultságot UAC ablakkal
        import sys
        params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
        if getattr(sys, 'frozen', False):
            # PyInstaller exe
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        else:
            # Python script
            script = sys.argv[0]
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)
        sys.exit()

    app = DriverCleanerApp()
    app.mainloop()
