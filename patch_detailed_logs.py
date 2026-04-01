import sys
import os

target = 'driver_tool.py'
with open(target, 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if 'def _start_restore_thread(' in line:
        start_idx = i
        break

if start_idx == -1:
    print('Could not find _start_restore_thread')
    sys.exit(1)

for i in range(start_idx + 1, len(lines)):
    if lines[i].startswith('    def ') or lines[i].startswith('if __name__'):
        end_idx = i
        break

if end_idx == -1:
    end_idx = len(lines)

new_func = """    def _start_restore_thread(self, online, source_dir, target_dir):
        prog_win = tk.Toplevel(self)
        title_txt = "Élő rendszer frissítése..." if online else f"Offline WinPE Integrálás: {target_dir}"
        prog_win.title(title_txt)
        prog_win.geometry("900x650")
        prog_win.minsize(700, 480)
        prog_win.transient(self)
        prog_win.grab_set()

        lbl_txt = ("Illesztőprogramok rátelepítése a jelenlegi gépre...\\nKérlek várj!" 
                   if online else f"Illesztőprogramok befűzése a(z) {target_dir} meghajtóra...\\nEz eltarthat egy darabig!")
        lbl = ttk.Label(prog_win, text=lbl_txt, justify=tk.CENTER, font=("Segoe UI", 11, "bold"))
        lbl.pack(pady=5)

        progress = ttk.Progressbar(prog_win, orient=tk.HORIZONTAL, mode='indeterminate')
        progress.pack(pady=5, fill=tk.X, padx=20)
        progress.start(15)
        
        status_lbl = ttk.Label(prog_win, text="Folyamat indítása... Készülünk az illesztőprogramokra...", font=("Segoe UI", 9))
        status_lbl.pack(pady=2)

        # Hatalmas szovegdoboz, hogy a felhasznalo LASSA a folyamatot reszletesen
        log_frame = tk.Frame(prog_win, bg="#000000")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 10), bg="#1E1E1E", fg="#00FF00")
        log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        log_text.config(yscrollcommand=scrollbar.set)
        
        # Auto-scroll toggler
        auto_scroll = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(prog_win, text="Automatikus görgetés (Auto-Scroll)", variable=auto_scroll)
        chk.pack(anchor="w", padx=20)
        
        close_btn = ttk.Button(prog_win, text="Bezárás", command=prog_win.destroy, state=tk.DISABLED)
        close_btn.pack(pady=10)

        def worker():
            log_handle = None
            try:
                # Ebbe a log fajlba 100% bizonyossággal beleteszi a kimenetet, meg a crash is bekerül!
                debug_log_path = os.path.join(os.path.abspath('.'), 'driver_tool_gui_restore.log')
                log_handle = open(debug_log_path, 'w', encoding='utf-8')
            except Exception:
                pass

            def write_log(msg):
                import logging
                if log_handle:
                    try:
                        log_handle.write(msg + '\\n')
                        log_handle.flush() # AZONNALI Iras a pendriverra!
                    except:
                        pass
                logging.info(f"[RESTORE] {msg}")
                def gui_up():
                    try:
                        log_text.insert(tk.END, msg + '\\n')
                        if auto_scroll.get():
                            log_text.see(tk.END)
                        short_txt = msg.strip() if len(msg.strip()) < 100 else msg.strip()[:97] + "..."
                        status_lbl.config(text=short_txt)
                    except:
                        pass
                self.after(0, gui_up)

            try:
                write_log(f"=== {'ONLINE' if online else 'OFFLINE'} DRIVER RESTORE INDÍTÁSA ===")
                write_log(f"Forrás könyvtár: {source_dir}")
                write_log(f"Célpont / OS Meghajtó: {target_dir}")
                write_log("Felkészülés a parancsok futtatására...")

                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                import os
                
                # Szigorú backslash konverzió a natív parancssori eszközöknek
                norm_source = os.path.normpath(source_dir).replace('/', '\\\\')
                norm_target = os.path.normpath(target_dir).replace('/', '\\\\')

                if online:
                    cmd = ['pnputil', '/add-driver', f"{norm_source}\\\\\\*.inf", '/subdirs', '/install']
                else:
                    cmd = ['dism', f'/Image:{norm_target}', '/Add-Driver', f'/Driver:{norm_source}', '/Recurse', '/ForceUnsigned']
                
                write_log(f"Végrehajtandó parancssor: {' '.join(cmd)}")

                process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, 
                                           startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW, errors='replace')
                
                for line in process.stdout:
                    line = line.strip()
                    if not line: continue
                    write_log(line)

                process.wait()
                write_log(f"\\n--- Alapfolyamat befejeződött, visszatérési kód (Return Code): {process.returncode} ---")

                if True:
                    if online:
                        write_log("Hardverváltozások keresése és eszközök frissítése az Eszközkezelőben...")
                        import time
                        time.sleep(1.5)
                        scan_proc = subprocess.run(['pnputil', '/scan-devices'], startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW, capture_output=True, text=True, errors='replace')
                        write_log("SCAN_DEVICES Kész! Kimenet:\\n" + scan_proc.stdout)
                        time.sleep(3.5)
                    else:
                        import shutil
                        temp_drivers_dir_target = os.path.join(target_dir, "TempRunDrivers")
                        write_log(f"Készül a Boot-idejű automatikus PnP telepítő... Ideiglenes fájlfa másolása:\\n {source_dir} -> {temp_drivers_dir_target}")
                        
                        if os.path.exists(temp_drivers_dir_target):
                            shutil.rmtree(temp_drivers_dir_target, ignore_errors=True)
                        try:
                            shutil.copytree(source_dir, temp_drivers_dir_target, dirs_exist_ok=True)
                            write_log(f"Másolás SIKERES a TempRunDrivers alá.")
                        except Exception as ee:
                            write_log(f"HIBA a temp mappa másolása közben: {ee}")

                        programdata_dir = os.path.join(target_dir, "ProgramData")
                        os.makedirs(programdata_dir, exist_ok=True)
                            
                        bat_path = os.path.join(programdata_dir, "auto_pnputil_scan.bat")
                        bat_content = "@echo off\\r\\n" \\
                                      "set LOGFILE=\\\"%SystemDrive%\\\\Users\\\\Public\\\\driver_startup_log.txt\\\"\\r\\n" \\
                                      "echo ---------------------------------------- >> %LOGFILE%\\r\\n" \\
                                      "echo [%DATE% %TIME%] Boot elotti SYSTEM telepites service (No UAC! Azonnali!) >> %LOGFILE%\\r\\n" \\
                                      "echo [%DATE% %TIME%] Ideiglenes szerviz torlese a registrybol is... >> %LOGFILE%\\r\\n" \\
                                      "sc delete DriverRestoreSvc >> %LOGFILE% 2>&1\\r\\n" \\
                                      "echo [%DATE% %TIME%] Varakozas a Windows PlugAndPlay szolgaltatasara (15 sec max)... >> %LOGFILE%\\r\\n" \\
                                      "ping 127.0.0.1 -n 15 > nul\\r\\n" \\
                                      "echo [%DATE% %TIME%] Driverek betoltese (Csendes mod)... kerlek varj! >> %LOGFILE%\\r\\n" \\
                                      "pnputil /add-driver \\\"%SystemDrive%\\\\TempRunDrivers\\\\*.inf\\\" /subdirs /install >> %LOGFILE% 2>&1\\r\\n" \\
                                      "echo [%DATE% %TIME%] pnputil scan-devices inditasa... >> %LOGFILE%\\r\\n" \\
                                      "pnputil /scan-devices >> %LOGFILE% 2>&1\\r\\n" \\
                                      "echo [%DATE% %TIME%] Ideiglenes mappak torlese... >> %LOGFILE%\\r\\n" \\
                                      "rd /s /q \\\"%SystemDrive%\\\\TempRunDrivers\\\" >> %LOGFILE% 2>&1\\r\\n" \\
                                      "echo [%DATE% %TIME%] Befejezve. Torlom a scriptet. >> %LOGFILE%\\r\\n" \\
                                      "ping 127.0.0.1 -n 3 > nul\\r\\n" \\
                                      "(goto) 2>nul & del \\\"%~f0\\\"\\r\\n"
                        with open(bat_path, "w", encoding="utf-8") as f:
                            f.write(bat_content)
                        write_log(f"BAT fájl regisztrálva a rendszerbe: {bat_path}")

                        hive_path = os.path.join(target_dir, "Windows", "System32", "config", "SYSTEM")
                        if os.path.exists(hive_path):
                            write_log(f"Offline registry beinjektálása a HKLM\\\\OFFLINE_SYSTEM hive-ba ({hive_path})...")
                            import winreg
                            try:
                                subprocess.run(['reg', 'load', 'HKLM\\\\OFFLINE_SYSTEM', hive_path], check=True, capture_output=True, text=True, errors='replace', creationflags=subprocess.CREATE_NO_WINDOW)
                                try:
                                    key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, r'OFFLINE_SYSTEM\\ControlSet001\\Services\\DriverRestoreSvc')
                                    winreg.SetValueEx(key, 'Type', 0, winreg.REG_DWORD, 16)
                                    winreg.SetValueEx(key, 'Start', 0, winreg.REG_DWORD, 2)
                                    winreg.SetValueEx(key, 'ErrorControl', 0, winreg.REG_DWORD, 1)
                                    bat_target_path = r'%SystemDrive%\\ProgramData\\auto_pnputil_scan.bat'
                                    winreg.SetValueEx(key, 'ImagePath', 0, winreg.REG_EXPAND_SZ, bat_target_path)
                                    winreg.SetValueEx(key, 'ObjectName', 0, winreg.REG_SZ, 'LocalSystem')
                                    winreg.CloseKey(key)
                                    write_log("Sikeresen felprogramoztuk az Offline Windows Registry-t az auto-installhoz (Boot Service)!")
                                except Exception as rx:
                                    write_log("HIBA A REGISTRY VÁLTOZÓK ÍRÁSÁNÁL: " + str(rx))
                                finally:
                                    subprocess.run(['reg', 'unload', 'HKLM\\\\OFFLINE_SYSTEM'], check=False, capture_output=True, text=True, errors='replace', creationflags=subprocess.CREATE_NO_WINDOW)
                            except subprocess.CalledProcessError as e:
                                write_log(f"REGISTRY MOUNT SIKERTELEN: Hibakód: {e.returncode}. Kimenet: {e.output}")
                        else:
                            write_log(f"FIGYELEM MS HIBA: Az offline rendszerleíró adatbázis nem található itt: {hive_path}")

                self.after(0, lambda: progress.stop())
                write_log("==== MINDEN FOLYAMAT BEFEJEZŐDÖTT ====")
                self.after(0, lambda: status_lbl.config(text="A visszaállítás befejeződött (az ablak biztonságosan bezárható)!"))
                
                def finish_state():
                    try:
                        close_btn.config(state=tk.NORMAL)
                        if hasattr(self, 'refresh_drivers'): self.refresh_drivers()
                        
                        # Set log frame color based on result code to visually inform user
                        if process.returncode == 0:
                            log_text.config(fg="#00FF00")   # green
                        else:
                            log_text.config(fg="#FFFF00")   # yellow/warning

                    except: pass
                self.after(0, finish_state)

            except Exception as e:
                import traceback
                error_msg = f"KATASZTRÓFÁLIS PROGRAMHIBA: {e}\\n{traceback.format_exc()}"
                write_log(error_msg)
                
                def crash_gui():
                    try:
                        progress.stop()
                        status_lbl.config(text="KRITIKUS HIBA TÖRTÉNT! NÉZD MEG A LOGOT!", foreground="red")
                        log_text.config(fg="red")
                        close_btn.config(state=tk.NORMAL)
                    except: pass
                self.after(0, crash_gui)
            finally:
                if log_handle:
                    try:
                        log_handle.close()
                    except: pass

        import threading
        threading.Thread(target=worker, daemon=True).start()
"""

final_output = lines[:start_idx] + [new_func + '\n'] + lines[end_idx:]
with open(target, 'w', encoding='utf-8') as f:
    f.writelines(final_output)

print('Patched successfully!')
