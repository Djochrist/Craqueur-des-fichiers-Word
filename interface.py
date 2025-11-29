import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.scrolledtext as scrolledtext

from core import analyser_fichier

def run_app():
    root = tk.Tk()
    root.title("Craqueur de fichiers Word")
    root.geometry("780x520")

    # --- Top frame: path selection ---
    frm_top = ttk.Frame(root, padding=8)
    frm_top.pack(fill=tk.X)

    ttk.Label(frm_top, text="Fichier sélectionné :").pack(side=tk.LEFT, padx=(0,6))
    var_path = tk.StringVar()
    ent_path = ttk.Entry(frm_top, textvariable=var_path, width=72)
    ent_path.pack(side=tk.LEFT, expand=True, fill=tk.X)

    def append_log(msg):
        def _append():
            txt_logs.config(state=tk.NORMAL)
            txt_logs.insert(tk.END, msg + "\n")
            txt_logs.see(tk.END)
            txt_logs.config(state=tk.DISABLED)
        root.after(0, _append)

    def browse_file():
        path = filedialog.askopenfilename(
            title="Sélectionner un fichier Word",
            filetypes=[("Word files", "*.docx;*.doc"), ("All files", "*.*")]
        )
        if path:
            var_path.set(path)
            append_log(f"[UI] Fichier sélectionné : {path}")

    btn_browse = ttk.Button(frm_top, text="Parcourir", command=browse_file)
    btn_browse.pack(side=tk.LEFT, padx=6)

    # --- Middle: logs ---
    frm_middle = ttk.Frame(root, padding=8)
    frm_middle.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frm_middle, text="Logs :").pack(anchor=tk.W)
    txt_logs = scrolledtext.ScrolledText(frm_middle, height=20, state=tk.DISABLED)
    txt_logs.pack(fill=tk.BOTH, expand=True, pady=4)

    # --- Bottom: controls & progress ---
    frm_bottom = ttk.Frame(root, padding=8)
    frm_bottom.pack(fill=tk.X)

    ttk.Label(frm_bottom, text="Longueur max brute :").pack(side=tk.LEFT, padx=(0,6))
    var_brute_len = tk.IntVar(value=4)
    spin_brute = ttk.Spinbox(frm_bottom, from_=1, to=6, width=4, textvariable=var_brute_len)
    spin_brute.pack(side=tk.LEFT)

    progress = ttk.Progressbar(frm_bottom, orient="horizontal", mode="determinate", length=460)
    progress.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.X)

    def set_progress(val):
        progress['value'] = val

    def set_progress_safe(val):
        root.after(0, lambda: set_progress(val))

    # Worker thread
    def worker(chemin, max_brute):
        try:
            ok = analyser_fichier(chemin, append_log, set_progress_safe, max_brute_length=max_brute)
            if ok:
                append_log("[RESULT] Mot de passe trouvé / décryption réussie.")
                root.after(0, lambda: messagebox.showinfo("Terminé", "Mot de passe trouvé / décryption réussie."))
            else:
                append_log("[RESULT] Mot de passe non trouvé.")
                root.after(0, lambda: messagebox.showwarning("Terminé", "Mot de passe non trouvé."))
        except Exception as e:
            append_log(f"[EXCEPTION] {e}")
            root.after(0, lambda: messagebox.showerror("Erreur", f"Exception pendant l'analyse : {e}"))
        finally:
            root.after(0, lambda: (btn_analyze.config(state=tk.NORMAL), btn_browse.config(state=tk.NORMAL), spin_brute.config(state='normal')))

    def analyze():
        chemin = var_path.get().strip()
        if not chemin:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un fichier .doc/.docx avant d'analyser.")
            return
        # Préparer UI
        btn_analyze.config(state=tk.DISABLED)
        btn_browse.config(state=tk.DISABLED)
        spin_brute.config(state='disabled')
        txt_logs.config(state=tk.NORMAL)
        txt_logs.delete('1.0', tk.END)
        txt_logs.config(state=tk.DISABLED)
        set_progress(0)
        append_log(f"[UI] Lancement de l'analyse pour : {chemin}")
        max_brute = int(var_brute_len.get())
        t = threading.Thread(target=worker, args=(chemin, max_brute), daemon=True)
        t.start()

    btn_analyze = ttk.Button(frm_bottom, text="Analyser le fichier", command=analyze)
    btn_analyze.pack(side=tk.LEFT, padx=6)

    root.mainloop()

if __name__ == "__main__":
    run_app()
