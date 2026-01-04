import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk


SCRIPTS_DIR = Path(__file__).parent / "scripts"
SCRIPTS_DIR.mkdir(exist_ok=True)
TEMPLATE_FILE = SCRIPTS_DIR / "template_config.txt"


def list_profiles() -> list[str]:
    return sorted(p.name for p in SCRIPTS_DIR.glob("*.txt"))


class ScriptEditor(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("MDF-e Profile Editor")
        self.geometry("700x520")
        self.resizable(True, True)

        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(
            self,
            values=list_profiles(),
            textvariable=self.profile_var,
            state="readonly",
        )
        self.profile_combo.bind("<<ComboboxSelected>>", self.load_profile)
        self.profile_combo.pack(fill="x", padx=12, pady=8)

        controls = ttk.Frame(self)
        controls.pack(fill="x", padx=12)

        ttk.Button(controls, text="Load / Refresh", command=self.refresh_list).pack(side="left")
        ttk.Button(controls, text="New Profile", command=self.new_profile).pack(side="left", padx=6)
        ttk.Button(controls, text="New From Template", command=self.new_profile_from_template).pack(side="left", padx=6)
        ttk.Button(controls, text="Save", command=self.save_profile).pack(side="left", padx=6)
        ttk.Button(controls, text="Save As", command=self.save_as_profile).pack(side="left", padx=6)

        self.editor = tk.Text(self, wrap="none", undo=True)
        self.editor.pack(fill="both", expand=True, padx=12, pady=6)

        scrollbar_y = ttk.Scrollbar(self.editor, orient="vertical", command=self.editor.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.editor.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=self.editor.xview)
        scrollbar_x.pack(side="bottom", fill="x")
        self.editor.configure(xscrollcommand=scrollbar_x.set)

        self.load_default()

    def refresh_list(self) -> None:
        current = self.profile_var.get()
        self.profile_combo.config(values=list_profiles())
        if current in list_profiles():
            self.profile_var.set(current)

    def load_default(self) -> None:
        profiles = list_profiles()
        if not profiles:
            messagebox.showinfo("Info", "No profiles found. Use New Profile to create one.")
            return
        self.profile_var.set(profiles[0])
        self.load_profile()

    def load_profile(self, event=None) -> None:
        profile = self.profile_var.get()
        if not profile:
            return
        path = SCRIPTS_DIR / profile
        if path.exists():
            self.editor.delete("1.0", "end")
            self.editor.insert("1.0", path.read_text(encoding="utf-8"))
            self.title(f"MDF-e Profile Editor — {profile}")

    def new_profile(self) -> None:
        name = simpledialog.askstring("New profile", "Enter profile file name (with .txt):")
        if not name:
            return
        if not name.endswith(".txt"):
            name += ".txt"
        path = SCRIPTS_DIR / name
        if path.exists():
            messagebox.showerror("Error", f"{name} already exists")
            return
        path.write_text("# New profile\n", encoding="utf-8")
        self.profile_var.set(name)
        self.refresh_list()
        self.load_profile()

    def new_profile_from_template(self) -> None:
        name = simpledialog.askstring("New profile", "Enter profile file name (with .txt):")
        if not name:
            return
        if not name.endswith(".txt"):
            name += ".txt"
        path = SCRIPTS_DIR / name
        if path.exists():
            messagebox.showerror("Error", f"{name} already exists")
            return
        template = TEMPLATE_FILE.read_text(encoding="utf-8") if TEMPLATE_FILE.exists() else "# New profile\n"
        path.write_text(template, encoding="utf-8")
        self.profile_var.set(name)
        self.refresh_list()
        self.load_profile()

    def save_profile(self) -> None:
        profile = self.profile_var.get()
        if not profile:
            messagebox.showwarning("Warning", "Select or create a profile first.")
            return
        path = SCRIPTS_DIR / profile
        path.write_text(self.editor.get("1.0", "end-1c"), encoding="utf-8")
        messagebox.showinfo("Saved", f"{profile} updated")

    def save_as_profile(self) -> None:
        name = simpledialog.askstring("Save as", "Enter new profile name (with .txt):")
        if not name:
            return
        if not name.endswith(".txt"):
            name += ".txt"
        path = SCRIPTS_DIR / name
        if path.exists():
            if not messagebox.askyesno("Overwrite", f"{name} already exists. Overwrite?"):
                return
        path.write_text(self.editor.get("1.0", "end-1c"), encoding="utf-8")
        self.profile_var.set(name)
        self.refresh_list()
        messagebox.showinfo("Saved", f"{name} created")


if __name__ == "__main__":
    app = ScriptEditor()
    app.mainloop()