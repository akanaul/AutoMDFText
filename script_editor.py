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
        self.title("Editor de Perfis MDF-e")
        self.geometry("860x560")
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

        ttk.Button(controls, text="Carregar / Atualizar", command=self.refresh_list).pack(side="left")
        ttk.Button(controls, text="Novo Perfil", command=self.new_profile).pack(side="left", padx=6)
        ttk.Button(controls, text="Novo do Template", command=self.new_profile_from_template).pack(side="left", padx=6)
        ttk.Button(controls, text="Assistente", command=self.wizard_create_script).pack(side="left", padx=6)
        ttk.Button(controls, text="Salvar", command=self.save_profile).pack(side="left", padx=6)
        ttk.Button(controls, text="Salvar Como", command=self.save_as_profile).pack(side="left", padx=6)

        self.editor = tk.Text(self, wrap="none", undo=True)
        self.editor.pack(fill="both", expand=True, padx=12, pady=6)

        scrollbar_y = ttk.Scrollbar(self.editor, orient="vertical", command=self.editor.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.editor.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(self, orient="horizontal", command=self.editor.xview)
        scrollbar_x.pack(side="bottom", fill="x")
        self.editor.configure(xscrollcommand=scrollbar_x.set)

        self.load_default()
        self.bind("<Return>", self.invoke_focused)

    def refresh_list(self) -> None:
        current = self.profile_var.get()
        self.profile_combo.config(values=list_profiles())
        if current in list_profiles():
            self.profile_var.set(current)

    def load_default(self) -> None:
        profiles = list_profiles()
        if not profiles:
            messagebox.showinfo("Info", "Nenhum perfil encontrado. Use Novo Perfil para criar um.")
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
            self.title(f"Editor de Perfis MDF-e — {profile}")

    def new_profile(self) -> None:
        name = simpledialog.askstring("Novo perfil", "Digite o nome do arquivo (com .txt):")
        if not name:
            return
        if not name.endswith(".txt"):
            name += ".txt"
        path = SCRIPTS_DIR / name
        if path.exists():
            messagebox.showerror("Erro", f"{name} ja existe")
            return
        path.write_text("# New profile\n", encoding="utf-8")
        self.profile_var.set(name)
        self.refresh_list()
        self.load_profile()

    def new_profile_from_template(self) -> None:
        name = simpledialog.askstring("Novo perfil", "Digite o nome do arquivo (com .txt):")
        if not name:
            return
        if not name.endswith(".txt"):
            name += ".txt"
        path = SCRIPTS_DIR / name
        if path.exists():
            messagebox.showerror("Erro", f"{name} ja existe")
            return
        if TEMPLATE_FILE.exists():
            template_text = TEMPLATE_FILE.read_text(encoding="utf-8")
            template_text = self.blank_template_values(template_text)
        else:
            template_text = "# Novo perfil\n"
        path.write_text(template_text, encoding="utf-8")
        self.profile_var.set(name)
        self.refresh_list()
        self.load_profile()

    def save_profile(self) -> None:
        profile = self.profile_var.get()
        if not profile:
            messagebox.showwarning("Aviso", "Selecione ou crie um perfil primeiro.")
            return
        path = SCRIPTS_DIR / profile
        path.write_text(self.editor.get("1.0", "end-1c"), encoding="utf-8")
        messagebox.showinfo("Salvo", f"{profile} atualizado")

    def save_as_profile(self) -> None:
        name = simpledialog.askstring("Salvar como", "Digite o novo nome do perfil (com .txt):")
        if not name:
            return
        if not name.endswith(".txt"):
            name += ".txt"
        path = SCRIPTS_DIR / name
        if path.exists():
            if not messagebox.askyesno("Substituir", f"{name} ja existe. Substituir?"):
                return
        path.write_text(self.editor.get("1.0", "end-1c"), encoding="utf-8")
        self.profile_var.set(name)
        self.refresh_list()
        messagebox.showinfo("Salvo", f"{name} criado")

    def wizard_create_script(self) -> None:
        if not TEMPLATE_FILE.exists():
            messagebox.showerror("Erro", "Template nao encontrado: template_config.txt")
            return

        dialog = tk.Toplevel(self)
        dialog.title("Assistente de Scripts")
        dialog.geometry("640x520")
        dialog.transient(self)
        dialog.grab_set()
        dialog.bind("<Return>", lambda event: submit())
        dialog.bind("<Escape>", lambda event: dialog.destroy())

        font = ("Segoe UI", 12)
        form = ttk.Frame(dialog, padding=12)
        form.pack(fill="both", expand=True)

        def digits_only_max(max_len: int):
            def _validator(value: str) -> bool:
                return value.isdigit() and len(value) <= max_len or value == ""
            return _validator

        def letters_only_max(max_len: int):
            def _validator(value: str) -> bool:
                return value.isalpha() and len(value) <= max_len or value == ""
            return _validator

        def letters_spaces_no_double(max_len: int):
            def _validator(value: str) -> bool:
                if value == "":
                    return True
                if len(value) > max_len:
                    return False
                if "  " in value:
                    return False
                for ch in value:
                    if not (ch.isalpha() or ch == " "):
                        return False
                return True
            return _validator


        fields = [
            ("uf_carregamento", "UF de carregamento"),
            ("uf_descarga", "UF de descarga"),
            ("municipio_carregamento", "Municipio de carregamento"),
            ("ncm_primary", "NCM 1 (mais usado)"),
            ("ncm_secondary", "NCM 2"),
            ("ncm_tertiary", "NCM 3"),
            ("contratante_cnpj", "CNPJ da unidade"),
            ("cep_origem", "CEP de origem"),
            ("cep_destino", "CEP de destino"),
            ("frete_valor", "Valor do frete"),
        ]

        entries: dict[str, tk.Entry] = {}
        entry_vars: dict[str, tk.StringVar] = {}
        for row, (key, label) in enumerate(fields):
            ttk.Label(form, text=label, font=font).grid(row=row, column=0, sticky="w", pady=6)
            entry_var = tk.StringVar()
            entry = tk.Entry(form, font=font, textvariable=entry_var)
            if key in {"cep_origem", "cep_destino"}:
                entry.config(
                    validate="key",
                    validatecommand=(dialog.register(digits_only_max(8)), "%P"),
                )
            elif key in {"uf_carregamento", "uf_descarga"}:
                entry.config(
                    validate="key",
                    validatecommand=(dialog.register(letters_only_max(2)), "%P"),
                )
            elif key == "contratante_cnpj":
                entry.config(
                    validate="key",
                    validatecommand=(dialog.register(digits_only_max(14)), "%P"),
                )
            elif key == "municipio_carregamento":
                entry.config(
                    validate="key",
                    validatecommand=(dialog.register(letters_spaces_no_double(60)), "%P"),
                )
            elif key in {"ncm_primary", "ncm_secondary", "ncm_tertiary"}:
                entry.config(
                    validate="key",
                    validatecommand=(dialog.register(digits_only_max(8)), "%P"),
                )
            elif key == "frete_valor":
                entry.config(
                    validate="key",
                    validatecommand=(dialog.register(digits_only_max(10)), "%P"),
                )
            entry.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=6)
            entries[key] = entry
            entry_vars[key] = entry_var

        def format_frete_value(*_args) -> None:
            var = entry_vars.get("frete_valor")
            if not var:
                return
            raw = var.get()
            digits = "".join(ch for ch in raw if ch.isdigit())
            if not digits:
                return
            digits = digits[-10:]
            if len(digits) <= 2:
                formatted = digits
            else:
                formatted = f"{digits[:-2]}.{digits[-2:]}"
            if raw != formatted:
                var.set(formatted)

        if "frete_valor" in entry_vars:
            entry_vars["frete_valor"].trace_add("write", format_frete_value)

        if fields:
            entries[fields[0][0]].focus_set()

        form.columnconfigure(1, weight=1)

        def submit() -> None:
            values = {key: entry.get().strip() for key, entry in entries.items()}
            frete_raw = values.get("frete_valor", "")
            if frete_raw:
                digits = "".join(ch for ch in frete_raw if ch.isdigit())
                if len(digits) <= 2:
                    values["frete_valor"] = f"0.{digits.zfill(2)}"
                else:
                    values["frete_valor"] = f"{digits[:-2]}.{digits[-2:]}"
            missing = [label for key, label in fields if not values.get(key)]
            if missing:
                messagebox.showwarning("Dados incompletos", "Preencha todos os campos antes de continuar.")
                return

            script_name = simpledialog.askstring(
                "Nome do script",
                "Digite o nome do arquivo (com .txt):",
                parent=dialog,
            )
            if not script_name:
                return
            if not script_name.endswith(".txt"):
                script_name += ".txt"

            target_path = SCRIPTS_DIR / script_name
            if target_path.exists():
                if not messagebox.askyesno("Substituir", f"{script_name} ja existe. Substituir?", parent=dialog):
                    return

            template_text = TEMPLATE_FILE.read_text(encoding="utf-8")
            updated_text = self.apply_template_replacements(template_text, values)
            target_path.write_text(updated_text, encoding="utf-8")

            dialog.destroy()
            self.profile_var.set(script_name)
            self.refresh_list()
            self.load_profile()
            messagebox.showinfo("Criado", f"{script_name} criado na pasta scripts")

        actions = ttk.Frame(form)
        actions.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(actions, text="Cancelar", command=dialog.destroy).pack(side="right")
        ttk.Button(actions, text="Criar", command=submit).pack(side="right", padx=6)

    @staticmethod
    def invoke_focused(event) -> None:
        widget = event.widget
        if hasattr(widget, "invoke"):
            widget.invoke()

    @staticmethod
    def apply_template_replacements(template_text: str, values: dict[str, str]) -> str:
        lines = []
        for line in template_text.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in line:
                lines.append(line)
                continue
            key = line.split("=", 1)[0].strip()
            if key in values:
                indent = line[: len(line) - len(line.lstrip())]
                lines.append(f"{indent}{key} = {values[key]}")
            else:
                lines.append(line)
        if template_text.endswith("\n"):
            return "\n".join(lines) + "\n"
        return "\n".join(lines)

    @staticmethod
    def blank_template_values(template_text: str) -> str:
        lines = []
        for line in template_text.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in line:
                lines.append(line)
                continue
            key = line.split("=", 1)[0].strip()
            indent = line[: len(line) - len(line.lstrip())]
            lines.append(f"{indent}{key} = ")
        if template_text.endswith("\n"):
            return "\n".join(lines) + "\n"
        return "\n".join(lines)


if __name__ == "__main__":
    app = ScriptEditor()
    app.mainloop()