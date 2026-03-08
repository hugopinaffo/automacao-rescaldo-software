import os
from tkinter import filedialog

import customtkinter as ctk

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

from automation_core import AutomacaoErro, executar_automacao


class AutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.title("Automação Rescaldo Softwares")
        self.resizable(False, False)
        self.arquivo_path = None
        self.arquivo_principal = os.getenv("PLANILHA_PRINCIPAL_PATH")

        self.setup_ui()

    def setup_ui(self):
        self.eval("tk::PlaceWindow . center")
        self.grid_columnconfigure(1, weight=1)

        labels_text = ["Requisição", "WO", "Software", "CVE"]
        self.entries = []

        for i, text in enumerate(labels_text):
            ctk.CTkLabel(self, text=text).grid(
                row=i, column=0, sticky="e", padx=15, pady=10
            )
            ent = ctk.CTkEntry(self, width=300)
            ent.grid(row=i, column=1, sticky="w", padx=15, pady=10)
            self.entries.append(ent)

        self.entry_req, self.entry_wo, self.entry_soft, self.entry_cve = self.entries
        self.entry_req.focus_set()

        self.setup_key_bindings()

        # --- Planilha Principal ---
        ctk.CTkLabel(self, text="Planilha Principal").grid(
            row=4, column=0, sticky="e", padx=15, pady=10
        )

        self.btn_sel_principal = ctk.CTkButton(
            self,
            text="Selecionar arquivo",
            command=self.selecionar_principal,
            fg_color="transparent",
            border_width=1,
            corner_radius=8,
        )
        self.btn_sel_principal.grid(row=4, column=1, sticky="w", padx=15, pady=10)

        texto_principal = "Nenhum arquivo selecionado"
        cor_texto_principal = "gray"
        if self.arquivo_principal:
            nome_p = os.path.basename(self.arquivo_principal)
            texto_principal = nome_p if len(nome_p) <= 25 else nome_p[:22] + "..."
            cor_texto_principal = ("black", "white")

        self.lbl_principal = ctk.CTkLabel(
            self,
            text=texto_principal,
            font=("Segoe UI", 11),
            text_color=cor_texto_principal,
        )
        self.lbl_principal.grid(row=4, column=1, sticky="e", padx=(0, 20))

        # --- Máquinas ---
        ctk.CTkLabel(self, text="Máquinas").grid(
            row=5, column=0, sticky="e", padx=15, pady=10
        )

        self.btn_sel_arquivo = ctk.CTkButton(
            self,
            text="Selecionar arquivo",
            command=self.selecionar_arquivo,
            fg_color="transparent",
            border_width=1,
            corner_radius=8,
        )
        self.btn_sel_arquivo.grid(row=5, column=1, sticky="w", padx=15, pady=10)

        self.lbl_arquivo = ctk.CTkLabel(
            self,
            text="Nenhum arquivo selecionado",
            font=("Segoe UI", 11),
            text_color="gray",
        )
        self.lbl_arquivo.grid(row=5, column=1, sticky="e", padx=(0, 20))

        # --- Executar ---
        self.btn_executar = ctk.CTkButton(
            self,
            text="Executar Automação",
            command=self.executar,
            fg_color="#2e7d32",
            hover_color="#1b5e20",
            height=40,
            corner_radius=8,
        )
        self.btn_executar.grid(
            row=6, column=0, columnspan=2, pady=20, padx=20, sticky="ew"
        )

        self.lbl_status = ctk.CTkLabel(self, text="", font=("Segoe UI", 12, "bold"))
        self.lbl_status.grid(row=7, column=0, columnspan=2, pady=10)

    def setup_key_bindings(self):
        def on_enter(event, idx):
            if idx < len(self.entries) - 1:
                self.entries[idx + 1].focus_set()
            else:
                self.btn_sel_arquivo.focus_set()
                self.btn_sel_arquivo.invoke()

        for i, ent in enumerate(self.entries):
            ent.bind("<Return>", lambda e, idx=i: on_enter(e, idx))
            ent.bind(
                "<Shift-Return>",
                lambda e, idx=i: self.entries[max(0, idx - 1)].focus_set(),
            )

    def selecionar_principal(self):
        caminho = filedialog.askopenfilename(
            title="Planilha Principal", filetypes=[("Excel", "*.xlsx *.xlsm")]
        )
        if caminho:
            self.arquivo_principal = caminho
            nome = os.path.basename(caminho)
            nome_resumido = nome if len(nome) <= 25 else nome[:22] + "..."
            self.lbl_principal.configure(
                text=nome_resumido, text_color=("black", "white")
            )

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Arquivo de hostnames", filetypes=[("Excel", "*.xlsx *.xlsm")]
        )
        if caminho:
            self.arquivo_path = caminho
            nome = os.path.basename(caminho)
            nome_resumido = nome if len(nome) <= 25 else nome[:22] + "..."
            self.lbl_arquivo.configure(
                text=nome_resumido, text_color=("black", "white")
            )

    def executar(self):
        if not getattr(self, "arquivo_principal", None):
            self.lbl_status.configure(
                text="Por favor, selecione a planilha principal.",
                text_color="#f84848",
            )
            return

        if not self.arquivo_path:
            self.lbl_status.configure(
                text="Por favor, selecione um arquivo de máquinas primeiro.",
                text_color="#f84848",
            )
            return

        self.btn_executar.configure(state="disabled")
        self.lbl_status.configure(text="Processando...", text_color="yellow")
        self.update_idletasks()

        # Desacopla a execução imediata para permitir refresh da UI do TKinter primeiro
        self.after(50, self._processar_automacao)

    def _processar_automacao(self):
        try:
            total = executar_automacao(
                self.entry_req.get(),
                self.entry_wo.get(),
                self.entry_soft.get(),
                self.entry_cve.get(),
                self.arquivo_path,
                self.arquivo_principal,
            )
            self.lbl_status.configure(
                text=f"Sucesso! {total} máquinas processadas.", text_color="#2e7d32"
            )
            self.after(3000, self.destroy)
        except AutomacaoErro as e:
            self.lbl_status.configure(text=str(e), text_color="#f84848")
        except Exception as e:
            self.lbl_status.configure(
                text=f"Erro Inesperado: {type(e).__name__}", text_color="#d32f2f"
            )
        finally:
            self.btn_executar.configure(state="normal")


if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()
