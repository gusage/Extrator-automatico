import customtkinter as ctk
from tkinter import filedialog
import threading
import os

# Importa as funções do extrator
from Extrator import processar_pasta

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Extrator de Cartão Ponto")
        self.geometry("520x400")
        self.resizable(False, False)

        self.caminho_pdf = ""
        self.pasta_saida = ""

        # Título
        ctk.CTkLabel(self, text="Extrator de Cartão Ponto", font=("Arial", 20, "bold")).pack(pady=(30, 5))

        # Seleção de PDF
        frame_pdf = ctk.CTkFrame(self, fg_color="transparent")
        frame_pdf.pack(fill="x", padx=30)
        ctk.CTkButton(frame_pdf, text="Selecionar pasta de PDFs", width=140, command=self.selecionar_pasta_pdfs).pack(side="left")
        self.label_pdf = ctk.CTkLabel(frame_pdf, text="Nenhum arquivo selecionado", text_color="gray")
        self.label_pdf.pack(side="left", padx=10)

        # Seleção de pasta de saída
        frame_pasta = ctk.CTkFrame(self, fg_color="transparent")
        frame_pasta.pack(fill="x", padx=30, pady=15)
        ctk.CTkButton(frame_pasta, text="Pasta de destino", width=140, command=self.selecionar_pasta).pack(side="left")
        self.label_pasta = ctk.CTkLabel(frame_pasta, text="Nenhuma pasta selecionada", text_color="gray")
        self.label_pasta.pack(side="left", padx=10)

        # Botão principal
        self.btn_gerar = ctk.CTkButton(
            self, text="Gerar Planilhas", height=45,
            font=("Arial", 14, "bold"), command=self.iniciar_processamento
        )
        self.btn_gerar.pack(padx=30, fill="x", pady=(10, 15))

        # Barra de progresso
        self.progresso = ctk.CTkProgressBar(self, width=460)
        self.progresso.set(0)
        self.progresso.pack(padx=30)

        # Status
        self.label_status = ctk.CTkLabel(self, text="Aguardando...", text_color="gray")
        self.label_status.pack(pady=10)

        # Log
        self.log = ctk.CTkTextbox(self, height=100, state="disabled")
        self.log.pack(padx=30, fill="x", pady=(0, 20))

    def selecionar_pasta_pdfs(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_pdfs = pasta
            pdfs = [f for f in os.listdir(pasta) if f.endswith(".pdf")]
            self.label_pdf.configure(
                text=f"{len(pdfs)} PDF(s) encontrado(s)", 
                text_color="white"
            )

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.pasta_saida = pasta
            self.label_pasta.configure(text=pasta, text_color="white")

    def adicionar_log(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def iniciar_processamento(self):
        if not self.pasta_pdfs:
            self.label_status.configure(text="⚠️ Selecione uma pasta de PDFs primeiro.", text_color="orange")
            return
        if not self.pasta_saida:
            self.label_status.configure(text="⚠️ Selecione a pasta de destino.", text_color="orange")
            return

        self.btn_gerar.configure(state="disabled")
        self.progresso.set(0)
        self.label_status.configure(text="Processando...", text_color="white")

        thread = threading.Thread(target=self.processar, daemon=True)
        thread.start()

    def processar(self):
        def atualizar(atual, total, nome):
            self.progresso.set(atual / total)
            self.label_status.configure(text=f"Processando {atual}/{total}...")
            self.adicionar_log(f"✓ {nome}")

        try:
            processar_pasta(self.pasta_pdfs, self.pasta_saida, callback=atualizar)
            self.label_status.configure(text="✅ Concluído!", text_color="green")
            os.startfile(self.pasta_saida)
        except Exception as e:
            self.label_status.configure(text=f"❌ Erro: {e}", text_color="red")
        finally:
            self.btn_gerar.configure(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()