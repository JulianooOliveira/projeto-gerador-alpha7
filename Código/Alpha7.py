import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.font as tkfont
import os
import threading
import pandas as pd


class FileGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Arquivo Alpha7")
        self.root.resizable(False, False)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        x = (screen_width - 600) // 2
        y = (screen_height - 350) // 2

        self.root.geometry(f"580x220+{x}+{y}")

        self.planilha_path_tabloide = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        frame_tabloide = tk.Frame(
            self.root, borderwidth=2, relief="solid", padx=10, pady=10
        )
        frame_tabloide.pack(padx=10, pady=10, fill="both", expand=True)

        titulo_tabloide = "Gerador de Arquivo Alpha7"
        my_font_titulo = tkfont.Font(size=14, weight="bold")
        label_titulo_tabloide = tk.Label(
            frame_tabloide, text=titulo_tabloide, font=my_font_titulo
        )
        label_titulo_tabloide.grid(row=0, column=0, columnspan=3, pady=5)

        label_text_tabloide = (
            "Selecione a planilha base para gerar o Arquivo_Promoção.txt"
        )
        my_font_text = tkfont.Font(size=10)
        label_text_tabloide = tk.Label(
            frame_tabloide, text=label_text_tabloide, font=my_font_text
        )
        label_text_tabloide.grid(row=1, column=0, columnspan=3, pady=5)

        tk.Label(frame_tabloide, text="Selecione o arquivo:").grid(
            row=2, column=0, sticky="w", pady=10
        )
        tk.Entry(
            frame_tabloide, textvariable=self.planilha_path_tabloide, width=50
        ).grid(row=2, column=1, pady=10, padx=5)
        tk.Button(
            frame_tabloide, text="Procurar Arquivos", command=self.browse_file_tabloide
        ).grid(row=2, column=2, padx=5, pady=10)

        tk.Button(
            frame_tabloide,
            text="Gerar Arquivo Alpha7",
            command=self.generate_alpha7_file,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            width=20,
            height=2,
        ).grid(row=3, column=0, columnspan=3, pady=10)

    def generate_txt_alpha7_file(self):
        try:
            # Tentar ler o arquivo Excel
            try:
                # Primeira tentativa: ler sem especificar engine
                base = pd.read_excel(
                    self.planilha_path_tabloide.get(),
                    header=3,
                    usecols=["CÓDIGO DE BARRAS", "PREÇO DE VENDA"],
                )
            except ImportError:
                # Segunda tentativa: tentar com engine openpyxl
                base = pd.read_excel(
                    self.planilha_path_tabloide.get(),
                    header=3,
                    usecols=["CÓDIGO DE BARRAS", "PREÇO DE VENDA"],
                    engine="openpyxl",
                )

            # Converter código de barras para inteiro
            base["CÓDIGO DE BARRAS"] = base["CÓDIGO DE BARRAS"].astype("int64")

            # Converter preço para float - lidando com formato "R$ 16,99"
            def converter_preço(valor):
                if isinstance(valor, str):
                    # Remover "R$", espaços e converter vírgula para ponto
                    valor = valor.replace("R$", "").replace(" ", "").replace(",", ".")
                return float(valor)

            base["PREÇO DE VENDA"] = base["PREÇO DE VENDA"].apply(converter_preço)

            # Preparar estrutura para Alpha7
            base.insert(0, "A", "A")
            base.insert(1, "Barra1", "|")
            base.insert(3, "Barra3", "|")
            base.insert(4, "Barra4", "|")
            base.insert(5, "Barra5", "|")

            excel_directory = os.path.dirname(self.planilha_path_tabloide.get())
            promo_file_path = os.path.join(excel_directory, "Arquivo_Promoção.txt")

            # Concatenate values without any separator
            output_text = base.to_string(
                index=False, header=False, index_names=False, col_space=0
            )

            # Remove spaces from the resulting string
            output_text = output_text.replace(" ", "")

            # Write the modified string to the file
            with open(promo_file_path, "w", encoding="utf-8") as file:
                file.write(output_text)
            return True

        except Exception as e:
            error_msg = f"Erro ao processar o arquivo!\n\nDetalhes: {str(e)}"

            # Mensagem mais específica para erro de dependência
            if "xlrd" in str(e).lower():
                error_msg += "\n\nDica: Instale a biblioteca necessária executando no prompt:\npip install xlrd openpyxl"

            messagebox.showerror("Erro!", error_msg)
            return False

    def browse_file_tabloide(self):
        filename = filedialog.askopenfilename(
            title="Selecione a planilha base",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.planilha_path_tabloide.set(filename)

    def generate_alpha7_file(self):
        try:
            path = self.planilha_path_tabloide.get()
            if not path or not os.path.exists(path):
                messagebox.showerror("Erro!", "Por favor, selecione um arquivo válido.")
                return False
        except:
            messagebox.showerror("Erro!", "Por favor, selecione um arquivo.")
            return False

        # Mostrar tela de carregamento
        loading_window = tk.Toplevel(self.root)
        loading_window.title("Carregando...")
        loading_window.resizable(False, False)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 300) // 2
        y = (screen_height - 100) // 2
        loading_window.geometry(f"580x220+{x}+{y}")

        font_size = tkfont.Font(size=12)
        label = tk.Label(
            loading_window, text="Gerando arquivo_promocao.txt...", font=font_size
        )
        label.pack(expand=True)

        loading_window.update()

        # Executar em thread separada para não travar a interface
        def worker():
            success = self.generate_txt_alpha7_file()
            loading_window.destroy()
            if success:
                messagebox.showinfo(
                    "Sucesso", "arquivo_promocao.txt gerado com sucesso!"
                )
            else:
                messagebox.showerror("Erro", "Falha ao gerar o arquivo.")

        threading.Thread(target=worker).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = FileGenerator(root)
    root.mainloop()
