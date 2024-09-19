import os
import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import win32com.client
import pyperclip
import re

class PedidoGerenciador:
    def __init__(self, root, base_dir):
        self.root = root
        self.base_dir = base_dir
        self.root.title("Gerenciador de Pedidos por Cliente")

        # Treeview para Clientes e Pedidos
        self.tree = ttk.Treeview(root, columns=("Nome"), show="tree", selectmode="browse")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Bind para duplo clique para abrir a pasta da cliente ou os pedidos
        self.tree.bind("<Double-1>", self.abrir_item)

        # Scrollbar
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side=tk.LEFT, fill="y")
        self.tree.configure(yscroll=self.scrollbar.set)

        # Botão de Refresh
        self.refresh_button = tk.Button(root, text="Refresh", command=self.atualizar_lista)
        self.refresh_button.pack(pady=5)

        # Botão de Processar Dados
        self.process_button = tk.Button(root, text="Processar Dados", command=self.processar_dados)
        self.process_button.pack(pady=5)

        # Botão de Marcar e Salvar
        self.marcar_button = tk.Button(root, text="Marcar e Salvar", command=self.marcar_e_salvar)
        self.marcar_button.pack(pady=5)

        # Carrega a lista de clientes e pedidos
        self.atualizar_lista()

    def atualizar_lista(self):
        """Atualiza a lista de clientes e seus pedidos."""
        self.tree.delete(*self.tree.get_children())  # Limpa a Treeview

        for cliente in os.listdir(self.base_dir):
            cliente_path = os.path.join(self.base_dir, cliente)

            if os.path.isdir(cliente_path):
                # Adiciona o cliente como um nó na Treeview
                cliente_node = self.tree.insert("", "end", text=cliente, open=False)

                # Adiciona os pedidos e atalhos desse cliente como sub-itens
                itens = os.listdir(cliente_path)
                for item in itens:
                    item_path = os.path.join(cliente_path, item)
                    if os.path.isdir(item_path) or item.endswith(".lnk"):
                        self.tree.insert(cliente_node, "end", text=item)

    def abrir_item(self, event):
        """Abre a pasta da cliente ou o pedido/atalho com duplo clique."""
        item_id = self.tree.selection()
        item_text = self.tree.item(item_id, "text")
        parent_id = self.tree.parent(item_id)

        if parent_id:  # Se houver um parent, é um pedido ou atalho
            cliente_text = self.tree.item(parent_id, "text")
            item_path = os.path.join(self.base_dir, cliente_text, item_text)

            if os.path.exists(item_path):
                if item_text.endswith(".lnk"):
                    # Abre atalhos .lnk
                    subprocess.run(['cmd', '/c', item_path], shell=True)
                else:
                    # Abre pastas de pedidos
                    os.startfile(item_path)
        else:
            # Se não houver parent, é uma cliente
            cliente_path = os.path.join(self.base_dir, item_text)
            if os.path.exists(cliente_path):
                os.startfile(cliente_path)  # Abre a pasta da cliente

    def resolver_atalho(self, caminho_atalho):
        """Resolve o caminho real de um atalho .lnk."""
        shell = win32com.client.Dispatch("WScript.Shell")
        atalho = shell.CreateShortcut(caminho_atalho)
        return atalho.TargetPath

    def processar_dados(self):
        """Processa dados da pasta ou atalho selecionado, calculando valores e formatando a saída."""
        item_id = self.tree.selection()  # Obtém o item selecionado
        if not item_id:
            tk.messagebox.showwarning("Seleção", "Nenhum item selecionado.")
            return

        item_text = self.tree.item(item_id, "text")
        parent_id = self.tree.parent(item_id)

        if parent_id:  # Se houver um parent, é um pedido ou atalho
            cliente_text = self.tree.item(parent_id, "text")
            item_path = os.path.join(self.base_dir, cliente_text, item_text)
        else:  # Se não houver parent, é uma cliente
            item_path = os.path.join(self.base_dir, item_text)

        if not os.path.exists(item_path):
            tk.messagebox.showerror("Erro", "O item selecionado não existe.")
            return

        # Resolver o caminho se for um atalho
        if item_text.endswith(".lnk"):
            try:
                item_path = self.resolver_atalho(item_path)
            except Exception as e:
                tk.messagebox.showerror("Erro", f"Erro ao resolver o atalho: {e}")
                return

        # Processa a pasta e calcula os valores
        total_valores = self.listar_pastas_e_somar_valores(item_path)
        tk.messagebox.showinfo("Processamento Completo", f"Total de valores processados: R$ {total_valores:.2f}")

    def listar_pastas_e_somar_valores(self, caminho_base):
        """Lista as pastas e soma os valores encontrados nos arquivos de texto."""
        total_valores = 0.0
        lista_dados = []

        # Expressão regular para capturar o valor monetário no nome do arquivo
        padrao_valor = re.compile(r"R\$(\d+(?:\.\d{1,2})?)", re.IGNORECASE)

        for item in os.listdir(caminho_base):
            item_path = os.path.join(caminho_base, item)

            # Se for um atalho, resolver o caminho real
            if item.endswith(".lnk"):
                try:
                    item_path = self.resolver_atalho(item_path)
                    item = item[:-4]  # Remover a extensão .lnk do nome para exibição
                except Exception as e:
                    print(f"Erro ao resolver o atalho {item}: {e}")
                    continue

            # Verifica se o caminho resolvido é uma pasta
            if os.path.isdir(item_path):
                arquivos_txt = [f for f in os.listdir(item_path) if f.endswith('.txt')]  # Procura arquivos .txt

                for arquivo in arquivos_txt:
                    match = padrao_valor.search(arquivo)  # Procura o valor monetário no nome do arquivo
                    if match:
                        try:
                            valor_str = match.group(1)
                            valor_float = float(valor_str.replace(",", "."))  # Converte o valor para float
                            total_valores += valor_float
                            lista_dados.append((item, valor_float))
                        except ValueError:
                            print(f"Erro ao converter o valor '{valor_str}' no arquivo {arquivo}.")
                    else:
                        print(f"Nenhum valor monetário encontrado no nome do arquivo: {arquivo}")

        # Construir o resultado em uma string
        resultado = f"{'Nome':<40} {'Valor':>10}\n" + "-" * 50 + "\n"
        for nome, valor in lista_dados:
            resultado += f"{nome:<40} R$ {valor:>10.2f}\n"

        resultado += "\n" + "-" * 50 + "\n"
        resultado += f"{'Total:':<40} R$ {total_valores:>10.2f}\n"

        # Exibir o resultado e copiar para a área de transferência
        messagebox.showinfo("Resultado", resultado)
        pyperclip.copy(resultado)

        return total_valores

    def marcar_e_salvar(self):
        """Marca a pasta selecionada e salva seu nome em um arquivo TXT."""
        item_id = self.tree.selection()  # Obtém o item selecionado
        if not item_id:
            tk.messagebox.showwarning("Seleção", "Nenhum item selecionado.")
            return

        item_text = self.tree.item(item_id, "text")
        parent_id = self.tree.parent(item_id)

        if parent_id:  # Se houver um parent, é um pedido ou atalho
            cliente_text = self.tree.item(parent_id, "text")
            item_path = os.path.join(self.base_dir, cliente_text, item_text)
        else:  # Se não houver parent, é uma cliente
            item_path = os.path.join(self.base_dir, item_text)

        if not os.path.exists(item_path):
            tk.messagebox.showerror("Erro", "O item selecionado não existe.")
            return

        # Grava o nome da pasta selecionada no arquivo TXT
        try:
            with open(r"C:\Users\leand\OneDrive\Desktop\Aplicativos\Pedidos Pagos.txt", "a") as file:
                file.write(f"{item_text}\n")
            tk.messagebox.showinfo("Sucesso", f"{item_text} foi adicionado ao arquivo Pedidos Pagos.txt.")
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro ao salvar o nome no arquivo: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    base_dir = r"C:\Users\leand\3D Objects"  # Substitua pelo caminho da sua base de clientes e pedidos
    app = PedidoGerenciador(root, base_dir)
    root.mainloop()
