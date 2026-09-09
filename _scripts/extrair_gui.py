#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aplicação gráfica (desktop) de extração de comprovantes.

- Extrai um arquivo ou um lote (pasta) usando o extractor da categoria.
- Campos PRÉ-PREENCHIDOS pela classificação do extrator e EDITÁVEIS.
- Filtro por categoria, remoção de itens e pré-visualização da imagem.
- Gera o dados_extraidos.json e regenera o comprovantes.csv (para o VBA).

Executar: python extrair_gui.py  (ou pelo atalho "Extrair Comprovantes")
"""
import os
import sys
import json
import subprocess
import hashlib
import datetime
import csv
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from extrair_categoria import EXTRACTORS, CONFIG, BASE  # noqa: E402
from extrair_ferramenta import extrair_um, extrair_lote  # noqa: E402
from classificacao import classificar_despesa  # noqa: E402
from indices_categoria import mesclar_indice_categoria  # noqa: E402

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    HAS_PIL = False

CATEGORIAS = list(EXTRACTORS.keys())

COR_FUNDO = "#eef3f8"
COR_PAINEL = "#f8fafc"
COR_TEXTO = "#172033"
COR_MUTED = "#64748b"
COR_BORDA = "#cbd5e1"
COR_AZUL = "#2563eb"
COR_VERDE = "#16803c"

CORES_CATEGORIA = {
    "Alimentação": "#f6d9b8",
    "Comunicação e tecnologia": "#c9def8",
    "Lazer e convivência": "#ead2f4",
    "Moradia": "#cfe6cf",
    "Saúde": "#f4cbd6",
    "Transporte": "#cbe8e5",
    "Educação": "#fff0b8",
    "Vestuário e higiene": "#e8d1eb",
}

DESPESAS = {
    "Alimentação": ["Supermercado e alimentação domiciliar",
                    "Suplementos e alimentação especial",
                    "Alimentação escolar / cantina"],
    "Moradia": ["Condomínio – quota-parte", "IPTU – quota-parte",
                "Energia elétrica – quota-parte", "Água e esgoto – quota-parte",
                "Gás – quota-parte", "Internet residencial – quota-parte",
                "Aluguel ou financiamento – quota-parte do adolescente",
                "Outras despesas comprovadas"],
    "Saúde": ["Medicamentos contínuos", "Consultas médicas particulares",
              "Tratamento odontológico", "Exames médicos",
              "Óculos, lentes e itens ortopédicos", "Psicólogo / psiquiatra"],
    "Transporte": ["Aplicativos de transporte", "Transporte escolar",
                   "Passagens de transporte coletivo",
                   "Combustível para deslocamentos"],
    "Educação": ["Mensalidade escolar", "Curso preparatório / profissionalizante",
                 "Livros didáticos e paradidáticos", "Cursos de música, arte ou cultura",
                 "Matrícula escolar"],
    "Vestuário e higiene": ["Roupas", "Calçados", "Cabeleireiro / cuidados pessoais",
                            "Produtos de higiene pessoal"],
    "Lazer e convivência": ["Assinaturas de streaming – quota-parte",
                            "Cinema, parques e passeios", "Viagens e férias"],
    "Comunicação e tecnologia": ["Plano de telefonia celular", "Aparelho celular"],
}


class App:
    def __init__(self, root):
        self.root = root
        root.title("Extração de Comprovantes")
        root.geometry("1180x700")
        root.minsize(1040, 660)
        root.configure(bg=COR_FUNDO)
        self.resultados = []        # lista de dicts (dados de cada comprovante)
        self.idx_atual = -1         # índice do item selecionado na lista completa
        self.foto = None

        self.var_entrada = tk.StringVar()
        self.var_categoria = tk.StringVar()
        self.var_filtro = tk.StringVar(value="Todas")
        self.var_data = tk.StringVar()
        self.var_valor = tk.StringVar()
        self.var_pagador = tk.StringVar()
        self.var_recebedor = tk.StringVar()
        self.var_tipo = tk.StringVar()
        self.var_despesa = tk.StringVar()
        self.var_periodicidade = tk.StringVar(value="Mensal")
        self.var_status = tk.StringVar()

        self._configurar_estilo()
        self._construir()

    def _configurar_estilo(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Segoe UI", 10), background=COR_FUNDO, foreground=COR_TEXTO)
        style.configure("App.TFrame", background=COR_FUNDO)
        style.configure("Panel.TFrame", background=COR_PAINEL, relief="flat")
        style.configure("Header.TLabel", background=COR_FUNDO, foreground=COR_TEXTO,
                        font=("Segoe UI Semibold", 18))
        style.configure("Subtle.TLabel", background=COR_FUNDO, foreground=COR_MUTED)
        style.configure("Panel.TLabel", background=COR_PAINEL, foreground=COR_TEXTO)
        style.configure("Status.TLabel", background="#dbeafe", foreground="#1e3a8a",
                        padding=(10, 6))
        style.configure("TEntry", fieldbackground="#ffffff", bordercolor=COR_BORDA,
                        lightcolor=COR_BORDA, darkcolor=COR_BORDA)
        style.configure("TCombobox", fieldbackground="#ffffff", bordercolor=COR_BORDA,
                        arrowcolor=COR_TEXTO)
        style.configure("TButton", padding=(12, 6), borderwidth=1)
        style.configure("Primary.TButton", background=COR_AZUL, foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#1d4ed8")])
        style.configure("Success.TButton", background=COR_VERDE, foreground="#ffffff")
        style.map("Success.TButton", background=[("active", "#166534")])
        style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff",
                        foreground=COR_TEXTO, rowheight=24, bordercolor=COR_BORDA)
        style.configure("Treeview.Heading", background="#1f4e79", foreground="#ffffff",
                        font=("Segoe UI Semibold", 10), padding=(6, 5))

    # ================= construção da interface =================
    def _construir(self):
        container = ttk.Frame(self.root, style="App.TFrame", padding=10)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Extração de Comprovantes", style="Header.TLabel").pack(anchor="w")
        ttk.Label(container, text="Extraia, revise e gere a fila para importação no Excel.",
                  style="Subtle.TLabel").pack(anchor="w", pady=(2, 8))

        topo = ttk.Frame(container, style="Panel.TFrame", padding=10)
        topo.pack(fill="x")
        topo.columnconfigure(1, weight=1)
        ttk.Label(topo, text="Arquivo ou pasta", style="Panel.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(topo, textvariable=self.var_entrada).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(topo, text="Arquivo...", command=self.selecionar_arquivo).grid(row=0, column=2, padx=2)
        ttk.Button(topo, text="Pasta...", command=self.selecionar_pasta).grid(row=0, column=3, padx=2)
        ttk.Label(topo, text="Categoria", style="Panel.TLabel").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Combobox(topo, textvariable=self.var_categoria, values=CATEGORIAS,
                     width=40, state="readonly").grid(row=1, column=1, sticky="w", padx=8, pady=(10, 0))
        ttk.Button(topo, text="Extrair", command=self.extrair,
                   style="Primary.TButton").grid(row=1, column=2, columnspan=2, sticky="ew", padx=2, pady=(10, 0))

        # barra de filtro + ações
        barra = ttk.Frame(container, style="App.TFrame", padding=(0, 8, 0, 6))
        barra.pack(fill="x")
        ttk.Label(barra, text="Filtrar categoria:", style="Subtle.TLabel").pack(side="left")
        self.cbo_filtro = ttk.Combobox(barra, textvariable=self.var_filtro,
                                       values=["Todas"] + CATEGORIAS, width=24, state="readonly")
        self.cbo_filtro.pack(side="left", padx=4)
        self.cbo_filtro.bind("<<ComboboxSelected>>", lambda e: self.atualizar_treeview())
        ttk.Button(barra, text="Ver comprovante", command=self.abrir_comprovante).pack(side="left", padx=4)
        ttk.Button(barra, text="Remover item", command=self.remover).pack(side="left", padx=4)

        # meio: tabela + preview
        meio = ttk.Frame(container, style="App.TFrame")
        meio.pack(fill="both", expand=True)
        self._construir_tabela(meio)
        self._construir_preview(meio)

        # base: campos editáveis + ações
        base = ttk.Frame(container, style="Panel.TFrame", padding=10)
        base.pack(fill="x", pady=(8, 0))
        self._construir_campos(base)

        ttk.Label(container, textvariable=self.var_status, style="Status.TLabel").pack(fill="x", pady=(6, 0))

    def _construir_tabela(self, pai):
        cols = ("arquivo", "data", "valor", "recebedor", "despesa", "periodicidade")
        self.tree = ttk.Treeview(pai, columns=cols, show="headings", selectmode="browse", height=8)
        for c, t in zip(cols, ("Arquivo", "Data", "Valor", "Recebedor", "Despesa", "Periodicidade")):
            self.tree.heading(c, text=t)
            self.tree.column(c, width={"arquivo": 200, "data": 90, "valor": 80,
                                       "recebedor": 150, "despesa": 200, "periodicidade": 90}[c])
        for categoria, cor in CORES_CATEGORIA.items():
            self.tree.tag_configure(categoria, background=cor, foreground=COR_TEXTO)
        self.tree.pack(side="left", fill="both", expand=True)
        barra = ttk.Scrollbar(pai, orient="vertical", command=self.tree.yview)
        barra.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=barra.set)
        self.tree.bind("<<TreeviewSelect>>", self.ao_selecionar)

    def _construir_preview(self, pai):
        painel = ttk.Frame(pai, width=290, style="Panel.TFrame", padding=8)
        painel.pack(side="right", fill="y", padx=6)
        painel.pack_propagate(False)
        ttk.Label(painel, text="Pré-visualização", style="Panel.TLabel",
                  font=("Segoe UI Semibold", 11)).pack(anchor="w")
        self.lbl_preview = ttk.Label(painel, text="Selecione um item", anchor="center",
                                     justify="center", style="Panel.TLabel")
        self.lbl_preview.pack(fill="both", expand=True, pady=4)

    def _construir_campos(self, base):
        f = ttk.Frame(base, style="Panel.TFrame")
        f.pack(fill="x")
        for col in (1, 3, 5):
            f.columnconfigure(col, weight=1)

        def campo(linha, col, label, var, width=22):
            ttk.Label(f, text=label, style="Panel.TLabel").grid(
                row=linha, column=col * 2, sticky="w", padx=(0, 4), pady=2
            )
            ttk.Entry(f, textvariable=var, width=width).grid(
                row=linha, column=col * 2 + 1, sticky="ew", padx=(0, 10), pady=2
            )

        campo(0, 0, "Data:", self.var_data, 18)
        campo(0, 1, "Valor:", self.var_valor, 14)
        ttk.Label(f, text="Periodicidade:", style="Panel.TLabel").grid(row=0, column=4, sticky="w", padx=(0, 4), pady=2)
        ttk.Combobox(f, textvariable=self.var_periodicidade,
                     values=["Mensal", "Anual", "Eventual"], width=18,
                     state="readonly").grid(row=0, column=5, sticky="ew", padx=(0, 10), pady=2)

        campo(1, 0, "Pagador:", self.var_pagador, 24)
        campo(1, 1, "Recebedor:", self.var_recebedor, 24)
        campo(2, 0, "Tipo:", self.var_tipo, 24)
        ttk.Label(f, text="Despesa:", style="Panel.TLabel").grid(row=2, column=2, sticky="w", padx=(0, 4), pady=2)
        self.cbo_despesa = ttk.Combobox(f, textvariable=self.var_despesa, width=28)
        self.cbo_despesa.grid(row=2, column=3, columnspan=3, sticky="ew", padx=(0, 10), pady=2)

        botoes = ttk.Frame(base, style="Panel.TFrame")
        botoes.pack(fill="x", pady=(6, 0))
        ttk.Button(botoes, text="Aplicar edição", command=self.aplicar).pack(side="left", padx=(0, 6))
        ttk.Button(botoes, text="Salvar tudo (JSON + CSV)", command=self.salvar,
                   style="Success.TButton").pack(side="left")
    # ================= ações =================
    def selecionar_arquivo(self):
        p = filedialog.askopenfilename(filetypes=[("Comprovantes", "*.png *.jpg *.jpeg *.pdf"), ("Todos", "*.*")])
        if p:
            self.var_entrada.set(p)

    def selecionar_pasta(self):
        p = filedialog.askdirectory()
        if p:
            self.var_entrada.set(p)

    def extrair(self):
        entrada = self.var_entrada.get().strip()
        cat = self.var_categoria.get()
        if not entrada or not cat:
            messagebox.showwarning("Atenção", "Informe o arquivo/pasta e a categoria.")
            return
        if not os.path.exists(entrada):
            messagebox.showerror("Erro", "Caminho não encontrado.")
            return
        try:
            if os.path.isfile(entrada):
                self.resultados = [self._completar(extrair_um(entrada, cat), cat)]
            else:
                self.resultados = [self._completar(r, cat) for r in extrair_lote(entrada, cat)]
        except Exception as e:
            messagebox.showerror("Erro na extração", str(e))
            return
        self.idx_atual = -1
        self.var_filtro.set("Todas")
        self.atualizar_treeview()
        self.var_status.set(f"{len(self.resultados)} comprovante(s) extraído(s).")

    def _completar(self, r, cat):
        r["despesa"] = classificar_despesa(cat, r)
        r["periodicidade"] = "Anual" if any(k in r["despesa"].lower()
                                             for k in ("iptu", "enem", "matrícula")) else "Mensal"
        return r

    def _itens_visiveis(self):
        f = self.var_filtro.get()
        return [(i, r) for i, r in enumerate(self.resultados) if f == "Todas" or r["categoria"] == f]

    def atualizar_treeview(self):
        self.tree.delete(*self.tree.get_children())
        for i, r in self._itens_visiveis():
            self.tree.insert("", "end", iid=str(i), values=(
                r["arquivo"], r.get("data") or "",
                f"{r['valor']:.2f}" if r.get("valor") is not None else "",
                r.get("recebedor") or "", r.get("despesa") or "",
                r.get("periodicidade") or "Mensal"),
                tags=(r.get("categoria") or "",))

    def ao_selecionar(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        self.idx_atual = int(sel[0])
        r = self.resultados[self.idx_atual]
        self.var_data.set(r.get("data") or "")
        self.var_valor.set(f"{r['valor']:.2f}" if r.get("valor") is not None else "")
        self.var_pagador.set(r.get("pagador") or "")
        self.var_recebedor.set(r.get("recebedor") or "")
        self.var_tipo.set(r.get("tipo") or "")
        self.var_despesa.set(r.get("despesa") or "")
        self.var_periodicidade.set(r.get("periodicidade") or "Mensal")
        self.cbo_despesa["values"] = DESPESAS.get(r["categoria"], [])
        self.mostrar_preview(r.get("caminho"))

    def mostrar_preview(self, caminho):
        if not HAS_PIL:
            self.lbl_preview.config(image="", text="(PIL não instalado)")
            return
        if not caminho or not os.path.exists(caminho):
            self.lbl_preview.config(image="", text="(sem arquivo)")
            return
        try:
            img = Image.open(caminho)
            img.thumbnail((220, 260))
            self.foto = ImageTk.PhotoImage(img)
            self.lbl_preview.config(image=self.foto, text="")
        except Exception:
            self.lbl_preview.config(image="", text="(PDF — use 'Ver comprovante')")

    def abrir_comprovante(self):
        if self.idx_atual < 0:
            messagebox.showinfo("Atenção", "Selecione um comprovante.")
            return
        caminho = self.resultados[self.idx_atual].get("caminho")
        if not caminho or not os.path.exists(caminho):
            messagebox.showwarning("Atenção", "Arquivo não encontrado.")
            return
        if sys.platform.startswith("win"):
            os.startfile(caminho)
        elif sys.platform == "darwin":
            subprocess.run(["open", caminho])
        else:
            subprocess.run(["xdg-open", caminho])

    def remover(self):
        if self.idx_atual < 0:
            messagebox.showinfo("Atenção", "Selecione um comprovante.")
            return
        if not messagebox.askyesno("Remover", "Remover este comprovante da lista?"):
            return
        del self.resultados[self.idx_atual]
        self.idx_atual = -1
        self.atualizar_treeview()
        self.var_status.set("Item removido.")

    def aplicar(self):
        if self.idx_atual < 0:
            messagebox.showinfo("Atenção", "Selecione um comprovante na lista.")
            return
        r = self.resultados[self.idx_atual]
        r["data"] = self.var_data.get().strip() or None
        try:
            r["valor"] = round(float(self.var_valor.get().replace(",", ".")), 2)
        except ValueError:
            r["valor"] = None
        r["pagador"] = self.var_pagador.get().strip() or None
        r["recebedor"] = self.var_recebedor.get().strip() or None
        r["tipo"] = self.var_tipo.get().strip() or None
        r["despesa"] = self.var_despesa.get().strip()
        r["periodicidade"] = self.var_periodicidade.get()
        self.atualizar_treeview()
        self.var_status.set("Edição aplicada.")

    def salvar(self):
        if not self.resultados:
            messagebox.showinfo("Atenção", "Nada para salvar.")
            return
        self.aplicar()
        entrada = self.var_entrada.get().strip()
        pasta_indice = entrada if os.path.isdir(entrada) else os.path.dirname(entrada)
        if pasta_indice and os.path.isdir(pasta_indice):
            _, _, normalizados = mesclar_indice_categoria(
                pasta_indice, self.var_categoria.get(), self.resultados
            )
            self.resultados = normalizados
        try:
            self.salvar_fila_selecionada()
            self.var_status.set("Salvo. fila JSON/CSV gerada somente com a selecao atual.")
        except Exception as e:
            self.var_status.set("Falha ao gerar fila selecionada: " + str(e))
            messagebox.showerror("Erro ao salvar", str(e))
            return
        messagebox.showinfo("Concluido", "Dados salvos.")

    def salvar_fila_selecionada(self):
        itens = []
        for ident, r in enumerate(self.resultados, 1):
            caminho = r.get("caminho") or r.get("arquivo") or ""
            arquivo_rel = self._relativo_base(caminho)
            data = r.get("data") or ""
            mes = data[:7] if data else ""
            item = {
                "id": ident,
                "status": "pendente",
                "categoria": r.get("categoria") or self.var_categoria.get(),
                "despesa": r.get("despesa") or classificar_despesa(self.var_categoria.get(), r),
                "data": data,
                "mes": mes,
                "valor": r.get("valor"),
                "pagador": r.get("pagador") or "",
                "recebedor": r.get("recebedor") or "",
                "tipo": r.get("tipo") or "",
                "arquivo": arquivo_rel,
                "hash": self._hash_arquivo(caminho),
                "periodicidade": r.get("periodicidade") or "Mensal",
            }
            itens.append(item)

        payload = {
            "meta": {
                "gerado_em": datetime.datetime.now().isoformat(timespec="seconds"),
                "base": BASE,
                "total_comprovantes": len(itens),
                "fonte": "selecao da ferramenta grafica",
            },
            "comprovantes": itens,
        }
        with open(os.path.join(BASE, "comprovantes.json"), "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False, indent=2)

        campos = ["id", "categoria", "despesa", "data", "mes", "valor", "pagador",
                  "recebedor", "tipo", "arquivo", "hash", "periodicidade"]
        with open(os.path.join(BASE, "comprovantes.csv"), "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=campos, delimiter=";", extrasaction="ignore")
            writer.writeheader()
            writer.writerows(itens)

    def _relativo_base(self, caminho):
        if not caminho:
            return ""
        try:
            return os.path.relpath(caminho, BASE)
        except ValueError:
            return caminho

    def _hash_arquivo(self, caminho):
        try:
            with open(caminho, "rb") as fh:
                return hashlib.md5(fh.read()).hexdigest()
        except Exception:
            return ""


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

