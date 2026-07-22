import customtkinter as ctk
from tkcalendar import DateEntry
import sqlite3
import os
import calendar
from datetime import date, timedelta, datetime
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import MaxNLocator
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# CATEGORIAS PRÉ-DEFINIDAS (Fase 6)
CATEGORIAS = [
    "Sem categoria", "💼 Trabalho", "📚 Estudos", "🏠 Casa", "💰 Financeiro", "❤️ Pessoal"
]

# NOMES EM PORTUGUÊS PARA O CALENDÁRIO (Fase 5)
MESES_PT = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]
DIAS_SEMANA_PT = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

lista_tarefas = []

entry_pesquisa = None
frame_lista_tarefas = None
indice_edicao = None
optionmenu_prioridade = None
optionmenu_categoria = None
optionmenu_ordenacao = None
checkbox_filtro_data_ativo = None
dateentry_filtro = None
checkbox_somente_favoritas = None

hoje = date.today()
ano_calendario = hoje.year
mes_calendario = hoje.month
dia_selecionado = None


# ============================================================
#  BANCO DE DADOS (SQLite) — persistência das tarefas em disco
# ============================================================
CAMINHO_BD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gpriority.db")
CAMINHO_EXPORTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exportacoes")

conexao_bd = sqlite3.connect(CAMINHO_BD)
conexao_bd.row_factory = sqlite3.Row


def criar_tabela():
    conexao_bd.execute("""
        CREATE TABLE IF NOT EXISTS tarefas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            titulo TEXT NOT NULL,
            descricao TEXT,
            prioridade TEXT,
            data TEXT,
            concluida INTEGER NOT NULL DEFAULT 0,
            favorita INTEGER NOT NULL DEFAULT 0,
            fixada INTEGER NOT NULL DEFAULT 0,
            categoria TEXT NOT NULL DEFAULT 'Sem categoria',
            tags TEXT NOT NULL DEFAULT ''
        )
    """)
    # Migração: adiciona colunas novas em bancos criados por versões anteriores do app
    colunas_existentes = {linha["name"] for linha in conexao_bd.execute("PRAGMA table_info(tarefas)")}
    if "categoria" not in colunas_existentes:
        conexao_bd.execute("ALTER TABLE tarefas ADD COLUMN categoria TEXT NOT NULL DEFAULT 'Sem categoria'")
    if "tags" not in colunas_existentes:
        conexao_bd.execute("ALTER TABLE tarefas ADD COLUMN tags TEXT NOT NULL DEFAULT ''")
    conexao_bd.commit()


def carregar_tarefas_do_bd():
    cursor = conexao_bd.execute("SELECT * FROM tarefas ORDER BY id ASC")
    tarefas = []
    for linha in cursor.fetchall():
        tarefas.append({
            "id": linha["id"],
            "titulo": linha["titulo"],
            "descricao": linha["descricao"] or "",
            "prioridade": linha["prioridade"] or "",
            "data": linha["data"] or "",
            "concluida": bool(linha["concluida"]),
            "favorita": bool(linha["favorita"]),
            "fixada": bool(linha["fixada"]),
            "categoria": linha["categoria"] or "Sem categoria",
            "tags": linha["tags"] or "",
        })
    return tarefas


def recarregar_tarefas():
    lista_tarefas[:] = carregar_tarefas_do_bd()


def inserir_tarefa_bd(tarefa):
    cursor = conexao_bd.execute(
        """INSERT INTO tarefas (titulo, descricao, prioridade, data, concluida, favorita, fixada, categoria, tags)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (
            tarefa["titulo"],
            tarefa["descricao"],
            tarefa["prioridade"],
            tarefa["data"],
            int(tarefa["concluida"]),
            int(tarefa["favorita"]),
            int(tarefa["fixada"]),
            tarefa["categoria"],
            tarefa["tags"],
        ),
    )
    conexao_bd.commit()
    return cursor.lastrowid


def atualizar_tarefa_bd(id_tarefa, tarefa):
    conexao_bd.execute(
        """UPDATE tarefas
           SET titulo=?, descricao=?, prioridade=?, data=?, concluida=?, favorita=?, fixada=?, categoria=?, tags=?
           WHERE id=?""",
        (
            tarefa["titulo"],
            tarefa["descricao"],
            tarefa["prioridade"],
            tarefa["data"],
            int(tarefa["concluida"]),
            int(tarefa["favorita"]),
            int(tarefa["fixada"]),
            tarefa["categoria"],
            tarefa["tags"],
            id_tarefa,
        ),
    )
    conexao_bd.commit()


def atualizar_campo_bd(id_tarefa, campo, valor):
    # 'campo' nunca vem de entrada do usuário — só é chamado com estes 3 nomes fixos
    assert campo in ("concluida", "favorita", "fixada")
    conexao_bd.execute(f"UPDATE tarefas SET {campo}=? WHERE id=?", (int(valor), id_tarefa))
    conexao_bd.commit()


def excluir_tarefa_bd(id_tarefa):
    conexao_bd.execute("DELETE FROM tarefas WHERE id=?", (id_tarefa,))
    conexao_bd.commit()


#  CORES DA JANELA TEMAS E BOTOES #
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

#  CRIANDO E CONFIGURANDO JANELA PRINCIPAL #
janela = ctk.CTk()
janela.title("Gpriority")
janela.geometry("1400x800")
janela.minsize(1200, 700)
janela.grid_columnconfigure(0, weight=0)
janela.grid_columnconfigure(1, weight=1)
janela.grid_rowconfigure(1, weight=1)
pesquisa = ctk.StringVar()


# APAGAR FRAME PRINCIPAL
def limpar_frame_principal():
    for widget in frame_principal.winfo_children():
        widget.destroy()


#  FUNÇÕES DE NAVEGAÇÃO #
def clicar_dashboard():
    mostrar_dashboard()


#  CARD DE ESTATÍSTICA REUTILIZÁVEL (usado no Dashboard) #
def criar_card(parent, titulo_texto, valor_texto, cor_valor=None):
    frame_card = ctk.CTkFrame(
        parent, corner_radius=15, width=220, height=120, fg_color="#1F2937"
    )
    frame_card.grid_propagate(False)
    frame_card.grid_columnconfigure(0, weight=1)

    titulo_card = ctk.CTkLabel(
        frame_card, text=titulo_texto, font=("Arial", 16, "bold")
    )
    titulo_card.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")

    valor_card = ctk.CTkLabel(
        frame_card,
        text=valor_texto,
        font=("Arial", 40, "bold"),
        text_color=cor_valor if cor_valor else "#F9FAFB",
    )
    valor_card.grid(row=1, column=0, pady=(5, 15))

    return frame_card


#  GRÁFICO: TAREFAS CONCLUÍDAS x PENDENTES (rosca) #
def criar_grafico_status(parent):
    total = len(lista_tarefas)
    concluidas = sum(1 for t in lista_tarefas if t.get("concluida", False))
    pendentes = total - concluidas

    figura = Figure(figsize=(4.1, 3.0), dpi=100, facecolor="#1F2937")
    eixo = figura.add_subplot(111)
    eixo.set_facecolor("#1F2937")
    eixo.set_title("Tarefas Concluídas", color="#E5E7EB", fontsize=13, fontweight="bold")

    if total == 0:
        eixo.text(0.5, 0.5, "Sem tarefas ainda", ha="center", va="center", color="#9CA3AF", fontsize=12)
        eixo.axis("off")
    else:
        dados = [(v, r, c) for v, r, c in (
            (concluidas, "Concluídas", "#10B981"),
            (pendentes, "Pendentes", "#374151"),
        ) if v > 0]

        eixo.pie(
            [d[0] for d in dados],
            labels=[d[1] for d in dados],
            colors=[d[2] for d in dados],
            autopct="%1.0f%%",
            textprops={"color": "#E5E7EB", "fontsize": 11},
            wedgeprops={"width": 0.45, "edgecolor": "#1F2937"},
            startangle=90,
        )

    figura.tight_layout()

    canvas = FigureCanvasTkAgg(figura, master=parent)
    canvas.draw()
    return canvas.get_tk_widget()


#  GRÁFICO: TAREFAS POR PRIORIDADE (barras) #
def criar_grafico_prioridade(parent):
    contagem = {"Alta": 0, "Média": 0, "Baixa": 0}
    for tarefa in lista_tarefas:
        prioridade = tarefa["prioridade"].strip()
        if prioridade in contagem:
            contagem[prioridade] += 1

    figura = Figure(figsize=(4.1, 3.0), dpi=100, facecolor="#1F2937")
    eixo = figura.add_subplot(111)
    eixo.set_facecolor("#1F2937")
    eixo.set_title("Tarefas por Prioridade", color="#E5E7EB", fontsize=13, fontweight="bold")

    categorias = list(contagem.keys())
    valores = list(contagem.values())
    cores = ["#EF4444", "#F59E0B", "#10B981"]

    barras = eixo.bar(categorias, valores, color=cores)
    eixo.set_ylim(bottom=0)
    eixo.tick_params(colors="#9CA3AF")
    eixo.yaxis.set_major_locator(MaxNLocator(integer=True))
    for lado in ("top", "right"):
        eixo.spines[lado].set_visible(False)
    for lado in ("left", "bottom"):
        eixo.spines[lado].set_color("#374151")

    for barra in barras:
        altura = barra.get_height()
        if altura > 0:
            eixo.text(
                barra.get_x() + barra.get_width() / 2, altura, str(int(altura)),
                ha="center", va="bottom", color="#E5E7EB", fontsize=10
            )

    figura.tight_layout()

    canvas = FigureCanvasTkAgg(figura, master=parent)
    canvas.draw()
    return canvas.get_tk_widget()


def mostrar_dashboard():
    limpar_frame_principal()

    titulo_dashboard = ctk.CTkLabel(
        frame_principal, text="Dashboard", font=("Arial", 28, "bold")
    )
    titulo_dashboard.grid(row=0, column=0, padx=30, pady=30, sticky="w")

    total = len(lista_tarefas)
    concluidas = sum(1 for t in lista_tarefas if t.get("concluida", False))
    pendentes = total - concluidas
    alta_prioridade = sum(1 for t in lista_tarefas if "Alta" in t["prioridade"])
    produtividade = (concluidas / total * 100) if total > 0 else 0

    # LINHA DE CARDS (Total / Concluídas / Pendentes / Alta prioridade)
    frame_cards = ctk.CTkFrame(frame_principal, fg_color="transparent")
    frame_cards.grid(row=1, column=0, padx=30, pady=10, sticky="w")

    criar_card(frame_cards, "Total de Tarefas", str(total)).grid(row=0, column=0, padx=(0, 15))
    criar_card(frame_cards, "Concluídas", str(concluidas), "#10B981").grid(row=0, column=1, padx=15)
    criar_card(frame_cards, "Pendentes", str(pendentes), "#F59E0B").grid(row=0, column=2, padx=15)
    criar_card(frame_cards, "Alta Prioridade", str(alta_prioridade), "#EF4444").grid(row=0, column=3, padx=(15, 0))

    # PERCENTUAL DE PRODUTIVIDADE
    frame_produtividade = ctk.CTkFrame(frame_principal, corner_radius=15, fg_color="#1F2937")
    frame_produtividade.grid(row=2, column=0, padx=30, pady=(10, 10), sticky="w")

    label_produtividade = ctk.CTkLabel(
        frame_produtividade,
        text=f"📈 Produtividade: {produtividade:.0f}%  ({concluidas} de {total} tarefas concluídas)",
        font=("Arial", 16, "bold"),
    )
    label_produtividade.grid(row=0, column=0, padx=20, pady=(15, 8), sticky="w")

    barra_produtividade = ctk.CTkProgressBar(
        frame_produtividade, width=900, height=16, progress_color="#10B981"
    )
    barra_produtividade.set(produtividade / 100)
    barra_produtividade.grid(row=1, column=0, padx=20, pady=(0, 18), sticky="w")

    # GRÁFICOS
    frame_graficos = ctk.CTkFrame(frame_principal, fg_color="transparent")
    frame_graficos.grid(row=3, column=0, padx=30, pady=(0, 30), sticky="w")

    frame_grafico_status = ctk.CTkFrame(
        frame_graficos, corner_radius=15, fg_color="#1F2937", width=430, height=320
    )
    frame_grafico_status.grid(row=0, column=0, padx=(0, 15))
    frame_grafico_status.grid_propagate(False)
    criar_grafico_status(frame_grafico_status).pack(fill="both", expand=True, padx=10, pady=10)

    frame_grafico_prioridade = ctk.CTkFrame(
        frame_graficos, corner_radius=15, fg_color="#1F2937", width=430, height=320
    )
    frame_grafico_prioridade.grid(row=0, column=1, padx=(15, 0))
    frame_grafico_prioridade.grid_propagate(False)
    criar_grafico_prioridade(frame_grafico_prioridade).pack(fill="both", expand=True, padx=10, pady=10)


#FUNÇÃO EDITAR TAREFA
def editar_tarefa(indice):
    global indice_edicao

    indice_edicao=indice

    nova_tarefa()


#FUNÇÃO NOVA TAREFA (garante modo de criação, não edição)
def clicar_nova_tarefa():
    global indice_edicao
    indice_edicao = None
    nova_tarefa()


#FUNÇÃO DE EXCLUIR TAREFA
def excluir_tarefa(indice):
    id_tarefa = lista_tarefas[indice]["id"]
    excluir_tarefa_bd(id_tarefa)
    recarregar_tarefas()
    atualizar_lista()

#CRIANDO FUNÇÃO DE TAREFA CONCLUIDA
def alternar_concluida (indice):
    tarefa = lista_tarefas[indice]
    novo_status = not tarefa.get("concluida", False)
    atualizar_campo_bd(tarefa["id"], "concluida", novo_status)
    recarregar_tarefas()
    atualizar_lista()


#FUNÇÃO DE FAVORITAR TAREFA
def alternar_favorita(indice):
    tarefa = lista_tarefas[indice]
    novo_status = not tarefa.get("favorita", False)
    atualizar_campo_bd(tarefa["id"], "favorita", novo_status)
    recarregar_tarefas()
    atualizar_lista()


#FUNÇÃO DE FIXAR TAREFA NO TOPO
def alternar_fixada(indice):
    tarefa = lista_tarefas[indice]
    novo_status = not tarefa.get("fixada", False)
    atualizar_campo_bd(tarefa["id"], "fixada", novo_status)
    recarregar_tarefas()
    atualizar_lista()


#FUNÇÃO PARA LIMPAR TODOS OS FILTROS/ORDENAÇÃO
def limpar_filtros():
    entry_pesquisa.delete(0, "end")
    optionmenu_prioridade.set("Todas")
    optionmenu_categoria.set("Todas")
    checkbox_filtro_data_ativo.deselect()
    optionmenu_ordenacao.set("Mais recente")
    checkbox_somente_favoritas.deselect()
    atualizar_lista()


# FUNÇÃO QUE APLICA PESQUISA + FILTROS + ORDENAÇÃO E REDESENHA A LISTA
def atualizar_lista(event=None):
    for widget in frame_lista_tarefas.winfo_children():
        widget.destroy()

    texto_pesquisa = entry_pesquisa.get().lower()
    filtro_prioridade = optionmenu_prioridade.get()
    filtro_categoria = optionmenu_categoria.get()
    filtro_data = dateentry_filtro.get() if checkbox_filtro_data_ativo.get() else None
    ordenacao = optionmenu_ordenacao.get()
    somente_favoritas = bool(checkbox_somente_favoritas.get())

    desenhar_cards(
        frame_lista_tarefas,
        texto_pesquisa,
        filtro_prioridade,
        filtro_categoria,
        filtro_data,
        ordenacao,
        somente_favoritas,
    )


def desenhar_cards(
    frame_lista_tarefas,
    texto_pesquisa="",
    filtro_prioridade="Todas",
    filtro_categoria="Todas",
    filtro_data=None,
    ordenacao="Mais recente",
    somente_favoritas=False,
):
    # Mantém junto de cada tarefa o seu índice real em lista_tarefas,
    # necessário para editar/excluir/favoritar mesmo depois de filtrar e ordenar
    itens = list(enumerate(lista_tarefas))

    # A pesquisa casa tanto com o título quanto com as tags da tarefa
    itens = [
        (i, t) for i, t in itens
        if texto_pesquisa in t["titulo"].lower() or texto_pesquisa in t.get("tags", "").lower()
    ]

    if filtro_prioridade and filtro_prioridade != "Todas":
        itens = [
            (i, t) for i, t in itens
            if t["prioridade"].strip() == filtro_prioridade.strip()
        ]

    if filtro_categoria and filtro_categoria != "Todas":
        itens = [(i, t) for i, t in itens if t.get("categoria", "Sem categoria") == filtro_categoria]

    if filtro_data:
        itens = [(i, t) for i, t in itens if t["data"] == filtro_data]

    if somente_favoritas:
        itens = [(i, t) for i, t in itens if t.get("favorita", False)]

    # ORDENAÇÃO
    ordem_prioridade = {"alta": 0, "média": 1, "baixa": 2}
    if ordenacao == "Mais antiga":
        itens.sort(key=lambda item: item[0])
    elif ordenacao == "Prioridade":
        itens.sort(
            key=lambda item: ordem_prioridade.get(
                item[1]["prioridade"].strip().lower(), 99
            )
        )
    elif ordenacao == "Ordem alfabética":
        itens.sort(key=lambda item: item[1]["titulo"].lower())
    else:  # "Mais recente" (padrão)
        itens.sort(key=lambda item: item[0], reverse=True)

    # Tarefas fixadas sempre no topo (sort estável preserva a ordenação já aplicada acima)
    itens.sort(key=lambda item: not item[1].get("fixada", False))

    for posicao, (indice, tarefa) in enumerate(itens):

        linha = posicao // 2
        coluna = posicao % 2

        # Verifica se a tarefa está concluída
        está_concluida = tarefa.get("concluida", False)
        está_favorita = tarefa.get("favorita", False)
        está_fixada = tarefa.get("fixada", False)

        # Se concluída: usa cinza esbranquiçado (#374151). Se pendente: usa escuro (#1F2937)
        cor_fundo = "#374151" if está_concluida else "#1F2937"

        # 1. FRAME DO CARD (com a cor dinâmica e destaque se fixada)
        frame_tarefa = ctk.CTkFrame(
            frame_lista_tarefas,
            corner_radius=15,
            fg_color=cor_fundo,
            border_width=2 if está_fixada else 0,
            border_color="#F59E0B"
        )
        frame_tarefa.grid(
            row=linha, column=coluna, padx=10, pady=10, sticky="nsew"
        )
        frame_tarefa.grid_columnconfigure(0, weight=1)
        frame_tarefa.grid_columnconfigure(1, weight=0)

        # 2. COR DA PRIORIDADE
        if "Alta" in tarefa["prioridade"]:
            cor_prioridade = "#EF4444"
        elif "Média" in tarefa["prioridade"]:
            cor_prioridade = "#F59E0B"
        else:
            cor_prioridade = "#10B981"

        # 3. BOLINHA + TÍTULO
        bolinha_prioridade = ctk.CTkLabel(
            frame_tarefa, text="●", text_color=cor_prioridade, font=("Arial", 20)
        )
        bolinha_prioridade.grid(row=0, column=0, padx=(15, 0), pady=(15, 5), sticky="w")

        titulo_tarefa = ctk.CTkLabel(
            frame_tarefa, text=tarefa["titulo"], font=ctk.CTkFont(size=18, weight="bold")
        )
        titulo_tarefa.grid(row=0, column=0, padx=(38, 15), pady=(15, 5), sticky="w")

        # 3.1 BOTÕES DE FAVORITAR E FIXAR (canto superior direito do card)
        frame_acoes_topo = ctk.CTkFrame(frame_tarefa, fg_color="transparent")
        frame_acoes_topo.grid(row=0, column=1, padx=(0, 15), pady=(10, 5), sticky="ne")

        botao_favoritar = ctk.CTkButton(
            frame_acoes_topo,
            text="⭐" if está_favorita else "☆",
            width=32,
            height=28,
            fg_color="transparent",
            hover_color="#4B5563",
            command=lambda i=indice: alternar_favorita(i)
        )
        botao_favoritar.grid(row=0, column=0, padx=2)

        botao_fixar = ctk.CTkButton(
            frame_acoes_topo,
            text="📌" if está_fixada else "📍",
            width=32,
            height=28,
            fg_color="#F59E0B" if está_fixada else "transparent",
            hover_color="#4B5563",
            command=lambda i=indice: alternar_fixada(i)
        )
        botao_fixar.grid(row=0, column=1, padx=2)

        # 4. DESCRIÇÃO
        descricao_tarefa = ctk.CTkLabel(
            frame_tarefa,
            text=tarefa["descricao"].strip(),
            font=ctk.CTkFont(size=14),
            text_color="#D1D5DB",
            wraplength=350,
            justify="left"
        )
        descricao_tarefa.grid(row=1, column=0, columnspan=2, padx=15, pady=(0, 10), sticky="w")

        # 5. PRIORIDADE E DATA
        prioridade_tarefa = ctk.CTkLabel(
            frame_tarefa, text=f"Prioridade: {tarefa['prioridade']}  |  Data: {tarefa['data']}", font=ctk.CTkFont(size=13), text_color="#9CA3AF"
        )
        prioridade_tarefa.grid(row=2, column=0, columnspan=2, padx=15, pady=(0, 10), sticky="w")

        # 5.1 CATEGORIA E TAGS (Fase 6 e 7) — só aparece se a tarefa tiver algo pra mostrar
        linha_botoes = 3
        info_extra = []
        if tarefa.get("categoria") and tarefa["categoria"] != "Sem categoria":
            info_extra.append(tarefa["categoria"])
        if tarefa.get("tags"):
            tags_formatadas = " ".join(f"#{t.strip()}" for t in tarefa["tags"].split(",") if t.strip())
            if tags_formatadas:
                info_extra.append(tags_formatadas)

        if info_extra:
            extra_tarefa = ctk.CTkLabel(
                frame_tarefa,
                text="   |   ".join(info_extra),
                font=ctk.CTkFont(size=12),
                text_color="#93C5FD",
                wraplength=350,
                justify="left"
            )
            extra_tarefa.grid(row=3, column=0, columnspan=2, padx=15, pady=(0, 10), sticky="w")
            linha_botoes = 4

        # 6. CHECKBOX DE CONCLUÍDA
        check_concluida = ctk.CTkCheckBox(
            frame_tarefa,
            text="Concluída",
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda i=indice: alternar_concluida(i)
        )
        check_concluida.grid(row=linha_botoes, column=0, padx=15, pady=(0, 15), sticky="w")

        # Se a tarefa já estiver concluída, deixa o checkbox marcado
        if está_concluida:
            check_concluida.select()


        #CRIANDO BOTAO EDITAR TAREFA
        botao_editar = ctk.CTkButton(
            frame_tarefa,
            text="✏ Editar",
            width=90,
            height=30,
            command=lambda i=indice: editar_tarefa(i)
        )
        botao_editar.grid(
            row=linha_botoes,
            column=2,
            padx=15,
            pady=(0,15),
            sticky="e"

        )
        # 7. BOTÃO EXCLUIR
        botao_excluir = ctk.CTkButton(
            frame_tarefa,
            text="🗑️ Excluir",
            fg_color="#DC2626",
            hover_color="#991B1B",
            width=90,
            height=30,
            command=lambda i=indice: excluir_tarefa(i)
        )
        botao_excluir.grid(row=linha_botoes, column=1, padx=15, pady=(0, 15), sticky="e")
    
        


#  FUNÇÃO DO BOTAO TAREFAS #
def abrir_tarefas():
    global entry_pesquisa, frame_lista_tarefas
    global optionmenu_prioridade, optionmenu_categoria, optionmenu_ordenacao
    global checkbox_filtro_data_ativo, dateentry_filtro, checkbox_somente_favoritas

    limpar_frame_principal()

    titulo_tarefas = ctk.CTkLabel(
        frame_principal, text="Tarefas", font=("Arial", 28, "bold")
    )
    titulo_tarefas.grid(row=0, column=0, padx=30, pady=30, sticky="w")
#BOTAO NOVA TAREFA #
    botao_nova_tarefa = ctk.CTkButton(
        frame_principal,
        text="➕ Nova Tarefa",
        width=180,
        command=clicar_nova_tarefa,
    )
    botao_nova_tarefa.grid(
        row=1, column=0, padx=30, pady=(10, 20), sticky="w"
    )

    # FRAME COM PESQUISA, FILTROS E ORDENAÇÃO
    frame_filtros = ctk.CTkFrame(frame_principal, fg_color="transparent")
    frame_filtros.grid(row=2, column=0, columnspan=2, padx=30, pady=(0, 15), sticky="w")

    #ENTRY PESQUISA (casa com título e tags)
    entry_pesquisa = ctk.CTkEntry(
        frame_filtros,
        width=220,
        placeholder_text="🔍 Pesquisar por título ou #tag...",
    )
    entry_pesquisa.grid(row=0, column=0, padx=(0, 10), pady=5, sticky="w")
    entry_pesquisa.bind("<KeyRelease>", atualizar_lista)

    # FILTRO POR PRIORIDADE
    optionmenu_prioridade = ctk.CTkOptionMenu(
        frame_filtros,
        values=["Todas", " Alta", " Média", " Baixa"],
        width=120,
        command=atualizar_lista,
    )
    optionmenu_prioridade.set("Todas")
    optionmenu_prioridade.grid(row=0, column=1, padx=10, pady=5, sticky="w")

    # FILTRO POR CATEGORIA
    optionmenu_categoria = ctk.CTkOptionMenu(
        frame_filtros,
        values=["Todas"] + CATEGORIAS,
        width=150,
        command=atualizar_lista,
    )
    optionmenu_categoria.set("Todas")
    optionmenu_categoria.grid(row=0, column=2, padx=10, pady=5, sticky="w")

    # ORDENAÇÃO
    optionmenu_ordenacao = ctk.CTkOptionMenu(
        frame_filtros,
        values=["Mais recente", "Mais antiga", "Prioridade", "Ordem alfabética"],
        width=160,
        command=atualizar_lista,
    )
    optionmenu_ordenacao.set("Mais recente")
    optionmenu_ordenacao.grid(row=0, column=3, padx=10, pady=5, sticky="w")

    # FILTRO POR DATA (segunda linha)
    checkbox_filtro_data_ativo = ctk.CTkCheckBox(
        frame_filtros, text="Filtrar por data", command=atualizar_lista
    )
    checkbox_filtro_data_ativo.grid(row=1, column=0, padx=(0, 10), pady=5, sticky="w")

    dateentry_filtro = DateEntry(
        frame_filtros, width=10, date_pattern="dd/mm/yyyy", justify="center"
    )
    dateentry_filtro.grid(row=1, column=1, padx=10, pady=5, sticky="w")
    dateentry_filtro.bind("<<DateEntrySelected>>", atualizar_lista)

    # SOMENTE FAVORITAS
    checkbox_somente_favoritas = ctk.CTkCheckBox(
        frame_filtros, text="⭐ Somente favoritas", command=atualizar_lista
    )
    checkbox_somente_favoritas.grid(row=1, column=2, padx=10, pady=5, sticky="w")

    # LIMPAR FILTROS
    botao_limpar_filtros = ctk.CTkButton(
        frame_filtros, text="🧹 Limpar", width=90, command=limpar_filtros
    )
    botao_limpar_filtros.grid(row=1, column=3, padx=10, pady=5, sticky="w")

    # FRAME COM SCROLL PARA A LISTA
    frame_lista_tarefas = ctk.CTkScrollableFrame(
        frame_principal, fg_color="transparent", width=900, height=460
    )
    frame_lista_tarefas.grid(
        row=3, column=0, columnspan=2, padx=30, pady=(10, 20), sticky="nsew"
    )
    frame_lista_tarefas.grid_columnconfigure(0, weight=1)
    frame_lista_tarefas.grid_columnconfigure(1, weight=1)

    # RENDERIZAÇÃO DOS CARDS DE TAREFA (já aplicando pesquisa/filtros/ordenação atuais)
    atualizar_lista()


#  FASE 5 — CALENDÁRIO MENSAL  #
def tarefas_do_dia(dia):
    data_str = f"{dia:02d}/{mes_calendario:02d}/{ano_calendario}"
    return [t for t in lista_tarefas if t["data"] == data_str]


def mudar_mes_calendario(delta):
    global ano_calendario, mes_calendario, dia_selecionado
    mes_calendario += delta
    if mes_calendario > 12:
        mes_calendario = 1
        ano_calendario += 1
    elif mes_calendario < 1:
        mes_calendario = 12
        ano_calendario -= 1
    dia_selecionado = None
    abrir_calendario()


def selecionar_dia_calendario(dia):
    global dia_selecionado
    dia_selecionado = dia
    abrir_calendario()


def abrir_calendario():
    limpar_frame_principal()

    titulo_calendario = ctk.CTkLabel(
        frame_principal, text="Calendário", font=("Arial", 28, "bold")
    )
    titulo_calendario.grid(row=0, column=0, padx=30, pady=30, sticky="w")

    frame_conteudo = ctk.CTkFrame(frame_principal, fg_color="transparent")
    frame_conteudo.grid(row=1, column=0, padx=30, pady=(0, 30), sticky="nw")

    # PAINEL DO CALENDÁRIO (esquerda)
    frame_calendario = ctk.CTkFrame(frame_conteudo, corner_radius=15, fg_color="#1F2937")
    frame_calendario.grid(row=0, column=0, padx=(0, 20), sticky="n")

    # NAVEGAÇÃO DE MÊS
    frame_nav = ctk.CTkFrame(frame_calendario, fg_color="transparent")
    frame_nav.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
    frame_nav.grid_columnconfigure(1, weight=1)

    botao_mes_anterior = ctk.CTkButton(
        frame_nav, text="◀", width=40, command=lambda: mudar_mes_calendario(-1)
    )
    botao_mes_anterior.grid(row=0, column=0, sticky="w")

    label_mes = ctk.CTkLabel(
        frame_nav,
        text=f"{MESES_PT[mes_calendario - 1]} {ano_calendario}",
        font=("Arial", 20, "bold"),
        width=260,
    )
    label_mes.grid(row=0, column=1)

    botao_mes_proximo = ctk.CTkButton(
        frame_nav, text="▶", width=40, command=lambda: mudar_mes_calendario(1)
    )
    botao_mes_proximo.grid(row=0, column=2, sticky="e")

    # CABEÇALHO DOS DIAS DA SEMANA
    frame_dias_semana = ctk.CTkFrame(frame_calendario, fg_color="transparent")
    frame_dias_semana.grid(row=1, column=0, padx=20, pady=(0, 5), sticky="ew")
    for coluna, nome_dia in enumerate(DIAS_SEMANA_PT):
        ctk.CTkLabel(
            frame_dias_semana, text=nome_dia, font=("Arial", 13, "bold"),
            text_color="#9CA3AF", width=82
        ).grid(row=0, column=coluna, padx=2)

    # GRADE DE DIAS DO MÊS
    frame_grade = ctk.CTkFrame(frame_calendario, fg_color="transparent")
    frame_grade.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="ew")

    semanas = calendar.monthcalendar(ano_calendario, mes_calendario)
    for linha, semana in enumerate(semanas):
        for coluna, dia in enumerate(semana):
            if dia == 0:
                ctk.CTkFrame(
                    frame_grade, fg_color="transparent", width=82, height=64
                ).grid(row=linha, column=coluna, padx=2, pady=2)
                continue

            tarefas_dia = tarefas_do_dia(dia)
            é_hoje = (
                dia == hoje.day and mes_calendario == hoje.month and ano_calendario == hoje.year
            )
            é_selecionado = (dia == dia_selecionado)

            if é_selecionado:
                cor_fundo = "#2563EB"
            elif é_hoje:
                cor_fundo = "#374151"
            else:
                cor_fundo = "#111827"

            botao_dia = ctk.CTkButton(
                frame_grade,
                text=f"{dia}\n{'●' if tarefas_dia else ''}",
                width=82,
                height=64,
                fg_color=cor_fundo,
                hover_color="#4B5563",
                text_color="#F59E0B" if (tarefas_dia and not é_selecionado) else "#F9FAFB",
                font=("Arial", 14, "bold"),
                command=lambda d=dia: selecionar_dia_calendario(d)
            )
            botao_dia.grid(row=linha, column=coluna, padx=2, pady=2)

    # PAINEL DE TAREFAS DO DIA SELECIONADO (direita)
    frame_painel = ctk.CTkFrame(
        frame_conteudo, corner_radius=15, fg_color="#1F2937", width=340, height=440
    )
    frame_painel.grid(row=0, column=1, sticky="n")
    frame_painel.grid_propagate(False)

    if dia_selecionado is None:
        ctk.CTkLabel(
            frame_painel,
            text="Selecione um dia\npara ver as tarefas",
            font=("Arial", 15), text_color="#9CA3AF", justify="center"
        ).grid(row=0, column=0, padx=20, pady=40)
    else:
        ctk.CTkLabel(
            frame_painel,
            text=f"{dia_selecionado:02d}/{mes_calendario:02d}",
            font=("Arial", 22, "bold")
        ).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        tarefas_dia = tarefas_do_dia(dia_selecionado)
        if not tarefas_dia:
            ctk.CTkLabel(
                frame_painel, text="Nenhuma tarefa para este dia.", text_color="#9CA3AF"
            ).grid(row=1, column=0, padx=20, pady=10, sticky="w")
        else:
            for indice_linha, tarefa in enumerate(tarefas_dia, start=1):
                if "Alta" in tarefa["prioridade"]:
                    cor = "#EF4444"
                elif "Média" in tarefa["prioridade"]:
                    cor = "#F59E0B"
                else:
                    cor = "#10B981"

                texto_item = f"•  {tarefa['titulo']}"
                if tarefa.get("concluida"):
                    texto_item += "  ✓"

                ctk.CTkLabel(
                    frame_painel,
                    text=texto_item,
                    font=("Arial", 14),
                    text_color=cor,
                    anchor="w",
                    justify="left",
                    wraplength=290,
                ).grid(row=indice_linha, column=0, padx=20, pady=4, sticky="w")


#  FUNÇÃO NOVA TAREFA #
def nova_tarefa():
    limpar_frame_principal()

    # Formulário em 2 colunas dentro de um frame com scroll (rede de segurança
    # pra janelas pequenas). Duas colunas mantêm a altura total baixa o
    # suficiente pra "Data de Vencimento" e "Salvar Tarefa" ficarem visíveis
    # sem precisar rolar — e pro popup do calendário ter espaço pra abrir.
    frame_form = ctk.CTkScrollableFrame(frame_principal, fg_color="transparent", width=1040, height=650)
    frame_form.grid(row=0, column=0, padx=20, pady=20, sticky="nw")

    titulo = ctk.CTkLabel(
        frame_form, text="Nova Tarefa", font=("Arial", 28, "bold")
    )
    titulo.grid(row=0, column=0, columnspan=2, padx=20, pady=(10, 20), sticky="w")

    # COLUNA ESQUERDA: Título, Descrição e Tags
    frame_coluna_esquerda = ctk.CTkFrame(frame_form, fg_color="transparent")
    frame_coluna_esquerda.grid(row=1, column=0, padx=(20, 10), pady=0, sticky="n")

    label_titulo = ctk.CTkLabel(
        frame_coluna_esquerda, text="Título da Tarefa", font=("Arial", 18, "bold")
    )
    label_titulo.grid(row=0, column=0, padx=10, pady=(5, 5), sticky="w")

    entry_titulo = ctk.CTkEntry(
        frame_coluna_esquerda, width=460, placeholder_text="Digite o título..."
    )
    entry_titulo.grid(row=1, column=0, padx=10, pady=(0, 20), sticky="w")

    label_descricao = ctk.CTkLabel(
        frame_coluna_esquerda, text="Descrição", font=ctk.CTkFont(size=18, weight="bold")
    )
    label_descricao.grid(row=2, column=0, padx=10, pady=(0, 5), sticky="w")

    textbox_descricao = ctk.CTkTextbox(
        frame_coluna_esquerda, width=460, height=150
    )
    textbox_descricao.grid(row=3, column=0, padx=10, pady=(0, 20), sticky="w")

    label_tags = ctk.CTkLabel(
        frame_coluna_esquerda, text="Tags", font=ctk.CTkFont(size=18, weight="bold")
    )
    label_tags.grid(row=4, column=0, padx=10, pady=(0, 5), sticky="w")

    entry_tags = ctk.CTkEntry(
        frame_coluna_esquerda, width=460, placeholder_text="Ex: Python, Faculdade, Urgente"
    )
    entry_tags.grid(row=5, column=0, padx=10, pady=(0, 20), sticky="w")

    # COLUNA DIREITA: Prioridade, Categoria e Data
    frame_coluna_direita = ctk.CTkFrame(frame_form, fg_color="transparent")
    frame_coluna_direita.grid(row=1, column=1, padx=(10, 20), pady=0, sticky="n")

    label_prioridade = ctk.CTkLabel(
        frame_coluna_direita, text="Prioridade", font=ctk.CTkFont(size=18, weight="bold")
    )
    label_prioridade.grid(row=0, column=0, padx=10, pady=(5, 5), sticky="w")

    combobox_prioridade = ctk.CTkOptionMenu(
        frame_coluna_direita,
        values=[" Alta", " Média", " Baixa"],
        width=260,
    )
    combobox_prioridade.grid(row=1, column=0, padx=10, pady=(0, 20), sticky="w")

    label_categoria = ctk.CTkLabel(
        frame_coluna_direita, text="Categoria", font=ctk.CTkFont(size=18, weight="bold")
    )
    label_categoria.grid(row=2, column=0, padx=10, pady=(0, 5), sticky="w")

    combobox_categoria = ctk.CTkOptionMenu(
        frame_coluna_direita,
        values=CATEGORIAS,
        width=260,
    )
    combobox_categoria.set("Sem categoria")
    combobox_categoria.grid(row=3, column=0, padx=10, pady=(0, 20), sticky="w")

    label_data = ctk.CTkLabel(
        frame_coluna_direita,
        text="Data de Vencimento",
        font=ctk.CTkFont(size=18, weight="bold"),
    )
    label_data.grid(row=4, column=0, padx=10, pady=(0, 5), sticky="w")

    campo_data = DateEntry(
        frame_coluna_direita, width=18, date_pattern="dd/mm/yyyy", justify="center"
    )
    campo_data.grid(row=5, column=0, padx=10, pady=(0, 20), sticky="w")

    # Se estiver editando, preenche TODOS os campos com os dados atuais da tarefa
    if indice_edicao is not None:
        tarefa_atual = lista_tarefas[indice_edicao]
        entry_titulo.insert(0, tarefa_atual["titulo"])
        textbox_descricao.insert("1.0", tarefa_atual["descricao"])
        combobox_prioridade.set(tarefa_atual["prioridade"])
        combobox_categoria.set(tarefa_atual.get("categoria") or "Sem categoria")
        entry_tags.insert(0, tarefa_atual.get("tags", ""))
        campo_data.delete(0, "end")
        campo_data.insert(0, tarefa_atual["data"])

    def salvar_tarefa():
        titulo_txt = entry_titulo.get()
        descricao_txt = textbox_descricao.get("1.0", "end-1c")
        prioridade_txt = combobox_prioridade.get()
        categoria_txt = combobox_categoria.get()
        tags_txt = ", ".join(
            tag.strip().lstrip("#") for tag in entry_tags.get().split(",") if tag.strip()
        )
        data_txt = campo_data.get()

        if not titulo_txt.strip():
            return  # Evita salvar se o título estiver em branco

        tarefa = {
            "titulo": titulo_txt,
            "descricao": descricao_txt,
            "prioridade": prioridade_txt,
            "data": data_txt,
            "categoria": categoria_txt,
            "tags": tags_txt,
            "concluida": False,
            "favorita": False,
            "fixada": False
        }
        global indice_edicao
        if indice_edicao is None:
          inserir_tarefa_bd(tarefa)
        else:
          # Preserva o status de conclusão, favorito e fixado ao editar
          tarefa_anterior = lista_tarefas[indice_edicao]
          tarefa["concluida"] = tarefa_anterior.get("concluida", False)
          tarefa["favorita"] = tarefa_anterior.get("favorita", False)
          tarefa["fixada"] = tarefa_anterior.get("fixada", False)
          atualizar_tarefa_bd(tarefa_anterior["id"], tarefa)

        recarregar_tarefas()
        indice_edicao = None  # Reseta o modo de edição
        abrir_tarefas()  # Atualiza e exibe a lista com o novo card
    # BOTÃO SALVAR TAREFA

    botao_salvar = ctk.CTkButton(
        frame_form,
        text="💾 Salvar Tarefa",
        width=200,
        command=salvar_tarefa,
    )
    botao_salvar.grid(row=2, column=0, columnspan=2, padx=20, pady=(10, 20), sticky="w")
   
    

#  FASE 8 — NOTIFICAÇÕES  #
def parse_data_tarefa(data_str):
    try:
        dia, mes, ano = data_str.split("/")
        return date(int(ano), int(mes), int(dia))
    except (ValueError, AttributeError):
        return None


def tarefas_vencidas():
    hoje_atual = date.today()
    return [
        t for t in lista_tarefas
        if not t.get("concluida") and (d := parse_data_tarefa(t["data"])) and d < hoje_atual
    ]


def tarefas_de_amanha():
    amanha = date.today() + timedelta(days=1)
    return [
        t for t in lista_tarefas
        if not t.get("concluida") and parse_data_tarefa(t["data"]) == amanha
    ]


def formatar_lista_notificacao(tarefas, limite=6):
    linhas = [f"•  {t['titulo']}" for t in tarefas[:limite]]
    if len(tarefas) > limite:
        linhas.append(f"…  e mais {len(tarefas) - limite}")
    return linhas


def abrir_notificacoes():
    vencidas = tarefas_vencidas()
    de_amanha = tarefas_de_amanha()

    janela_aviso = ctk.CTkToplevel(janela)
    janela_aviso.title("Notificações")
    janela_aviso.geometry("420x420")
    janela_aviso.configure(fg_color="#111827")
    janela_aviso.transient(janela)
    janela_aviso.grab_set()

    ctk.CTkLabel(
        janela_aviso, text="🔔 Notificações", font=("Arial", 20, "bold")
    ).pack(padx=20, pady=(20, 10), anchor="w")

    if not vencidas and not de_amanha:
        ctk.CTkLabel(
            janela_aviso,
            text="Nenhuma pendência no momento.\nTudo em dia! ✅",
            font=("Arial", 14), text_color="#9CA3AF", justify="left"
        ).pack(padx=20, pady=10, anchor="w")
    else:
        frame_scroll = ctk.CTkScrollableFrame(janela_aviso, fg_color="transparent", width=380, height=280)
        frame_scroll.pack(padx=20, pady=(0, 10), fill="both", expand=True)

        if vencidas:
            plural = "s" if len(vencidas) != 1 else ""
            ctk.CTkLabel(
                frame_scroll,
                text=f"⚠️ Você possui {len(vencidas)} tarefa{plural} vencida{plural}.",
                font=("Arial", 15, "bold"), text_color="#EF4444", anchor="w", justify="left"
            ).pack(anchor="w", pady=(0, 5))
            for linha in formatar_lista_notificacao(vencidas):
                ctk.CTkLabel(frame_scroll, text=linha, text_color="#FCA5A5", anchor="w").pack(anchor="w", padx=10, pady=1)
            ctk.CTkLabel(frame_scroll, text="", height=10).pack()

        if de_amanha:
            ctk.CTkLabel(
                frame_scroll,
                text="📅 Amanhã vence:",
                font=("Arial", 15, "bold"), text_color="#F59E0B", anchor="w", justify="left"
            ).pack(anchor="w", pady=(0, 5))
            for linha in formatar_lista_notificacao(de_amanha):
                ctk.CTkLabel(frame_scroll, text=linha, text_color="#FCD34D", anchor="w").pack(anchor="w", padx=10, pady=1)

    ctk.CTkButton(janela_aviso, text="Fechar", command=janela_aviso.destroy).pack(pady=(5, 15))


def verificar_notificacoes_iniciais():
    if tarefas_vencidas() or tarefas_de_amanha():
        abrir_notificacoes()


#  FASE 9 — EXPORTAÇÃO  #
def exportar_excel():
    if not lista_tarefas:
        return False, "Não há tarefas para exportar."
    try:
        os.makedirs(CAMINHO_EXPORTS, exist_ok=True)
        nome_arquivo = f"tarefas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho = os.path.join(CAMINHO_EXPORTS, nome_arquivo)

        workbook = Workbook()
        planilha = workbook.active
        planilha.title = "Tarefas"

        cabecalho = ["Título", "Descrição", "Prioridade", "Categoria", "Tags", "Data", "Concluída"]
        planilha.append(cabecalho)
        for celula in planilha[1]:
            celula.font = Font(bold=True, color="FFFFFF")
            celula.fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
            celula.alignment = Alignment(horizontal="center")

        for tarefa in lista_tarefas:
            planilha.append([
                tarefa["titulo"],
                tarefa["descricao"],
                tarefa["prioridade"].strip(),
                tarefa["categoria"],
                tarefa["tags"],
                tarefa["data"],
                "Sim" if tarefa["concluida"] else "Não",
            ])

        for coluna, largura in zip(planilha.columns, [30, 45, 12, 18, 25, 12, 10]):
            planilha.column_dimensions[coluna[0].column_letter].width = largura

        workbook.save(caminho)
        try:
            os.startfile(caminho)
        except OSError:
            pass
        return True, f"Exportado: {nome_arquivo}"
    except Exception as erro:
        return False, f"Erro ao exportar: {erro}"


def gerar_pdf_relatorio():
    """Gera o PDF do relatório e devolve (sucesso, mensagem, caminho)."""
    if not lista_tarefas:
        return False, "Não há tarefas para exportar.", None
    try:
        os.makedirs(CAMINHO_EXPORTS, exist_ok=True)
        nome_arquivo = f"tarefas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        caminho = os.path.join(CAMINHO_EXPORTS, nome_arquivo)

        linhas = [
            [t["titulo"], t["prioridade"].strip(), t["categoria"], t["data"], "Sim" if t["concluida"] else "Não"]
            for t in lista_tarefas
        ]
        cabecalho = ["Título", "Prioridade", "Categoria", "Data", "Concluída"]

        altura_fig = max(2.5, 0.45 * (len(linhas) + 2))
        figura = Figure(figsize=(11, altura_fig))
        eixo = figura.add_subplot(111)
        eixo.axis("off")
        eixo.set_title(
            f"Relatório de Tarefas — GPriority  ({datetime.now().strftime('%d/%m/%Y %H:%M')})",
            fontsize=13, fontweight="bold"
        )

        tabela = eixo.table(cellText=linhas, colLabels=cabecalho, loc="center", cellLoc="left")
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(9)
        tabela.scale(1, 1.5)

        figura.tight_layout()
        figura.savefig(caminho, format="pdf")
        return True, f"Exportado: {nome_arquivo}", caminho
    except Exception as erro:
        return False, f"Erro ao exportar: {erro}", None


def exportar_pdf():
    sucesso, mensagem, caminho = gerar_pdf_relatorio()
    if sucesso and caminho:
        try:
            os.startfile(caminho)
        except OSError:
            pass
    return sucesso, mensagem


def imprimir_relatorio():
    sucesso, mensagem, caminho = gerar_pdf_relatorio()
    if not sucesso or not caminho:
        return sucesso, mensagem
    try:
        os.startfile(caminho, "print")
        return True, "Enviado para impressão."
    except OSError:
        try:
            os.startfile(caminho)
            return True, f"Não achei uma impressora padrão — abrindo {os.path.basename(caminho)} pra imprimir manualmente."
        except OSError as erro:
            return False, f"Erro ao imprimir: {erro}"


def abrir_relatorios():
    limpar_frame_principal()

    titulo_relatorios = ctk.CTkLabel(
        frame_principal, text="Relatórios", font=("Arial", 28, "bold")
    )
    titulo_relatorios.grid(row=0, column=0, padx=30, pady=30, sticky="w")

    total = len(lista_tarefas)
    label_info = ctk.CTkLabel(
        frame_principal,
        text=f"Exporte o relatório completo com as {total} tarefa(s) cadastradas.",
        font=("Arial", 15), text_color="#9CA3AF"
    )
    label_info.grid(row=1, column=0, padx=30, pady=(0, 25), sticky="w")

    frame_botoes = ctk.CTkFrame(frame_principal, fg_color="transparent")
    frame_botoes.grid(row=2, column=0, padx=30, pady=0, sticky="w")

    label_status = ctk.CTkLabel(frame_principal, text="", font=("Arial", 13))
    label_status.grid(row=3, column=0, padx=30, pady=(20, 0), sticky="w")

    def executar(funcao_exportacao):
        sucesso, mensagem = funcao_exportacao()
        label_status.configure(text=mensagem, text_color="#10B981" if sucesso else "#F59E0B")

    botao_excel = ctk.CTkButton(
        frame_botoes, text="⬇ Exportar Excel", width=220, height=45,
        fg_color="#15803D", hover_color="#166534",
        command=lambda: executar(exportar_excel)
    )
    botao_excel.grid(row=0, column=0, padx=(0, 15), pady=10)

    botao_pdf = ctk.CTkButton(
        frame_botoes, text="⬇ Exportar PDF", width=220, height=45,
        fg_color="#B91C1C", hover_color="#7F1D1D",
        command=lambda: executar(exportar_pdf)
    )
    botao_pdf.grid(row=0, column=1, padx=15, pady=10)

    botao_imprimir = ctk.CTkButton(
        frame_botoes, text="🖨 Imprimir", width=220, height=45,
        fg_color="#1D4ED8", hover_color="#1E3A8A",
        command=lambda: executar(imprimir_relatorio)
    )
    botao_imprimir.grid(row=0, column=2, padx=(15, 0), pady=10)


def clicar_config():
    print("Painel de configurações")


#  ESTRUTURA DE ESTRUTURAÇÃO DE FRAMES (TOPO E MENU) #
frame_topo = ctk.CTkFrame(janela, height=70, fg_color="#1E293B")
frame_topo.grid(row=0, column=0, columnspan=2, sticky="ew")
frame_topo.grid_columnconfigure(1, weight=1)

frame_menu = ctk.CTkFrame(janela, width=250, fg_color="#0F172A")
frame_menu.grid(row=1, column=0, sticky="nsew")
frame_menu.grid_columnconfigure(0, weight=1)

frame_principal = ctk.CTkFrame(janela, fg_color="#111827")
frame_principal.grid(row=1, column=1, sticky="nsew")

# TOPBAR
titulo = ctk.CTkLabel(
    frame_topo, text="GPriority", font=("Arial", 24, "bold")
)
titulo.grid(row=0, column=0, padx=20, pady=10, sticky="w")

frame_topo_direita = ctk.CTkFrame(frame_topo, fg_color="transparent")
frame_topo_direita.grid(row=0, column=1, padx=20, pady=10, sticky="e")

usuario = ctk.CTkLabel(frame_topo_direita, text="Usuário: Walace", font=("Arial", 16))
usuario.grid(row=0, column=0, padx=(0, 15))

botao_notificacoes = ctk.CTkButton(
    frame_topo_direita, text="🔔", width=40, command=abrir_notificacoes
)
botao_notificacoes.grid(row=0, column=1)

# BOTOES MENU LATERAL
botao_dashboard = ctk.CTkButton(
    frame_menu, text="🏠 Dashboard", command=clicar_dashboard
)
botao_dashboard.grid(row=0, column=0, padx=20, pady=10, sticky="ew")

botao_tarefas = ctk.CTkButton(
    frame_menu, text="📋 Tarefas", command=abrir_tarefas
)
botao_tarefas.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

botao_calendario = ctk.CTkButton(
    frame_menu, text="📅 Calendário", command=abrir_calendario
)
botao_calendario.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

botao_relatorio = ctk.CTkButton(
    frame_menu, text="📊 Relatórios", command=abrir_relatorios
)
botao_relatorio.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

botao_config = ctk.CTkButton(
    frame_menu, text="⚙ Configurações", command=clicar_config
)
botao_config.grid(row=4, column=0, padx=20, pady=10, sticky="ew")


# FECHA A CONEXÃO COM O BANCO AO FECHAR A JANELA
def ao_fechar_janela():
    conexao_bd.close()
    janela.destroy()


janela.protocol("WM_DELETE_WINDOW", ao_fechar_janela)

# Carrega as tarefas salvas no banco SQLite e inicializa direto no Dashboard
criar_tabela()
recarregar_tarefas()
mostrar_dashboard()

# Avisa sobre tarefas vencidas/vencendo amanhã pouco depois de abrir a janela
janela.after(400, verificar_notificacoes_iniciais)

janela.mainloop()