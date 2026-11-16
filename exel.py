"""Aplicacao de gestao de estoque com MySQL, Tkinter e Google Sheets."""

import os
import sys
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk

import gspread
import matplotlib
import mysql.connector
import pandas as pd
from google.oauth2.service_account import Credentials
matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


DB_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "user": os.getenv("DB_USER", "root"),
    "password": os.getenv("DB_PASSWORD", ""),
    "database": os.getenv("DB_NAME", "estoque_gestao"),
}
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "")
GOOGLE_SHEETS_ID = os.getenv("GOOGLE_SHEETS_ID", "")
GOOGLE_SHEETS_NAME = "relatorio_estoque"
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PRODUCTS_SQL = """
    SELECT p.id, p.nome, p.categoria, p.preco_unitario, p.estoque_inicial,
           p.estoque_inicial + COALESCE(SUM(
               CASE WHEN m.tipo = 'Entrada' THEN m.quantidade ELSE -m.quantidade END
           ), 0) AS estoque_atual
    FROM produtos AS p
    LEFT JOIN movimentacoes AS m ON m.produto_id = p.id
    GROUP BY p.id, p.nome, p.categoria, p.preco_unitario, p.estoque_inicial
    ORDER BY p.id
"""
MOVEMENTS_SQL = """
    SELECT m.id, m.data_movimento, p.nome AS produto, p.categoria, m.tipo,
           m.quantidade, m.valor_unitario, m.observacao, m.produto_id
    FROM movimentacoes AS m
    JOIN produtos AS p ON p.id = m.produto_id
    ORDER BY m.data_movimento DESC, m.id DESC
"""


def connect_db():
    return mysql.connector.connect(**DB_CONFIG)


def query_frame(cursor, query, params=None):
    cursor.execute(query, params or ())
    columns = [column[0] for column in cursor.description]
    return pd.DataFrame(cursor.fetchall(), columns=columns)


def current_stock(cursor, product_id, excluding_movement_id=None):
    exclusion = ""
    params = [product_id]
    if excluding_movement_id is not None:
        exclusion = " AND id <> %s"
        params = [excluding_movement_id, product_id]
    cursor.execute(
        f"""
        SELECT estoque_inicial + COALESCE((
            SELECT SUM(CASE WHEN tipo = 'Entrada' THEN quantidade ELSE -quantidade END)
            FROM movimentacoes
            WHERE produto_id = produtos.id{exclusion}
        ), 0) AS saldo
        FROM produtos
        WHERE id = %s
        """,
        params,
    )
    row = cursor.fetchone()
    if row is None:
        raise ValueError("O produto selecionado nao existe.")
    return int(row[0])


def read_frames():
    connection = connect_db()
    try:
        cursor = connection.cursor()
        products = query_frame(cursor, PRODUCTS_SQL)
        movements = query_frame(cursor, MOVEMENTS_SQL)
        cursor.close()
    finally:
        connection.close()
    return products, movements


def export_frames_to_sheets(consolidated, summary):
    if not GOOGLE_CREDENTIALS_FILE:
        raise ValueError(
            "Configure a variavel GOOGLE_CREDENTIALS_FILE com o caminho "
            "do JSON da conta de servico do Google."
        )
    credentials_path = Path(GOOGLE_CREDENTIALS_FILE)
    if not credentials_path.is_file():
        raise FileNotFoundError(f"Arquivo de credenciais nao encontrado: {credentials_path}")

    credentials = Credentials.from_service_account_file(
        str(credentials_path), scopes=GOOGLE_SCOPES
    )
    client = gspread.authorize(credentials)
    if GOOGLE_SHEETS_ID:
        spreadsheet = client.open_by_key(GOOGLE_SHEETS_ID)
    else:
        try:
            spreadsheet = client.open(GOOGLE_SHEETS_NAME)
        except gspread.SpreadsheetNotFound:
            spreadsheet = client.create(GOOGLE_SHEETS_NAME)

    for title, frame in (
        ("Dados Consolidados", consolidated),
        ("Resumo Gerencial", summary),
    ):
        try:
            worksheet = spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            if title == "Dados Consolidados" and len(spreadsheet.worksheets()) == 1:
                worksheet = spreadsheet.sheet1
                worksheet.update_title(title)
            else:
                worksheet = spreadsheet.add_worksheet(
                    title=title, rows=max(len(frame) + 1, 100), cols=max(len(frame.columns), 10)
                )
        worksheet.clear()
        values = [frame.columns.tolist()] + frame_to_sheet_values(frame)
        worksheet.update(
            range_name="A1",
            values=values,
            value_input_option="USER_ENTERED",
        )
        format_worksheet(worksheet, frame)
    return spreadsheet.url


def format_worksheet(worksheet, frame):
    row_count = max(len(frame) + 1, 2)
    column_count = max(len(frame.columns), 1)
    worksheet.resize(rows=row_count, cols=column_count)
    last_cell = gspread.utils.rowcol_to_a1(row_count, column_count)
    last_header = gspread.utils.rowcol_to_a1(1, column_count)

    worksheet.format(
        f"A1:{last_cell}",
        {
            "textFormat": {
                "fontFamily": "Arial",
                "fontSize": 10,
                "foregroundColor": {"red": 0.12, "green": 0.18, "blue": 0.15},
            },
            "verticalAlignment": "MIDDLE",
        },
    )
    worksheet.format(
        f"A1:{last_header}",
        {
            "backgroundColor": {"red": 0.09, "green": 0.24, "blue": 0.19},
            "textFormat": {
                "fontFamily": "Arial",
                "fontSize": 10,
                "bold": True,
                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
            },
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "wrapStrategy": "WRAP",
        },
    )
    worksheet.format(
        f"A2:{last_cell}",
        {
            "verticalAlignment": "MIDDLE",
            "borders": {
                "bottom": {
                    "style": "SOLID",
                    "color": {"red": 0.88, "green": 0.91, "blue": 0.89},
                }
            },
        },
    )
    worksheet.freeze(rows=1)
    worksheet.columns_auto_resize(0, column_count)
    metadata = worksheet.spreadsheet.fetch_sheet_metadata(
        params={"fields": "sheets(properties(sheetId),bandedRanges(bandedRangeId))"}
    )
    sheet_metadata = next(
        sheet for sheet in metadata["sheets"]
        if sheet["properties"]["sheetId"] == worksheet.id
    )
    requests = [
        {
            "deleteBanding": {"bandedRangeId": band["bandedRangeId"]}
        }
        for band in sheet_metadata.get("bandedRanges", [])
    ]
    requests.extend(
        [
            {
                "addBanding": {
                    "bandedRange": {
                        "range": {
                            "sheetId": worksheet.id,
                            "startRowIndex": 0,
                            "endRowIndex": row_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": column_count,
                        },
                        "rowProperties": {
                            "headerColor": {"red": 0.09, "green": 0.24, "blue": 0.19},
                            "firstBandColor": {"red": 1, "green": 1, "blue": 1},
                            "secondBandColor": {"red": 0.94, "green": 0.96, "blue": 0.94},
                        },
                    }
                }
            },
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": worksheet.id,
                        "gridProperties": {"hideGridlines": True},
                    },
                    "fields": "gridProperties.hideGridlines",
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "ROWS",
                        "startIndex": 0,
                        "endIndex": 1,
                    },
                    "properties": {"pixelSize": 42},
                    "fields": "pixelSize",
                }
            },
        ]
    )
    worksheet.batch_update({"requests": requests})

    for column_index, column_name in enumerate(frame.columns, start=1):
        column_letter = gspread.utils.rowcol_to_a1(1, column_index)[:-1]
        column_range = f"{column_letter}2:{column_letter}{row_count}"
        normalized_name = str(column_name).lower()
        if any(word in normalized_name for word in ("preco", "valor", "total")):
            worksheet.format(
                column_range,
                {"numberFormat": {"type": "CURRENCY", "pattern": '"R$" #,##0.00'}},
            )
        elif "data" in normalized_name:
            worksheet.format(
                column_range,
                {"numberFormat": {"type": "DATE", "pattern": "dd/mm/yyyy"}},
            )
        elif pd.api.types.is_numeric_dtype(frame[column_name]):
            worksheet.format(
                column_range,
                {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.##"}},
            )


def frame_to_sheet_values(frame):
    values = []
    for row in frame.itertuples(index=False, name=None):
        converted = []
        for value in row:
            if pd.isna(value):
                converted.append("")
            elif isinstance(value, (date, datetime)):
                converted.append(value.isoformat())
            elif isinstance(value, Decimal):
                converted.append(float(value))
            elif hasattr(value, "item"):
                converted.append(value.item())
            else:
                converted.append(value)
        values.append(converted)
    return values


class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Estoque | Gestao inteligente")
        self.root.geometry("1240x800")
        self.root.minsize(1000, 680)
        self.colors = {
            "background": "#F2F5F3",
            "surface": "#FFFFFF",
            "ink": "#172923",
            "muted": "#73827B",
            "green": "#24634D",
            "green_dark": "#173D30",
            "green_light": "#E7F1EB",
            "border": "#E0E7E2",
            "red": "#B5473B",
        }
        self.products = pd.DataFrame()
        self.movements = pd.DataFrame()
        self._configure_theme()
        self._build_ui()
        self.refresh_data()

    def _configure_theme(self):
        self.root.configure(bg=self.colors["background"])
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            "TFrame",
            background=self.colors["background"],
        )
        style.configure(
            "TLabel",
            background=self.colors["background"],
            foreground=self.colors["ink"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "TButton",
            padding=(12, 8),
            background=self.colors["surface"],
            foreground=self.colors["ink"],
            bordercolor=self.colors["border"],
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "TButton",
            background=[("active", self.colors["green_light"]), ("pressed", "#D9E8DE")],
            foreground=[("disabled", self.colors["muted"])],
        )
        style.configure(
            "Accent.TButton",
            padding=(14, 9),
            background=self.colors["green"],
            foreground="#FFFFFF",
            bordercolor=self.colors["green"],
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#1D513E"), ("pressed", self.colors["green_dark"])],
            foreground=[("active", "#FFFFFF")],
        )
        style.configure(
            "Danger.TButton",
            padding=(12, 8),
            background="#FBEDEC",
            foreground=self.colors["red"],
            bordercolor="#F1D5D1",
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#F5DEDA"), ("pressed", "#EED1CC")],
        )
        style.configure(
            "TNotebook",
            background=self.colors["background"],
            borderwidth=0,
            tabmargins=(0, 0, 0, 0),
        )
        style.configure(
            "TNotebook.Tab",
            padding=(18, 11),
            background="#E7ECE8",
            foreground=self.colors["muted"],
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", self.colors["surface"]), ("active", "#EDF3EF")],
            foreground=[("selected", self.colors["green_dark"]), ("active", self.colors["ink"])],
        )
        style.configure(
            "Treeview",
            background=self.colors["surface"],
            fieldbackground=self.colors["surface"],
            foreground=self.colors["ink"],
            rowheight=36,
            borderwidth=0,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#EAF0EC",
            foreground=self.colors["muted"],
            padding=(10, 11),
            relief="flat",
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", self.colors["green"])],
            foreground=[("selected", "#FFFFFF")],
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#DFE9E2")],
        )
        style.configure(
            "Vertical.TScrollbar",
            background="#DCE5DF",
            troughcolor=self.colors["surface"],
            bordercolor=self.colors["surface"],
            arrowsize=12,
        )

    def _build_ui(self):
        header = tk.Frame(self.root, bg=self.colors["green_dark"], padx=28, pady=20)
        header.pack(fill="x")
        brand = tk.Frame(header, bg=self.colors["green_dark"])
        brand.pack(side="left", fill="y")
        tk.Label(
            brand,
            text="ESTOQUE",
            bg=self.colors["green_dark"],
            fg="#FFFFFF",
            font=("Segoe UI", 19, "bold"),
        ).pack(anchor="w")
        tk.Label(
            brand,
            text="Gestao simples. Decisoes melhores.",
            bg=self.colors["green_dark"],
            fg="#BBD0C5",
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(2, 0))

        header_actions = tk.Frame(header, bg=self.colors["green_dark"])
        header_actions.pack(side="right", pady=5)
        ttk.Button(
            header_actions,
            text="Atualizar dados",
            command=self.refresh_data,
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            header_actions,
            text="Exportar para Sheets",
            style="Accent.TButton",
            command=self.export_to_sheets,
        ).pack(side="left")

        self.metrics = tk.Frame(self.root, bg=self.colors["background"], padx=24, pady=18)
        self.metrics.pack(fill="x")
        self.metric_values = {}
        self._create_metric_card(self.metrics, "PRODUTOS CADASTRADOS", "0", "products")
        self._create_metric_card(self.metrics, "UNIDADES EM ESTOQUE", "0", "stock")
        self._create_metric_card(self.metrics, "MOVIMENTACOES", "0", "movements")

        self.tabs = ttk.Notebook(self.root, padding=0)
        self.tabs.pack(fill="both", expand=True, padx=24, pady=(0, 22))

        self.products_tab = ttk.Frame(self.tabs, padding=(16, 16))
        self.movements_tab = ttk.Frame(self.tabs, padding=(16, 16))
        self.analytics_tab = ttk.Frame(self.tabs, padding=(16, 16))
        self.tabs.add(self.products_tab, text="Produtos")
        self.tabs.add(self.movements_tab, text="Movimentacoes")
        self.tabs.add(self.analytics_tab, text="Graficos")
        self._build_products_tab()
        self._build_movements_tab()
        self._build_analytics_tab()

    def _create_metric_card(self, parent, title, value, key):
        card = tk.Frame(
            parent,
            bg=self.colors["surface"],
            highlightbackground=self.colors["border"],
            highlightthickness=1,
            padx=18,
            pady=13,
        )
        card.pack(side="left", fill="both", expand=True, padx=(0, 12))
        tk.Label(
            card,
            text=title,
            bg=self.colors["surface"],
            fg=self.colors["muted"],
            font=("Segoe UI", 8, "bold"),
        ).pack(anchor="w")
        value_label = tk.Label(
            card,
            text=value,
            bg=self.colors["surface"],
            fg=self.colors["green_dark"],
            font=("Segoe UI", 21, "bold"),
        )
        value_label.pack(anchor="w", pady=(4, 0))
        self.metric_values[key] = value_label

    def _build_products_tab(self):
        actions = ttk.Frame(self.products_tab, padding=(0, 0, 0, 12))
        actions.pack(fill="x")
        ttk.Label(actions, text="Catalogo de produtos", font=("Segoe UI", 14, "bold")).pack(side="left", padx=(0, 16))
        ttk.Button(
            actions,
            text="+ Novo produto",
            style="Accent.TButton",
            command=lambda: self.open_form("product"),
        ).pack(side="right", padx=(8, 0))
        ttk.Button(
            actions,
            text="Editar",
            command=lambda: self.edit_selected("product"),
        ).pack(side="right")
        ttk.Button(
            actions,
            text="Excluir",
            style="Danger.TButton",
            command=lambda: self.delete_selected("product"),
        ).pack(side="right", padx=8)
        columns = ("id", "nome", "categoria", "preco", "inicial", "atual")
        self.products_tree = self._make_tree(self.products_tab, columns, {
            "id": ("ID", 55), "nome": ("Produto", 220), "categoria": ("Categoria", 150),
            "preco": ("Preco unitario", 130), "inicial": ("Estoque inicial", 120),
            "atual": ("Estoque atual", 120),
        })

    def _build_movements_tab(self):
        actions = ttk.Frame(self.movements_tab, padding=(0, 0, 0, 12))
        actions.pack(fill="x")
        ttk.Label(actions, text="Historico de movimentacoes", font=("Segoe UI", 14, "bold")).pack(side="left", padx=(0, 16))
        ttk.Button(
            actions,
            text="+ Nova movimentacao",
            style="Accent.TButton",
            command=lambda: self.open_form("movement"),
        ).pack(side="right", padx=(8, 0))
        ttk.Button(
            actions,
            text="Editar",
            command=lambda: self.edit_selected("movement"),
        ).pack(side="right")
        ttk.Button(
            actions,
            text="Excluir",
            style="Danger.TButton",
            command=lambda: self.delete_selected("movement"),
        ).pack(side="right", padx=8)
        columns = ("id", "data", "produto", "categoria", "tipo", "quantidade", "valor", "observacao")
        self.movements_tree = self._make_tree(self.movements_tab, columns, {
            "id": ("ID", 45), "data": ("Data", 100), "produto": ("Produto", 170),
            "categoria": ("Categoria", 110), "tipo": ("Tipo", 75), "quantidade": ("Quantidade", 85),
            "valor": ("Valor unitario", 110), "observacao": ("Observacao", 200),
        })

    def _build_analytics_tab(self):
        actions = ttk.Frame(self.analytics_tab, padding=(0, 0, 0, 12))
        actions.pack(fill="x")
        ttk.Label(actions, text="Visao analitica", font=("Segoe UI", 14, "bold")).pack(side="left")
        ttk.Button(
            actions,
            text="Atualizar graficos",
            style="Accent.TButton",
            command=self.refresh_charts,
        ).pack(side="right")
        self.chart_frame = ttk.Frame(self.analytics_tab)
        self.chart_frame.pack(fill="both", expand=True)
        self.chart_canvas = None

    @staticmethod
    def _make_tree(parent, columns, headings):
        wrapper = ttk.Frame(
            parent,
            padding=1,
            style="TFrame",
        )
        wrapper.pack(fill="both", expand=True)
        tree = ttk.Treeview(wrapper, columns=columns, show="headings", selectmode="browse")
        yscroll = ttk.Scrollbar(wrapper, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(wrapper, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        wrapper.rowconfigure(0, weight=1)
        wrapper.columnconfigure(0, weight=1)
        for key in columns:
            label, width = headings[key]
            tree.heading(key, text=label)
            tree.column(key, width=width, minwidth=45, anchor="w")
        tree.tag_configure("even", background="#FFFFFF")
        tree.tag_configure("odd", background="#F6F8F6")
        return tree

    def refresh_data(self):
        try:
            self.products, self.movements = read_frames()
            self._fill_products()
            self._fill_movements()
            self._update_metrics()
        except (mysql.connector.Error, ValueError) as error:
            messagebox.showerror("Erro ao consultar o banco", str(error), parent=self.root)

    def _update_metrics(self):
        product_count = len(self.products)
        stock_count = int(self.products["estoque_atual"].sum()) if product_count else 0
        movement_count = len(self.movements)
        self.metric_values["products"].configure(text=f"{product_count:,}".replace(",", "."))
        self.metric_values["stock"].configure(text=f"{stock_count:,}".replace(",", "."))
        self.metric_values["movements"].configure(text=f"{movement_count:,}".replace(",", "."))

    def _fill_products(self):
        self.products_tree.delete(*self.products_tree.get_children())
        for index, row in enumerate(self.products.itertuples(index=False)):
            self.products_tree.insert("", "end", iid=str(row.id), tags=("even" if index % 2 == 0 else "odd",), values=(
                row.id, row.nome, row.categoria, f"R$ {float(row.preco_unitario):.2f}",
                row.estoque_inicial, row.estoque_atual,
            ))

    def _fill_movements(self):
        self.movements_tree.delete(*self.movements_tree.get_children())
        for index, row in enumerate(self.movements.itertuples(index=False)):
            self.movements_tree.insert("", "end", iid=str(row.id), tags=("even" if index % 2 == 0 else "odd",), values=(
                row.id, row.data_movimento, row.produto, row.categoria, row.tipo,
                row.quantidade, f"R$ {float(row.valor_unitario):.2f}", row.observacao,
            ))

    def selected_id(self, tree, label):
        selection = tree.selection()
        if not selection:
            messagebox.showinfo("Selecao necessaria", f"Selecione {label}.", parent=self.root)
            return None
        return int(selection[0])

    def edit_selected(self, kind):
        tree = self.products_tree if kind == "product" else self.movements_tree
        record_id = self.selected_id(tree, "um registro para editar")
        if record_id is None:
            return
        frame = self.products if kind == "product" else self.movements
        record = frame.loc[frame["id"] == record_id].iloc[0].to_dict()
        self.open_form(kind, record)

    def open_form(self, kind, record=None):
        is_product = kind == "product"
        if not is_product and self.products.empty:
            messagebox.showinfo("Sem produtos", "Cadastre um produto antes de registrar movimentacoes.", parent=self.root)
            return
        window = tk.Toplevel(self.root)
        window.title(("Editar" if record else "Cadastrar") + (" produto" if is_product else " movimentacao"))
        window.transient(self.root)
        window.grab_set()
        form = ttk.Frame(window, padding=16)
        form.pack(fill="both", expand=True)
        widgets = {}
        if is_product:
            fields = [
                ("nome", "Nome", "entry"),
                ("categoria", "Categoria", "entry"),
                ("preco_unitario", "Preco unitario", "entry"),
                ("estoque_inicial", "Estoque inicial", "entry"),
            ]
        else:
            product_options = [
                (int(row.id), f"{int(row.id)} - {row.nome}")
                for row in self.products.itertuples(index=False)
            ]
            fields = [
                ("produto_id", "Produto", "product"),
                ("tipo", "Tipo", "type"),
                ("quantidade", "Quantidade", "entry"),
                ("data_movimento", "Data (AAAA-MM-DD)", "entry"),
                ("valor_unitario", "Valor unitario", "entry"),
                ("observacao", "Observacao", "entry"),
            ]
        for row_index, (key, label, widget_type) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=row_index, column=0, sticky="w", pady=5)
            initial = record.get(key, "") if record else ""
            if widget_type == "product":
                widget = ttk.Combobox(form, state="readonly", values=[name for _, name in product_options], width=34)
                chosen = next((name for product_id, name in product_options if product_id == initial), "")
                widget.set(chosen)
            elif widget_type == "type":
                widget = ttk.Combobox(form, state="readonly", values=("Entrada", "Saida"), width=34)
                widget.set(initial or "Entrada")
            else:
                if isinstance(initial, (date, datetime)):
                    initial = initial.isoformat()
                widget = ttk.Entry(form, width=37)
                widget.insert(0, str(initial))
            widget.grid(row=row_index, column=1, sticky="ew", pady=5, padx=(10, 0))
            widgets[key] = widget
        form.columnconfigure(1, weight=1)

        def save():
            try:
                values = {}
                for key, widget in widgets.items():
                    if key == "produto_id":
                        selection = widget.get().split(" - ", 1)
                        if not selection or not selection[0]:
                            raise ValueError("Selecione um produto.")
                        values[key] = int(selection[0])
                    else:
                        values[key] = widget.get().strip()
                if is_product:
                    values["preco_unitario"] = self.parse_money(values["preco_unitario"])
                    values["estoque_inicial"] = self.parse_int(values["estoque_inicial"], "Estoque inicial", minimum=0)
                    if not values["nome"] or len(values["nome"]) > 100:
                        raise ValueError("Informe o nome do produto (ate 100 caracteres).")
                    if not values["categoria"] or len(values["categoria"]) > 60:
                        raise ValueError("Informe a categoria (ate 60 caracteres).")
                    action = "atualizar" if record else "cadastrar"
                    if not messagebox.askyesno(
                        "Confirmar operacao",
                        f"Confirma {action} o produto '{values['nome']}'?",
                        parent=window,
                    ):
                        return
                    self.save_product(values, record.get("id") if record else None)
                else:
                    values["quantidade"] = self.parse_int(values["quantidade"], "Quantidade", minimum=1)
                    values["valor_unitario"] = self.parse_money(values["valor_unitario"])
                    try:
                        datetime.strptime(values["data_movimento"], "%Y-%m-%d")
                    except ValueError as error:
                        raise ValueError("A data deve estar no formato AAAA-MM-DD.") from error
                    if len(values["observacao"]) > 255:
                        raise ValueError("A observacao deve ter ate 255 caracteres.")
                    product_name = self.products.loc[
                        self.products["id"] == values["produto_id"], "nome"
                    ].iloc[0]
                    action = "atualizar" if record else "registrar"
                    if not messagebox.askyesno(
                        "Confirmar operacao",
                        f"Confirma {action} {values['tipo'].lower()} de "
                        f"{values['quantidade']} unidade(s) de '{product_name}'?",
                        parent=window,
                    ):
                        return
                    self.save_movement(values, record.get("id") if record else None)
                window.destroy()
                self.refresh_data()
            except (ValueError, mysql.connector.Error) as error:
                messagebox.showerror("Nao foi possivel salvar", str(error), parent=window)

        buttons = ttk.Frame(form)
        buttons.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="Cancelar", command=window.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(buttons, text="Salvar", command=save).pack(side="right")
        window.bind("<Return>", lambda _event: save())
        window.wait_window()

    @staticmethod
    def parse_int(value, label, minimum):
        try:
            result = int(value)
        except (TypeError, ValueError) as error:
            raise ValueError(f"{label} deve ser um numero inteiro.") from error
        if result < minimum:
            raise ValueError(f"{label} deve ser no minimo {minimum}.")
        return result

    @staticmethod
    def parse_money(value):
        try:
            result = Decimal(value.replace(",", "."))
        except (InvalidOperation, AttributeError) as error:
            raise ValueError("Informe um valor monetario valido.") from error
        if not result.is_finite() or result < 0:
            raise ValueError("O valor monetario nao pode ser negativo.")
        return result

    def save_product(self, values, product_id):
        connection = connect_db()
        try:
            cursor = connection.cursor()
            if product_id is None:
                cursor.execute(
                    """INSERT INTO produtos (nome, categoria, preco_unitario, estoque_inicial)
                       VALUES (%s, %s, %s, %s)""",
                    (values["nome"], values["categoria"], values["preco_unitario"], values["estoque_inicial"]),
                )
            else:
                cursor.execute(
                    """SELECT COALESCE(SUM(CASE WHEN tipo = 'Entrada' THEN quantidade ELSE -quantidade END), 0)
                       FROM movimentacoes WHERE produto_id = %s""",
                    (product_id,),
                )
                movement_balance = int(cursor.fetchone()[0])
                if values["estoque_inicial"] + movement_balance < 0:
                    raise ValueError("Esse estoque inicial deixaria o saldo atual negativo.")
                cursor.execute(
                    """UPDATE produtos SET nome=%s, categoria=%s, preco_unitario=%s,
                       estoque_inicial=%s WHERE id=%s""",
                    (values["nome"], values["categoria"], values["preco_unitario"], values["estoque_inicial"], product_id),
                )
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def save_movement(self, values, movement_id):
        connection = connect_db()
        try:
            cursor = connection.cursor()
            cursor.execute("SELECT id FROM produtos WHERE id=%s FOR UPDATE", (values["produto_id"],))
            if cursor.fetchone() is None:
                raise ValueError("O produto selecionado nao existe.")
            if movement_id is not None:
                cursor.execute(
                    "SELECT produto_id FROM movimentacoes WHERE id=%s FOR UPDATE",
                    (movement_id,),
                )
                old = cursor.fetchone()
                if old is None:
                    raise ValueError("A movimentacao selecionada nao existe mais.")
                old_product_id = int(old[0])
                old_balance = current_stock(cursor, old_product_id, movement_id)
                if old_balance < 0:
                    raise ValueError("A alteracao deixaria o estoque do produto anterior negativo.")
            available = current_stock(cursor, values["produto_id"], movement_id)
            if values["tipo"] == "Saida" and values["quantidade"] > available:
                raise ValueError(f"Estoque insuficiente. Disponivel: {available}.")
            if movement_id is None:
                cursor.execute(
                    """INSERT INTO movimentacoes
                       (produto_id, tipo, quantidade, data_movimento, valor_unitario, observacao)
                       VALUES (%s, %s, %s, %s, %s, %s)""",
                    (values["produto_id"], values["tipo"], values["quantidade"],
                     values["data_movimento"], values["valor_unitario"], values["observacao"]),
                )
            else:
                cursor.execute(
                    """UPDATE movimentacoes SET produto_id=%s, tipo=%s, quantidade=%s,
                       data_movimento=%s, valor_unitario=%s, observacao=%s WHERE id=%s""",
                    (values["produto_id"], values["tipo"], values["quantidade"],
                     values["data_movimento"], values["valor_unitario"], values["observacao"], movement_id),
                )
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def delete_selected(self, kind):
        tree = self.products_tree if kind == "product" else self.movements_tree
        record_id = self.selected_id(tree, "um registro para excluir")
        if record_id is None:
            return
        label = "produto" if kind == "product" else "movimentacao"
        if not messagebox.askyesno("Confirmar exclusao", f"Deseja excluir este {label}?", parent=self.root):
            return
        connection = None
        try:
            connection = connect_db()
            cursor = connection.cursor()
            if kind == "product":
                cursor.execute("DELETE FROM produtos WHERE id=%s", (record_id,))
            else:
                cursor.execute("SELECT produto_id FROM movimentacoes WHERE id=%s FOR UPDATE", (record_id,))
                record = cursor.fetchone()
                if record is None:
                    raise ValueError("A movimentacao selecionada nao existe mais.")
                if current_stock(cursor, int(record[0]), record_id) < 0:
                    raise ValueError("A exclusao deixaria o estoque negativo.")
                cursor.execute("DELETE FROM movimentacoes WHERE id=%s", (record_id,))
            connection.commit()
            self.refresh_data()
        except (mysql.connector.Error, ValueError) as error:
            if connection is not None:
                connection.rollback()
            messagebox.showerror(
                "Nao foi possivel excluir",
                "O produto pode ter movimentacoes vinculadas. Exclua ou ajuste os vinculos primeiro.\n\n"
                + str(error),
                parent=self.root,
            )
        finally:
            if connection is not None:
                connection.close()

    def build_reports(self):
        connection = connect_db()
        try:
            cursor = connection.cursor()
            consolidated = query_frame(
                cursor,
                """SELECT m.id AS movimentacao_id, m.data_movimento, p.id AS produto_id,
                          p.nome AS produto, p.categoria, m.tipo, m.quantidade,
                          m.valor_unitario, m.quantidade * m.valor_unitario AS valor_total,
                          m.observacao
                   FROM movimentacoes AS m
                   JOIN produtos AS p ON p.id = m.produto_id
                   ORDER BY m.data_movimento, m.id""",
            )
            products = query_frame(cursor, PRODUCTS_SQL)
            movements = query_frame(cursor, MOVEMENTS_SQL)
            cursor.close()
            if products.empty:
                by_category = pd.DataFrame(
                    columns=["indicador", "grupo", "produtos", "unidades_em_estoque", "valor_estoque"]
                )
            else:
                by_category = products.groupby("categoria", as_index=False).agg(
                    produtos=("id", "count"),
                    unidades_em_estoque=("estoque_atual", "sum"),
                )
                category_value = products.assign(
                    valor_estoque=products["preco_unitario"].astype(float)
                    * products["estoque_atual"].astype(int)
                ).groupby("categoria")["valor_estoque"].sum()
                by_category["valor_estoque"] = by_category["categoria"].map(category_value)
                by_category.insert(0, "indicador", "Estoque por categoria")
                by_category = by_category.rename(columns={"categoria": "grupo"})
            by_type = movements.groupby("tipo", as_index=False).agg(
                quantidade_total=("quantidade", "sum"),
                media_por_movimentacao=("quantidade", "mean"),
                movimentacoes=("id", "count"),
            ) if not movements.empty else pd.DataFrame(
                columns=["tipo", "quantidade_total", "media_por_movimentacao", "movimentacoes"]
            )
            if not by_type.empty:
                by_type.insert(0, "indicador", "Movimentacoes por tipo")
                by_type = by_type.rename(columns={"tipo": "grupo"})
            summary = pd.concat([by_category, by_type], ignore_index=True, sort=False)
        finally:
            connection.close()
        return consolidated, summary

    def export_to_sheets(self):
        if not messagebox.askyesno(
            "Confirmar exportacao",
            "Deseja enviar os dados atuais para o Google Sheets?\n\n"
            "As abas Dados Consolidados e Resumo Gerencial serao atualizadas "
            "e formatadas.",
            parent=self.root,
        ):
            return
        try:
            consolidated, summary = self.build_reports()
            url = export_frames_to_sheets(consolidated, summary)
            messagebox.showinfo("Exportacao concluida", f"Relatorio enviado ao Google Sheets:\n{url}", parent=self.root)
        except (mysql.connector.Error, gspread.exceptions.GSpreadException, OSError, ValueError) as error:
            messagebox.showerror("Falha na exportacao", str(error), parent=self.root)

    def refresh_charts(self):
        try:
            connection = connect_db()
            try:
                cursor = connection.cursor()
                category = query_frame(
                    cursor,
                    """SELECT p.categoria, SUM(m.quantidade) AS quantidade
                       FROM movimentacoes AS m
                       JOIN produtos AS p ON p.id = m.produto_id
                       GROUP BY p.categoria ORDER BY p.categoria""",
                )
                monthly = query_frame(
                    cursor,
                    """SELECT DATE_FORMAT(data_movimento, '%Y-%m') AS periodo,
                              SUM(CASE WHEN tipo='Entrada' THEN quantidade ELSE -quantidade END) AS saldo
                       FROM movimentacoes GROUP BY DATE_FORMAT(data_movimento, '%Y-%m')
                       ORDER BY periodo"""
                )
                cursor.close()
            finally:
                connection.close()
            figure = Figure(figsize=(10, 4.8), dpi=100)
            figure.patch.set_facecolor(self.colors["surface"])
            axes = figure.subplots(1, 2)
            for axis in axes:
                axis.set_facecolor(self.colors["surface"])
                axis.spines["top"].set_visible(False)
                axis.spines["right"].set_visible(False)
                axis.spines["left"].set_color(self.colors["border"])
                axis.spines["bottom"].set_color(self.colors["border"])
                axis.tick_params(colors=self.colors["muted"], labelsize=9)
                axis.grid(axis="y", color=self.colors["border"], linewidth=0.8, alpha=0.8)
                axis.set_axisbelow(True)
            axes[0].bar(
                category["categoria"],
                category["quantidade"],
                color=self.colors["green"],
                width=0.58,
                zorder=3,
            )
            axes[0].set_title("Movimentacoes por categoria", loc="left", pad=14, fontweight="bold", color=self.colors["ink"])
            axes[0].set_ylabel("Unidades movimentadas", color=self.colors["muted"])
            axes[0].tick_params(axis="x", rotation=25)
            month_positions = list(range(len(monthly)))
            monthly_balance = monthly["saldo"].tolist()
            axes[1].plot(
                month_positions,
                monthly_balance,
                marker="o",
                markersize=6,
                linewidth=2.5,
                color="#C27A3A",
            )
            axes[1].fill_between(month_positions, monthly_balance, color="#C27A3A", alpha=0.12)
            axes[1].set_title("Saldo mensal de movimentacoes", loc="left", pad=14, fontweight="bold", color=self.colors["ink"])
            axes[1].set_ylabel("Entradas - saidas", color=self.colors["muted"])
            axes[1].set_xticks(month_positions, monthly["periodo"].tolist())
            axes[1].tick_params(axis="x", rotation=35)
            figure.tight_layout()
            if self.chart_canvas is not None:
                self.chart_canvas.get_tk_widget().destroy()
            self.chart_canvas = FigureCanvasTkAgg(figure, master=self.chart_frame)
            self.chart_canvas.draw()
            self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)
            figure.savefig(Path(__file__).with_name("graficos_estoque.png"))
        except (mysql.connector.Error, OSError, ValueError) as error:
            messagebox.showerror("Falha ao gerar graficos", str(error), parent=self.root)


def main():
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
