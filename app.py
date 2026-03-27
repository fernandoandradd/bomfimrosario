import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(
    page_title="Consolidador de Estoque",
    page_icon="📦",
    layout="wide",
)

st.title("📦 Consolidador de Patrimônio em Estoque")
st.markdown("Consolida planilhas de inventário com múltiplas páginas em uma única tabela limpa.")

uploaded_file = st.file_uploader(
    "Envie o arquivo Excel (.xlsx) de inventário",
    type=["xlsx"],
    help="Arquivo exportado do C-Plus com registro de inventário",
)

# ── Mapeamento fixo das colunas (índices 0-based) ──────────────────────────
COL_MAP = {
    0:  "Código",
    1:  "Produto",
    9:  "UN",
    10: "Qtde",
    12: "Preço Custo",
    14: "Custo Final",
    16: "Valor Total",
}

# Palavras-chave que identificam linhas de cabeçalho / rodapé a descartar
SKIP_KEYWORDS = {"Código", "Software C-Plus", "Registro de inventário", "Page "}


def is_data_row(row: tuple) -> bool:
    """Retorna True se a linha contiver dados reais de produto."""
    codigo = row[0]
    produto = row[1]

    # Linha toda vazia
    if codigo is None and produto is None:
        return False

    # Converter para string para checar palavras-chave
    codigo_str = str(codigo) if codigo is not None else ""
    produto_str = str(produto) if produto is not None else ""

    for kw in SKIP_KEYWORDS:
        if kw in codigo_str or kw in produto_str:
            return False

    # Verificar se "Page X of Y" está em qualquer coluna
    for cell in row:
        if cell is not None and isinstance(cell, str) and cell.strip().startswith("Page "):
            return False

    # Pular linhas de data (datetime)
    if hasattr(codigo, "strftime"):
        return False

    # Pular linhas de cabeçalho de seção (ex: "Todas as seções")
    for cell in row:
        if cell is not None and isinstance(cell, str) and "Todas" in cell:
            return False

    # Se código existe e é alfanumérico, é dado válido
    if codigo is not None and str(codigo).strip():
        return True

    return False


def process_file(file) -> pd.DataFrame:
    """Lê o arquivo Excel e extrai apenas as linhas de dados consolidadas."""
    wb = load_workbook(file, read_only=True, data_only=True)
    ws = wb.active

    records = []
    for row in ws.iter_rows(values_only=True):
        if is_data_row(row):
            record = {col_name: row[col_idx] for col_idx, col_name in COL_MAP.items()}
            records.append(record)

    wb.close()

    df = pd.DataFrame(records, columns=list(COL_MAP.values()))

    # Garantir tipos corretos
    df["Código"] = df["Código"].astype(str).str.strip()
    df["Produto"] = df["Produto"].astype(str).str.strip()
    df["UN"] = df["UN"].astype(str).str.strip().replace("None", "")
    df["Qtde"] = pd.to_numeric(df["Qtde"], errors="coerce").fillna(0)
    df["Preço Custo"] = pd.to_numeric(df["Preço Custo"], errors="coerce").fillna(0)
    df["Custo Final"] = pd.to_numeric(df["Custo Final"], errors="coerce").fillna(0)
    df["Valor Total"] = pd.to_numeric(df["Valor Total"], errors="coerce").fillna(0)

    return df


def to_excel_download(df: pd.DataFrame) -> bytes:
    """Gera bytes de um arquivo Excel formatado para download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Estoque Consolidado")

        # Ajustar largura das colunas
        ws = writer.sheets["Estoque Consolidado"]
        col_widths = {
            "A": 12,   # Código
            "B": 55,   # Produto
            "C": 8,    # UN
            "D": 10,   # Qtde
            "E": 14,   # Preço Custo
            "F": 14,   # Custo Final
            "G": 16,   # Valor Total
        }
        for col_letter, width in col_widths.items():
            ws.column_dimensions[col_letter].width = width

        # Formato numérico para colunas monetárias
        from openpyxl.styles import numbers
        for row_cells in ws.iter_rows(min_row=2, min_col=5, max_col=7):
            for cell in row_cells:
                cell.number_format = '#,##0.0000'

    return output.getvalue()


# ── Processamento ──────────────────────────────────────────────────────────
if uploaded_file is not None:
    with st.spinner("Processando arquivo..."):
        df = process_file(uploaded_file)

    st.success(f"✅ **{len(df):,}** produtos extraídos com sucesso!")

    # ── Métricas resumo ────────────────────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total de Itens", f"{len(df):,}")
    with col2:
        st.metric("Qtde Total em Estoque", f"{df['Qtde'].sum():,.0f}")
    with col3:
        st.metric("Valor Total do Estoque", f"R$ {df['Valor Total'].sum():,.2f}")
    with col4:
        custo_medio = df.loc[df['Qtde'] > 0, 'Custo Final'].mean()
        st.metric("Custo Final Médio", f"R$ {custo_medio:,.4f}")

    st.divider()

    # ── Filtros ────────────────────────────────────────────────────────
    col_busca, col_un = st.columns([3, 1])
    with col_busca:
        busca = st.text_input(
            "🔍 Buscar por código ou produto",
            placeholder="Digite para filtrar...",
        )
    with col_un:
        unidades = ["Todos"] + sorted(df["UN"].unique().tolist())
        un_filtro = st.selectbox("Unidade", unidades)

    df_filtrado = df.copy()
    if busca:
        mask = (
            df_filtrado["Código"].str.contains(busca, case=False, na=False)
            | df_filtrado["Produto"].str.contains(busca, case=False, na=False)
        )
        df_filtrado = df_filtrado[mask]
    if un_filtro != "Todos":
        df_filtrado = df_filtrado[df_filtrado["UN"] == un_filtro]

    st.markdown(f"**Exibindo {len(df_filtrado):,} de {len(df):,} registros**")

    # ── Tabela ─────────────────────────────────────────────────────────
    st.dataframe(
        df_filtrado,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Código": st.column_config.TextColumn("Código", width="small"),
            "Produto": st.column_config.TextColumn("Produto", width="large"),
            "UN": st.column_config.TextColumn("UN", width="small"),
            "Qtde": st.column_config.NumberColumn("Qtde", format="%d"),
            "Preço Custo": st.column_config.NumberColumn(
                "Preço Custo", format="R$ %.4f"
            ),
            "Custo Final": st.column_config.NumberColumn(
                "Custo Final", format="R$ %.4f"
            ),
            "Valor Total": st.column_config.NumberColumn(
                "Valor Total", format="R$ %.2f"
            ),
        },
        height=600,
    )

    # ── Download ───────────────────────────────────────────────────────
    st.divider()
    excel_bytes = to_excel_download(df_filtrado)
    st.download_button(
        label="⬇️ Baixar planilha consolidada (.xlsx)",
        data=excel_bytes,
        file_name="estoque_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
else:
    st.info("👆 Envie o arquivo Excel de inventário para começar.")
