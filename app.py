import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from io import BytesIO

st.set_page_config(
    page_title="Gestão de Estoque — Físico x Fiscal",
    page_icon="📦",
    layout="wide",
)

# ══════════════════════════════════════════════════════════════════════════════
# MAPEAMENTOS DE COLUNAS — cada relatório do C-Plus tem layout diferente
# ══════════════════════════════════════════════════════════════════════════════

# Relatório de Loja (Físico):
#   Dados reais nas colunas: Código(0), Produto(1), UN(9), Qtde(10),
#   Preço Custo(12), Custo Final(14), Valor Total(16)
COL_MAP_FISICO = {
    "Código":      0,
    "Produto":     1,
    "UN":          9,
    "Qtde":       10,
    "Preço Custo":12,
    "Custo Final":14,
    "Valor Total":16,
}

# Relatório Fiscal:
#   Dados reais nas colunas: Código(0), NCM(1), Produto(2), UN(10),
#   Último Custo(12), Qtde(13), Custo Médio(14), Valor Total(16)
COL_MAP_FISCAL = {
    "Código":       0,
    "NCM":          1,
    "Produto":      2,
    "UN":          10,
    "Último Custo":12,
    "Qtde":        13,
    "Custo Médio": 14,
    "Valor Total": 16,
}

SKIP_KEYWORDS = {"Código", "Codigo", "Software C-Plus", "Registro de inventário", "Page "}
OUTPUT_COLS_FISICO = ["Código", "Produto", "UN", "Qtde", "Preço Custo", "Custo Final", "Valor Total"]
OUTPUT_COLS_FISCAL = ["Código", "NCM", "Produto", "UN", "Qtde", "Último Custo", "Custo Médio", "Valor Total"]


# ══════════════════════════════════════════════════════════════════════════════
# FUNÇÕES DE PROCESSAMENTO
# ══════════════════════════════════════════════════════════════════════════════

def is_data_row(row: tuple) -> bool:
    """Verifica se a linha contém dados reais de produto."""
    codigo = row[0]
    if codigo is None:
        return False
    if hasattr(codigo, "strftime"):
        return False
    codigo_str = str(codigo)
    for kw in SKIP_KEYWORDS:
        if kw in codigo_str:
            return False
    for cell in row:
        if cell is not None and isinstance(cell, str):
            s = cell.strip()
            if s.startswith("Page ") or "Todas" in s or "Registro de" in s or "Software C-Plus" in s:
                return False
    if str(codigo).strip():
        return True
    return False


def detect_layout(file) -> str:
    """Detecta se o arquivo é FÍSICO ou FISCAL pelo cabeçalho."""
    wb = load_workbook(file, read_only=True)
    ws = wb.active
    for row in ws.iter_rows(max_row=15, values_only=True):
        cells = [str(c).strip().lower() if c else "" for c in row]
        joined = " ".join(cells)
        if "ncm" in joined:
            wb.close()
            file.seek(0)
            return "fiscal"
        if "custo final" in joined:
            wb.close()
            file.seek(0)
            return "fisico"

    # Fallback: checar se col 1 dos dados parece NCM (8 dígitos numéricos)
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 7:
            continue
        if is_data_row(row):
            col1 = str(row[1]).strip() if row[1] else ""
            wb.close()
            file.seek(0)
            if col1.isdigit() and len(col1) == 8:
                return "fiscal"
            else:
                return "fisico"
    wb.close()
    file.seek(0)
    return "fisico"


def process_cplus_file(file, col_map: dict, output_cols: list) -> pd.DataFrame:
    """Processa arquivo bruto do C-Plus usando o mapeamento de colunas correto."""
    wb = load_workbook(file, read_only=True, data_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(values_only=True):
        if is_data_row(row):
            record = {}
            for col_name, col_idx in col_map.items():
                record[col_name] = row[col_idx] if col_idx < len(row) else None
            records.append(record)
    wb.close()

    df = pd.DataFrame(records)
    if df.empty:
        return df

    # Limpar tipos
    df["Código"] = df["Código"].astype(str).str.strip()
    if "Produto" in df.columns:
        df["Produto"] = df["Produto"].astype(str).str.strip().replace({"None": ""})
    if "NCM" in df.columns:
        df["NCM"] = df["NCM"].astype(str).str.strip().replace({"None": ""})
    if "UN" in df.columns:
        df["UN"] = df["UN"].fillna("").astype(str).str.strip().replace({"None": "", "nan": ""})

    numeric_cols = [c for c in df.columns if c not in ("Código", "Produto", "NCM", "UN")]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    existing_cols = [c for c in output_cols if c in df.columns]
    return df[existing_cols]


def smart_load(file):
    """Carrega arquivo detectando automaticamente o layout."""
    layout = detect_layout(file)
    if layout == "fiscal":
        df = process_cplus_file(file, COL_MAP_FISCAL, OUTPUT_COLS_FISCAL)
    else:
        df = process_cplus_file(file, COL_MAP_FISICO, OUTPUT_COLS_FISICO)
    return df, layout


def to_excel_download(dfs: dict[str, pd.DataFrame]) -> bytes:
    """Gera Excel com múltiplas abas formatado."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            if df.empty:
                continue
            name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=name)
            ws = writer.sheets[name]
            header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=11)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            for col_cells in ws.columns:
                max_len = 0
                for cell in col_cells:
                    max_len = max(max_len, len(str(cell.value or "")))
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 55)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════════════════════════════════════

st.title("📦 Gestão de Estoque — Físico x Fiscal")

tab1, tab2 = st.tabs(["📋 Consolidar Planilha", "🔍 Comparar Físico x Fiscal"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — CONSOLIDADOR
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Consolidar planilha de inventário")
    st.markdown(
        "Envie o arquivo bruto exportado do C-Plus (físico ou fiscal). "
        "O sistema detecta automaticamente o layout e consolida todas as páginas."
    )

    file_consolidar = st.file_uploader("📁 Arquivo de inventário (.xlsx)", type=["xlsx"], key="consolidar")

    if file_consolidar:
        with st.spinner("Detectando layout e processando..."):
            df_cons, layout_cons = smart_load(file_consolidar)

        tipo_label = "FISCAL" if layout_cons == "fiscal" else "FÍSICO"
        st.success(f"✅ Layout detectado: **{tipo_label}** — **{len(df_cons):,}** produtos extraídos!")

        c1, c2, c3 = st.columns(3)
        c1.metric("Total de Itens", f"{len(df_cons):,}")
        c2.metric("Qtde em Estoque", f"{df_cons['Qtde'].sum():,.0f}")
        c3.metric("Valor Total", f"R$ {df_cons['Valor Total'].sum():,.2f}")

        st.divider()
        busca = st.text_input("🔍 Buscar por código ou produto", key="busca_cons")
        df_show = df_cons
        if busca:
            mask = df_show["Código"].str.contains(busca, case=False, na=False)
            if "Produto" in df_show.columns:
                mask = mask | df_show["Produto"].str.contains(busca, case=False, na=False)
            df_show = df_show[mask]

        st.markdown(f"**Exibindo {len(df_show):,} de {len(df_cons):,} registros**")
        st.dataframe(df_show, use_container_width=True, hide_index=True, height=500)

        excel = to_excel_download({"Estoque Consolidado": df_show})
        st.download_button(
            "⬇️ Baixar consolidado (.xlsx)", excel,
            file_name=f"estoque_{tipo_label.lower()}_consolidado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — COMPARAÇÃO FÍSICO x FISCAL
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Comparar estoque Físico x Fiscal")
    st.markdown(
        "Envie os dois arquivos brutos do C-Plus. O sistema detecta automaticamente "
        "qual é o físico e qual é o fiscal, consolida ambos e gera a análise de divergências."
    )

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        file_fisico = st.file_uploader("📦 Estoque FÍSICO (.xlsx)", type=["xlsx"], key="fisico")
    with col_up2:
        file_fiscal = st.file_uploader("📑 Estoque FISCAL (.xlsx)", type=["xlsx"], key="fiscal")

    if file_fisico and file_fiscal:
        with st.spinner("Consolidando e comparando arquivos..."):
            df_fis, layout_fis = smart_load(file_fisico)
            df_fisc, layout_fisc = smart_load(file_fiscal)

        st.info(
            f"📦 Físico: **{len(df_fis):,}** itens (layout: {layout_fis.upper()})  |  "
            f"📑 Fiscal: **{len(df_fisc):,}** itens (layout: {layout_fisc.upper()})"
        )

        # ── Preparar para merge ────────────────────────────────────────
        fis = df_fis[["Código", "Produto", "Qtde", "Valor Total"]].copy()
        fis.columns = ["Código", "Produto_Físico", "Qtde_Físico", "VT_Físico"]

        fisc_merge_cols = ["Código", "Produto", "Qtde", "Valor Total"]
        if "NCM" in df_fisc.columns:
            fisc_merge_cols.insert(1, "NCM")
        fisc = df_fisc[fisc_merge_cols].copy()
        rename_map = {"Produto": "Produto_Fiscal", "Qtde": "Qtde_Fiscal", "Valor Total": "VT_Fiscal"}
        fisc = fisc.rename(columns=rename_map)

        # ── Merge por Código ───────────────────────────────────────────
        merged = pd.merge(fis, fisc, on="Código", how="outer", indicator=True)

        merged["Qtde_Físico"] = merged["Qtde_Físico"].fillna(0)
        merged["Qtde_Fiscal"] = merged["Qtde_Fiscal"].fillna(0)
        merged["VT_Físico"] = merged["VT_Físico"].fillna(0)
        merged["VT_Fiscal"] = merged["VT_Fiscal"].fillna(0)

        # Produto consolidado
        merged["Produto"] = merged["Produto_Físico"].fillna("").replace({"": None})
        merged["Produto"] = merged["Produto"].fillna(merged.get("Produto_Fiscal", ""))
        merged["Produto"] = merged["Produto"].fillna("")

        # Divergências
        merged["Dif_Qtde"] = merged["Qtde_Físico"] - merged["Qtde_Fiscal"]
        merged["Dif_Valor"] = merged["VT_Físico"] - merged["VT_Fiscal"]

        def classificar(row):
            if row["_merge"] == "left_only":
                return "Só no Físico"
            elif row["_merge"] == "right_only":
                return "Só no Fiscal"
            elif row["Dif_Qtde"] > 0:
                return "Físico > Fiscal"
            elif row["Dif_Qtde"] < 0:
                return "Fiscal > Físico"
            else:
                return "Iguais"

        merged["Situação"] = merged.apply(classificar, axis=1)

        display_cols = ["Código", "Produto"]
        if "NCM" in merged.columns:
            display_cols.append("NCM")
        display_cols += ["Qtde_Físico", "Qtde_Fiscal", "Dif_Qtde",
                         "VT_Físico", "VT_Fiscal", "Dif_Valor", "Situação"]
        result = merged[display_cols].copy()

        # ── MÉTRICAS ──────────────────────────────────────────────────
        st.divider()
        st.markdown("### 📊 Resumo Geral")

        so_fisico = result[result["Situação"] == "Só no Físico"]
        so_fiscal = result[result["Situação"] == "Só no Fiscal"]
        fisico_maior = result[result["Situação"] == "Físico > Fiscal"]
        fiscal_maior = result[result["Situação"] == "Fiscal > Físico"]
        iguais = result[result["Situação"] == "Iguais"]

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Iguais", f"{len(iguais):,}")
        m2.metric("Só no Físico", f"{len(so_fisico):,}")
        m3.metric("Só no Fiscal", f"{len(so_fiscal):,}")
        m4.metric("Físico > Fiscal", f"{len(fisico_maior):,}")
        m5.metric("Fiscal > Físico", f"{len(fiscal_maior):,}")

        st.divider()
        st.markdown("### 💰 Impacto em Valores")

        vt_fis = result["VT_Físico"].sum()
        vt_fisc = result["VT_Fiscal"].sum()
        dif_total = vt_fis - vt_fisc

        v1, v2, v3 = st.columns(3)
        v1.metric("Valor Total Físico", f"R$ {vt_fis:,.2f}")
        v2.metric("Valor Total Fiscal", f"R$ {vt_fisc:,.2f}")
        v3.metric(
            "Diferença (Físico − Fiscal)", f"R$ {dif_total:,.2f}",
            delta=f"R$ {dif_total:,.2f}",
            delta_color="normal" if dif_total >= 0 else "inverse",
        )

        v4, v5 = st.columns(2)
        v4.metric("Valor só no Físico (sem registro fiscal)", f"R$ {so_fisico['VT_Físico'].sum():,.2f}")
        v5.metric("Valor só no Fiscal (sem contagem física)", f"R$ {so_fiscal['VT_Fiscal'].sum():,.2f}")

        # ── FILTROS E TABELA ───────────────────────────────────────────
        st.divider()
        st.markdown("### 📋 Detalhamento")

        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            busca_comp = st.text_input("🔍 Buscar por código ou produto", key="busca_comp")
        with col_f2:
            situacoes = ["Todos"] + sorted(result["Situação"].unique().tolist())
            sit_filtro = st.selectbox("Filtrar por situação", situacoes)

        df_view = result.copy()
        if busca_comp:
            mask = (
                df_view["Código"].str.contains(busca_comp, case=False, na=False)
                | df_view["Produto"].astype(str).str.contains(busca_comp, case=False, na=False)
            )
            df_view = df_view[mask]
        if sit_filtro != "Todos":
            df_view = df_view[df_view["Situação"] == sit_filtro]

        df_view = df_view.sort_values("Dif_Valor", key=abs, ascending=False)

        st.markdown(f"**Exibindo {len(df_view):,} de {len(result):,} registros**")

        st.dataframe(
            df_view,
            use_container_width=True,
            hide_index=True,
            height=600,
            column_config={
                "Código": st.column_config.TextColumn("Código", width="small"),
                "Produto": st.column_config.TextColumn("Produto", width="large"),
                "NCM": st.column_config.TextColumn("NCM", width="small"),
                "Qtde_Físico": st.column_config.NumberColumn("Qtde Físico", format="%d"),
                "Qtde_Fiscal": st.column_config.NumberColumn("Qtde Fiscal", format="%d"),
                "Dif_Qtde": st.column_config.NumberColumn("Dif. Qtde", format="%d"),
                "VT_Físico": st.column_config.NumberColumn("VT Físico", format="R$ %.2f"),
                "VT_Fiscal": st.column_config.NumberColumn("VT Fiscal", format="R$ %.2f"),
                "Dif_Valor": st.column_config.NumberColumn("Dif. Valor", format="R$ %.2f"),
                "Situação": st.column_config.TextColumn("Situação", width="medium"),
            },
        )

        # ── DOWNLOAD ───────────────────────────────────────────────────
        st.divider()
        sheets = {
            "Comparativo Completo": df_view,
            "Só no Físico": so_fisico,
            "Só no Fiscal": so_fiscal,
            "Físico maior": fisico_maior,
            "Fiscal maior": fiscal_maior,
            "Iguais": iguais,
        }
        excel_comp = to_excel_download(sheets)
        st.download_button(
            "⬇️ Baixar relatório completo (.xlsx)",
            excel_comp,
            file_name="comparativo_fisico_x_fiscal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
