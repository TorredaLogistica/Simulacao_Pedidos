import io
import re
import unicodedata
from pathlib import Path
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Simulação de Pedidos SIMCARD", page_icon="📦", layout="wide")

CUSTO_POR_PEDIDO = 75.0
MESES = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}


def normalizar(v):
    s = "" if pd.isna(v) else str(v).strip().upper()
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")


def brl(v):
    return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner=False)
def ler_guia_parquet(arq, nome_guia, colunas=None):
    """Lê uma guia lógica do Parquet único criado a partir do Excel."""
    df = pd.read_parquet(
        arq,
        engine="pyarrow",
        filters=[("__guia_origem__", "==", nome_guia)],
    )
    df = df.drop(columns=["__guia_origem__", "__linha_origem__"], errors="ignore")
    df = df.dropna(axis=1, how="all")
    if colunas is not None:
        faltantes = [coluna for coluna in colunas if coluna not in df.columns]
        if faltantes:
            raise ValueError(
                f"Coluna(s) não encontrada(s) na guia {nome_guia}: {', '.join(faltantes)}"
            )
        df = df[colunas].copy()
    return df


@st.cache_data(show_spinner=False)
def ler_base(arq):
    return ler_guia_parquet(arq, "Base").astype("string")


def localizar_coluna(df, nomes, posicao_excel=None):
    mapa = {normalizar(c): c for c in df.columns}
    for nome in nomes:
        if normalizar(nome) in mapa:
            return mapa[normalizar(nome)]
    if posicao_excel and df.shape[1] >= posicao_excel:
        return df.columns[posicao_excel - 1]
    raise ValueError(f"Coluna não encontrada: {nomes[0]}")


def ano_mes(serie):
    dt = pd.to_datetime(serie, errors="coerce")
    faltantes = dt.isna() & serie.notna()
    if faltantes.any():
        dt_alt = pd.to_datetime(serie.loc[faltantes].astype(str), errors="coerce", dayfirst=True)
        dt.loc[faltantes] = dt_alt
    return dt.dt.year.astype("Int64"), dt.dt.month.astype("Int64")


def numero_br(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("R$", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return 0.0

def carregar_descricoes_tipo_doc(caminho):
    if not caminho.exists():
        return {}
    if caminho.suffix.lower() == ".parquet":
        tabela = ler_guia_parquet(caminho, "Tipo de Material").astype("string")
    else:
        tabela = pd.read_excel(caminho, sheet_name=0, dtype=str, engine="openpyxl")
    if tabela.shape[1] < 2:
        return {}
    codigos = tabela.iloc[:, 0].fillna("").astype(str).str.strip()
    descricoes = tabela.iloc[:, 1].fillna("").astype(str).str.strip()
    return {
        codigo: descricao
        for codigo, descricao in zip(codigos, descricoes)
        if codigo and descricao
    }

@st.cache_data(show_spinner=False)
def ler_detalhamento_material(arq):
    colunas = ["Documento SD", "Item (SD)", "Material", "Texto breve material",
               "Quantidade prevista", "UMB", "TIPO DA OV"]
    d = ler_guia_parquet(arq, "Detalhamento", colunas=colunas).astype("string")
    for c in ["Documento SD", "Item (SD)", "Material"]:
        d[c] = d[c].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    d["Quantidade prevista"] = d["Quantidade prevista"].map(numero_br)
    return d


@st.cache_data(show_spinner=False)
def resumo_quantidade_por_pedido(detalhe):
    return (detalhe.groupby("Documento SD", as_index=False)["Quantidade prevista"]
            .sum().rename(columns={"Documento SD": "pedido",
                                  "Quantidade prevista": "Quantidade Prevista"}))


@st.cache_data(show_spinner=False)
def excel_tabela(df, nome_aba):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        df.to_excel(w, sheet_name=nome_aba, index=False)
        ws = w.sheets[nome_aba]
        head = w.book.add_format({"bold": True, "bg_color": "#172B4D", "font_color": "white"})
        ws.freeze_panes(1, 0); ws.set_row(0, None, head)
        formato_inteiro = w.book.add_format({"num_format": "#,##0"})
        formato_moeda = w.book.add_format({"num_format": 'R$ #,##0.00;[Red](R$ #,##0.00);-'})
        for i, c in enumerate(df.columns):
            formato = None
            if c in ["Quantidade Prevista", "Quantidade prevista"]:
                formato = formato_inteiro
            elif c == "Custo de faturamento":
                formato = formato_moeda
            ws.set_column(i, i, min(max(len(str(c)) + 2, 14), 45), formato)
    return out.getvalue()


def preparar(df):
    """Prepara a nova base usando exclusivamente a guia Base."""
    colunas_necessarias = {
        "tipo_doc_vendas": ["Tipo Doc"],
        "loja": ["LOJA (SAP)", "CLIENTE"],
        "perfil_origem": ["SUB_CANAL"],
        "centro_distribuicao": ["CD Origem", "Codigo CD"],
        "pedido": ["PEDIDO"],
        "tipo_ov": ["Tipo da OV"],
        "competencia": ["Data Logística"],
        "custo_base": ["Custo da OV"],
    }
    encontradas = {
        destino: localizar_coluna(df, alternativas)
        for destino, alternativas in colunas_necessarias.items()
    }
    b = df[[encontradas[c] for c in colunas_necessarias]].copy()
    b.columns = list(colunas_necessarias)

    datas = pd.to_datetime(b["competencia"], errors="coerce")
    b["tipo_doc_vendas"] = b["tipo_doc_vendas"].fillna("").astype(str).str.strip()
    b["loja"] = b["loja"].fillna("").astype(str).str.strip()
    b["centro_distribuicao"] = b["centro_distribuicao"].fillna("").astype(str).str.strip().str.title()
    b["pedido"] = b["pedido"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    b["tipo_ov"] = b["tipo_ov"].map(normalizar)
    b["perfil_origem"] = b["perfil_origem"].map(normalizar)
    mapa_perfil = {
        "AGENTE AUTORIZADO": "AA",
        "LOJA PROPRIA": "LP",
        "AA": "AA",
        "LP": "LP",
    }
    b["perfil"] = b["perfil_origem"].map(mapa_perfil).fillna("")
    b["ano"] = datas.dt.year.astype("Int64")
    b["mes_num"] = datas.dt.month.astype("Int64")
    b["custo_faturamento"] = b["custo_base"].map(numero_br)

    # A quantidade prevista será incorporada posteriormente a partir da guia Detalhamento,
    # vinculando Documento SD ao campo PEDIDO da guia Base.

    mascara_perfil = b["perfil"].isin(["AA", "LP"])
    mascara_chaves = b["loja"].ne("") & b["pedido"].ne("")
    mascara_data = datas.notna() & datas.le(pd.Timestamp(2026, 7, 31))
    mascara_custo = b["custo_faturamento"].gt(0)
    mascara_tipo_ov = b["tipo_ov"].ne("")
    auditoria = {
        "linhas_base": int(len(b)),
        "tipos_ov_validos": int(mascara_tipo_ov.sum()),
        "perfil_aa_lp": int(mascara_perfil.sum()),
        "datas_validas": int(mascara_data.fillna(False).sum()),
        "custos_validos": int(mascara_custo.sum()),
    }
    # Tipo da OV, ano e mês não são fixados aqui. A seleção é feita pelo usuário nos filtros.
    b = b[
        mascara_perfil & mascara_chaves & mascara_data
        & mascara_custo & mascara_tipo_ov
    ].copy()
    b["mes"] = b["mes_num"].map(MESES)
    return b, "Data Logística (coluna AU)", auditoria

def resumo_mensal(base):
    if base.empty:
        return pd.DataFrame(columns=["Mês do pedido", "Contagem Número do pedido", "Soma de Custo Faturamento"])
    max_mes = int(base["mes_num"].max())
    idx = pd.DataFrame({"mes_num": range(1, max_mes + 1)})
    agg = base.groupby("mes_num", as_index=False).agg(
        qtd=("pedido", "count"),
        custo=("custo_faturamento", "sum"),
    )
    r = idx.merge(agg, how="left", on="mes_num").fillna({"qtd": 0, "custo": 0})
    r["qtd"] = r["qtd"].astype(int)
    r["Mês do pedido"] = r["mes_num"].map(MESES)
    r["Contagem Número do pedido"] = r["qtd"]
    r["Soma de Custo Faturamento"] = r["custo"]
    return r[["Mês do pedido", "Contagem Número do pedido", "Soma de Custo Faturamento"]]

def calcular_curva_abc(base):
    """Classifica cada pedido pela Quantidade prevista da guia Detalhamento."""
    colunas = ["Curva ABC", "Quantidade de lojas", "Quantidade de pedidos", "% do volume"]
    detalhe_colunas = ["perfil", "loja", "pedido", "qtd_simcards_pedido", "Curva ABC"]
    if base.empty or "qtd_simcards_pedido" not in base.columns:
        return pd.DataFrame(columns=colunas), pd.DataFrame(columns=detalhe_colunas)

    detalhe = base[["perfil", "loja", "pedido", "qtd_simcards_pedido"]].copy()
    detalhe["qtd_simcards_pedido"] = pd.to_numeric(
        detalhe["qtd_simcards_pedido"], errors="coerce"
    ).fillna(0.0)

    # Criterio por quantidade prevista de SIMCARDs em cada pedido:
    # A = ate 10 | B = 11 a 50 | C = 51 a 200 | D = acima de 200.
    def faixa_abc(quantidade):
        if quantidade <= 10:
            return "A"
        if quantidade <= 50:
            return "B"
        if quantidade <= 200:
            return "C"
        return "D"

    detalhe["Curva ABC"] = detalhe["qtd_simcards_pedido"].map(faixa_abc)
    volume_total = float(detalhe["qtd_simcards_pedido"].sum())

    resumo = (
        detalhe.groupby("Curva ABC", as_index=False)
        .agg(
            **{
                "Quantidade de lojas": ("loja", "nunique"),
                "Quantidade de pedidos": ("pedido", "count"),
                "volume": ("qtd_simcards_pedido", "sum"),
            }
        )
    )
    resumo["% do volume"] = resumo["volume"] / volume_total if volume_total > 0 else 0.0
    ordem = pd.DataFrame({"Curva ABC": ["A", "B", "C", "D"]})
    resumo = ordem.merge(resumo, on="Curva ABC", how="left").fillna(
        {"Quantidade de lojas": 0, "Quantidade de pedidos": 0, "volume": 0, "% do volume": 0}
    )
    resumo["Quantidade de lojas"] = resumo["Quantidade de lojas"].astype(int)
    resumo["Quantidade de pedidos"] = resumo["Quantidade de pedidos"].astype(int)
    return resumo[colunas], detalhe[detalhe_colunas]

def calcular(base, resumo):
    total_pedidos = int(resumo["Contagem Número do pedido"].sum())
    meses = max(len(resumo), 1)
    atual_pedidos_mes = total_pedidos / meses
    atual_pedidos_ano = atual_pedidos_mes * 12
    atual_custo_mes = float(resumo["Soma de Custo Faturamento"].sum()) / meses
    atual_custo_ano = atual_custo_mes * 12

    por_loja = base.groupby(["perfil", "loja"]).agg(
        pedidos_total=("pedido", "count"),
        custo_unitario=("custo_faturamento", lambda x: float(x.mode().iloc[0]) if not x.mode().empty else float(x.mean())),
    ).reset_index()
    por_loja["media_atual"] = por_loja["pedidos_total"] / meses
    # Limites: nunca criar pedidos que não existem no ritmo atual.
    por_loja["c1"] = por_loja["media_atual"].clip(upper=1)
    por_loja["c2"] = por_loja["media_atual"].clip(upper=2)

    c1_pm = float(por_loja["c1"].sum())
    c1_cm = float((por_loja["c1"] * por_loja["custo_unitario"]).sum())
    c2_pm = float(por_loja["c2"].sum())
    c2_cm = float((por_loja["c2"] * por_loja["custo_unitario"]).sum())

    linhas = [
        ["Cenário Atual", atual_pedidos_mes, atual_custo_mes, atual_pedidos_ano, atual_custo_ano, 0.0],
        ["Cenário 1 - até 1 pedido/mês", c1_pm, c1_cm, c1_pm*12, c1_cm*12, max(atual_custo_ano-c1_cm*12, 0.0)],
        ["Cenário 2 - até 2 pedidos/mês", c2_pm, c2_cm, c2_pm*12, c2_cm*12, max(atual_custo_ano-c2_cm*12, 0.0)],
    ]
    return pd.DataFrame(linhas, columns=["Cenário", "Pedidos/mês", "Custo/mês", "Pedidos/ano", "Custo/ano", "Economia/ano"])

def excel_saida(resumo, cenarios, curva_abc, detalhe_curva, detalhe):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        resumo.to_excel(w, "Cenario Atual", index=False)
        cenarios.to_excel(w, "Comparativo Financeiro", index=False)
        curva_abc.to_excel(w, "Curva ABC", index=False)
        detalhe_curva.to_excel(w, "Curva ABC por Loja", index=False)
        detalhe.to_excel(w, "Pedidos Unicos", index=False)
        wb = w.book
        head = wb.add_format({"bold":True,"bg_color":"#172B4D","font_color":"white"})
        moeda = wb.add_format({"num_format":'R$ #,##0.00;[Red](R$ #,##0.00);-'})
        for nome in w.sheets:
            ws = w.sheets[nome]
            ws.freeze_panes(1,0)
            ws.set_row(0,None,head)
            ws.set_column(0,ws.dim_colmax,22)
        w.sheets["Cenario Atual"].set_column(2,2,22,moeda)
        w.sheets["Comparativo Financeiro"].set_column(2,2,18,moeda)
        w.sheets["Comparativo Financeiro"].set_column(4,5,18,moeda)
        w.sheets["Curva ABC"].set_column(3,3,14,wb.add_format({"num_format":"0%"}))
        w.sheets["Curva ABC por Loja"].set_column(3,3,22)
    return out.getvalue()


st.markdown("""
<style>
.block-container{padding-top:1.2rem;padding-bottom:2rem}
[data-testid="stMetric"]{background:#fff;border:1px solid #dce3ed;border-radius:10px;padding:14px;overflow:visible}
[data-testid="stMetricValue"]{overflow:visible}
[data-testid="stMetricValue"] > div{
    font-size:clamp(1.55rem,2.25vw,2.5rem);
    line-height:1.15;
    white-space:nowrap;
    overflow:visible;
    text-overflow:clip;
}
.titulo{font-size:2rem;font-weight:800;color:#172B4D;margin:0}.sub{font-size:2.5rem;font-weight:600;color:#687386;margin-top:0;line-height:1.25}
</style>
""", unsafe_allow_html=True)
st.markdown('<p class="titulo">Simulação de Pedidos SIMCARD</p>', unsafe_allow_html=True)
st.markdown('<p class="sub">Simulação de frequência e custos por período | nova base</p>', unsafe_allow_html=True)

# O Parquet único é carregado automaticamente da mesma pasta do aplicativo.
arq = Path(__file__).resolve().parent / "Base_Pedidos_AAs_PLs_2026.parquet"
if not arq.exists():
    st.error("A base 'Base_Pedidos_AAs_PLs_2026.parquet' não foi encontrada na pasta do aplicativo.")
    st.stop()

try:
    with st.spinner("Processando a base..."):
        df = ler_base(arq)
        pedidos, coluna_competencia, auditoria = preparar(df)
        detalhe_material_base = ler_detalhamento_material(arq)
        qtd_por_pedido = resumo_quantidade_por_pedido(detalhe_material_base).rename(
            columns={"Quantidade Prevista": "qtd_simcards_pedido"}
        )
        pedidos = pedidos.merge(qtd_por_pedido, on="pedido", how="left")
        pedidos["qtd_simcards_pedido"] = pedidos["qtd_simcards_pedido"].fillna(0.0)
except Exception as e:
    st.error(f"Erro ao processar a base: {e}")
    st.stop()

if pedidos.empty:
    st.warning("Nenhum pedido válido foi encontrado para os filtros selecionados.")
    st.error(
        "Diagnóstico: "
        f"linhas={auditoria['linhas_base']} | "
        f"Tipo da OV válido={auditoria['tipos_ov_validos']} | "
        f"AA/LP={auditoria['perfil_aa_lp']} | "
        f"data válida={auditoria['datas_validas']} | "
        f"custo válido={auditoria['custos_validos']}"
    )
    st.stop()

st.sidebar.markdown("## Filtros")

arquivo_descricoes = arq
descricoes_tipo_doc = carregar_descricoes_tipo_doc(arquivo_descricoes)

tipos_ov_lista = sorted(v for v in pedidos["tipo_ov"].unique() if v)
tipos_ov = st.sidebar.multiselect(
    "Tipo da OV",
    tipos_ov_lista,
    default=["SIMCARD"] if "SIMCARD" in tipos_ov_lista else [],
    placeholder="Todos os tipos",
)
b_tipo_ov = pedidos if not tipos_ov else pedidos[pedidos["tipo_ov"].isin(tipos_ov)]

anos_lista = sorted(int(v) for v in b_tipo_ov["ano"].dropna().unique())
anos = st.sidebar.multiselect(
    "Ano",
    anos_lista,
    default=[],
    placeholder="Todos os anos",
)
b_ano = b_tipo_ov if not anos else b_tipo_ov[b_tipo_ov["ano"].isin(anos)]

meses_disponiveis = sorted(int(v) for v in b_ano["mes_num"].dropna().unique())
meses_opcoes = [MESES[m] for m in meses_disponiveis]
meses_selecionados = st.sidebar.multiselect(
    "Mês",
    meses_opcoes,
    default=[],
    placeholder="Todos os meses",
)
meses_numeros = [m for m in meses_disponiveis if MESES[m] in meses_selecionados]
b_periodo = b_ano if not meses_numeros else b_ano[b_ano["mes_num"].isin(meses_numeros)]

pedido_pesquisa = st.sidebar.text_input("Pesquisar pedido", placeholder="Digite o número do pedido").strip()
b_pedido = b_periodo if not pedido_pesquisa else b_periodo[
    b_periodo["pedido"].astype(str).str.contains(pedido_pesquisa, case=False, na=False, regex=False)
]
apenas_pedidos_unicos = st.sidebar.checkbox(
    "Considerar apenas pedidos únicos", value=True,
    help="Quando ativado, cada número de pedido é considerado somente uma vez."
)
b_unicidade = (b_pedido.drop_duplicates(subset=["pedido"], keep="first").copy()
                if apenas_pedidos_unicos else b_pedido)

tipos_doc_lista = sorted(v for v in b_unicidade["tipo_doc_vendas"].unique() if v)
tipos_doc = st.sidebar.multiselect(
    "Tipo Doc Vendas",
    tipos_doc_lista,
    default=[],
    placeholder="Selecione um ou mais tipos",
    format_func=lambda codigo: f"{codigo} - {descricoes_tipo_doc[codigo]}" if codigo in descricoes_tipo_doc else codigo,
)
b_tipo = b_unicidade if not tipos_doc else b_unicidade[b_unicidade["tipo_doc_vendas"].isin(tipos_doc)]

centros_lista = sorted(v for v in b_tipo["centro_distribuicao"].unique() if v)
centros_distribuicao = st.sidebar.multiselect(
    "Centro Distribuição",
    centros_lista,
    default=[],
    placeholder="Selecione um ou mais centros",
)
b_centro = b_tipo if not centros_distribuicao else b_tipo[b_tipo["centro_distribuicao"].isin(centros_distribuicao)]

perfis = sorted(b_centro["perfil"].unique())
# A opção Todos é o padrão. O usuário pode selecionar AA ou LP quando necessário.
perfil = st.sidebar.selectbox(
    "Perfil AA ou LP",
    ["Todos"] + perfis,
    index=0,
)
bp = b_centro if perfil == "Todos" else b_centro[b_centro["perfil"].eq(perfil)]

lojas_lista = sorted(bp["loja"].unique())
lojas = st.sidebar.multiselect(
    "Identificação da loja",
    lojas_lista,
    default=[],
    placeholder="Selecione uma ou mais lojas",
)
bf = bp if not lojas else bp[bp["loja"].isin(lojas)]

resumo = resumo_mensal(bf)
cenarios = calcular(bf, resumo)
curva_abc, detalhe_curva = calcular_curva_abc(bf)
a, c1, c2 = cenarios.iloc[0], cenarios.iloc[1], cenarios.iloc[2]

# Base exclusiva dos cards de resumo dos cenários: ignora somente o filtro de mês.
b_cenario_pedido = b_ano if not pedido_pesquisa else b_ano[
    b_ano["pedido"].astype(str).str.contains(
        pedido_pesquisa, case=False, na=False, regex=False
    )
]
b_cenario_unicidade = (
    b_cenario_pedido.drop_duplicates(subset=["pedido"], keep="first").copy()
    if apenas_pedidos_unicos else b_cenario_pedido
)
b_cenario_tipo = b_cenario_unicidade if not tipos_doc else b_cenario_unicidade[
    b_cenario_unicidade["tipo_doc_vendas"].isin(tipos_doc)
]
b_cenario_centro = b_cenario_tipo if not centros_distribuicao else b_cenario_tipo[
    b_cenario_tipo["centro_distribuicao"].isin(centros_distribuicao)
]
b_cenario_perfil = b_cenario_centro if perfil == "Todos" else b_cenario_centro[
    b_cenario_centro["perfil"].eq(perfil)
]
bf_cenarios = b_cenario_perfil if not lojas else b_cenario_perfil[
    b_cenario_perfil["loja"].isin(lojas)
]
resumo_cenarios = resumo_mensal(bf_cenarios)
cenarios_fechados = calcular(bf_cenarios, resumo_cenarios)
a_fechado, c1_fechado, c2_fechado = cenarios_fechados.iloc[0], cenarios_fechados.iloc[1], cenarios_fechados.iloc[2]

st.caption(f"Filtro atual: Tipo da OV **{', '.join(tipos_ov) if tipos_ov else 'Todos'}** | Ano **{', '.join(map(str, anos)) if anos else 'Todos'}** | Mês **{', '.join(meses_selecionados) if meses_selecionados else 'Todos'}** | Perfil **{perfil}** | Loja(s): **{', '.join(lojas) if lojas else 'Todas'}** | Competência: **{coluna_competencia}**")

m1,m2,m3,m4 = st.columns(4)
m1.metric("Pedidos encontrados", f"{int(resumo['Contagem Número do pedido'].sum()):,}".replace(",", "."))
m2.metric("Lojas selecionadas", f"{int(bf['loja'].nunique()):,}".replace(",", "."))
m3.metric("Média mensal atual", f"{float(a['Pedidos/mês']):,.0f}".replace(",", "."))
m4.metric("Custo anual atual", brl(a["Custo/ano"]))

aba1,aba2,aba3,aba4 = st.tabs(["Cenário atual","Comparativo financeiro","Detalhamento","Análise Gerencial"])

with aba1:
    st.subheader("Cenário atual mês a mês")
    tabela = resumo.copy()
    total = pd.DataFrame([{"Mês do pedido":"Total Geral","Contagem Número do pedido":resumo["Contagem Número do pedido"].sum(),"Soma de Custo Faturamento":resumo["Soma de Custo Faturamento"].sum()}])
    linha_media = pd.DataFrame([{"Mês do pedido":"Média mês","Contagem Número do pedido":a["Pedidos/mês"],"Soma de Custo Faturamento":a["Custo/mês"]}])
    linha_ano = pd.DataFrame([{"Mês do pedido":"Média Ano","Contagem Número do pedido":a["Pedidos/ano"],"Soma de Custo Faturamento":a["Custo/ano"]}])
    tabela_exibicao = pd.concat([tabela,total,linha_media,linha_ano],ignore_index=True)

    col_t,col_g = st.columns([1.05,1])
    with col_t:
        st.dataframe(tabela_exibicao.style.format({
            "Contagem Número do pedido":lambda v:f"{v:,.0f}".replace(",","."),
            "Soma de Custo Faturamento":brl
        }),use_container_width=True,hide_index=True)
    with col_g:
        fig=go.Figure(go.Bar(x=resumo["Mês do pedido"],y=resumo["Contagem Número do pedido"],text=resumo["Contagem Número do pedido"],textposition="outside",marker_color="#2D6CDF"))
        fig.update_layout(title="Quantidade de pedidos por mês",xaxis_title="Mês",yaxis_title="Pedidos",showlegend=False,margin=dict(l=20,r=20,t=55,b=30))
        st.plotly_chart(fig,use_container_width=True)

    st.markdown("#### Resumo dos cenários para a seleção atual")
    r1,r2,r3 = st.columns(3)
    r1.info(f"**Cenário Atual**  \nMédia mês: **{a_fechado['Pedidos/mês']:.0f} pedidos | {brl(a_fechado['Custo/mês'])}**  \nMédia Ano: **{a_fechado['Pedidos/ano']:.0f} pedidos | {brl(a_fechado['Custo/ano'])}**")
    r2.success(f"**Cenário 1: até 1 pedido/mês**  \nMédia mês: **{c1_fechado['Pedidos/mês']:.0f} pedido | {brl(c1_fechado['Custo/mês'])}**  \nMédia Ano: **{c1_fechado['Pedidos/ano']:.0f} pedidos | {brl(c1_fechado['Custo/ano'])}**  \nEconomia Ano: **{brl(c1_fechado['Economia/ano'])}**  \nEconomia Mês: **{brl(c1_fechado['Economia/ano'] / 12)}**")
    r3.warning(f"**Cenário 2: até 2 pedidos/mês**  \nMédia mês: **{c2_fechado['Pedidos/mês']:.0f} pedidos | {brl(c2_fechado['Custo/mês'])}**  \nMédia Ano: **{c2_fechado['Pedidos/ano']:.0f} pedidos | {brl(c2_fechado['Custo/ano'])}**  \nEconomia Ano: **{brl(c2_fechado['Economia/ano'])}**  \nEconomia Mês: **{brl(c2_fechado['Economia/ano'] / 12)}**")

with aba2:
    st.subheader("Comparativo financeiro")
    ex = cenarios.copy()
    st.dataframe(ex.style.format({
        "Pedidos/mês":lambda v:f"{v:,.0f}".replace(",","."),
        "Custo/mês":brl,
        "Pedidos/ano":lambda v:f"{v:,.0f}".replace(",","."),
        "Custo/ano":brl,
        "Economia/ano":brl,
    }).map(lambda v:"background-color:#EAF7EF;color:#1F7A45;font-weight:bold" if isinstance(v,(int,float)) and v>0 else "",subset=["Economia/ano"]),use_container_width=True,hide_index=True)

    fig2=go.Figure(go.Bar(x=cenarios["Cenário"],y=cenarios["Custo/ano"],text=[brl(v) for v in cenarios["Custo/ano"]],textposition="outside",marker_color=["#2D6CDF","#1F9D55","#F39C12"]))
    fig2.update_layout(title="Custo anual por cenário",yaxis_title="Custo anual (R$)",showlegend=False,margin=dict(l=20,r=20,t=55,b=30))
    st.plotly_chart(fig2,use_container_width=True)

    st.markdown("#### Curva ABC do cenário atual")
    st.caption("Classificação pela coluna BG: A = até 10 SIMCARDs por pedido | B = 11 a 50 | C = 51 a 200 | D = acima de 200.")
    total_curva = pd.DataFrame([{
        "Curva ABC": "TOTAL",
        "Quantidade de lojas": int(curva_abc["Quantidade de lojas"].sum()),
        "Quantidade de pedidos": int(curva_abc["Quantidade de pedidos"].sum()),
        "% do volume": float(curva_abc["% do volume"].sum()),
    }])
    curva_exibicao = pd.concat([curva_abc, total_curva], ignore_index=True)
    col_curva_tabela, col_curva_grafico = st.columns([1.05, 1])
    with col_curva_tabela:
        st.dataframe(
            curva_exibicao.style.format({
                "Quantidade de lojas": lambda v: f"{v:,.0f}".replace(",", "."),
                "Quantidade de pedidos": lambda v: f"{v:,.0f}".replace(",", "."),
                "% do volume": lambda v: f"{v:.2%}".replace(".", ","),
            }),
            use_container_width=True,
            hide_index=True,
        )
    with col_curva_grafico:
        fig_curva = go.Figure(go.Bar(
            x=curva_abc["Curva ABC"],
            y=curva_abc["Quantidade de pedidos"],
            text=curva_abc["Quantidade de pedidos"],
            textposition="outside",
            marker_color=["#2D6CDF", "#1F9D55", "#F39C12", "#D64545"],
            customdata=curva_abc[["Quantidade de lojas", "% do volume"]],
            hovertemplate="Curva %{x}<br>Pedidos: %{y}<br>Lojas: %{customdata[0]}<br>% do volume: %{customdata[1]:.0%}<extra></extra>",
        ))
        fig_curva.update_layout(
            title="Pedidos por Curva ABC",
            xaxis_title="Curva ABC",
            yaxis_title="Quantidade de pedidos",
            showlegend=False,
            clickmode="event+select",
            margin=dict(l=20, r=20, t=55, b=30),
        )
        evento_curva = st.plotly_chart(
            fig_curva,
            use_container_width=True,
            key="grafico_curva_abc",
            on_select="rerun",
            selection_mode="points",
        )
    pontos = evento_curva.selection.points if evento_curva and evento_curva.selection else []
    curva_selecionada = str(pontos[0].get("x")) if pontos else None
    if curva_selecionada in ["A", "B", "C", "D"]:
        lojas_curva = detalhe_curva[detalhe_curva["Curva ABC"].eq(curva_selecionada)].copy()
        lojas_resumo = lojas_curva.groupby(["perfil", "loja"], as_index=False).agg(
            **{
                "Quantidade de pedidos": ("pedido", "count"),
                "Menor quantidade BG": ("qtd_simcards_pedido", "min"),
                "Maior quantidade BG": ("qtd_simcards_pedido", "max"),
            }
        ).sort_values(["Quantidade de pedidos", "loja"], ascending=[False, True])
        st.markdown(f"##### Lojas da Curva {curva_selecionada}")
        st.caption(f"Total exibido: {len(lojas_resumo)} loja(s). Clique em outra barra para atualizar.")
        st.dataframe(lojas_resumo, use_container_width=True, hide_index=True)
    else:
        st.info("Clique em uma barra A, B, C ou D para visualizar abaixo as lojas correspondentes.")

with aba3:
    visao = st.radio("Escolha a visão do detalhamento", ["Por pedido", "Por material"], horizontal=True)
    pedidos_filtrados = set(bf["pedido"].astype(str))

    if visao == "Por pedido":
        st.subheader("Pedidos únicos considerados")
        qtd_pedido = resumo_quantidade_por_pedido(detalhe_material_base)
        det = bf[["perfil", "loja", "pedido", "tipo_ov", "mes_num", "ano", "custo_faturamento"]].copy()
        det = det.merge(qtd_pedido, on="pedido", how="left")
        det["Quantidade Prevista"] = det["Quantidade Prevista"].fillna(0)
        det["MÊS/ANO"] = det["mes_num"].astype("Int64").astype(str).str.zfill(2) + "/" + det["ano"].astype("Int64").astype(str)
        det = det[["perfil", "loja", "pedido", "tipo_ov", "Quantidade Prevista", "MÊS/ANO", "custo_faturamento"]]
        det = det.rename(columns={
            "tipo_ov": "TIPO DA OV",
            "custo_faturamento": "Custo de faturamento",
        }).sort_values(["perfil", "loja", "pedido"])
        det_exibicao = det.copy()
        det_exibicao["Quantidade Prevista"] = det_exibicao["Quantidade Prevista"].map(
            lambda v: f"{float(v):,.0f}".replace(",", ".")
        )
        det_exibicao["Custo de faturamento"] = det_exibicao["Custo de faturamento"].map(brl)
        st.dataframe(det_exibicao, use_container_width=True, hide_index=True)
        st.download_button("Baixar visão por pedido em Excel", excel_tabela(det, "Por Pedido"),
                           "Detalhamento_por_Pedido.xlsx", use_container_width=True)
    else:
        st.subheader("Detalhamento por material")
        detalhe_material = detalhe_material_base[
            detalhe_material_base["Documento SD"].isin(pedidos_filtrados)
        ].sort_values(["Documento SD", "Item (SD)"]).reset_index(drop=True)
        detalhe_material_exibicao = detalhe_material.copy()
        detalhe_material_exibicao["Quantidade prevista"] = detalhe_material_exibicao["Quantidade prevista"].map(
            lambda v: f"{float(v):,.0f}".replace(",", ".")
        )
        st.dataframe(detalhe_material_exibicao, use_container_width=True, hide_index=True)
        st.download_button("Baixar visão por material em Excel",
                           excel_tabela(detalhe_material, "Por Material"),
                           "Detalhamento_por_Material.xlsx", use_container_width=True)

with aba4:
    st.subheader("Análise Gerencial para Tomada de Decisão")
    st.caption("Leitura automática da seleção atual. O conteúdo é atualizado sempre que os filtros forem alterados.")

    fmt_int = lambda valor: f"{float(valor):,.0f}".replace(",", ".")
    fmt_pct = lambda valor: f"{float(valor):.1%}".replace(".", ",")

    total_pedidos = int(resumo["Contagem Número do pedido"].sum())
    total_lojas = int(bf["loja"].nunique())
    custo_anual_atual = float(a["Custo/ano"])
    economia_c1 = float(c1["Economia/ano"])
    economia_c2 = float(c2["Economia/ano"])
    reducao_c1 = economia_c1 / custo_anual_atual if custo_anual_atual else 0.0
    reducao_c2 = economia_c2 / custo_anual_atual if custo_anual_atual else 0.0

    resumo_com_movimento = resumo[resumo["Contagem Número do pedido"] > 0].copy()
    if not resumo_com_movimento.empty:
        linha_pico = resumo_com_movimento.loc[
            resumo_com_movimento["Contagem Número do pedido"].idxmax()
        ]
        mes_pico = str(linha_pico["Mês do pedido"])
        pedidos_pico = int(linha_pico["Contagem Número do pedido"])
    else:
        mes_pico = "N/A"
        pedidos_pico = 0

    if not curva_abc.empty and curva_abc["Quantidade de pedidos"].sum() > 0:
        linha_curva = curva_abc.loc[curva_abc["Quantidade de pedidos"].idxmax()]
        curva_principal = str(linha_curva["Curva ABC"])
        participacao_curva = float(linha_curva["% do volume"])
    else:
        curva_principal = "N/A"
        participacao_curva = 0.0

    if reducao_c1 >= 0.20:
        nivel_oportunidade = "Alta"
        mensagem_oportunidade = "O cenário de até 1 pedido por loja/mês apresenta redução potencial relevante frente ao custo anual atual."
        exibidor_status = st.success
    elif reducao_c1 >= 0.08:
        nivel_oportunidade = "Moderada"
        mensagem_oportunidade = "O cenário de até 1 pedido por loja/mês apresenta redução potencial moderada frente ao custo anual atual."
        exibidor_status = st.warning
    else:
        nivel_oportunidade = "Baixa"
        mensagem_oportunidade = "O cenário de até 1 pedido por loja/mês apresenta redução potencial limitada frente ao custo anual atual."
        exibidor_status = st.info

    g1, g2, g3 = st.columns(3)
    g1.metric("Oportunidade estimada", nivel_oportunidade)
    g2.metric("Economia potencial máxima", brl(economia_c1), f"{fmt_pct(reducao_c1)} do custo atual")
    g3.metric("Curva com mais pedidos", f"Curva {curva_principal}", f"{fmt_pct(participacao_curva)} dos pedidos")

    exibidor_status(f"**Sinal gerencial:** {mensagem_oportunidade}")

    st.markdown("### Resumo executivo")
    st.markdown(
        f"""
A seleção atual reúne **{fmt_int(total_pedidos)} pedidos** em **{fmt_int(total_lojas)} lojas**, com média de
**{a['Pedidos/mês']:.0f} pedidos por mês** e custo anual projetado de **{brl(custo_anual_atual)}**.
O maior volume mensal ocorreu em **{mes_pico}**, com **{fmt_int(pedidos_pico)} pedidos**.
A **Curva {curva_principal}** concentra **{fmt_pct(participacao_curva)}** dos pedidos considerados.
"""
    )

    st.markdown("### Comparação para decisão")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.success(
            f"""
**Cenário 1: até 1 pedido por loja/mês**

- Custo anual estimado: **{brl(c1['Custo/ano'])}**
- Economia anual potencial: **{brl(economia_c1)}**
- Redução estimada: **{fmt_pct(reducao_c1)}**
- Pedidos anuais projetados: **{fmt_int(c1['Pedidos/ano'])}**
"""
        )
    with col_c2:
        st.warning(
            f"""
**Cenário 2: até 2 pedidos por loja/mês**

- Custo anual estimado: **{brl(c2['Custo/ano'])}**
- Economia anual potencial: **{brl(economia_c2)}**
- Redução estimada: **{fmt_pct(reducao_c2)}**
- Pedidos anuais projetados: **{fmt_int(c2['Pedidos/ano'])}**
"""
        )

    st.markdown("### Recomendação gerencial")
    if economia_c1 > economia_c2 and economia_c1 > 0:
        recomendacao = (
            "Priorizar a avaliação do Cenário 1, pois ele apresenta o maior potencial de economia. "
            "Antes da implantação, validar cobertura de estoque, capacidade de armazenagem e risco de ruptura por loja."
        )
    elif economia_c2 > 0:
        recomendacao = (
            "Considerar o Cenário 2 como alternativa de menor restrição operacional. "
            "Antes da implantação, validar cobertura de estoque e risco de ruptura nas lojas impactadas."
        )
    else:
        recomendacao = (
            "A seleção atual não apresenta economia projetada relevante nos cenários calculados. "
            "Recomenda-se revisar os filtros ou manter a frequência atual."
        )
    st.markdown(f"**Direcionamento sugerido:** {recomendacao}")

    st.markdown("### Plano de ação sugerido")
    st.markdown(
        f"""
1. **Priorizar a Curva {curva_principal}:** analisar primeiro as lojas que concentram a maior parcela dos pedidos.
2. **Revisar recorrência:** identificar lojas acima da frequência proposta e verificar as causas dos pedidos adicionais.
3. **Validar o risco operacional:** conferir estoque disponível, consumo médio e capacidade de armazenamento antes de reduzir a frequência.
4. **Executar um piloto:** aplicar o cenário escolhido a um grupo controlado de lojas e acompanhar custo, nível de serviço e possíveis rupturas.
5. **Reavaliar a decisão:** comparar os resultados do piloto com a economia estimada no indicador.
"""
    )

    st.info(
        "A análise é uma recomendação baseada nos registros e filtros atuais. "
        "Os valores são projeções e devem ser validados com as áreas responsáveis antes da decisão final."
    )

st.download_button("Baixar simulação em Excel",data=excel_saida(resumo,cenarios,curva_abc,detalhe_curva,bf),file_name="Simulacao_Pedidos_SIMCARD_2027.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
st.caption("Cenário 1 limita a frequência a até 1 pedido por loja/mês. Cenário 2 limita a frequência a até 2 pedidos por loja/mês. Os cenários usam o valor unitário da coluna DA de cada loja e funcionam como limites máximos.")
