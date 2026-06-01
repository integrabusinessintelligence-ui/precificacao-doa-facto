import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date, datetime
import uuid

BASE_DIR = Path(__file__).parent
DADOS_DIR = BASE_DIR / "dados"
ASSETS_DIR = BASE_DIR / "assets"

ELO_AZUL = ASSETS_DIR / "elo_azul.svg"

st.set_page_config(
    page_title="Precificação da DOA | Facto",
    page_icon=str(ELO_AZUL),
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# CORES FACTO
# =========================

VERMELHO = "#D34022"
AZUL = "#163246"
OFF_WHITE = "#EAE8DD"
BRANCO = "#FFFFFF"

# =========================
# CAMINHOS
# =========================

BASE_DIR = Path(__file__).parent
DADOS_DIR = BASE_DIR / "dados"
ASSETS_DIR = BASE_DIR / "assets"

DADOS_DIR.mkdir(exist_ok=True)
ASSETS_DIR.mkdir(exist_ok=True)

ARQ_PROJETOS = DADOS_DIR / "projetos.xlsx"
ARQ_PARF = DADOS_DIR / "parf_itens.xlsx"
ARQ_ICP_GIRO = DADOS_DIR / "icp_giro.xlsx"
ARQ_DOA = DADOS_DIR / "doa_resultados.xlsx"
ARQ_AUXILIARES = DADOS_DIR / "auxiliares.xlsx"
ARQ_CUSTOS_INDICADORES = DADOS_DIR / "custos_indicadores.xlsx"
ARQ_MODELO_PROPOSTA = DADOS_DIR / "modelo_proposta.docx"
SAIDAS_DIR = BASE_DIR / "saidas"
SAIDAS_DIR.mkdir(exist_ok=True)

LOGO_FACTO = ASSETS_DIR / "logo_facto.svg"

# =========================
# CSS
# =========================

st.markdown(
    f"""
    <style>

    .stApp {{
        background-color: white;
        color: {AZUL};
    }}

    .block-container {{
        background-color: white;
    }}

    p, span, label, div {{
        color: {AZUL};
    }}

    h1, h2, h3 {{
        color: {AZUL};
    }}

    section[data-testid="stSidebar"] {{
        background-color: {AZUL};
    }}

    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {{
        color: white !important;
    }}

    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] textarea,
    section[data-testid="stSidebar"] select,
    section[data-testid="stSidebar"] div[data-baseweb="select"] *,
    section[data-testid="stSidebar"] div[data-baseweb="input"] * {{
        color: {AZUL} !important;
    }}

    section[data-testid="stSidebar"] button {{
        color: {AZUL} !important;
        background-color: {OFF_WHITE} !important;
        border: 1px solid {OFF_WHITE} !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
    }}

    section[data-testid="stSidebar"] button:hover {{
        color: white !important;
        background-color: {VERMELHO} !important;
        border: 1px solid {VERMELHO} !important;
    }}

    .facto-title {{
        color: white;
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 4px;
    }}

    .facto-subtitle {{
        color: {OFF_WHITE};
        font-size: 16px;
    }}

    .facto-badge {{
        display: inline-block;
        background-color: {VERMELHO};
        color: white;
        padding: 6px 14px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
        margin-bottom: 10px;
    }}

    div[data-testid="stForm"] {{
        background-color: white;
        padding: 26px;
        border-radius: 22px;
        border: 1px solid #ddd;
        box-shadow: 0 6px 18px rgba(22, 50, 70, 0.10);
    }}

    .metric-card {{
        background: white;
        padding: 22px;
        border-radius: 18px;
        border: 1px solid #ddd;
        box-shadow: 0 4px 14px rgba(22, 50, 70, 0.08);
    }}

    .metric-label {{
        color: #5f6f7a;
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
    }}

    .metric-value {{
        color: {AZUL};
        font-size: 26px;
        font-weight: 800;
    }}

    </style>
    """,
    unsafe_allow_html=True
)

# =========================
# FUNÇÕES BASE
# =========================

def moeda(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def carregar_excel(path: Path, colunas: list) -> pd.DataFrame:
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame(columns=colunas)


def salvar_excel(df: pd.DataFrame, path: Path):
    df.to_excel(path, index=False)


def gerar_id():
    return str(uuid.uuid4())[:8]


def agora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def colunas_projetos():
    return [
        "id_projeto", "versao", "status",
        "nome_projeto", "coordenador", "email_coordenador",
        "instituicao_executora", "campus", "financiador",
        "tipo_financiador", "tipo_instrumento",
        "data_inicio", "data_fim", "prazo_meses",
        "valor_aprovado", "objeto_resumido", "area_tematica",
        "possui_bolsas", "possui_compras", "possui_concurso_ps",
        "possui_clt_rpa_pj", "possui_importacao",
        "responsavel_facto", "data_criacao", "data_atualizacao"
    ]


def colunas_parf():
    return [
        "id_item",
        "id_projeto", "versao", "grupo", "item",
        "modalidade", "categoria", "descricao",
        "quantidade", "meses", "valor_unitario",
        "percentual_adicional", "valor_adicional",
        "modalidade_contratacao", "total"
    ]


def colunas_icp_giro():
    return [
        "id_projeto", "versao", "tipo", "indicador",
        "nota", "justificativa"
    ]


def colunas_doa():
    return [
        "id_projeto", "versao", "componente", "setor", "atividade",
        "quantidade_operacional", "tempo_base_h", "horas_estimadas",
        "custo_hora_pessoal", "custo_hora_adm", "custo_hora_total",
        "fator_icp", "fator_giro", "fator_complexidade",
        "valor_calculado", "valor_ajustado",
        "fator_limitador", "data_calculo"
    ]


def carregar_projetos():
    return carregar_excel(ARQ_PROJETOS, colunas_projetos())


def carregar_parf():
    df = carregar_excel(ARQ_PARF, colunas_parf())

    # Compatibilidade com arquivos antigos sem id_item
    if "id_item" not in df.columns:
        df.insert(0, "id_item", [gerar_id() for _ in range(len(df))])
        if not df.empty:
            salvar_excel(df, ARQ_PARF)

    # Garante todas as colunas esperadas
    for coluna in colunas_parf():
        if coluna not in df.columns:
            df[coluna] = None

    return df[colunas_parf()]


def carregar_icp_giro():
    return carregar_excel(ARQ_ICP_GIRO, colunas_icp_giro())


def carregar_doa():
    return carregar_excel(ARQ_DOA, colunas_doa())


def calcular_prazo_meses(inicio, fim):
    if not inicio or not fim:
        return 0
    return max(1, (fim.year - inicio.year) * 12 + (fim.month - inicio.month) + 1)



def carregar_auxiliares():
    """Carrega a planilha auxiliares.xlsx, padronizando nomes de abas e colunas."""
    aux = {}

    if not ARQ_AUXILIARES.exists():
        return aux

    try:
        xls = pd.ExcelFile(ARQ_AUXILIARES)
        for aba in xls.sheet_names:
            df = pd.read_excel(ARQ_AUXILIARES, sheet_name=aba)
            df.columns = df.columns.astype(str).str.strip()
            aux[aba.strip()] = df
    except Exception:
        pass

    return aux


def grupos_parf():
    return [
        "Bolsas",
        "Celetistas",
        "RPA",
        "Material de consumo",
        "Material permanente",
        "Serviço PJ",
        "Importação",
        "Diárias",
        "Passagens",
        "Ressarcimento",
        "Contrapartida",
        "Prospecção"
    ]


def grupos_contratacao():
    return ["Material de consumo", "Material permanente", "Serviço PJ"]


def limpar_lista(valores):
    proibidos = [
        "nan", "none", "subelementos de despesa", "equipamentos e materiais permanentes",
        "material de consumo", "serviços de terceiros", "tipo de despesa"
    ]
    lista = []
    for valor in valores:
        if pd.isna(valor):
            continue
        texto = str(valor).strip()
        if not texto:
            continue
        if texto.lower() in proibidos:
            continue
        if texto not in lista:
            lista.append(texto)
    return lista


def obter_opcoes_bolsas(aux):
    df = aux.get("CLASSIFICAÇÃO DE BOLSAS", pd.DataFrame()).copy()
    if df.empty:
        return df
    df.columns = df.columns.astype(str).str.strip()
    return df


def obter_valor_bolsa(df_bolsas, modalidade, categoria):
    if df_bolsas.empty or "Valor\xa0R$" not in df_bolsas.columns:
        return 0.0

    filtro = pd.Series([True] * len(df_bolsas))
    if "Modalidade" in df_bolsas.columns:
        filtro = filtro & (df_bolsas["Modalidade"].astype(str) == str(modalidade))
    if "Categoria" in df_bolsas.columns:
        filtro = filtro & (df_bolsas["Categoria"].astype(str) == str(categoria))

    valores = pd.to_numeric(df_bolsas.loc[filtro, "Valor\xa0R$"], errors="coerce").dropna()
    if valores.empty:
        return 0.0
    return float(valores.iloc[0])


def obter_lista_despesas(aux, grupo):
    df = aux.get("CLASSIFICAÇÃO DE DESPESAS", pd.DataFrame()).copy()
    if df.empty:
        return []

    # Estrutura da planilha: col. 1 permanente, col. 2 consumo, col. 3 serviço PJ.
    mapa_colunas = {
        "Material permanente": 1,
        "Material de consumo": 2,
        "Serviço PJ": 3,
    }

    pos = mapa_colunas.get(grupo)
    if pos is None or pos >= len(df.columns):
        return []

    return limpar_lista(df.iloc[:, pos].tolist())


def obter_lista_diarias(aux):
    df = aux.get("CLASSIFICAÇÃO E VALORES DE DIÁR", pd.DataFrame()).copy()
    if df.empty or len(df.columns) < 3:
        return pd.DataFrame(columns=["tipo", "valor"])

    base = pd.DataFrame({
        "tipo": df.iloc[:, 1],
        "valor": pd.to_numeric(df.iloc[:, 2], errors="coerce")
    })
    base = base.dropna(subset=["tipo", "valor"])
    base["tipo"] = base["tipo"].astype(str).str.strip()
    return base


def obter_lista_passagens(aux):
    df = aux.get("CLASSIFICAÇÃO PASSAGENS", pd.DataFrame()).copy()
    if df.empty:
        # Tolerância para aba com espaço final no nome, caso venha sem strip em algum ambiente.
        df = aux.get("CLASSIFICAÇÃO PASSAGENS ", pd.DataFrame()).copy()

    if df.empty:
        return []

    primeira_coluna = df.columns[0]
    return limpar_lista(df[primeira_coluna].tolist())

def salvar_ou_atualizar_projeto(dados):
    df = carregar_projetos()

    for campo in ["data_inicio", "data_fim", "data_criacao", "data_atualizacao"]:
        if campo in dados and dados[campo] is not None:
            dados[campo] = str(dados[campo])

    if not df.empty:
        filtro = (
            (df["id_projeto"].astype(str) == str(dados["id_projeto"])) &
            (df["versao"].astype(int) == int(dados["versao"]))
        )
    else:
        filtro = pd.Series([], dtype=bool)

    if filtro.any():
        idx = df.index[filtro][0]
        for k, v in dados.items():
            if k in df.columns:
                df.at[idx, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)

    salvar_excel(df, ARQ_PROJETOS)


def obter_projeto_atual():
    id_projeto = st.session_state.get("id_projeto")
    versao = st.session_state.get("versao")

    if not id_projeto or not versao:
        return None

    df = carregar_projetos()

    if df.empty:
        return None

    filtro = (
        (df["id_projeto"].astype(str) == str(id_projeto)) &
        (df["versao"].astype(int) == int(versao))
    )

    if filtro.any():
        return df.loc[filtro].iloc[0].to_dict()

    return None


def mostrar_card_proposta():
    id_projeto = st.session_state.get("id_projeto")
    versao = st.session_state.get("versao")

    if not id_projeto or not versao:
        return

    projeto = obter_projeto_atual() or {}

    nome = projeto.get("nome_projeto", "Proposta ainda sem nome")
    status = projeto.get("status", st.session_state.get("status", "Rascunho"))
    valor = projeto.get("valor_aprovado", 0)

    with st.container(border=True):
        st.caption("PROPOSTA EM EDIÇÃO")
        st.markdown(f"### {nome}")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.markdown(f"**ID**  \n{id_projeto}")
        with col2:
            st.markdown(f"**Versão**  \n{versao}")
        with col3:
            st.markdown(f"**Status**  \n{status}")
        with col4:
            st.markdown(f"**Valor previsto**  \n{moeda(valor)}")


def atualizar_item_parf(id_item, dados_atualizados):
    df = carregar_parf()

    if df.empty or "id_item" not in df.columns:
        return False

    filtro = df["id_item"].astype(str) == str(id_item)

    if not filtro.any():
        return False

    idx = df.index[filtro][0]

    for k, v in dados_atualizados.items():
        if k in df.columns:
            df.at[idx, k] = v

    salvar_excel(df, ARQ_PARF)
    return True


def excluir_item_parf(id_item):
    df = carregar_parf()

    if df.empty or "id_item" not in df.columns:
        return False

    filtro = df["id_item"].astype(str) == str(id_item)

    if not filtro.any():
        return False

    df = df.loc[~filtro].copy()
    salvar_excel(df, ARQ_PARF)
    return True


def criar_nova_versao(id_projeto, versao_origem):
    nova_versao = int(versao_origem) + 1

    dfp = carregar_projetos()

    origem = dfp[
        (dfp["id_projeto"].astype(str) == str(id_projeto)) &
        (dfp["versao"].astype(int) == int(versao_origem))
    ]

    if origem.empty:
        return None

    novo = origem.iloc[0].copy()
    novo["versao"] = nova_versao
    novo["status"] = "Renegociada"
    novo["data_criacao"] = agora()
    novo["data_atualizacao"] = agora()

    dfp = pd.concat([dfp, pd.DataFrame([novo])], ignore_index=True)
    salvar_excel(dfp, ARQ_PROJETOS)

    for arquivo, colunas in [
        (ARQ_PARF, colunas_parf()),
        (ARQ_ICP_GIRO, colunas_icp_giro()),
        (ARQ_DOA, colunas_doa())
    ]:
        df = carregar_excel(arquivo, colunas)

        if df.empty:
            continue

        dados_origem = df[
            (df["id_projeto"].astype(str) == str(id_projeto)) &
            (df["versao"].astype(int) == int(versao_origem))
        ].copy()

        if not dados_origem.empty:
            dados_origem["versao"] = nova_versao
            df = pd.concat([df, dados_origem], ignore_index=True)
            salvar_excel(df, arquivo)

    return nova_versao



# =========================
# FUNÇÕES DOA
# =========================

def carregar_custos_indicadores():
    """Carrega custos_indicadores.xlsx, que deve estar na pasta dados."""
    bases = {}

    if not ARQ_CUSTOS_INDICADORES.exists():
        return bases

    try:
        xls = pd.ExcelFile(ARQ_CUSTOS_INDICADORES)
        for aba in xls.sheet_names:
            df = pd.read_excel(ARQ_CUSTOS_INDICADORES, sheet_name=aba)
            df.columns = df.columns.astype(str).str.strip()
            bases[aba.strip()] = df
    except Exception:
        pass

    return bases


def normalizar_texto(valor):
    return str(valor or "").strip().lower()


def buscar_atividade(tempo_df, atividade):
    """Busca uma atividade na aba tempo_pessoas pelo nome exato; se não achar, tenta por contém."""
    if tempo_df.empty or "atividade" not in tempo_df.columns:
        return None

    alvo = normalizar_texto(atividade)
    temp = tempo_df.copy()
    temp["atividade_norm"] = temp["atividade"].apply(normalizar_texto)

    exata = temp[temp["atividade_norm"] == alvo]
    if not exata.empty:
        return exata.iloc[0].to_dict()

    contem = temp[temp["atividade_norm"].str.contains(alvo, na=False, regex=False)]
    if not contem.empty:
        return contem.iloc[0].to_dict()

    return None


def componente_por_setor(setor):
    mapa = {
        "Compras": "Gerenciamento de compras e contratações",
        "SelectID": "Gerenciamento de Bolsas/Auxílios e/ou PS e Concursos",
        "Projetos": "Gerenciamento de Projetos",
        "Financeiro": "Gerenciamento administrativo e financeiro",
        "Gente e Gestão": "Gestão institucional e governança",
    }
    return mapa.get(str(setor), str(setor))


def fator_por_nota(nota):
    """Escala 1 a 5: 1=0,90; 2=1,00; 3=1,10; 4=1,20; 5=1,30."""
    try:
        nota = float(nota)
    except Exception:
        nota = 3.0
    nota = max(1.0, min(5.0, nota))
    return 0.8 + (nota * 0.1)


def obter_fatores_complexidade():
    """Calcula fatores de ICP e Giro a partir das notas salvas para a proposta atual."""
    df = carregar_icp_giro()

    if df.empty or not st.session_state.get("id_projeto"):
        return 1.10, 1.10, 1.10, 3.0, 3.0

    filtro = (
        (df["id_projeto"].astype(str) == str(st.session_state.get("id_projeto"))) &
        (df["versao"].astype(int) == int(st.session_state.get("versao")))
    )
    df_atual = df[filtro].copy()

    if df_atual.empty:
        return 1.10, 1.10, 1.10, 3.0, 3.0

    df_atual["nota"] = pd.to_numeric(df_atual["nota"], errors="coerce")

    media_icp = df_atual.loc[df_atual["tipo"].astype(str).str.upper() == "ICP", "nota"].mean()
    media_giro = df_atual.loc[df_atual["tipo"].astype(str).str.upper() == "GIRO", "nota"].mean()

    if pd.isna(media_icp):
        media_icp = 3.0
    if pd.isna(media_giro):
        media_giro = 3.0

    fator_icp = fator_por_nota(media_icp)
    fator_giro = fator_por_nota(media_giro)

    # Peso metodológico inicial: ICP 40% e Giro 60%.
    fator_complexidade = (fator_icp * 0.40) + (fator_giro * 0.60)

    return fator_icp, fator_giro, fator_complexidade, float(media_icp), float(media_giro)


def calcular_custo_hora_setor(bases):
    pessoal = bases.get("custos_pessoal", pd.DataFrame()).copy()
    adm = bases.get("custos_adm", pd.DataFrame()).copy()

    if pessoal.empty:
        return pd.DataFrame(columns=["setor", "custo_hora_pessoal", "custo_hora_adm", "custo_hora_total"])

    pessoal.columns = pessoal.columns.astype(str).str.strip()
    pessoal["ch_mensal"] = pd.to_numeric(pessoal.get("ch_mensal"), errors="coerce").fillna(0)
    pessoal["custo_pessoa_mensal"] = pd.to_numeric(pessoal.get("custo_pessoa_mensal"), errors="coerce").fillna(0)

    setor = pessoal.groupby("setor", as_index=False).agg({
        "ch_mensal": "sum",
        "custo_pessoa_mensal": "sum"
    })

    setor["custo_hora_pessoal"] = setor.apply(
        lambda r: r["custo_pessoa_mensal"] / r["ch_mensal"] if r["ch_mensal"] else 0,
        axis=1
    )

    total_horas = setor["ch_mensal"].sum()

    if not adm.empty and "valor" in adm.columns and total_horas:
        adm["valor"] = pd.to_numeric(adm["valor"], errors="coerce").fillna(0)
        custo_hora_adm = adm["valor"].sum() / total_horas
    else:
        custo_hora_adm = 0.0

    setor["custo_hora_adm"] = custo_hora_adm
    setor["custo_hora_total"] = setor["custo_hora_pessoal"] + setor["custo_hora_adm"]

    return setor[["setor", "custo_hora_pessoal", "custo_hora_adm", "custo_hora_total"]]


def adicionar_linha_doa(linhas, tempo_df, custo_hora_df, atividade, quantidade, fator_icp, fator_giro, fator_complexidade, componente_forcado=None):
    if quantidade is None:
        quantidade = 0
    try:
        quantidade = float(quantidade)
    except Exception:
        quantidade = 0
    if quantidade <= 0:
        return

    info = buscar_atividade(tempo_df, atividade)
    if not info:
        return

    setor = info.get("setor")
    tempo_base = float(info.get("tempo_base_h") or 0)
    horas_estimadas = quantidade * tempo_base

    custo_setor = custo_hora_df[custo_hora_df["setor"].astype(str) == str(setor)]

    if custo_setor.empty:
        custo_hora_pessoal = 0.0
        custo_hora_adm = 0.0
        custo_hora_total = 0.0
    else:
        custo_hora_pessoal = float(custo_setor["custo_hora_pessoal"].iloc[0] or 0)
        custo_hora_adm = float(custo_setor["custo_hora_adm"].iloc[0] or 0)
        custo_hora_total = float(custo_setor["custo_hora_total"].iloc[0] or 0)

    valor_calculado = horas_estimadas * custo_hora_total * fator_complexidade

    linhas.append({
        "componente": componente_forcado or componente_por_setor(setor),
        "setor": setor,
        "atividade": info.get("atividade"),
        "quantidade_operacional": quantidade,
        "tempo_base_h": tempo_base,
        "horas_estimadas": horas_estimadas,
        "custo_hora_pessoal": custo_hora_pessoal,
        "custo_hora_adm": custo_hora_adm,
        "custo_hora_total": custo_hora_total,
        "fator_icp": fator_icp,
        "fator_giro": fator_giro,
        "fator_complexidade": fator_complexidade,
        "valor_calculado": valor_calculado,
    })


def gerar_memoria_doa(df_parf_atual, bases):
    tempo_df = bases.get("tempo_pessoas", pd.DataFrame()).copy()
    if not tempo_df.empty:
        tempo_df.columns = tempo_df.columns.astype(str).str.strip()

    custo_hora_df = calcular_custo_hora_setor(bases)
    fator_icp, fator_giro, fator_complexidade, media_icp, media_giro = obter_fatores_complexidade()

    linhas = []

    for _, row in df_parf_atual.iterrows():
        grupo = str(row.get("grupo", ""))
        modalidade_contratacao = str(row.get("modalidade_contratacao", ""))

        qtd = float(pd.to_numeric(row.get("quantidade", 0), errors="coerce") or 0)
        meses = float(pd.to_numeric(row.get("meses", 0), errors="coerce") or 0)
        if meses <= 0:
            meses = 1

        if grupo == "Bolsas":
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Implementação de bolsas", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Gestão de contrato de bolsas", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Pagamento de bolsa", qtd * meses, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Gestão de pagamentos de bolsas", meses, fator_icp, fator_giro, fator_complexidade)

        elif grupo == "Celetistas":
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Seleção de CLT", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Contratação CLT", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Gestão mensal de contratos (por pessoa)", qtd * meses, fator_icp, fator_giro, fator_complexidade)

        elif grupo == "RPA":
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Seleção de RPA (Projetos)", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Gestão de Contratos RPA (Projetos)", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Pagamento de RPA", qtd, fator_icp, fator_giro, fator_complexidade)

        elif grupo in ["Material de consumo", "Material permanente", "Serviço PJ"]:
            if "Inexigibilidade" in modalidade_contratacao:
                atividade_compra = "Inexigibilidade"
            elif "Licitação" in modalidade_contratacao or "Dispensa" in modalidade_contratacao:
                atividade_compra = "Licitação/Dispensa"
            else:
                atividade_compra = "Compra direta"

            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, atividade_compra, qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Pagamento de fornecedor", qtd, fator_icp, fator_giro, fator_complexidade)

            if grupo == "Material permanente":
                adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Gestão de patrimônio", qtd, fator_icp, fator_giro, fator_complexidade)
            if grupo == "Serviço PJ":
                adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Gestão de contratos de compras e contratações", qtd, fator_icp, fator_giro, fator_complexidade)

        elif grupo == "Importação":
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Importação", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Pagamento de fornecedor", qtd, fator_icp, fator_giro, fator_complexidade)

        elif grupo == "Passagens":
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Número de passagens", qtd, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Pagamento de fornecedor", qtd, fator_icp, fator_giro, fator_complexidade)

        elif grupo == "Diárias":
            quantidade_operacional = qtd * meses
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Diárias (Projetos)", quantidade_operacional, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Pagamento de diária", quantidade_operacional, fator_icp, fator_giro, fator_complexidade)

        elif grupo in ["Ressarcimento", "Contrapartida"]:
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Gerenciamento orçamentário e financeiro (Projetos)", 1, fator_icp, fator_giro, fator_complexidade)
            adicionar_linha_doa(linhas, tempo_df, custo_hora_df, "Entrada de receita (notas)", 1, fator_icp, fator_giro, fator_complexidade)

        elif grupo == "Prospecção":
            adicionar_linha_doa(
                linhas, tempo_df, custo_hora_df, "Comunicação institucional", 1,
                fator_icp, fator_giro, fator_complexidade,
                componente_forcado="Prospecção e articulação institucional"
            )

    memoria = pd.DataFrame(linhas)
    return memoria, custo_hora_df, media_icp, media_giro, fator_icp, fator_giro, fator_complexidade


def formatar_tabela_moeda(df, colunas_moeda):
    visual = df.copy()
    for coluna in colunas_moeda:
        if coluna in visual.columns:
            visual[coluna] = visual[coluna].apply(moeda)
    return visual




# =========================
# FUNÇÕES PROPOSTA COMERCIAL
# =========================

def nome_seguro_arquivo(texto):
    texto = str(texto or "proposta").strip().lower()
    substituicoes = {
        "á": "a", "à": "a", "ã": "a", "â": "a",
        "é": "e", "ê": "e",
        "í": "i",
        "ó": "o", "õ": "o", "ô": "o",
        "ú": "u",
        "ç": "c"
    }
    for origem, destino in substituicoes.items():
        texto = texto.replace(origem, destino)
    permitido = []
    for c in texto:
        if c.isalnum():
            permitido.append(c)
        elif c in [" ", "-", "_"]:
            permitido.append("_")
    texto = "".join(permitido)
    while "__" in texto:
        texto = texto.replace("__", "_")
    return texto.strip("_") or "proposta"


def adicionar_paragrafo_apos(paragrafo, texto="", estilo=None):
    from docx.text.paragraph import Paragraph
    from docx.oxml import OxmlElement

    novo_p = OxmlElement("w:p")
    paragrafo._p.addnext(novo_p)
    novo_paragrafo = Paragraph(novo_p, paragrafo._parent)
    if estilo:
        novo_paragrafo.style = estilo
    if texto:
        novo_paragrafo.add_run(texto)
    return novo_paragrafo


def inserir_tabela_apos_paragrafo(documento, paragrafo, dados, estilo="Table Grid"):
    """Insere tabela depois de um parágrafo do modelo e devolve a tabela."""
    linhas = len(dados)
    colunas = len(dados[0]) if dados else 1
    tabela = documento.add_table(rows=linhas, cols=colunas)
    tabela.style = estilo

    for i, linha in enumerate(dados):
        for j, valor in enumerate(linha):
            celula = tabela.cell(i, j)
            celula.text = str(valor if valor is not None else "")
            for par in celula.paragraphs:
                for run in par.runs:
                    run.font.size = None
                    if i == 0:
                        run.bold = True

    paragrafo._p.addnext(tabela._tbl)
    return tabela


def localizar_paragrafo_por_texto(documento, texto_busca):
    alvo = str(texto_busca).strip().lower()
    for paragrafo in documento.paragraphs:
        texto = paragrafo.text.strip().lower()
        if alvo in texto:
            return paragrafo
    return None


def obter_parf_atual_df():
    df_parf = carregar_parf()
    if df_parf.empty or not id_atual:
        return pd.DataFrame(columns=colunas_parf())
    return df_parf[
        (df_parf["id_projeto"].astype(str) == str(id_atual)) &
        (df_parf["versao"].astype(int) == int(versao_atual))
    ].copy()


def obter_resumo_parf_proposta(df_parf_atual):
    if df_parf_atual.empty:
        return pd.DataFrame(columns=["Rubrica", "Quantidade", "Valor previsto"])

    temp = df_parf_atual.copy()
    temp["quantidade"] = pd.to_numeric(temp["quantidade"], errors="coerce").fillna(0)
    temp["total"] = pd.to_numeric(temp["total"], errors="coerce").fillna(0)

    resumo = temp.groupby("grupo", as_index=False).agg({
        "quantidade": "sum",
        "total": "sum"
    })
    resumo = resumo.rename(columns={
        "grupo": "Rubrica",
        "quantidade": "Quantidade",
        "total": "Valor previsto"
    })
    return resumo


def obter_resumo_doa_proposta(df_parf_atual):
    """Retorna tabela de DOA agrupada por componente. Usa resultado salvo; se não houver, recalcula."""
    df_doa = carregar_doa()
    if not df_doa.empty:
        filtro = (
            (df_doa["id_projeto"].astype(str) == str(id_atual)) &
            (df_doa["versao"].astype(int) == int(versao_atual))
        )
        df_salvo = df_doa[filtro].copy()
        if not df_salvo.empty:
            df_salvo["valor_calculado"] = pd.to_numeric(df_salvo["valor_calculado"], errors="coerce").fillna(0)
            df_salvo["valor_ajustado"] = pd.to_numeric(df_salvo["valor_ajustado"], errors="coerce").fillna(0)
            resumo = df_salvo.groupby("componente", as_index=False).agg({
                "valor_calculado": "sum",
                "valor_ajustado": "sum"
            })
            resumo = resumo.rename(columns={
                "componente": "Componente da DOA",
                "valor_calculado": "Valor calculado",
                "valor_ajustado": "Valor final"
            })
            return resumo

    # Recalcula quando o usuário ainda não clicou em "Salvar resultado da DOA".
    if df_parf_atual.empty or not ARQ_CUSTOS_INDICADORES.exists():
        return pd.DataFrame(columns=["Componente da DOA", "Valor calculado", "Valor final"])

    bases_custos = carregar_custos_indicadores()
    memoria, *_ = gerar_memoria_doa(df_parf_atual, bases_custos)
    if memoria.empty:
        return pd.DataFrame(columns=["Componente da DOA", "Valor calculado", "Valor final"])

    projeto = obter_projeto_atual() or {}
    valor_projeto = float(projeto.get("valor_aprovado", 0) or 0)
    if valor_projeto <= 0:
        valor_projeto = float(pd.to_numeric(df_parf_atual["total"], errors="coerce").fillna(0).sum())

    resumo = memoria.groupby("componente", as_index=False)["valor_calculado"].sum()
    total_calculado = float(resumo["valor_calculado"].sum())
    limite_legal = valor_projeto * 0.15
    fator_limitador = min(1, limite_legal / total_calculado) if total_calculado > 0 else 1
    resumo["valor_ajustado"] = resumo["valor_calculado"] * fator_limitador

    resumo = resumo.rename(columns={
        "componente": "Componente da DOA",
        "valor_calculado": "Valor calculado",
        "valor_ajustado": "Valor final"
    })
    return resumo


def gerar_proposta_docx():
    try:
        from docx import Document
    except Exception as exc:
        st.error("Biblioteca python-docx não instalada. Rode: pip install python-docx")
        return None

    if not ARQ_MODELO_PROPOSTA.exists():
        st.error("Modelo não encontrado. Salve o arquivo como dados/modelo_proposta.docx.")
        return None

    projeto = obter_projeto_atual() or {}
    df_parf_atual = obter_parf_atual_df()
    resumo_parf = obter_resumo_parf_proposta(df_parf_atual)
    resumo_doa = obter_resumo_doa_proposta(df_parf_atual)

    documento = Document(str(ARQ_MODELO_PROPOSTA))

    # 2. Objeto — quadro com dados do projeto
    par_objeto = localizar_paragrafo_por_texto(documento, "2. Objeto")
    if par_objeto:
        dados_projeto = [
            ["Campo", "Informação"],
            ["Projeto", projeto.get("nome_projeto", "")],
            ["Coordenador", projeto.get("coordenador", "")],
            ["E-mail do coordenador", projeto.get("email_coordenador", "")],
            ["Instituição executora", projeto.get("instituicao_executora", "")],
            ["Campus", projeto.get("campus", "")],
            ["Financiador", projeto.get("financiador", "")],
            ["Tipo de financiador", projeto.get("tipo_financiador", "")],
            ["Tipo de instrumento", projeto.get("tipo_instrumento", "")],
            ["Vigência", f"{projeto.get('data_inicio', '')} a {projeto.get('data_fim', '')}"],
            ["Prazo estimado", f"{projeto.get('prazo_meses', '')} meses"],
            ["Valor aprovado/previsto", moeda(projeto.get("valor_aprovado", 0))],
            ["Área temática", projeto.get("area_tematica", "")],
        ]
        inserir_tabela_apos_paragrafo(documento, par_objeto, dados_projeto)

    # 9. Serviços e Itens Executados — resumo das rubricas do PARF
    par_servicos = localizar_paragrafo_por_texto(documento, "9. Serviços e Itens Executados")
    if par_servicos:
        dados_rubricas = [["Rubrica", "Quantidade", "Valor previsto"]]
        if not resumo_parf.empty:
            for _, row in resumo_parf.iterrows():
                dados_rubricas.append([
                    row.get("Rubrica", ""),
                    f"{float(row.get('Quantidade', 0) or 0):,.0f}".replace(",", "."),
                    moeda(row.get("Valor previsto", 0))
                ])
            dados_rubricas.append([
                "Total",
                "",
                moeda(resumo_parf["Valor previsto"].sum())
            ])
        else:
            dados_rubricas.append(["Sem rubricas cadastradas", "", moeda(0)])
        inserir_tabela_apos_paragrafo(documento, par_servicos, dados_rubricas)

    # 11. Detalhamento dos valores de DOA — tabela final da DOA
    par_doa = localizar_paragrafo_por_texto(documento, "11. Detalhamento dos valores de DOA")
    if par_doa:
        dados_doa = [["Componente da DOA", "Valor calculado", "Valor final"]]
        if not resumo_doa.empty:
            for _, row in resumo_doa.iterrows():
                dados_doa.append([
                    row.get("Componente da DOA", ""),
                    moeda(row.get("Valor calculado", 0)),
                    moeda(row.get("Valor final", 0))
                ])
            dados_doa.append([
                "Total",
                moeda(resumo_doa["Valor calculado"].sum()),
                moeda(resumo_doa["Valor final"].sum())
            ])
        else:
            dados_doa.append(["DOA ainda não calculada", moeda(0), moeda(0)])
        inserir_tabela_apos_paragrafo(documento, par_doa, dados_doa)

    nome_base = nome_seguro_arquivo(projeto.get("nome_projeto", "proposta"))
    saida = SAIDAS_DIR / f"proposta_{nome_base}_v{versao_atual}_{id_atual}.docx"
    documento.save(str(saida))
    return saida

# =========================
# CABEÇALHO
# =========================

col_logo, col_titulo = st.columns([1, 5])

with col_logo:
    if LOGO_FACTO.exists():
        st.image(str(LOGO_FACTO), width=150)
    else:
        st.markdown("### Facto")

with col_titulo:
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, {AZUL} 0%, #244a63 100%);
            padding: 24px 32px;
            border-radius: 0 0 24px 24px;
            margin-bottom: 28px;
        ">
            <div class="facto-badge">FACTO • Plataforma Institucional</div>
            <div class="facto-title">Precificação de Projetos e DOA</div>
            <div class="facto-subtitle">
                Cadastro → PARF → ICP/Giro → DOA → Proposta Comercial
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================
# SIDEBAR
# =========================

st.sidebar.title("Navegação")

acao = st.sidebar.radio(
    "Ação",
    ["Nova proposta", "Abrir proposta existente"]
)

df_projetos = carregar_projetos()

if acao == "Nova proposta":
    criar = st.sidebar.button(
        "➕ Criar nova proposta",
        use_container_width=True,
        type="primary"
    )

    if criar:
        st.session_state["id_projeto"] = gerar_id()
        st.session_state["versao"] = 1
        st.session_state["status"] = "Rascunho"
        st.rerun()

else:
    if df_projetos.empty:
        st.sidebar.info("Nenhuma proposta cadastrada.")
    else:
        df_projetos = df_projetos.copy()
        df_projetos["label"] = (
            df_projetos["nome_projeto"].fillna("Sem nome").astype(str) +
            " | ID " + df_projetos["id_projeto"].astype(str) +
            " | v" + df_projetos["versao"].astype(str) +
            " | " + df_projetos["status"].astype(str)
        )

        selecionado = st.sidebar.selectbox(
            "Selecione uma proposta",
            df_projetos["label"].tolist()
        )

        linha = df_projetos[df_projetos["label"] == selecionado].iloc[0]

        abrir = st.sidebar.button(
            "Abrir proposta",
            use_container_width=True
        )

        if abrir:
            st.session_state["id_projeto"] = str(linha["id_projeto"])
            st.session_state["versao"] = int(linha["versao"])
            st.session_state["status"] = str(linha["status"])
            st.sidebar.success("Proposta carregada.")
            st.rerun()

        nova_versao_btn = st.sidebar.button(
            "Criar nova versão para renegociação",
            use_container_width=True
        )

        if nova_versao_btn:
            nova = criar_nova_versao(str(linha["id_projeto"]), int(linha["versao"]))

            if nova:
                st.session_state["id_projeto"] = str(linha["id_projeto"])
                st.session_state["versao"] = nova
                st.session_state["status"] = "Renegociada"
                st.sidebar.success(f"Versão {nova} criada.")
                st.rerun()

pagina = st.sidebar.radio(
    "Etapas",
    [
        "1. Cadastro do Projeto",
        "2. PARF",
        "3. ICP e Giro",
        "4. DOA",
        "5. Proposta Comercial",
        "6. Histórico"
    ]
)

id_atual = st.session_state.get("id_projeto")
versao_atual = st.session_state.get("versao")

if id_atual:
    st.sidebar.markdown("---")
    st.sidebar.write(f"**Projeto atual:** {id_atual}")
    st.sidebar.write(f"**Versão:** {versao_atual}")

mostrar_card_proposta()

# =========================
# PÁGINA 1 — CADASTRO
# =========================

if pagina == "1. Cadastro do Projeto":

    st.subheader("1. Cadastro do Projeto")

    if not id_atual:
        st.warning("Crie uma nova proposta ou abra uma proposta existente na barra lateral.")
        st.stop()

    projeto = obter_projeto_atual() or {}

    with st.form("form_cadastro"):
        col1, col2 = st.columns(2)

        with col1:
            nome_projeto = st.text_input("Nome do projeto", value=projeto.get("nome_projeto", ""))
            coordenador = st.text_input("Coordenador", value=projeto.get("coordenador", ""))
            email_coordenador = st.text_input("E-mail do coordenador", value=projeto.get("email_coordenador", ""))
            instituicao_executora = st.text_input("Instituição executora", value=projeto.get("instituicao_executora", ""))
            campus = st.text_input("Campus", value=projeto.get("campus", ""))
            financiador = st.text_input("Financiador", value=projeto.get("financiador", ""))

            tipo_financiador = st.selectbox(
                "Tipo de financiador",
                ["Público Federal", "Público Estadual", "Público Municipal", "Privado", "Organismo Internacional", "Outro"],
                index=0
            )

            tipo_instrumento = st.selectbox(
                "Tipo de instrumento",
                ["Convênio", "Contrato", "TED", "Termo de Cooperação", "Emenda", "Outro"],
                index=0
            )

        with col2:
            data_inicio = st.date_input("Data de início", value=date.today())
            data_fim = st.date_input("Data de fim", value=date.today())
            prazo_meses = calcular_prazo_meses(data_inicio, data_fim)

            st.info(f"Prazo estimado: {prazo_meses} meses")

            valor_aprovado = st.number_input(
                "Valor aprovado / previsto do projeto",
                min_value=0.0,
                step=1000.0,
                format="%.2f",
                value=float(projeto.get("valor_aprovado", 0) or 0)
            )

            area_tematica = st.selectbox(
                "Área temática",
                ["Ensino", "Pesquisa", "Extensão", "Inovação", "Desenvolvimento institucional", "Concurso/PS", "Outra"],
                index=0
            )

            responsavel_facto = st.text_input("Responsável interno Facto", value=projeto.get("responsavel_facto", ""))

            status = st.selectbox(
                "Status da proposta",
                ["Rascunho", "Enviada", "Não aprovada", "Renegociada", "Aprovada"],
                index=0
            )

        objeto_resumido = st.text_area("Objeto resumido", value=projeto.get("objeto_resumido", ""))

        st.markdown("### Características operacionais previstas")

        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:
            possui_bolsas = st.checkbox("Bolsas/Auxílios", value=bool(projeto.get("possui_bolsas", False)))
        with c2:
            possui_compras = st.checkbox("Compras/Contratações", value=bool(projeto.get("possui_compras", False)))
        with c3:
            possui_concurso_ps = st.checkbox("Concurso/PS", value=bool(projeto.get("possui_concurso_ps", False)))
        with c4:
            possui_clt_rpa_pj = st.checkbox("CLT/RPA/PJ", value=bool(projeto.get("possui_clt_rpa_pj", False)))
        with c5:
            possui_importacao = st.checkbox("Importação", value=bool(projeto.get("possui_importacao", False)))

        salvar = st.form_submit_button("Salvar cadastro")

    if salvar:
        dados = {
            "id_projeto": id_atual,
            "versao": versao_atual,
            "status": status,
            "nome_projeto": nome_projeto,
            "coordenador": coordenador,
            "email_coordenador": email_coordenador,
            "instituicao_executora": instituicao_executora,
            "campus": campus,
            "financiador": financiador,
            "tipo_financiador": tipo_financiador,
            "tipo_instrumento": tipo_instrumento,
            "data_inicio": data_inicio,
            "data_fim": data_fim,
            "prazo_meses": prazo_meses,
            "valor_aprovado": valor_aprovado,
            "objeto_resumido": objeto_resumido,
            "area_tematica": area_tematica,
            "possui_bolsas": possui_bolsas,
            "possui_compras": possui_compras,
            "possui_concurso_ps": possui_concurso_ps,
            "possui_clt_rpa_pj": possui_clt_rpa_pj,
            "possui_importacao": possui_importacao,
            "responsavel_facto": responsavel_facto,
            "data_criacao": projeto.get("data_criacao", agora()),
            "data_atualizacao": agora()
        }

        salvar_ou_atualizar_projeto(dados)
        st.success("Cadastro salvo com sucesso.")
        st.rerun()

# =========================
# PÁGINA 2 — PARF
# =========================

elif pagina == "2. PARF":

    st.subheader("2. PARF — Estrutura Orçamentária")

    if not id_atual:
        st.warning("Crie ou abra uma proposta antes de preencher o PARF.")
        st.stop()

    aux = carregar_auxiliares()
    df_parf = carregar_parf()

    st.info("Nesta primeira versão, os itens são salvos em formato de tabela. Depois refinamos a tela por blocos.")

    grupo = st.selectbox(
        "Grupo de rubrica",
        grupos_parf()
    )

    modalidade = ""
    categoria = ""
    item = ""
    descricao = ""
    percentual_adicional = 0.0
    valor_adicional = 0.0
    modalidade_contratacao = ""

    with st.form("form_parf_item"):
        col1, col2, col3 = st.columns(3)

        with col1:
            valor_sugerido = 0.0

            if grupo == "Bolsas":
                df_bolsas = obter_opcoes_bolsas(aux)

                if not df_bolsas.empty and "Modalidade" in df_bolsas.columns:
                    modalidades = sorted(
                        df_bolsas["Modalidade"]
                        .dropna()
                        .astype(str)
                        .unique()
                        .tolist()
                    )

                    modalidade = st.selectbox(
                        "Modalidade da bolsa",
                        modalidades if modalidades else ["Não informado"]
                    )

                    if "Categoria" in df_bolsas.columns:
                        categorias = sorted(
                            df_bolsas.loc[
                                df_bolsas["Modalidade"].astype(str) == modalidade,
                                "Categoria"
                            ]
                            .dropna()
                            .astype(str)
                            .unique()
                            .tolist()
                        )
                    else:
                        categorias = []

                    categoria = st.selectbox(
                        "Categoria da bolsa",
                        categorias if categorias else ["Não informado"]
                    )

                    valor_sugerido = obter_valor_bolsa(df_bolsas, modalidade, categoria)

                else:
                    modalidade = st.text_input("Modalidade da bolsa")
                    categoria = st.text_input("Categoria da bolsa")

                item = st.text_input("Item", value=categoria)
                descricao = st.text_input("Descrição", value="Bolsa")

            elif grupo in grupos_contratacao():
                modalidade = st.selectbox(
                    "Modalidade",
                    [grupo]
                )

                opcoes = obter_lista_despesas(aux, grupo)
                categoria = st.selectbox(
                    "Categoria",
                    opcoes if opcoes else ["Não informado"]
                )

                item = st.text_input("Item específico")
                descricao = st.text_input("Descrição")

            elif grupo == "Diárias":
                df_diarias = obter_lista_diarias(aux)
                opcoes = df_diarias["tipo"].tolist() if not df_diarias.empty else []

                modalidade = st.selectbox(
                    "Tipo de diária/auxílio",
                    opcoes if opcoes else ["Não informado"]
                )

                categoria = "Diárias e auxílio financeiro para viagem"
                item = st.text_input("Item", value=modalidade)
                descricao = st.text_input("Descrição", value="Diária/Auxílio financeiro para viagem")

                if not df_diarias.empty:
                    valores = df_diarias.loc[df_diarias["tipo"] == modalidade, "valor"]
                    if not valores.empty:
                        valor_sugerido = float(valores.iloc[0])

            elif grupo == "Passagens":
                opcoes = obter_lista_passagens(aux)

                modalidade = st.selectbox(
                    "Tipo de passagem",
                    opcoes if opcoes else ["Não informado"]
                )

                categoria = "Passagens"
                item = st.text_input("Item", value=modalidade)
                descricao = st.text_input("Descrição", value="Passagem")

            elif grupo == "Celetistas":
                modalidade = st.selectbox(
                    "Modalidade",
                    ["CLT - Celetista"]
                )
                categoria = st.text_input("Cargo/Função")
                item = st.text_input("Item", value=categoria)
                descricao = st.text_input("Descrição", value="Contratação celetista")

            else:
                modalidade = st.text_input("Modalidade")
                categoria = st.text_input("Categoria")
                item = st.text_input("Item")
                descricao = st.text_input("Descrição")

        with col2:
            quantidade = st.number_input("Quantidade", min_value=0.0, step=1.0, value=1.0)
            meses = st.number_input("Meses", min_value=0.0, step=1.0, value=1.0)

        with col3:
            valor_unitario = st.number_input(
                "Valor unitário",
                min_value=0.0,
                step=100.0,
                format="%.2f",
                value=float(valor_sugerido or 0.0)
            )

            if grupo == "RPA":
                percentual_adicional = 20.0
                st.info("INSS patronal considerado: 20%")

            elif grupo == "Importação":
                percentual_adicional = 60.0
                st.info("Despesas acessórias de importação consideradas: 60%")

            if grupo in grupos_contratacao():
                marcacao = st.radio(
                    "Tratamento da contratação",
                    ["Automático", "Dispensa", "Inexigibilidade"],
                    horizontal=True
                )
            else:
                marcacao = "Automático"

        subtotal = quantidade * max(1, meses) * valor_unitario
        valor_adicional = subtotal * (percentual_adicional / 100)
        total = subtotal + valor_adicional

        if grupo in grupos_contratacao():
            if marcacao == "Dispensa":
                modalidade_contratacao = "Dispensa"
            elif marcacao == "Inexigibilidade":
                modalidade_contratacao = "Inexigibilidade"
            elif total > 40000:
                modalidade_contratacao = "Licitação"
            else:
                modalidade_contratacao = "Compra direta"
        else:
            modalidade_contratacao = ""

        st.markdown(f"**Total estimado do item:** {moeda(total)}")

        if modalidade_contratacao:
            st.markdown(f"**Modalidade estimada:** {modalidade_contratacao}")

        salvar_item = st.form_submit_button("Adicionar item ao PARF")

    if salvar_item:
        novo = {
            "id_item": gerar_id(),
            "id_projeto": id_atual,
            "versao": versao_atual,
            "grupo": grupo,
            "item": item,
            "modalidade": modalidade,
            "categoria": categoria,
            "descricao": descricao,
            "quantidade": quantidade,
            "meses": meses,
            "valor_unitario": valor_unitario,
            "percentual_adicional": percentual_adicional,
            "valor_adicional": valor_adicional,
            "modalidade_contratacao": modalidade_contratacao,
            "total": total
        }

        df_parf = pd.concat([df_parf, pd.DataFrame([novo])], ignore_index=True)
        salvar_excel(df_parf, ARQ_PARF)
        st.success("Item adicionado ao PARF.")
        st.rerun()

    st.divider()

    df_parf = carregar_parf()

    if df_parf.empty:
        df_atual = pd.DataFrame(columns=colunas_parf())
    else:
        df_atual = df_parf[
            (df_parf["id_projeto"].astype(str) == str(id_atual)) &
            (df_parf["versao"].astype(int) == int(versao_atual))
        ]

    if df_atual.empty:
        st.info("Nenhum item cadastrado no PARF ainda.")
    else:
        st.subheader("Itens do PARF")

        resumo = df_atual.groupby("grupo", as_index=False)["total"].sum()
        total_execucao = resumo["total"].sum()

        df_visual = df_atual.rename(columns={
            "grupo": "Grupo",
            "item": "Item",
            "modalidade": "Modalidade",
            "categoria": "Categoria",
            "descricao": "Descrição",
            "quantidade": "Quantidade",
            "meses": "Meses",
            "valor_unitario": "Valor unitário",
            "percentual_adicional": "% adicional",
            "valor_adicional": "Valor adicional",
            "modalidade_contratacao": "Modalidade de contratação",
            "total": "Total"
        })

        colunas_exibir = [
            "Grupo",
            "Item",
            "Modalidade",
            "Categoria",
            "Descrição",
            "Quantidade",
            "Meses",
            "Valor unitário",
            "% adicional",
            "Valor adicional",
            "Modalidade de contratação",
            "Total"
        ]

        for coluna in ["Valor unitário", "Valor adicional", "Total"]:
            df_visual[coluna] = df_visual[coluna].apply(moeda)

        st.dataframe(
            df_visual[colunas_exibir],
            use_container_width=True,
            hide_index=True
        )

        col1, col2 = st.columns([2, 1])

        with col1:
            resumo_visual = resumo.rename(columns={
                "grupo": "Grupo",
                "total": "Total"
            })

            resumo_visual["Total"] = resumo_visual["Total"].apply(moeda)

            st.markdown("### Resumo por rubrica")
            st.dataframe(
                resumo_visual,
                use_container_width=True,
                hide_index=True
            )

        with col2:
            st.markdown(
                f"""
                <div class="metric-card">
                    <div class="metric-label">Valor total de execução do projeto</div>
                    <div class="metric-value">{moeda(total_execucao)}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        st.divider()
        st.subheader("Editar ou excluir item do PARF")

        df_edicao = df_atual.copy()

        # Garante id_item para itens antigos
        if "id_item" not in df_edicao.columns:
            df_edicao["id_item"] = [gerar_id() for _ in range(len(df_edicao))]

        df_edicao["label_edicao"] = (
            df_edicao["grupo"].fillna("").astype(str)
            + " | "
            + df_edicao["item"].fillna("Sem item").astype(str)
            + " | "
            + df_edicao["total"].apply(moeda)
        )

        item_escolhido = st.selectbox(
            "Selecione um item salvo",
            df_edicao["label_edicao"].tolist()
        )

        linha_item = df_edicao[df_edicao["label_edicao"] == item_escolhido].iloc[0]

        with st.form("form_editar_item_parf"):
            col_e1, col_e2, col_e3 = st.columns(3)

            grupos_lista = grupos_parf()

            grupo_atual = str(linha_item.get("grupo", "Bolsas") or "Bolsas")
            index_grupo = grupos_lista.index(grupo_atual) if grupo_atual in grupos_lista else 0

            with col_e1:
                novo_grupo = st.selectbox(
                    "Grupo",
                    grupos_lista,
                    index=index_grupo
                )

                nova_modalidade = st.text_input(
                    "Modalidade",
                    value=str(linha_item.get("modalidade", "") or "")
                )

                nova_categoria = st.text_input(
                    "Categoria",
                    value=str(linha_item.get("categoria", "") or "")
                )

                novo_item = st.text_input(
                    "Item",
                    value=str(linha_item.get("item", "") or "")
                )

                nova_descricao = st.text_input(
                    "Descrição",
                    value=str(linha_item.get("descricao", "") or "")
                )

            with col_e2:
                nova_quantidade = st.number_input(
                    "Quantidade",
                    min_value=0.0,
                    step=1.0,
                    value=float(linha_item.get("quantidade", 0) or 0)
                )

                novos_meses = st.number_input(
                    "Meses",
                    min_value=0.0,
                    step=1.0,
                    value=float(linha_item.get("meses", 0) or 0)
                )

                novo_valor_unitario = st.number_input(
                    "Valor unitário",
                    min_value=0.0,
                    step=100.0,
                    format="%.2f",
                    value=float(linha_item.get("valor_unitario", 0) or 0)
                )

            with col_e3:
                novo_percentual_adicional = st.number_input(
                    "% adicional",
                    min_value=0.0,
                    step=1.0,
                    value=float(linha_item.get("percentual_adicional", 0) or 0)
                )

                nova_modalidade_contratacao = st.text_input(
                    "Modalidade de contratação",
                    value=str(linha_item.get("modalidade_contratacao", "") or "")
                )

            novo_subtotal = nova_quantidade * max(1, novos_meses) * novo_valor_unitario
            novo_valor_adicional = novo_subtotal * (novo_percentual_adicional / 100)
            novo_total = novo_subtotal + novo_valor_adicional

            st.markdown(f"**Novo total estimado:** {moeda(novo_total)}")

            col_btn1, col_btn2 = st.columns(2)

            with col_btn1:
                salvar_edicao = st.form_submit_button("Salvar alteração")

            with col_btn2:
                excluir_item = st.form_submit_button("Excluir item")

        if salvar_edicao:
            dados_atualizados = {
                "grupo": novo_grupo,
                "item": novo_item,
                "modalidade": nova_modalidade,
                "categoria": nova_categoria,
                "descricao": nova_descricao,
                "quantidade": nova_quantidade,
                "meses": novos_meses,
                "valor_unitario": novo_valor_unitario,
                "percentual_adicional": novo_percentual_adicional,
                "valor_adicional": novo_valor_adicional,
                "modalidade_contratacao": nova_modalidade_contratacao,
                "total": novo_total
            }

            ok = atualizar_item_parf(linha_item["id_item"], dados_atualizados)

            if ok:
                st.success("Item atualizado com sucesso.")
                st.rerun()
            else:
                st.error("Não foi possível atualizar o item.")

        if excluir_item:
            ok = excluir_item_parf(linha_item["id_item"])

            if ok:
                st.success("Item excluído com sucesso.")
                st.rerun()
            else:
                st.error("Não foi possível excluir o item.")

# =========================
# PÁGINA 3 — ICP/GIRO
# =========================

elif pagina == "3. ICP e Giro":

    st.subheader("3. ICP e Giro")

    if not id_atual:
        st.warning("Crie ou abra uma proposta antes de preencher ICP/Giro.")
        st.stop()

    escala = {
        1: "Muito baixa complexidade",
        2: "Baixa complexidade",
        3: "Média complexidade",
        4: "Alta complexidade",
        5: "Muito alta complexidade"
    }

    indicadores_icp = [
        "Complexidade do financiador",
        "Maturidade do projeto",
        "Dependência de articulação institucional",
        "Sensibilidade institucional/reputacional",
        "Complexidade normativa/jurídica",
        "Intensidade de acompanhamento gerencial",
        "Risco operacional",
        "Dependência de conhecimento técnico especializado"
    ]

    indicadores_giro = [
        "Giro em compras e contratações",
        "Giro em bolsas, auxílios e seleções",
        "Giro em projetos",
        "Giro administrativo e financeiro",
        "Giro em gestão institucional e governança"
    ]

    with st.expander("Ver escala explicativa"):
        for nota, texto in escala.items():
            st.write(f"**{nota}** — {texto}")

    registros = []

    st.markdown("### ICP")

    for indicador in indicadores_icp:
        col1, col2 = st.columns([1, 2])
        with col1:
            nota = st.slider(indicador, 1, 5, 3, key=f"icp_{indicador}")
        with col2:
            justificativa = st.text_input(f"Justificativa — {indicador}", key=f"just_icp_{indicador}")

        registros.append({
            "id_projeto": id_atual,
            "versao": versao_atual,
            "tipo": "ICP",
            "indicador": indicador,
            "nota": nota,
            "justificativa": justificativa
        })

    st.markdown("### Giro")

    for indicador in indicadores_giro:
        col1, col2 = st.columns([1, 2])
        with col1:
            nota = st.slider(indicador, 1, 5, 3, key=f"giro_{indicador}")
        with col2:
            justificativa = st.text_input(f"Justificativa — {indicador}", key=f"just_giro_{indicador}")

        registros.append({
            "id_projeto": id_atual,
            "versao": versao_atual,
            "tipo": "GIRO",
            "indicador": indicador,
            "nota": nota,
            "justificativa": justificativa
        })

    if st.button("Salvar ICP e Giro"):
        df = carregar_icp_giro()

        if not df.empty:
            df = df[
                ~(
                    (df["id_projeto"].astype(str) == str(id_atual)) &
                    (df["versao"].astype(int) == int(versao_atual))
                )
            ]

        df = pd.concat([df, pd.DataFrame(registros)], ignore_index=True)
        salvar_excel(df, ARQ_ICP_GIRO)
        st.success("ICP e Giro salvos com sucesso.")

    df_medias = pd.DataFrame(registros)

    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("ICP médio", round(df_medias[df_medias["tipo"] == "ICP"]["nota"].mean(), 2))
    with col_b:
        st.metric("Giro médio", round(df_medias[df_medias["tipo"] == "GIRO"]["nota"].mean(), 2))

# =========================
# PÁGINA 4 — DOA
# =========================

elif pagina == "4. DOA":

    st.subheader("4. DOA — Cálculo Operacional")

    if not id_atual:
        st.warning("Crie ou abra uma proposta antes de calcular a DOA.")
        st.stop()

    if not ARQ_CUSTOS_INDICADORES.exists():
        st.error("Arquivo de custos não encontrado. Salve a planilha como dados/custos_indicadores.xlsx.")
        st.stop()

    bases_custos = carregar_custos_indicadores()

    df_parf = carregar_parf()
    df_parf_atual = df_parf[
        (df_parf["id_projeto"].astype(str) == str(id_atual)) &
        (df_parf["versao"].astype(int) == int(versao_atual))
    ].copy()

    if df_parf_atual.empty:
        st.warning("Nenhum item foi lançado no PARF desta proposta. Preencha o PARF antes de calcular a DOA.")
        st.stop()

    projeto = obter_projeto_atual() or {}
    valor_projeto = float(projeto.get("valor_aprovado", 0) or 0)

    if valor_projeto <= 0:
        valor_projeto = float(pd.to_numeric(df_parf_atual["total"], errors="coerce").fillna(0).sum())
        st.info("Como o valor aprovado/previsto do projeto não foi informado, o limite de 15% foi calculado sobre o total do PARF.")

    memoria, custo_hora_df, media_icp, media_giro, fator_icp, fator_giro, fator_complexidade = gerar_memoria_doa(df_parf_atual, bases_custos)

    if memoria.empty:
        st.error("Não foi possível gerar memória de cálculo. Verifique se a aba tempo_pessoas contém as atividades esperadas.")
        st.stop()

    resumo_componentes = memoria.groupby("componente", as_index=False)["valor_calculado"].sum()

    total_calculado = float(resumo_componentes["valor_calculado"].sum())
    limite_legal = valor_projeto * 0.15
    fator_limitador = min(1, limite_legal / total_calculado) if total_calculado > 0 else 1
    total_ajustado = total_calculado * fator_limitador
    percentual_calculado = total_calculado / valor_projeto if valor_projeto else 0
    deficit_operacional = max(0, total_calculado - limite_legal)

    resumo_componentes["valor_ajustado"] = resumo_componentes["valor_calculado"] * fator_limitador

    st.markdown("### Resultado da DOA")

    c1, c2, c3 = st.columns(3)
    c4, c5, c6 = st.columns(3)

    with c1:
        st.metric("DOA calculada", moeda(total_calculado))
    with c2:
        st.metric("Percentual calculado", f"{percentual_calculado:.2%}".replace(".", ","))
    with c3:
        st.metric("Limite legal — 15%", moeda(limite_legal))
    with c4:
        st.metric("DOA aplicável", moeda(total_ajustado))
    with c5:
        st.metric("Fator limitador", f"{fator_limitador:.2%}".replace(".", ","))
    with c6:
        st.metric("Déficit operacional", moeda(deficit_operacional))

    st.markdown("### Fatores ICP e Giro aplicados")

    col_i1, col_i2, col_i3 = st.columns(3)
    with col_i1:
        st.metric("ICP médio", round(media_icp, 2))
        st.caption(f"Fator ICP aplicado: {fator_icp:.2f}".replace(".", ","))
    with col_i2:
        st.metric("Giro médio", round(media_giro, 2))
        st.caption(f"Fator Giro aplicado: {fator_giro:.2f}".replace(".", ","))
    with col_i3:
        st.metric("Fator ponderado", f"{fator_complexidade:.2f}".replace(".", ","))
        st.caption("Peso inicial: ICP 40% e Giro 60%")

    st.markdown("### Composição da DOA por componente")

    resumo_visual = resumo_componentes.rename(columns={
        "componente": "Componente da DOA",
        "valor_calculado": "Valor calculado",
        "valor_ajustado": "Valor ajustado pelo limite"
    })
    resumo_visual = formatar_tabela_moeda(resumo_visual, ["Valor calculado", "Valor ajustado pelo limite"])

    st.dataframe(resumo_visual, use_container_width=True, hide_index=True)

    with st.expander("Ver memória de cálculo detalhada"):
        memoria_visual = memoria.rename(columns={
            "componente": "Componente",
            "setor": "Setor",
            "atividade": "Atividade",
            "quantidade_operacional": "Quantidade operacional",
            "tempo_base_h": "Tempo base (h)",
            "horas_estimadas": "Horas estimadas",
            "custo_hora_pessoal": "Custo hora pessoal",
            "custo_hora_adm": "Custo hora administrativo",
            "custo_hora_total": "Custo hora total",
            "fator_complexidade": "Fator ICP/Giro",
            "valor_calculado": "Valor calculado"
        })

        memoria_visual = formatar_tabela_moeda(
            memoria_visual,
            ["Custo hora pessoal", "Custo hora administrativo", "Custo hora total", "Valor calculado"]
        )

        st.dataframe(memoria_visual, use_container_width=True, hide_index=True)

    with st.expander("Ver custo/hora por centro de custo"):
        custo_visual = custo_hora_df.rename(columns={
            "setor": "Setor",
            "custo_hora_pessoal": "Custo hora pessoal",
            "custo_hora_adm": "Custo hora administrativo rateado",
            "custo_hora_total": "Custo hora total"
        })
        custo_visual = formatar_tabela_moeda(
            custo_visual,
            ["Custo hora pessoal", "Custo hora administrativo rateado", "Custo hora total"]
        )
        st.dataframe(custo_visual, use_container_width=True, hide_index=True)

    with st.expander("Explicações das faixas ICP e Giro"):
        bases = carregar_custos_indicadores()

        st.markdown("#### ICP")
        df_icp = bases.get("ICP", pd.DataFrame())
        if not df_icp.empty:
            st.dataframe(df_icp, use_container_width=True, hide_index=True)
        else:
            st.info("Aba ICP não encontrada no arquivo custos_indicadores.xlsx.")

        st.markdown("#### Giro")
        df_giro = bases.get("GIRO", pd.DataFrame())
        if not df_giro.empty:
            st.dataframe(df_giro, use_container_width=True, hide_index=True)
        else:
            st.info("Aba GIRO não encontrada no arquivo custos_indicadores.xlsx.")

    if st.button("Salvar resultado da DOA", type="primary"):
        df_doa = carregar_doa()

        if not df_doa.empty:
            df_doa = df_doa[
                ~(
                    (df_doa["id_projeto"].astype(str) == str(id_atual)) &
                    (df_doa["versao"].astype(int) == int(versao_atual))
                )
            ]

        memoria_salvar = memoria.copy()
        memoria_salvar["id_projeto"] = id_atual
        memoria_salvar["versao"] = versao_atual
        memoria_salvar["valor_ajustado"] = memoria_salvar["valor_calculado"] * fator_limitador
        memoria_salvar["fator_limitador"] = fator_limitador
        memoria_salvar["data_calculo"] = agora()

        for coluna in colunas_doa():
            if coluna not in memoria_salvar.columns:
                memoria_salvar[coluna] = None

        memoria_salvar = memoria_salvar[colunas_doa()]
        df_doa = pd.concat([df_doa, memoria_salvar], ignore_index=True)
        salvar_excel(df_doa, ARQ_DOA)

        st.success("Resultado da DOA salvo com sucesso.")

# =========================
# PÁGINA 5 — PROPOSTA
# =========================

elif pagina == "5. Proposta Comercial":

    st.subheader("5. Proposta Comercial")

    if not id_atual:
        st.warning("Crie ou abra uma proposta antes de gerar o documento.")
        st.stop()

    st.markdown(
        "A proposta será gerada a partir do modelo `dados/modelo_proposta.docx`, "
        "incluindo os dados do projeto, o resumo das rubricas do PARF e o detalhamento final da DOA."
    )

    if not ARQ_MODELO_PROPOSTA.exists():
        st.error("Modelo não encontrado. Salve o arquivo como dados/modelo_proposta.docx.")
        st.stop()

    df_parf_atual = obter_parf_atual_df()
    if df_parf_atual.empty:
        st.warning("Nenhum item foi lançado no PARF desta proposta. Preencha o PARF antes de gerar a proposta.")
    else:
        st.markdown("### Prévia das informações que serão inseridas")

        resumo_parf = obter_resumo_parf_proposta(df_parf_atual)
        resumo_parf_visual = resumo_parf.copy()
        if not resumo_parf_visual.empty:
            resumo_parf_visual["Valor previsto"] = resumo_parf_visual["Valor previsto"].apply(moeda)
        st.markdown("#### Resumo das rubricas do PARF")
        st.dataframe(resumo_parf_visual, use_container_width=True, hide_index=True)

        resumo_doa = obter_resumo_doa_proposta(df_parf_atual)
        resumo_doa_visual = resumo_doa.copy()
        if not resumo_doa_visual.empty:
            resumo_doa_visual["Valor calculado"] = resumo_doa_visual["Valor calculado"].apply(moeda)
            resumo_doa_visual["Valor final"] = resumo_doa_visual["Valor final"].apply(moeda)
        st.markdown("#### Detalhamento da DOA")
        st.dataframe(resumo_doa_visual, use_container_width=True, hide_index=True)

    if st.button("Gerar proposta", type="primary"):
        caminho_proposta = gerar_proposta_docx()
        if caminho_proposta and caminho_proposta.exists():
            st.session_state["ultima_proposta_docx"] = str(caminho_proposta)
            st.success("Proposta gerada com sucesso.")

    caminho_ultima = st.session_state.get("ultima_proposta_docx")
    if caminho_ultima and Path(caminho_ultima).exists():
        with open(caminho_ultima, "rb") as f:
            st.download_button(
                label="Baixar proposta em Word",
                data=f,
                file_name=Path(caminho_ultima).name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

# =========================
# PÁGINA 6 — HISTÓRICO
# =========================

elif pagina == "6. Histórico":

    st.subheader("6. Histórico de Propostas")

    df = carregar_projetos()

    if df.empty:
        st.info("Nenhuma proposta cadastrada.")
    else:
        st.dataframe(df, use_container_width=True)