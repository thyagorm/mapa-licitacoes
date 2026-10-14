import streamlit as st
import pandas as pd
import time
import os
import re
import json
import hashlib
import unicodedata
from io import BytesIO
from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from pypdf import PdfReader
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Cliente Supabase
try:
    from supabase import create_client
    supabase_disponivel = True
except ImportError:
    supabase_disponivel = False

# 1. Estrutura de Extração
class ItemLicitacao(BaseModel):
    numero_item: str = Field(description="Número do item ou lote (ex: '1', '02')")
    descricao: str = Field(description="Descrição completa do medicamento no edital")
    principio_ativo_identificado: str = Field(description="Substância ativa genérica identificada")
    unidade: str = Field(description="Unidade de medida (ex: AMP, FA, COMP, FR)")
    quantidade: float = Field(description="Quantidade numérica float")
    valor_referencia_unitario: float = Field(description="Valor unitário de referência float (0.0 se não houver)")

class ListaItens(BaseModel):
    itens: list[ItemLicitacao]

# 2. Configurações da Página
st.set_page_config(page_title="Radar Farma - Licitações", layout="wide", page_icon="💊")
st.title("🎯 Analisador de Editais & Mapa de Preços")
st.caption("🚀 Pipeline Ativo: v4.2 - Modelos Validados na Conta (3.5-Lite / 3.8-Flash) & Cota Otimizada")

# Chaves de API e Conexões
api_key = ""
if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = str(st.secrets["GEMINI_API_KEY"]).strip().strip("'").strip('"')
else:
    api_key_input = st.sidebar.text_input("Chave Gemini API", type="password")
    if api_key_input:
        api_key = api_key_input.strip().strip("'").strip('"')

supabase_client = None
if supabase_disponivel and "SUPABASE_URL" in st.secrets and "SUPABASE_KEY" in st.secrets:
    try:
        supabase_client = create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])
    except Exception:
        supabase_client = None

# 3. Base de Dados e Normalização
ARQUIVO_PORTFOLIO = "portfolio_laboratorios.xlsx"

STOP_WORDS_FARMA = {
    "ACETATO", "CLORIDRATO", "SULFATO", "FOSFATO", "DIPROPIONATO", "BROMETO", 
    "CITRATO", "SODICO", "SODICA", "POTASSICO", "POTASSICA", "MALEATO", 
    "SUCCINATO", "HEMISSULFATO", "MESILATO", "TARTARATO", "GLICONATO", "PO", 
    "SOLUCAO", "INJETAVEL", "COMPRIMIDO", "SUSPENSAO", "GOTAS", "CAPSULA", 
    "FRASCO", "AMPOLA", "BISNAGA", "CREME", "POMADA", "XAROPE", "MG", "ML", "MCG", "UI"
}

# Lista estrita de modelos verificados e ativos na sua chave da API
MODELOS_ATIVOS_CONTA = [
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.1-flash-lite",
    "gemini-2.5-flash-lite",
    "gemini-3.5-flash",
    "gemini-3.6-flash",
    "gemini-2.5-flash"
]

def normalizar_texto(texto):
    if not texto or pd.isna(texto):
        return ""
    nfkd = unicodedata.normalize('NFKD', str(texto))
    texto_sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    texto_limpo = re.sub(r'[^A-Z0-9\s]', ' ', texto_sem_acento.upper())
    return " ".join(texto_limpo.split())

@st.cache_data(ttl=300, show_spinner=False)
def carregar_memoria_db():
    if not supabase_client:
        return {}
    try:
        res = supabase_client.table("memoria_medicamentos").select("*").execute()
        memoria = {}
        for row in res.data:
            memoria[row["descricao_normalizada"]] = {
                "principio_ativo": row["principio_ativo"],
                "unidade": row.get("unidade", ""),
                "vezes_visto": row.get("vezes_visto", 1)
            }
        return memoria
    except Exception:
        return {}

def salvar_aprendizado_db(novos_itens):
    if not supabase_client:
        return
    rows_to_upsert = []
    for item in novos_itens:
        desc_norm = normalizar_texto(item.get("descricao", ""))
        princ_ativo = normalizar_texto(item.get("principio_ativo_identificado", ""))
        unidade = item.get("unidade", "")
        if desc_norm and princ_ativo:
            rows_to_upsert.append({
                "descricao_normalizada": desc_norm,
                "principio_ativo": princ_ativo,
                "unidade": unidade
            })
    if rows_to_upsert:
        try:
            supabase_client.table("memoria_medicamentos").upsert(rows_to_upsert).execute()
        except Exception:
            pass

MAPEAMENTO_LABS = {
    "Pint Pharma": ["PINT", "PINT PHARMA"],
    "Sanofi": ["SANOFI"],
    "Blau": ["BLAU"],
    "Eurofarma": ["EUROFARMA"],
    "Baxter": ["BAXTER"],
    "Biocon": ["BIOCON"],
    "Accord": ["ACCORD"],
    "Halex Istar": ["HALEX", "ISTAR"],
    "United Medical": ["UNITED MEDICAL", "UNITED"],
    "GSK": ["GSK", "GLAXO", "GLAXOSMITHKLINE"],
    "Aspen": ["ASPEN"]
}

@st.cache_data
def carregar_base_portfolio(caminho):
    if not os.path.exists(caminho):
        return None
    try:
        df = pd.read_excel(caminho)
        df.columns = [str(c).strip().upper() for c in df.columns]
        for col in ["SUBSTÂNCIA", "LABORATÓRIO", "PRODUTO", "APRESENTAÇÃO"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
        return df
    except Exception:
        return None

df_portfolio = carregar_base_portfolio(ARQUIVO_PORTFOLIO)
base_memoria = carregar_memoria_db()

# 4. Barra Lateral
st.sidebar.header("⚙️ Configurações de Busca")

modelo_preferido = st.sidebar.selectbox(
    "🤖 Modelo Principal (Ativos na Chave):",
    options=MODELOS_ATIVOS_CONTA,
    index=0,
    help="Modelos confirmados na sua chave. Se o principal estiver com fila (503) ou limite (429), o sistema transborda automaticamente para o próximo da lista."
)

segmento = st.sidebar.selectbox(
    "Segmento do Edital:",
    options=["Medicamentos", "Material Elétrico / Engenharia", "Personalizado / Geral"]
)

if segmento == "Medicamentos":
    st.sidebar.subheader("💊 Portfólio de Laboratórios")
    labs_base = [
        "Pint Pharma", "Sanofi", "Blau", "Eurofarma", "Baxter", 
        "Biocon", "Accord", "Halex Istar", "United Medical", "GSK", "Aspen"
    ]
    
    labs_selecionados = st.sidebar.multiselect(
        "Laboratórios ativos para cotação:",
        options=labs_base,
        default=labs_base
    )
    
    if df_portfolio is not None:
        st.sidebar.success(f"Base Portfólio: {len(df_portfolio)} produtos cadastrados.")
    else:
        st.sidebar.warning("Arquivo 'portfolio_laboratorios.xlsx' não encontrado na raiz.")

    if supabase_client:
        st.sidebar.success(f"🗄️ **Banco Supabase Conectado:** {len(base_memoria)} termos.")

    st.sidebar.markdown(
        "**Hierarquia Comercial Estratégica:**\n"
        "- 🥇 **Pint Pharma & Sanofi** prioridade máxima.\n"
        "- 🥈 **Blau** sobre **Eurofarma**.\n"
        "- 🚫 **Bloqueio:** Medley e genéricos da Sanofi eliminados."
    )

elif segmento == "Material Elétrico / Engenharia":
    portfolio_texto = st.sidebar.text_area(
        "Itens de interesse:",
        value="Material elétrico, cabeamento estruturado, iluminação LED, quadros de distribuição"
    )
else:
    portfolio_texto = st.sidebar.text_area("Itens de interesse:", value="")

arquivo_pdf = st.file_uploader("Arraste o PDF do Edital ou Termo de Referência aqui", type=["pdf"])

# ========================================================
# PRÉ-PROCESSADOR CIRÚRGICO (REDUÇÃO MÁXIMA DE TOKENS)
# ========================================================
@st.cache_data(show_spinner=False)
def analisar_e_filtrar_pdf(bytes_arquivo):
    leitor = PdfReader(BytesIO(bytes_arquivo))
    total_paginas = len(leitor.pages)
    
    termos_itens = [
        "ITEM", "QUANT", "QTD", "UNID", "VALOR", "PRECO", "REFERENCIA", 
        "PRINCIPIO ATIVO", "CONCENTRACAO", "LOTE", "MEDICAMENTO", "MG", "ML",
        "AMPOLA", "FRASCO", "COMPRIMIDO", "CAPSULA", "SOLUCAO", "INJETAVEL"
    ]
    
    paginas_selecionadas = []
    paginas_indices = []
    total_chars_extraidos = 0

    for i, pagina in enumerate(leitor.pages):
        raw_text = pagina.extract_text() or ""
        total_chars_extraidos += len(raw_text.strip())
        norm_text = normalizar_texto(raw_text)
        
        score = sum(1 for t in termos_itens if t in norm_text)
        tem_padrao_item = bool(re.search(r'\b(ITEM\s*\d+|LOTE\s*\d+|\d+\s+UNID|\d+\s+AMP|\d+\s+FR|\d+\s+COMP)\b', norm_text))
        
        if score >= 2 or tem_padrao_item:
            texto_enxuto = re.sub(r'[ \t]+', ' ', raw_text.strip())
            texto_enxuto = re.sub(r'\n{2,}', '\n', texto_enxuto)
            paginas_selecionadas.append(f"[PÁG {i+1}]\n{texto_enxuto}")
            paginas_indices.append(i + 1)

    eh_escaneado = (total_chars_extraidos < (total_paginas * 30)) and (total_chars_extraidos < 120)

    if not eh_escaneado:
        if len(paginas_selecionadas) == 0:
            for i, pagina in enumerate(leitor.pages):
                txt = (pagina.extract_text() or "").strip()
                if txt:
                    paginas_selecionadas.append(f"[PÁG {i+1}]\n{txt}")
                    paginas_indices.append(i + 1)

        texto_filtrado = "\n\n".join(paginas_selecionadas)
        if len(texto_filtrado) > 50000:
            texto_filtrado = texto_filtrado[:50000]
    else:
        texto_filtrado = ""

    return eh_escaneado, texto_filtrado, paginas_indices, total_paginas

# ========================================================
# ROTEADOR MULTI-MODELO (USA APENAS OS MODELOS VIVOS DA CHAVE)
# ========================================================
@st.cache_data(show_spinner=False)
def extrair_itens_roteamento_inteligente(pdf_bytes_hash, pdf_bytes, texto_edital, eh_escaneado, prompt_instrucao, modelo_primario, key_api):
    client = genai.Client(api_key=key_api)

    if eh_escaneado:
        conteudos = [
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            prompt_instrucao
        ]
    else:
        conteudos = [f"{prompt_instrucao}\n\nEDITAL:\n{texto_edital}"]

    fila_modelos = [modelo_primario] + [m for m in MODELOS_ATIVOS_CONTA if m != modelo_primario]
    ultimo_erro = None

    for mod in fila_modelos:
        try:
            resposta = client.models.generate_content(
                model=mod,
                contents=conteudos,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ListaItens,
                    temperature=0.1
                )
            )
            if resposta and resposta.text:
                return resposta.text, mod
        except Exception as err:
            ultimo_erro = err
            msg = str(err)
            if any(cod in msg for cod in ["429", "RESOURCE_EXHAUSTED", "503", "UNAVAILABLE", "500"]):
                time.sleep(1.5)
                continue
            else:
                raise err

    raise ultimo_erro

# ========================================================
# CRUZAMENTO INTELIGENTE: PINT PHARMA & HIERARQUIA RIGOROSA
# ========================================================
def enriquecer_com_portfolio(df_extraido, df_port, labs_escolhidos):
    if df_port is None:
        df_extraido["Laboratório Sugerido"] = "Sem base carregada"
        df_extraido["Produto / Marca Ref."] = "-"
        return df_extraido

    base = df_port.copy()
    base["SUBSTANCIA_NORM"] = base["SUBSTÂNCIA"].apply(normalizar_texto)
    base["LABORATORIO_NORM"] = base["LABORATÓRIO"].apply(normalizar_texto)
    base["PRODUTO_NORM"] = base["PRODUTO"].apply(normalizar_texto)
    base["APRESENTACAO_NORM"] = base["APRESENTAÇÃO"].apply(normalizar_texto) if "APRESENTAÇÃO" in base.columns else ""

    mascara_medley = (
        base["LABORATORIO_NORM"].str.contains("MEDLEY", na=False) |
        base["PRODUTO_NORM"].str.contains("MEDLEY", na=False)
    )

    mascara_generico_sanofi = (
        base["LABORATORIO_NORM"].str.contains("SANOFI", na=False) & (
            base["PRODUTO_NORM"].str.contains("GENERIC", na=False) |
            base["APRESENTACAO_NORM"].str.contains("GENERIC", na=False) |
            (base["PRODUTO_NORM"] == base["SUBSTANCIA_NORM"])
        )
    )

    base_valida = base[~(mascara_medley | mascara_generico_sanofi)]

    termos_busca_labs = []
    for lab in labs_escolhidos:
        termos_busca_labs.extend(MAPEAMENTO_LABS.get(lab, [normalizar_texto(lab)]))

    labs_atribuidos = []
    produtos_atribuidos = []

    for _, row in df_extraido.iterrows():
        substancia_edital = normalizar_texto(row["Princípio Ativo"])
        desc_completa = normalizar_texto(row["Descrição Completa Edital"])
        
        palavras_edital = [
            p for p in substancia_edital.split() 
            if len(p) > 3 and p not in STOP_WORDS_FARMA
        ]

        matches = base_valida[
            base_valida["SUBSTANCIA_NORM"].apply(lambda s: s != "" and (s in substancia_edital or substancia_edital in s)) |
            base_valida["PRODUTO_NORM"].apply(lambda p: p != "" and len(p) > 3 and (p in desc_completa or p in substancia_edital)) |
            base_valida["SUBSTANCIA_NORM"].apply(lambda s: any(p in s for p in palavras_edital) if palavras_edital else False)
        ]

        matches = matches[matches["LABORATORIO_NORM"].apply(
            lambda lab_nome: any(termo in lab_nome for termo in termos_busca_labs)
        )]

        if not matches.empty:
            labs_encontrados = matches["LABORATORIO_NORM"].unique().tolist()
            
            tem_pint = any("PINT" in l for l in labs_encontrados)
            tem_sanofi = any("SANOFI" in l for l in labs_encontrados)
            tem_united = any("UNITED" in l for l in labs_encontrados)
            tem_baxter = any("BAXTER" in l for l in labs_encontrados)
            tem_blau = any("BLAU" in l for l in labs_encontrados)
            tem_euro = any("EUROFARMA" in l for l in labs_encontrados)

            if tem_pint:
                lab_final_filtro = [l for l in labs_encontrados if "PINT" in l][0]
            elif tem_sanofi:
                lab_final_filtro = [l for l in labs_encontrados if "SANOFI" in l][0]
            elif tem_united:
                lab_final_filtro = [l for l in labs_encontrados if "UNITED" in l][0]
            elif tem_baxter:
                lab_final_filtro = [l for l in labs_encontrados if "BAXTER" in l][0]
            elif tem_blau:
                lab_final_filtro = [l for l in labs_encontrados if "BLAU" in l][0]
            elif tem_euro:
                lab_final_filtro = [l for l in labs_encontrados if "EUROFARMA" in l][0]
            else:
                lab_final_filtro = labs_encontrados[0]

            linha_escolhida = matches[matches["LABORATORIO_NORM"] == lab_final_filtro].iloc[0]
            labs_atribuidos.append(linha_escolhida["LABORATÓRIO"])
            produtos_atribuidos.append(linha_escolhida["PRODUTO"])
        else:
            labs_atribuidos.append("Não mapeado / Verificar")
            produtos_atribuidos.append("-")

    df_extraido["Laboratório Sugerido"] = labs_atribuidos
    df_extraido["Produto / Marca Ref."] = produtos_atribuidos
    return df_extraido

# Gerador Excel
def gerar_excel_estilizado(df_dados):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_dados.to_excel(writer, index=False, sheet_name="Mapa_de_Precos")
        ws = writer.sheets["Mapa_de_Precos"]
        ws.views.sheetView[0].showGridLines = True

        fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_cabecalho = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        
        fonte_corpo = Font(name="Calibri", size=10)
        fill_editavel = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        alinhamento_centro = Alignment(horizontal="center", vertical="center")
        alinhamento_esquerda = Alignment(horizontal="left", vertical="center")
        alinhamento_direita = Alignment(horizontal="right", vertical="center")
        
        borda_fina = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )

        num_linhas = len(df_dados)
        num_cols = len(df_dados.columns)
        colunas_nomes = list(df_dados.columns)

        for col_num in range(1, num_cols + 1):
            celula = ws.cell(row=1, column=col_num)
            celula.font = fonte_cabecalho
            celula.fill = fill_cabecalho
            celula.alignment = alinhamento_centro
            ws.row_dimensions[1].height = 28

        def get_col_letter(col_name):
            if col_name in colunas_nomes:
                return get_column_letter(colunas_nomes.index(col_name) + 1)
            return None

        col_qtd = get_col_letter("Qtd")
        col_vref = get_col_letter("Valor Ref. Unit. (R$)")
        col_custo = get_col_letter("Custo Aquisição (R$)")
        col_margem = get_col_letter("Margem Alvo (%)")

        for row_idx in range(2, num_linhas + 2):
            ws.row_dimensions[row_idx].height = 20
            for col_idx in range(1, num_cols + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                nome_col = colunas_nomes[col_idx - 1]
                cell.font = fonte_corpo
                cell.border = borda_fina

                if nome_col in ["Item", "Unidade"]:
                    cell.alignment = alinhamento_centro
                elif nome_col in ["Descrição Completa Edital", "Princípio Ativo", "Laboratório Sugerido", "Produto / Marca Ref."]:
                    cell.alignment = alinhamento_esquerda
                elif nome_col == "Qtd":
                    cell.alignment = alinhamento_direita
                    cell.number_format = "#,##0"
                elif nome_col in ["Valor Ref. Unit. (R$)", "Valor Total Estimado (R$)"]:
                    cell.alignment = alinhamento_direita
                    cell.number_format = "R$ #,##0.00"
                    if nome_col == "Valor Total Estimado (R$)" and col_qtd and col_vref:
                        cell.value = f"={col_qtd}{row_idx}*{col_vref}{row_idx}"
                elif nome_col == "Custo Aquisição (R$)":
                    cell.alignment = alinhamento_direita
                    cell.number_format = "R$ #,##0.00"
                    cell.fill = fill_editavel
                elif nome_col == "Margem Alvo (%)":
                    cell.alignment = alinhamento_direita
                    cell.number_format = "0.0%"
                    cell.value = 0.15
                elif nome_col == "Preço Proposta Unit. (R$)":
                    cell.alignment = alinhamento_direita
                    cell.number_format = "R$ #,##0.00"
                    if col_custo and col_margem:
                        cell.value = f"={col_custo}{row_idx}*(1+{col_margem}{row_idx})"

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

        if "Descrição Completa Edital" in colunas_nomes:
            ws.column_dimensions[get_col_letter("Descrição Completa Edital")].width = 42
        if "Princípio Ativo" in colunas_nomes:
            ws.column_dimensions[get_col_letter("Princípio Ativo")].width = 25
        if "Laboratório Sugerido" in colunas_nomes:
            ws.column_dimensions[get_col_letter("Laboratório Sugerido")].width = 28
        if "Produto / Marca Ref." in colunas_nomes:
            ws.column_dimensions[get_col_letter("Produto / Marca Ref.")].width = 22

    return output.getvalue()

# 5. Processamento Principal
if arquivo_pdf and api_key:
    if st.button("🚀 Processar Edital", type="primary"):
        barra_progresso = st.progress(0)
        texto_status = st.empty()
        
        try:
            pdf_bytes = arquivo_pdf.getvalue()
            pdf_hash = hashlib.sha256(pdf_bytes).hexdigest()
            
            # ETAPA 1: Filtragem Cirúrgica de Páginas
            barra_progresso.progress(25)
            texto_status.text("🔍 Filtrando páginas relevantes localmente (0 tokens)...")
            
            eh_escaneado, texto_edital, paginas_filtradas, total_pags = analisar_e_filtrar_pdf(pdf_bytes)

            if eh_escaneado:
                st.warning(f"📸 **PDF Escaneado Detectado ({total_pags} págs):** Usando leitura visual direta.")
            else:
                tokens_est = int(len(texto_edital) / 4)
                st.info(f"📄 **Filtro Ativo:** {len(paginas_filtradas)} de {total_pags} página(s) selecionada(s) (~{tokens_est:,} tokens de entrada).")

            # ETAPA 2: Prompt Ultra-Enxuto
            barra_progresso.progress(45)
            texto_status.text("⚙️ Preparando extração com mínimo consumo de TPM...")

            if segmento == "Medicamentos":
                prompt = (
                    "Extraia todos os medicamentos licitados neste documento para o JSON solicitado. "
                    "Em 'principio_ativo_identificado', informe apenas a substância ativa genérica "
                    "(ex: 'Enoxaparina Sódica', 'Propofol', 'Dipirona'). "
                    "Quantidades e valores unitários em float (0.0 se ausente)."
                )
            else:
                prompt = f"Extraia os itens referentes a: {portfolio_texto}. Retorne no JSON solicitado."

            # ETAPA 3: Extração com Roteamento Multi-Modelo (Cotas Independentes)
            barra_progresso.progress(65)
            texto_status.text(f"🤖 Consultando {modelo_preferido} (com transbordo automático de cota)...")

            json_resposta, modelo_usado = extrair_itens_roteamento_inteligente(
                pdf_hash, pdf_bytes, texto_edital, eh_escaneado, prompt, modelo_preferido, api_key
            )

            barra_progresso.progress(80)
            texto_status.text("📦 Estruturando itens extraídos...")
            dados = ListaItens.model_validate_json(json_resposta)

            if not dados.itens:
                barra_progresso.empty()
                texto_status.empty()
                st.warning("Nenhum item foi identificado no documento.")
            else:
                # ETAPA 4: Cruzamento Comercial Local (0 Tokens)
                barra_progresso.progress(90)
                texto_status.text("🧠 Cruzando com portfólio Excel localmente (Pint Pharma / Sanofi no topo)...")

                lista_dicts = [item.model_dump() for item in dados.itens]
                salvar_aprendizado_db(lista_dicts)

                df = pd.DataFrame(lista_dicts)
                df.columns = [
                    "Item", 
                    "Descrição Completa Edital", 
                    "Princípio Ativo", 
                    "Unidade", 
                    "Qtd", 
                    "Valor Ref. Unit. (R$)"
                ]

                if (df["Valor Ref. Unit. (R$)"] > 0).any():
                    df["Valor Total Estimado (R$)"] = df["Qtd"] * df["Valor Ref. Unit. (R$)"]

                total_extraidos_bruto = len(df)

                if segmento == "Medicamentos":
                    df = enriquecer_com_portfolio(df, df_portfolio, labs_selecionados)
                    df = df[df["Laboratório Sugerido"] != "Não mapeado / Verificar"].copy()

                if df.empty:
                    barra_progresso.empty()
                    texto_status.empty()
                    st.warning(f"⚠️ A IA extraiu **{total_extraidos_bruto} itens** do edital, mas nenhum deles pertence ao portfólio dos laboratórios selecionados.")
                    st.stop()

                df["Custo Aquisição (R$)"] = 0.0
                df["Margem Alvo (%)"] = 0.15
                df["Preço Proposta Unit. (R$)"] = 0.0

                try:
                    df["_item_num"] = pd.to_numeric(df["Item"].str.extract(r'(\d+)')[0], errors="coerce")
                    df = df.sort_values(by="_item_num").drop(columns=["_item_num"])
                except Exception:
                    pass

                # ETAPA 5: Excel
                barra_progresso.progress(96)
                texto_status.text("📊 Gerando planilha Excel executiva (.xlsx)...")
                excel_bytes = gerar_excel_estilizado(df)

                barra_progresso.progress(100)
                texto_status.text("✅ Concluído!")
                time.sleep(0.2)
                barra_progresso.empty()
                texto_status.empty()

                st.success(f"✅ De **{total_extraidos_bruto} itens** lidos no edital, **{len(df)} itens comerciais** deram match no portfólio via **{modelo_usado}**!")
                st.dataframe(df, use_container_width=True)

                st.download_button(
                    label="📥 Baixar Mapa de Preços Profissional (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Mapa_Precos_{arquivo_pdf.name.replace('.pdf', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            barra_progresso.empty()
            texto_status.empty()
            st.error(f"Erro no processamento: {str(e)}")

elif not api_key:
    st.info("Insira a sua chave de API na barra lateral ou configure nos secrets para iniciar a análise.")