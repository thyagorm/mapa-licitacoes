import io
import os
import re
import json
import hashlib
from io import BytesIO
import pandas as pd
import streamlit as st
import google.generativeai as genai
from google.generativeai import types
from pypdf import PdfReader
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ========================================================
# CONFIGURAÇÃO INICIAL DA PÁGINA
# ========================================================
st.set_page_config(
    page_title="STRIKE | Analisador de Editais & Mapa de Preços",
    page_icon="🎯",
    layout="wide"
)

# ========================================================
# CONSTANTES E DICIONÁRIOS DE DOMÍNIO (FARMA)
# ========================================================
STOP_WORDS_FARMA = {
    "solucao", "injetavel", "comprimido", "capsula", "frasco", "ampola",
    "revestido", "mg", "ml", "g", "UI", "fr/amp", "f/a", "suspensao",
    "oral", "topica", "oftalmica", "nasal", "pomada", "creme", "gel"
}

MAPEAMENTO_LABS = {
    "Pint Pharma": ["PINT", "PINT PHARMA"],
    "Sanofi": ["SANOFI", "SANOFI-AVENTIS", "GENERICO SANOFI"],
    "Blau": ["BLAU", "BLAU FARMACEUTICA"],
    "Eurofarma": ["EUROFARMA"],
    "United Medical": ["UNITED", "UNITED MEDICAL"],
    "Baxter": ["BAXTER"]
}

MODELOS_ATIVOS_CONTA = [
    "gemini-2.5-flash",
    "gemini-2.5-pro",
    "gemini-1.5-flash"
]

# ========================================================
# FUNÇÕES DE UTILIDADE E NORMALIZAÇÃO
# ========================================================
def normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    import unicodedata
    nfkd = unicodedata.normalize('NFKD', texto)
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return re.sub(r'[^a-zA-Z0-9\s]', ' ', sem_acento).upper()

# ========================================================
# ESQUEMA DE VALIDAÇÃO (PYDANTIC)
# ========================================================
from pydantic import BaseModel, Field

class ItemEdital(BaseModel):
    item: str = Field(description="Número ou identificador do item no edital")
    principio_ativo_identificado: str = Field(description="Substância ativa genérica ou descrição resumida do item")
    descricao_completa_edital: str = Field(description="Descrição original integral do item")
    unidade_fornecimento: str = Field(description="UN, FR, AMP, CX, COMP, etc.")
    quantidade_licitada: float = Field(description="Quantidade total licitada")
    valor_unitario_estimado: float = Field(description="Valor unitário estimado ou 0.0 se sigiloso")
    codigo_catmat_siga: str = Field(description="Código CATMAT/CATMAS/SIGA se houver")

class ListaItens(BaseModel):
    itens: list[ItemEdital]

# ========================================================
# CARREGAMENTO DE BASES LOCAIS (CACHE)
# ========================================================
@st.cache_data
def carregar_portfolio():
    caminho = "portfolio_laboratorios.xlsx"
    if os.path.exists(caminho):
        try:
            return pd.read_excel(caminho)
        except Exception:
            return None
    return None

df_portfolio = carregar_portfolio()

# ========================================================
# PRÉ-PROCESSADOR CIRÚRGICO DE PDF (CALIBRADO COM 6 EDIOTAIS)
# ========================================================
@st.cache_data(show_spinner=False)
def analisar_e_filtrar_pdf(bytes_arquivo):
    leitor = PdfReader(BytesIO(bytes_arquivo))
    total_paginas = len(leitor.pages)
    
    termos_cabecalho_tabela = [
        "CATMAT", "CATMAS", "CODIGO GHC", "CODIGO SIGA", "ID SIGA", "DESCRITIVO MV",
        "PRINCIPIO ATIVO", "FORMA FARMACEUTICA", "CONCENTRACAO/DOSAGEM",
        "UNIDADE DE FORNECIMENTO", "UNIDADE DE AQUISICAO", "QUANTIDADE TOTAL",
        "ESPECIFICACAO", "VALOR UNIT", "PRECO UNIT"
    ]
    
    termos_farmaceuticos = [
        "MG/ML", "MCG", "UI/ML", "AMPOLA", "FRASCO", "FR/AMP", "F/A",
        "COMPRIMIDO", "CAPSULA", "SOLUCAO INJETAVEL", "SUSPENSAO ORAL",
        "PO LIOFILO", "PO PARA SOLUCAO", "SERINGA", "BISNAGA", "DRAGEA", "BOLSA"
    ]
    
    termos_ruido_pesado = [
        "CODIGO IBGE", "CONSUMO MEDIO MENSAL", "CONSUMO ULTIMO TRIMESTRE",
        "CADASTRO DE RESERVA", "MINUTA DE TERMO DE CONTRATO", "MINUTA DE ATA",
        "INSERIR NOME DO CONTRATADO", "R$00,0000", "R$ 0,0000"
    ]
    
    termos_juridicos = [
        "DA HABILITACAO", "DOS RECURSOS", "DAS SANCOES ADMINISTRATIVAS",
        "DA IMPUGNACAO AO EDITAL", "DO JULGAMENTO DAS PROPOSTAS",
        "OBRIGACOES DO CONTRATANTE", "OBRIGACOES DO CONTRATADO", "DO FORO"
    ]
    
    paginas_candidatas = []
    total_chars_extraidos = 0

    for i, pagina in enumerate(leitor.pages):
        raw_text = pagina.extract_text() or ""
        total_chars_extraidos += len(raw_text.strip())
        norm_text = normalizar_texto(raw_text)
        
        if not norm_text:
            continue
            
        if "CODIGO IBGE" in norm_text or ("CONSUMO MEDIO MENSAL" in norm_text and "ULTIMO PRECO" in norm_text):
            continue
            
        pts_cabecalho = sum(3 for t in termos_cabecalho_tabela if t in norm_text)
        pts_farma = sum(2 for t in termos_farmaceuticos if t in norm_text)
        
        padroes_linha_item = len(re.findall(
            r'\b(LOTE:\s*\d+|CODIGO GHC:\s*\d+|BR0\d{6,7}|ID\s*\d{4,6}|\d+\s+UNIDADE|\d+\s+FR/AMP|\d+\s+AMP|\d+\s+COMP)\b',
            norm_text
        ))
        
        penalidade_ruido = sum(8 for tr in termos_ruido_pesado if tr in norm_text)
        penalidade_jur = sum(4 for tj in termos_juridicos if tj in norm_text)
        
        score = pts_cabecalho + pts_farma + (padroes_linha_item * 3) - penalidade_ruido
        if padroes_linha_item < 2:
            score -= penalidade_jur
            
        if score >= 5 or padroes_linha_item >= 2:
            texto_enxuto = re.sub(r'[ \t]+', ' ', raw_text.strip())
            texto_enxuto = re.sub(r'\n{2,}', '\n', texto_enxuto)
            paginas_candidatas.append((i + 1, score, f"[PÁG {i+1}]\n{texto_enxuto}"))

    eh_escaneado = (total_chars_extraidos < (total_paginas * 30)) and (total_chars_extraidos < 120)

    if not eh_escaneado:
        if not paginas_candidatas:
            for i, pagina in enumerate(leitor.pages):
                txt = (pagina.extract_text() or "").strip()
                if txt:
                    paginas_candidatas.append((i + 1, 1, f"[PÁG {i+1}]\n{txt}"))

        paginas_candidatas.sort(key=lambda x: x[1], reverse=True)
        
        MAX_PAGINAS_TR = 18
        MAX_CHARS_ENTRADA = 32000
        
        escolhidas = []
        chars_acumulados = 0
        for num_pag, sc, bloco_txt in paginas_candidatas:
            if len(escolhidas) >= MAX_PAGINAS_TR:
                break
            if chars_acumulados + len(bloco_txt) > MAX_CHARS_ENTRADA and len(escolhidas) > 0:
                break
            escolhidas.append((num_pag, bloco_txt))
            chars_acumulados += len(bloco_txt)
            
        escolhidas.sort(key=lambda x: x[0])
        paginas_indices = [x[0] for x in escolhidas]
        texto_filtrado = "\n\n".join(x[1] for x in escolhidas)
    else:
        paginas_indices = []
        texto_filtrado = ""

    return eh_escaneado, texto_filtrado, paginas_indices, total_paginas

# ========================================================
# ROTEADOR MULTI-MODELO (COM SUPORTE A 8192 TOKENS DE SAÍDA)
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

    import time
    for mod in fila_modelos:
        try:
            resposta = client.models.generate_content(
                model=mod,
                contents=conteudos,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ListaItens,
                    temperature=0.1,
                    max_output_tokens=8192
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
# CRUZAMENTO INTELIGENTE: HIERARQUIA RIGOROSA PINT & SANOFI
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
                lab
