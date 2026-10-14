import os
import re
import json
import fitz  # PyMuPDF (se usar pypdf/pdfplumber, o fallback abaixo também suporta)
import pandas as pd
import streamlit as st
import google.generativeai as genai

# =====================================================================
# CONFIGURAÇÃO DA API E CONSTANTES DE LIMITE DE TOKENS
# =====================================================================
API_KEY = os.getenv("GEMINI_API_KEY") or st.secrets.get("GEMINI_API_KEY", "")
if API_KEY:
    genai.configure(api_key=API_KEY)

# Limites padrão para proteger a cota da API e evitar estouro na 1ª tentativa
PADRAO_MAX_PAGINAS = 15
PADRAO_MAX_TOKENS_ENTRADA = 6500  # ~26.000 caracteres
MAX_OUTPUT_TOKENS = 8192


# =====================================================================
# 1. EXTRAÇÃO E FILTRAGEM INTELIGENTE DE PÁGINAS (RANKING POR DENSIDADE)
# =====================================================================
def extrair_paginas_pdf(arquivo_bytes: bytes) -> list[str]:
    """Extrai o texto página por página usando PyMuPDF (fitz)."""
    paginas = []
    with fitz.open(stream=arquivo_bytes, filetype="pdf") as doc:
        for pagina in doc:
            paginas.append(pagina.get_text("text") or "")
    return paginas


def pontuar_pagina_para_itens(texto: str) -> int:
    """
    Calcula uma pontuação para identificar páginas que contêm tabelas de itens
    (Termo de Referência / Anexo de Quantitativos) e descartar páginas jurídicas.
    """
    t_lower = texto.lower()
    pontos = 0

    # Termos de alto peso (fortes indicadores de tabela de licitação)
    termos_fortes = [
        "catmat", "catser", "código br", "valor unitário", "preco unitario",
        "preço unitário", "valor total", "quantidade", "unid.", "unidade de fornecimento",
        "unidade de medida", "especificação técnica", "termo de referência", "ampla concorrência",
        "cota reservada", "item", "lote"
    ]
    for termo in termos_fortes:
        if termo in t_lower:
            pontos += 3

    # Termos específicos do nicho saúde/medicamentos e materiais
    termos_saude = [
        "mg/ml", "comprimido", "ampola", "frasco", "cápsula", "seringa",
        "concentração", "forma farmacêutica", "anvisa", "cmed"
    ]
    for termo in termos_saude:
        if termo in t_lower:
            pontos += 4

    # Densidade de números e valores monetários (ex: R$ 12,50 ou quantidades)
    ocorrencias_moeda = len(re.findall(r"r\$\s*\d+[\.,]\d+", t_lower))
    ocorrencias_numeros = len(re.findall(r"\b\d{1,6}\b", t_lower))

    if ocorrencias_moeda >= 3:
        pontos += 6
    if ocorrencias_numeros >= 15:
        pontos += 4

    # Penalizar páginas puramente jurídicas (minuta de contrato, sanções, penalidades)
    termos_juridicos = [
        "cláusula primeira", "cláusula segunda", "foro da comarca",
        "sanções administrativas", "inexecução total ou parcial", "lei nº 14.133"
    ]
    penalidade = sum(2 for tj in termos_juridicos if tj in t_lower)

    # Só penaliza se não tiver forte presença de valores monetários
    if ocorrencias_moeda < 2:
        pontos -= penalidade

    return max(0, pontos)


def filtrar_paginas_com_limite_tokens(
    paginas: list[str],
    max_paginas: int = PADRAO_MAX_PAGINAS,
    max_tokens_entrada: int = PADRAO_MAX_TOKENS_ENTRADA
) -> tuple[str, int, int, int]:
    """
    Seleciona apenas as páginas mais relevantes respeitando o teto máximo
    de páginas e o limite máximo de tokens de entrada.
    Retorna: (texto_consolidado, qtd_paginas_selecionadas, total_paginas, tokens_estimados)
    """
    total_paginas = len(paginas)
    paginas_avaliadas = []

    for idx, texto in enumerate(paginas):
        if not texto.strip():
            continue
        score = pontuar_pagina_para_itens(texto)
        if score >= 4:  # Filtro mínimo de relevância
            paginas_avaliadas.append((idx, score, texto))

    # Se nenhuma página atingiu o score mínimo, pega as páginas com mais números como fallback
    if not paginas_avaliadas:
        for idx, texto in enumerate(paginas):
            if texto.strip():
                paginas_avaliadas.append((idx, 1, texto))

    # Ordena pelas páginas mais ricas em dados de tabela (maior score primeiro)
    paginas_avaliadas.sort(key=lambda x: x[1], reverse=True)

    # Seleciona até o limite de páginas e respeita o teto de tokens (1 token ≈ 4 chars)
    selecionadas = []
    caracteres_acumulados = 0
    limite_caracteres = max_tokens_entrada * 4

    for idx, score, texto in paginas_avaliadas:
        if len(selecionadas) >= max_paginas:
            break
        if caracteres_acumulados + len(texto) > limite_caracteres and len(selecionadas) > 0:
            break
        selecionadas.append((idx, texto))
        caracteres_acumulados += len(texto)

    # Reordena na ordem natural das páginas do documento para manter a leitura contínua
    selecionadas.sort(key=lambda x: x[0])

    blocos_texto = [f"\n--- PÁGINA {idx + 1} ---\n{txt.strip()}" for idx, txt in selecionadas]
    texto_final = "\n".join(blocos_texto)
    tokens_estimados = int(len(texto_final) / 4)

    return texto_final, len(selecionadas), total_paginas, tokens_estimados


# =====================================================================
# 2. PARSER SEGURO DE JSON (RECUPERAÇÃO DE CORTES POR MAX_OUTPUT_TOKENS)
# =====================================================================
def extrair_json_seguro(texto_resposta: str) -> list[dict]:
    """
    Limpa marcações Markdown e recupera objetos JSON mesmo que a resposta
    tenha sido truncada no final por atingir o limite de tokens de saída.
    """
    if not texto_resposta:
        return []

    limpo = texto_resposta.strip()
    limpo = re.sub(r"^```(?:json)?\s*", "", limpo)
    limpo = re.sub(r"\s*```$", "", limpo)

    # Tentativa 1: Parse direto
    try:
        dados = json.loads(limpo)
        if isinstance(dados, list):
            return dados
        if isinstance(dados, dict) and "itens" in dados:
            return dados["itens"]
    except json.JSONDecodeError:
        pass

    # Tentativa 2: Se o JSON cortou no meio por limite de output tokens,
    # recupera todos os objetos completos `{...}` fechados até o ponto do corte.
    ultimo_fechamento = limpo.rfind("}")
    if ultimo_fechamento != -1 and limpo.startswith("["):
        json_reconstruido = limpo[:ultimo_fechamento + 1] + "]"
        try:
            dados = json.loads(json_reconstruido)
            if isinstance(dados, list):
                return dados
        except json.JSONDecodeError:
            pass

    return []


# =====================================================================
# 3. CHAMADA À API GEMINI COM CONTROLE DE TOKENS E TRATAMENTO DE ERROS
# =====================================================================
def extrair_itens_edital_ia(texto_filtrado: str) -> tuple[list[dict], str | None]:
    """
    Envia o texto filtrado para o Gemini com limites estritos de tokens.
    Retorna: (lista_de_itens, mensagem_de_erro)
    """
    if not API_KEY:
        return [], "Chave GEMINI_API_KEY não configurada no ambiente."

    prompt = f"""
Você é um especialista em licitações públicas brasileiras (Lei 14.133/21).
Analise as páginas selecionadas do Edital / Termo de Referência abaixo e extraia TODOS os itens licitados.

Retorne APENAS um array JSON válido onde cada objeto representa um item com a seguinte estrutura exata:
[
  {{
    "item": "01",
    "nome": "Descrição completa e princípio ativo/especificação do item",
    "unidade": "UN / CX / FR / AMP / COMP",
    "quantidade": 100.0,
    "preco_proposta_unitario": 12.50,
    "codigo_catmat": "Código CATMAT/BR se houver ou null"
  }}
]

Regras importantes:
1. Se o valor unitário estimado não estiver explícito, preencha `preco_proposta_unitario` com 0.0.
2. Converta quantidades e valores para número (float), usando ponto como separador decimal.
3. Seja conciso na descrição (`nome`) para caber o máximo de itens na resposta (máximo 180 caracteres por descrição).

TEXTO DO EDITAL:
{texto_filtrado}
"""

    try:
        model = genai.GenerativeModel("gemini-2.5-flash")
        resposta = model.generate_content(
            prompt,
            generation_config=genai.types.GenerationConfig(
                temperature=0.1,
                max_output_tokens=MAX_OUTPUT_TOKENS,
                response_mime_type="application/json"
            )
        )

        itens = extrair_json_seguro(resposta.text)
        return itens, None

    except Exception as e:
        erro_str = str(e)
        if "429" in erro_str or "quota" in erro_str.lower() or "exhausted" in erro_str.lower():
            return [], (
                "🚨 Limite de requisições/tokens da API atingido (Erro 429). "
                "Reduza o limite máximo de tokens de entrada no controle deslizante ou aguarde 60 segundos."
            )
        return [], f"Erro ao consultar a API de IA: {erro_str}"


# =====================================================================
# 4. INTERFACE STREAMLIT
# =====================================================================
def renderizar_extrator_pdf():
    st.subheader("📄 Extrator Inteligente de Editais e Termos de Referência")

    # Controles de limite de tokens visíveis e ajustáveis
    with st.expander("⚙️ Configurações de Limite de Tokens (Proteção de Cota API)", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            max_tokens_input = st.slider(
                "Limite Máximo de Tokens de Entrada",
                min_value=2000,
                max_value=15000,
                value=PADRAO_MAX_TOKENS_ENTRADA,
                step=500,
                help="Reduzir este valor evita o erro de limite de requisição (429) e foca apenas nas páginas com tabelas."
            )
        with col2:
            max_paginas_input = st.slider(
                "Máximo de Páginas Analisadas por Envio",
                min_value=3,
                max_value=30,
                value=PADRAO_MAX_PAGINAS,
                step=1
            )

    arquivo_pdf = st.file_uploader(
        "Arraste o PDF do Edital ou Termo de Referência aqui",
        type=["pdf"]
    )

    if arquivo_pdf is not None:
        bytes_pdf = arquivo_pdf.read()
        paginas_brutas = extrair_paginas_pdf(bytes_pdf)

        texto_filtrado, paginas_usadas, total_paginas, tokens_estimados = filtrar_paginas_com_limite_tokens(
            paginas_brutas,
            max_paginas=max_paginas_input,
            max_tokens_entrada=max_tokens_input
        )

        if st.button("🚀 Processar Edital", type="primary"):
            st.info(
                f"📄 **Filtro Ativo:** {paginas_usadas} de {total_paginas} página(s) selecionada(s) "
                f"(~{tokens_estimados:,} tokens de entrada — Teto configurado: {max_tokens_input:,} tokens)."
            )

            with st.spinner("Extraindo itens e estruturando tabela..."):
                itens_extraidos, erro = extrair_itens_edital_ia(texto_filtrado)

            if erro:
                st.error(erro)
            elif not itens_extraidos:
                st.warning("Nenhum item foi identificado nas páginas selecionadas. Tente aumentar o limite de páginas nas configurações acima.")
            else:
                st.success(f"✅ {len(itens_extraidos)} item(ns) identificado(s) com sucesso!")
                df_itens = pd.DataFrame(itens_extraidos)
                st.dataframe(df_itens, use_container_width=True)

                # Guarda no session_state para envio à Mesa de Operação do STRIKE
                st.session_state["strike_itens_extraidos"] = itens_extraidos


if __name__ == "__main__":
    renderizar_extrator_pdf()