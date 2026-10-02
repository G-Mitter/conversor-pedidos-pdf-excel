#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversor de Pedidos de Compra (PDF -> Excel) - Versão 5.0 com OCR Inteligente
Desenvolvido para extrair com precisão dados de PDFs digitais e escaneados/digitalizados,
reconstruindo dados que ficam misturados ou fragmentados pelo OCR e organizando-os
no padrão pré-estabelecido de planilhas Excel estruturadas por filial.

Recursos:
1. Reconstrução Espacial 2D por Coordenadas (RapidOCR e Tesseract):
   - Evita colunas e palavras misturadas agrupando caixas delimitadoras pela linha horizontal física (Y)
     e ordenando da esquerda para a direita (X).
2. Costura Inteligente de Linhas Quebradas (Stream Stitcher):
   - Une automaticamente fragmentos de descrição, códigos de barras e preços divididos em várias linhas pelo OCR.
3. Cura Automática de Ruídos de OCR:
   - Corrige confusões clássicas de OCR em números ('O' ou 'o' em vez de '0', 'l'/'|' em vez de '1').
   - Descola códigos de barras concatenados com vírgulas ou códigos ERP.
4. Identificação Semântica da Empresa e Filial:
   - Detecta CNPJ, Razão Social, Número da Filial (mesmo a partir do sufixo do CNPJ /0002- ou da sigla da loja),
     Número do Pedido, Datas, Comprador, Condição de Pagamento e Totais Gerais.
5. Exportação Formatada no Padrão Pré-Estabelecido:
   - Cabeçalho executivo por filial com metadados completos.
   - Tabela de 8 colunas com formatação monetária e de quantidades.
   - Linhas zebra, totais e rodapé de observações.
"""

import io
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

# 1. Leitura e Renderização de PDF
HAS_PYMUPDF = False
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    pass

HAS_PDFPLUMBER = False
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    pass

HAS_PYPDF = False
try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    pass

# 2. Geração de Planilhas Excel
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 3. Processamento de Imagens (PIL)
HAS_PIL = False
try:
    from PIL import Image, ImageEnhance, ImageFilter
    HAS_PIL = True
except ImportError:
    pass

# 4. Motores de OCR para PDFs Escaneados
HAS_RAPIDOCR = False
RAPIDOCR_ERRO = ""
rapid_ocr_engine = None
try:
    from rapidocr_onnxruntime import RapidOCR
    rapid_ocr_engine = RapidOCR()
    HAS_RAPIDOCR = True
except Exception as e:
    # Qualquer falha ao iniciar o OCR (ex: modelos ausentes no .exe) não pode impedir o app de abrir
    RAPIDOCR_ERRO = f"{type(e).__name__}: {e}"

HAS_TESSERACT = False
try:
    import pytesseract
    tesseract_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.getenv("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
        os.path.join(os.getenv("USERPROFILE", ""), "AppData", "Local", "Programs", "Tesseract-OCR", "tesseract.exe"),
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract"
    ]
    for p in tesseract_paths:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            HAS_TESSERACT = True
            break
except ImportError:
    pass


# ==============================================================================
# SEÇÃO 1: TRATAMENTO DE IMAGEM E EXTRAÇÃO DE OCR COM RECONSTRUÇÃO ESPACIAL 2D
# ==============================================================================

def preprocessar_imagem_ocr(pil_img):
    """
    Otimiza a imagem escaneada para OCR:
    - Escala de cinza
    - Aumento de contraste (destaca texto contra fundo)
    - Filtro de nitidez (aguça caracteres numéricos finos)
    """
    if not HAS_PIL or not pil_img:
        return pil_img
    try:
        img = pil_img.convert('L')
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.0)
        img = img.filter(ImageFilter.SHARPEN)
        return img
    except Exception:
        return pil_img


def renderizar_pagina_para_imagens(caminho_pdf, page_idx, page_obj=None):
    """
    Renderiza ou extrai a imagem da página do PDF em alta resolução (300 DPI).
    """
    imagens = []

    # 1. PyMuPDF (fitz) - Renderizador de máxima nitidez e velocidade
    if HAS_PYMUPDF and caminho_pdf:
        try:
            doc = fitz.open(caminho_pdf)
            if page_idx < len(doc):
                pg = doc[page_idx]
                # Página que é uma única imagem (ERP em nuvem / scanner): usa a imagem embutida
                # na resolução original. Reamostrar para 300 DPI borra as letras e piora o OCR.
                infos = pg.get_image_info(xrefs=True)
                if len(infos) == 1 and pg.rotation == 0 and infos[0].get("xref"):
                    a, b, c, d = infos[0]["transform"][:4]
                    cobre = abs(fitz.Rect(infos[0]["bbox"])) >= 0.8 * abs(pg.rect)
                    if cobre and a > 0 and d > 0 and abs(b) < 1e-3 and abs(c) < 1e-3:
                        try:
                            dados = doc.extract_image(infos[0]["xref"])["image"]
                            img = Image.open(io.BytesIO(dados)).convert("RGB")
                            doc.close()
                            return [img]
                        except Exception as e:
                            print(f"[PyMuPDF imagem embutida warn] {e}")
                pix = pg.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                imagens.append(img)
                doc.close()
                return imagens
        except Exception as e:
            print(f"[PyMuPDF render warn] {e}")

    # 2. pdfplumber to_image()
    if HAS_PDFPLUMBER and caminho_pdf:
        try:
            with pdfplumber.open(caminho_pdf) as pdf:
                if page_idx < len(pdf.pages):
                    im_obj = pdf.pages[page_idx].to_image(resolution=300)
                    imagens.append(im_obj.original)
                    return imagens
        except Exception as e:
            print(f"[pdfplumber render warn] {e}")

    # 3. pypdf imagens embutidas (se for PDF envelopando imagem crua)
    if HAS_PIL and page_obj and hasattr(page_obj, "images"):
        try:
            for img_file in page_obj.images:
                pil_img = Image.open(io.BytesIO(img_file.data))
                imagens.append(pil_img)
        except Exception as e:
            print(f"[pypdf images warn] {e}")

    return imagens


def agrupar_caixas_em_linhas(boxes):
    """
    Agrupa as caixas {'x_min', 'x_max', 'y_center', 'height', 'text'} que compartilham a mesma
    faixa horizontal (linha Y). Devolve as linhas do topo ao rodapé, cada uma ordenada por X.
    """
    # Ordena inicialmente por Y
    boxes.sort(key=lambda b: (b['y_center'], b['x_min']))
    rows = []

    for b in boxes:
        placed = False
        for row in rows:
            row_y_center = sum(item['y_center'] for item in row) / len(row)
            avg_h = sum(item['height'] for item in row) / len(row)
            threshold = max(9.0, avg_h * 0.55)
            if abs(b['y_center'] - row_y_center) <= threshold:
                row.append(b)
                placed = True
                break
        if not placed:
            rows.append([b])

    # Ordena linhas verticalmente
    rows.sort(key=lambda r: sum(item['y_center'] for item in r) / len(r))

    # Dentro de cada linha, ordena da esquerda para a direita (X)
    for r in rows:
        r.sort(key=lambda item: item['x_min'])
    return rows


def reconstruir_linhas_por_coordenadas(boxes):
    """
    Reconstrói as linhas de texto físicas ordenadas no espaço visual 2D:
    Recebe caixas com {'x_min', 'x_max', 'y_center', 'height', 'text'}.
    1. Agrupa palavras que compartilham a mesma faixa horizontal (linha Y).
    2. Ordena as linhas do topo até o rodapé da página.
    3. Em cada linha, ordena os termos rigorosamente da esquerda para a direita (X).
    Isso impede que colunas ou blocos sejam lidos misturados!
    """
    linhas_texto = []
    for r in agrupar_caixas_em_linhas(boxes):
        linha_str = ' '.join(item['text'] for item in r if item['text'])
        if linha_str.strip():
            linhas_texto.append(linha_str.strip())

    return linhas_texto


def caixas_rapidocr(pil_img):
    """Executa o RapidOCR e devolve as caixas de texto com suas coordenadas."""
    import numpy as np
    result, _ = rapid_ocr_engine(np.array(pil_img))
    boxes = []
    for item in result or []:
        poly = item[0]
        txt = str(item[1]).strip()
        if not txt:
            continue
        xs = [p[0] for p in poly]
        ys = [p[1] for p in poly]
        x_min = min(xs)
        x_max = max(xs)
        y_min = min(ys)
        y_max = max(ys)
        boxes.append({
            'x_min': x_min,
            'x_max': x_max,
            'y_min': y_min,
            'y_max': y_max,
            'y_center': (y_min + y_max) / 2.0,
            'height': max(y_max - y_min, 10.0),
            'text': txt
        })
    return boxes


def executar_ocr_com_reconstrucao(pil_img):
    """
    Executa OCR em uma imagem PIL e reconstrói as linhas visualmente
    evitando que os dados fiquem misturados.
    """
    # 1. Tenta RapidOCR (motor 100% Python ONNX)
    if HAS_RAPIDOCR and rapid_ocr_engine:
        try:
            boxes = caixas_rapidocr(pil_img)
            if boxes:
                linhas = reconstruir_linhas_por_coordenadas(boxes)
                texto = "\n".join(linhas)
                if len(texto.strip()) > 30:
                    return texto
        except Exception as e:
            print(f"[RapidOCR reconstrução erro] {e}")

    # 2. Tenta Tesseract OCR com image_to_data (coordenadas completas de cada palavra)
    if HAS_TESSERACT and HAS_PIL:
        try:
            img_tratada = preprocessar_imagem_ocr(pil_img)
            try:
                data = pytesseract.image_to_data(img_tratada, lang='por+eng', output_type=pytesseract.Output.DICT)
            except Exception:
                data = pytesseract.image_to_data(img_tratada, output_type=pytesseract.Output.DICT)

            boxes = []
            n_words = len(data.get('text', []))
            for idx in range(n_words):
                w_txt = str(data['text'][idx]).strip()
                if not w_txt:
                    continue
                left = data['left'][idx]
                top = data['top'][idx]
                w = data['width'][idx]
                h = data['height'][idx]
                boxes.append({
                    'x_min': left,
                    'x_max': left + w,
                    'y_min': top,
                    'y_max': top + h,
                    'y_center': top + h / 2.0,
                    'height': max(float(h), 8.0),
                    'text': w_txt
                })

            if boxes:
                linhas = reconstruir_linhas_por_coordenadas(boxes)
                texto = "\n".join(linhas)
                if len(texto.strip()) > 30:
                    return texto
        except Exception as e:
            print(f"[Tesseract coordenadas erro] {e}")

        # Fallback Tesseract simples
        try:
            img_tratada = preprocessar_imagem_ocr(pil_img)
            config_tess = "--oem 3 --psm 6"
            try:
                txt = pytesseract.image_to_string(img_tratada, lang='por+eng', config=config_tess)
            except Exception:
                txt = pytesseract.image_to_string(img_tratada, config=config_tess)
            if txt and len(txt.strip()) > 30:
                return txt
        except Exception as e:
            print(f"[Tesseract simples erro] {e}")

    return ""


# ==============================================================================
# SEÇÃO 2: RECONSTRUÇÃO NATIVA DE PDF (DOCUMENTOS DIGITAIS)
# ==============================================================================

def extrair_texto_pagina_2d(page):
    """
    Reconstrói as linhas visualmente agrupando pedaços de texto pela coordenada Y vertical
    no leitor nativo pypdf.
    """
    linhas_y = {}

    def visitor_body(text, cm, tm, fontDict, fontSize):
        if text and text.strip():
            try:
                x = round(tm[4], 1)
                y = round(tm[5], 1)
            except Exception:
                return

            matched_y = None
            for ey in linhas_y:
                if abs(ey - y) <= 3.5:
                    matched_y = ey
                    break
            if matched_y is None:
                matched_y = y
                linhas_y[matched_y] = []
            linhas_y[matched_y].append((x, text))

    try:
        page.extract_text(visitor_text=visitor_body)
        if linhas_y:
            sorted_ys = sorted(linhas_y.keys(), reverse=True)
            reconstruidas = []
            for y in sorted_ys:
                row_items = sorted(linhas_y[y], key=lambda item: item[0])
                row_text = ' '.join(item[1].strip() for item in row_items if item[1].strip())
                if row_text:
                    reconstruidas.append(row_text)
            resultado = '\n'.join(reconstruidas)
            if len(resultado.strip()) > 30:
                return resultado
    except Exception:
        pass
    return None


def caixas_texto_nativo(pdf_path, page_idx):
    """Devolve os trechos de texto nativo da página (PyMuPDF) como caixas com coordenadas."""
    doc = fitz.open(pdf_path)
    try:
        if page_idx >= len(doc):
            return []
        blocos = doc[page_idx].get_text("dict")["blocks"]
    finally:
        doc.close()

    # Converte pontos do PDF para a escala de 300 DPI esperada pela reconstrução 2D
    k = 300 / 72.0
    boxes = []
    for b in blocos:
        for l in b.get("lines", []):
            for s in l["spans"]:
                txt = s["text"].strip()
                if not txt:
                    continue
                x0, y0, x1, y1 = s["bbox"]
                boxes.append({
                    'x_min': x0 * k,
                    'x_max': x1 * k,
                    'y_center': (y0 + y1) / 2.0 * k,
                    'height': max((y1 - y0) * k, 10.0),
                    'text': txt
                })
    return boxes


def extrair_texto_pagina_trechos(pdf_path, page_idx):
    """
    Lê o texto nativo com PyMuPDF mantendo cada trecho (célula) inteiro.
    O PyMuPDF respeita o recorte das células: ERPs que escrevem a descrição completa
    e escondem o excesso não têm o texto escondido misturado com a coluna seguinte
    (ex: 'TIJUCAS 193411' + 'CX 1' virando 'TIJUCAS 1C93X3 411' no pdfplumber/pypdf).
    """
    try:
        resultado = "\n".join(reconstruir_linhas_por_coordenadas(caixas_texto_nativo(pdf_path, page_idx)))
        if len(resultado.strip()) > 40:
            return resultado
    except Exception as e:
        print(f"[PyMuPDF trechos warn] {e}")
    return None


def extrair_texto_pagina(page, pdf_path=None, page_idx=0, forcar_ocr=False, callback_log=None):
    """
    Extrai texto da página utilizando métodos nativos de layout e, se for escaneada
    ou se 'forcar_ocr' estiver ativo, aplica OCR com reconstrução espacial 2D.
    """
    def log(msg):
        if callback_log:
            callback_log(msg)

    # Se NÃO forçado a usar OCR, tenta extrações nativas digitais
    if not forcar_ocr:
        # 0. PyMuPDF por trechos (respeita o recorte das células)
        if HAS_PYMUPDF and pdf_path:
            t = extrair_texto_pagina_trechos(pdf_path, page_idx)
            if t:
                return t, "Nativo (PyMuPDF Trechos)"

        # 1. pdfplumber modo layout nativo
        if HAS_PDFPLUMBER and pdf_path:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    if page_idx < len(pdf.pages):
                        t = pdf.pages[page_idx].extract_text(layout=True)
                        if t and len(t.strip()) > 40:
                            return t, "Nativo (pdfplumber Layout)"
            except Exception:
                pass

        # 2. pypdf Reconstrução 2D
        if HAS_PYPDF and page:
            t_2d = extrair_texto_pagina_2d(page)
            if t_2d and len(t_2d.strip()) > 40:
                return t_2d, "Nativo (pypdf Reconstrução 2D)"

            # 3. pypdf modo layout
            try:
                t = page.extract_text(extraction_mode="layout")
                if t and len(t.strip()) > 40:
                    return t, "Nativo (pypdf Layout)"
            except Exception:
                pass

            # 4. pypdf padrão
            try:
                t = page.extract_text()
                if t and len(t.strip()) > 40:
                    return t, "Nativo (pypdf Padrão)"
            except Exception:
                pass

    # 5. Se chegou aqui, é um documento escaneado (ou forçado OCR)
    log(f"   [Página {page_idx+1}] Acionando Motor de OCR com Reconstrução 2D...")

    if not HAS_RAPIDOCR and not HAS_TESSERACT:
        log("   [Aviso OCR] Nenhum motor de OCR instalado para ler este documento escaneado.")
        return "", "Sem OCR instalado"

    imagens = renderizar_pagina_para_imagens(pdf_path, page_idx, page)
    if not imagens:
        log("   [Aviso OCR] Não foi possível renderizar a página em imagem.")
        return "", "Falha na renderização de imagem"

    textos_ocr = []
    for idx_img, img in enumerate(imagens):
        txt_ocr = executar_ocr_com_reconstrucao(img)
        if txt_ocr:
            textos_ocr.append(txt_ocr)

    texto_final = "\n".join(textos_ocr)
    if len(texto_final.strip()) > 30:
        log(f"   [OCR Concluído] {len(texto_final.strip())} caracteres reconstruídos espacialmente!")
        return texto_final, "OCR Inteligente (2D Reconstruído)"

    return "", "Falha no OCR"


# ==============================================================================
# SEÇÃO 3: IDENTIFICAÇÃO SEMÂNTICA, NORMALIZAÇÃO E ENCAIXE NO PADRÃO
# ==============================================================================

def limpar_numero(val_str):
    """Converte números em formato monetário ou numérico brasileiro para float."""
    if not val_str:
        return 0.0
    try:
        s = str(val_str).strip().replace('.', '').replace(',', '.')
        return float(s)
    except Exception:
        return 0.0


def normalizar_linha_ocr(texto):
    """
    Remove ruídos comuns gerados por OCR em tabelas e descola códigos aglutinados:
    - Corrige confusões de 'O' e 'o' em valores com vírgula ou decimais (ex: 12,OOO -> 12,000 / 8,5O -> 8,50)
    - Elimina espaços espúrios entre vírgula/ponto e dígitos (ex: 6, 000 -> 6,000 / 20. 36 -> 20,36 / 360, 000 -> 360,000)
    - Converte ponto decimal isolado em valores monetários (ex: 3.46 -> 3,46 / 5.50 -> 5,50), preservando dimensões (1.3L, 22CM)
    - Remove ruídos no final da descrição antes da UM (ex: 'OOOE', ':', '|')
    - Descola códigos de barras EAN colados em vírgulas ou códigos ERP
    """
    if not texto:
        return ""
    t = texto.replace('\xa0', ' ').replace('|', ' ')

    # Corrige 'O' ou 'o' no lugar de zeros numéricos decimais
    t = re.sub(r'(\d+)[,\.](\d*)([Oo]+)(\d*)', lambda m: m.group(0).replace('O', '0').replace('o', '0'), t)
    t = re.sub(r'([Oo]+)[,\.](\d+)', lambda m: m.group(0).replace('O', '0').replace('o', '0'), t)

    # Elimina espaços espúrios entre vírgula/ponto e dígitos numéricos (6, 000 -> 6,000 | 20. 36 -> 20,36 | 360, 000 -> 360,000)
    t = re.sub(r'(\d+)[,\.]\s+(\d+)', r'\1,\2', t)

    # Converte ponto isolado de preço em vírgula (ex: 3.46 -> 3,46 | 5.50 -> 5,50), sem alterar dimensões (1.3L, 2.5L)
    t = re.sub(r'(?<=\s)(\d{1,4})\.(\d{2})(?=\s|$)', r'\1,\2', t)

    # Remove ruídos de final de descrição
    t = re.sub(r'\s+OOOE\b', '', t)
    t = re.sub(r':(?=\s+[A-Z]{2}|\s+\d)', '', t)

    # Descola código de barras colado em vírgula (ex: ,007891234567895 -> ,00 7891234567895)
    t = re.sub(r',(\d{2,3})(789\d{10}|790\d{10}|\d{13})', r',\1 \2', t)

    # Descola código de barras colado em código ERP (ex: 0123457891234567895 -> 012345 7891234567895)
    t = re.sub(r'(\d{4,8})(789\d{10}|790\d{10})', r'\1 \2', t)
    t = re.sub(r'(\.\d{2})(789\d{10}|790\d{10})', r'\1 \2', t)

    return t.strip()


def analisar_codigos_prefixo(prefixo_str):
    """
    Analisa os tokens que antecedem o código de barras (EAN-13) para separar:
    - Código do Fornecedor (opcional, ex: 1234.56, 7890, 0001.00, 9A123B)
    - Código ERP da Loja (obrigatório, numérico de 4 a 8 dígitos, ex: 012345, 987654)

    Tratamentos inteligentes:
    1. Corrige ERP fragmentado pelo OCR: se o último token tiver 1 ou 2 dígitos (ex: '1')
       e o penúltimo tiver 3 a 5 dígitos (ex: '98765'), o OCR quebrou o código em dois.
       Unifica: '98765' + '4' -> '987654'.
    2. Evita duplicação: se cod_forn for idêntico a codigo_erp, limpa cod_forn = "".
    3. Se houver apenas 1 token, ele é o código ERP (cod_forn fica vazio).
    """
    if not prefixo_str:
        return "", ""
    tokens = prefixo_str.strip().split()
    if not tokens:
        return "", ""

    # Se o último token for 1 ou 2 dígitos (ex: '4') e o penúltimo tiver 3 a 5 dígitos (ex: '98765')
    if len(tokens) >= 2 and len(tokens[-1]) <= 2 and tokens[-1].isdigit() and tokens[-2].isdigit() and len(tokens[-2]) in [3, 4, 5]:
        tokens[-2] = tokens[-2] + tokens[-1]
        tokens.pop()

    # Se houver apenas 1 token
    if len(tokens) == 1:
        return "", tokens[0]

    # Se houver 2 tokens idênticos (duplicação por quebra de linha ou repetição)
    if len(tokens) == 2 and tokens[0] == tokens[1]:
        return "", tokens[1]

    codigo_erp = tokens[-1]
    cod_forn = " ".join(tokens[:-1])

    if cod_forn == codigo_erp:
        cod_forn = ""

    return cod_forn, codigo_erp


def extrair_dados_cabecalho_rodape(linhas):
    """
    Extrai com tolerância semântica todos os dados cadastrais da empresa compradora e filial,
    mesmo que as linhas do OCR estejam fora da ordem tradicional.
    """
    info = {
        'filial_cod': '',
        'ped_compra': '',
        'data_pedido': '',
        'razao_social': '',
        'cnpj': '',
        'ie': '',
        'endereco': '',
        'telefone': '',
        'email': '',
        'data_entrega': '',
        'cond_pagamento': '',
        'comprador': '',
        'quant_itens': '',
        'total_geral': '',
        'observacoes': ''
    }

    texto_completo = "\n".join(linhas)

    # 1. CNPJ (com ou sem máscara)
    m_cnpj = re.search(r'\b(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\b', texto_completo)
    if not m_cnpj:
        m_cnpj = re.search(r'CNPJ[\s:\.]*([0-9\.\/-]{14,20})', texto_completo, re.IGNORECASE)
    if m_cnpj:
        info['cnpj'] = m_cnpj.group(1).strip()
        # Se a filial não foi encontrada explicitamente, deduz do CNPJ (ex: /0002- -> Filial 002)
        m_fil_cnpj = re.search(r'/(\d{4})-', info['cnpj'])
        if m_fil_cnpj and not info['filial_cod']:
            info['filial_cod'] = str(int(m_fil_cnpj.group(1))).zfill(3)

    # 2. Código da Filial
    m_filial = re.search(r'Filial[\s:\.]*([0-9A-Za-z]+)', texto_completo, re.IGNORECASE)
    if m_filial:
        f_val = m_filial.group(1).strip()
        if f_val.isdigit():
            info['filial_cod'] = str(int(f_val)).zfill(3)
        else:
            info['filial_cod'] = f_val

    # 3. Número do Pedido de Compra
    m_ped = re.search(r'(?:Ped\.?\s*Compra|Pedido|Ordem\s*de\s*Compra|OC)[\s\.:#Nº]*([0-9]{3,12})', texto_completo, re.IGNORECASE)
    if m_ped:
        info['ped_compra'] = m_ped.group(1).strip()

    # 4. Data do Pedido
    m_data = re.search(r'Data\s*(?:do\s*)?Pedido[\s:\.]*([\d\/]{6,10})', texto_completo, re.IGNORECASE)
    if m_data:
        info['data_pedido'] = m_data.group(1).strip()

    # 5. Inscrição Estadual
    m_ie = re.search(r'Insc\.?\s*Estad\.?[\s:\.]*([\d-]+)', texto_completo, re.IGNORECASE)
    if m_ie:
        info['ie'] = m_ie.group(1).strip()

    # 6. Razão Social (busca em linhas próximas do topo com palavras-chave empresariais)
    for l in linhas[:15]:
        l_s = l.strip()
        if any(k in l_s.upper() for k in ['LTDA', 'S.A.', 'S/A', 'EIRELI', ' ME ', 'COMERCIO', 'COM.']):
            # Isola a Razão Social retirando trechos como CNPJ: ... ou Página x de y
            rz = re.split(r'CNPJ|Insc|Data|P[áa]gina|RUA|AV\.', l_s, flags=re.IGNORECASE)[0].strip()
            if len(rz) > 3 and not info['razao_social']:
                info['razao_social'] = rz
                break

    # 7. Endereço, Telefone e E-mail
    for l in linhas[:20]:
        if any(k in l.upper() for k in ["RUA ", "AV. ", "RODOVIA ", "ALAMEDA ", "CEP:"]) and not info['endereco']:
            info['endereco'] = re.sub(r'P[áa]gina\s+\d+\s+de\s+\d+', '', l).strip()
        if "Telefone:" in l and not info['telefone']:
            m_tel = re.search(r'Telefone:\s*([\d\s()-]+)', l)
            if m_tel:
                info['telefone'] = m_tel.group(1).strip()
        if "Email:" in l and not info['email']:
            m_em = re.search(r'Email:\s*([^\s]+@[^\s]+)', l)
            if m_em:
                info['email'] = m_em.group(1).strip()

    # 8. Data de Entrega
    m_ent = re.search(r'Data\s*(?:de\s*)?Entrega[\s:\.]*([\d\/]{6,10})', texto_completo, re.IGNORECASE)
    if m_ent:
        info['data_entrega'] = m_ent.group(1).strip()

    # 9. Condição de Pagamento
    m_cond = re.search(r'Condi[çc][ãa]o\s*(?:de\s*)?Pagamento[\s:\.]*([^\n\r]+)', texto_completo, re.IGNORECASE)
    if m_cond:
        info['cond_pagamento'] = m_cond.group(1).split("Despesas")[0].split("Comprador")[0].strip()

    # 10. Comprador
    m_comp = re.search(r'Comprador[\s:\.]*([A-Za-zÀ-ÿ\s]+)', texto_completo, re.IGNORECASE)
    if m_comp:
        info['comprador'] = m_comp.group(1).split("Email")[0].split("Obs")[0].strip()

    # 11. Quantidade de Itens e Total Geral do Rodapé
    m_qi = re.search(r'Quant\.?\s*Itens[\s:\.]*(\d+)', texto_completo, re.IGNORECASE)
    if m_qi:
        info['quant_itens'] = m_qi.group(1).strip()

    m_tot = re.search(r'Total\s*Geral[^0-9\n\r]*?(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})', texto_completo, re.IGNORECASE)
    if m_tot:
        info['total_geral'] = m_tot.group(1).strip()

    # 12. Observações
    m_obs = re.search(r'Observa[çc][õo]es:\s*([^\n\r]+(?:(?:\n[^\n\r]+){1,4})?)', texto_completo, re.IGNORECASE)
    if m_obs:
        obs_raw = m_obs.group(1)
        for parada in ["Peso Total:", "Total Geral:", "Quant. Itens:", "Subst. Tributária:"]:
            if parada in obs_raw:
                obs_raw = obs_raw.split(parada)[0]
        info['observacoes'] = " ".join([l.strip() for l in obs_raw.splitlines() if l.strip()])

    return info


def extrair_produtos_inteligente(linhas):
    """
    Motor semântico de extração e costura de produtos:
    - Identifica produtos em linhas completas.
    - Costura linhas quebradas pelo OCR (Stream Stitcher).
    - Descola códigos e normaliza ruídos de OCR em decimais ('O' -> '0').
    - Utiliza analisar_codigos_prefixo() para separar Cód. Fornecedor e Código ERP sem fragmentação nem duplicidade.
    - Isola a Unidade de Medida (UM) real sem confundir com medidas da descrição (1.0 LT, 2.4L, 22CM, 25,5CM).
    - Impede que quantidade e preço vazem para dentro da descrição.
    - Valida e calcula totais matemáticos cruzados (total = quant * preco_unit).
    - Entrega rigorosamente no padrão de 8 colunas pré-estabelecido.
    """
    produtos = []
    buffer_linha = ""

    for l in linhas:
        l_norm = normalizar_linha_ocr(l)
        if not l_norm:
            continue

        # Ignora cabeçalhos de tabela e linhas de metadados
        if any(k in l_norm.lower() for k in [
            'código de barras', 'codigo de barras', 'preço unit', 'preco unit',
            'tot. subs', 'turno ent', 'tipo de frete', 'ordem de compra',
            'inscrição estadual', 'insc. estad', 'condição de pagamento',
            'telefone:', 'email:', 'status: aberto', 'endereço:', 'endereco:'
        ]):
            continue

        candidata = (buffer_linha + " " + l_norm).strip() if buffer_linha else l_norm

        # 1. Busca código de barras EAN (789..., 790... ou 13 dígitos)
        ean_match = re.search(r'(789\d{10}|790\d{10}|\b\d{13}\b)', candidata)
        texto_apos_ean = candidata[ean_match.end():].strip() if ean_match else candidata

        # Procura a Unidade de Medida (UM) de faturamento oficial: UN, CX, PC, KG, FD, RL, M2, PAR, JG, PCT, BD, SC
        # Deve estar perto dos números finais ou antes de '1' / quantidade (NÃO 'L' ou 'LT' isolado no meio do nome)
        m_um = re.search(r'\b(CX|UN|PC|KG|FD|RL|M2|PAR|JG|PÇ|SC|BD|PCT)(?:\s+1|\s+01)?\b(?=\s+\d)', texto_apos_ean, re.IGNORECASE)
        if not m_um:
            m_um = re.search(r'\b(CX|UN|PC|KG|FD|RL|M2|PAR|JG|PÇ|SC|BD|PCT)\b(?=\s+\d)', texto_apos_ean, re.IGNORECASE)
        if not m_um:
            m_um = re.search(r'\b(CX|UN|PC|KG|FD|RL|M2|PAR|JG|PÇ|SC|BD|PCT)(?:\s+1|\s+01)?\b', texto_apos_ean, re.IGNORECASE)

        if m_um:
            um = m_um.group(1).upper()
            desc_bruta = texto_apos_ean[:m_um.start()].strip()
            resto_nums = texto_apos_ean[m_um.start() + len(m_um.group(0)):].strip()
        else:
            um = "UN"
            m_num_inicio = re.search(r'(?<=\s)(\d{1,4}(?:\.\d{3})*,\d{2,4}|\d+,\d{2,4})(?=\s|$)', texto_apos_ean)
            if m_num_inicio:
                desc_bruta = texto_apos_ean[:m_num_inicio.start()].strip()
                resto_nums = texto_apos_ean[m_num_inicio.start():].strip()
            else:
                desc_bruta = texto_apos_ean
                resto_nums = ""

        # Extrai números no segmento numérico final
        nums = re.findall(r'\b\d{1,4}(?:\.\d{3})*,\d{2,4}\b|\b\d+,\d{2,4}\b', resto_nums)

        # Caso A: Linha com EAN e números
        if ean_match and len(nums) >= 2:
            ean = ean_match.group(1)
            prefixo = candidata[:ean_match.start()].strip()
            cod_forn, codigo_erp = analisar_codigos_prefixo(prefixo)

            desc = re.sub(r'[:\-\|]+$', '', desc_bruta).strip()
            desc = re.sub(r'^\d+\s+', '', desc).strip()
            desc = re.sub(r'\s+', ' ', desc).strip()

            quant_str = nums[0]
            preco_str = nums[1]
            total_str = nums[-1] if len(nums) >= 3 else ""

            q_num = limpar_numero(quant_str)
            p_num = limpar_numero(preco_str)
            t_num = limpar_numero(total_str) if total_str else round(q_num * p_num, 2)

            # Validação cruzada: se o preço estiver 0 mas tivermos total e quant
            if p_num == 0.0 and q_num > 0 and t_num > 0:
                p_num = round(t_num / q_num, 2)
                preco_str = f"{p_num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

            # Se total estiver 0 e tivermos quant e preco
            if t_num == 0.0 and q_num > 0 and p_num > 0:
                t_num = round(q_num * p_num, 2)
                total_str = f"{t_num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

            if not total_str:
                total_str = f"{t_num:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

            produtos.append({
                'cod_forn': cod_forn,
                'codigo_erp': codigo_erp,
                'cod_barras': ean,
                'descricao': desc,
                'um': um,
                'quant': quant_str,
                'preco_unit': preco_str,
                'total': total_str,
                'quant_num': q_num,
                'preco_num': p_num,
                'total_num': t_num
            })
            buffer_linha = ""

        # Caso B: Linha com EAN mas ainda sem os números (fragmentada pelo OCR)
        elif ean_match and len(nums) < 2:
            buffer_linha = candidata

        # Caso C: Linha SEM EAN, mas que pode ser produto de layout alternativo
        elif not ean_match and m_um and len(nums) >= 2 and not buffer_linha:
            texto_antes = candidata[:m_um.start()].strip()
            m_cod = re.match(r'^([A-Za-z0-9\.\-\/]{2,15})\s+(.+)', texto_antes)
            if m_cod:
                cod_erp = m_cod.group(1)
                desc = m_cod.group(2).strip()
            else:
                cod_erp = ""
                desc = texto_antes

            quant_str = nums[0]
            preco_str = nums[1]
            total_str = nums[-1] if len(nums) >= 3 else ""
            q_num = limpar_numero(quant_str)
            p_num = limpar_numero(preco_str)
            t_num = limpar_numero(total_str) if total_str else round(q_num * p_num, 2)

            produtos.append({
                'cod_forn': "",
                'codigo_erp': cod_erp,
                'cod_barras': "",
                'descricao': desc,
                'um': um,
                'quant': quant_str,
                'preco_unit': preco_str,
                'total': total_str,
                'quant_num': q_num,
                'preco_num': p_num,
                'total_num': t_num
            })
            buffer_linha = ""

        # Caso D: Linha que pode ser continuação do buffer anterior
        elif not ean_match and buffer_linha:
            buffer_linha = candidata

        # Caso E: Código fornecedor/ERP isolado antes da linha com EAN
        elif not ean_match and re.match(r'^[A-Za-z0-9\.\-\/\s]{2,25}$', l_norm):
            buffer_linha = l_norm

    return produtos


# ==============================================================================
# SEÇÃO 4: PROCESSO COMPLETO DE CONVERSÃO E GERAÇÃO DO EXCEL NO PADRÃO
# ==============================================================================

def extrair_pdf_para_excel(caminho_pdf, caminho_excel, forcar_ocr=False, callback_log=None):
    """
    Processa o PDF integralmente (página por página), executa OCR com reconstrução
    espacial 2D se necessário, identifica empresas/filiais e seus produtos cadastrados
    e gera a planilha Excel no padrão pré-estabelecido.
    """
    def log(msg):
        if callback_log:
            callback_log(msg)
        print(f"[PDF Conversor] {msg}")

    if not HAS_PYPDF and not HAS_PDFPLUMBER and not HAS_PYMUPDF:
        raise RuntimeError("Instale as bibliotecas para leitura de PDF:\npip install pypdf openpyxl")

    reader = None
    total_paginas = 0

    if HAS_PYPDF:
        reader = PdfReader(caminho_pdf)
        total_paginas = len(reader.pages)
    elif HAS_PDFPLUMBER:
        with pdfplumber.open(caminho_pdf) as p:
            total_paginas = len(p.pages)
    elif HAS_PYMUPDF:
        doc = fitz.open(caminho_pdf)
        total_paginas = len(doc)
        doc.close()

    log(f"Arquivo selecionado: '{os.path.basename(caminho_pdf)}'")
    log(f"Total de páginas a processar: {total_paginas}")

    # Diagnóstico dos motores disponíveis
    if HAS_RAPIDOCR:
        log("Motor OCR primário: RapidOCR (Ativo, 100% Python ONNX)")
    elif HAS_TESSERACT:
        log("Motor OCR primário: Tesseract OCR (Ativo no Windows)")
    else:
        log("Aviso: Nenhum motor OCR detectado. Se for PDF escaneado, instale: pip install rapidocr-onnxruntime")
    if RAPIDOCR_ERRO:
        log(f"Aviso: o RapidOCR falhou ao iniciar: {RAPIDOCR_ERRO}")

    if forcar_ocr:
        log("Opção ativada: FORÇAR OCR INTELIGENTE EM TODAS AS PÁGINAS.")

    pedidos_filiais = {}
    filial_chave_atual = None
    total_caracteres_geral = 0
    paginas_com_ocr = 0

    for i in range(total_paginas):
        page = reader.pages[i] if (reader and i < len(reader.pages)) else None
        texto, modo_extracao = extrair_texto_pagina(
            page, caminho_pdf, i, forcar_ocr=forcar_ocr, callback_log=log
        )
        chars = len(texto.strip()) if texto else 0
        total_caracteres_geral += chars

        if "OCR" in modo_extracao:
            paginas_com_ocr += 1

        log(f"Pág. {i+1}: {chars} caracteres obtidos via [{modo_extracao}].")

        if not texto or chars < 10:
            continue

        linhas = [l.strip() for l in texto.split('\n') if l.strip()]

        # 1. Identifica dados cadastrais da empresa/filial da página
        info_pagina = extrair_dados_cabecalho_rodape(linhas)

        # 2. Determina a chave identificadora da filial
        chave_candidata = ""
        if info_pagina['filial_cod']:
            chave_candidata = f"FILIAL_{info_pagina['filial_cod']}"
            if info_pagina['ped_compra']:
                chave_candidata += f"_PED_{info_pagina['ped_compra']}"
        elif info_pagina['cnpj']:
            chave_candidata = f"CNPJ_{info_pagina['cnpj']}"
        elif info_pagina['razao_social']:
            chave_candidata = f"RZ_{info_pagina['razao_social'][:20]}"

        if chave_candidata:
            filial_chave_atual = chave_candidata
            if filial_chave_atual not in pedidos_filiais:
                log(f"-> Filial Identificada na Pág. {i+1}: {filial_chave_atual} ({info_pagina['razao_social']})")
                pedidos_filiais[filial_chave_atual] = {
                    'filial_cod': info_pagina['filial_cod'],
                    'ped_compra': info_pagina['ped_compra'],
                    'data_pedido': info_pagina['data_pedido'],
                    'razao_social': info_pagina['razao_social'] or (f"Filial {info_pagina['filial_cod']}" if info_pagina['filial_cod'] else "Filial"),
                    'cnpj': info_pagina['cnpj'],
                    'ie': info_pagina['ie'],
                    'endereco': info_pagina['endereco'],
                    'telefone': info_pagina['telefone'],
                    'email': info_pagina['email'],
                    'data_entrega': info_pagina['data_entrega'],
                    'cond_pagamento': info_pagina['cond_pagamento'],
                    'comprador': info_pagina['comprador'],
                    'quant_itens': info_pagina['quant_itens'],
                    'total_geral': info_pagina['total_geral'],
                    'observacoes': info_pagina['observacoes'],
                    'produtos': []
                }
            else:
                for k, v in info_pagina.items():
                    if v and not pedidos_filiais[filial_chave_atual].get(k):
                        pedidos_filiais[filial_chave_atual][k] = v

        if not filial_chave_atual:
            filial_chave_atual = "FILIAL_001"
            pedidos_filiais[filial_chave_atual] = {
                'filial_cod': '001',
                'ped_compra': '',
                'data_pedido': '',
                'razao_social': 'Filial 001',
                'cnpj': '',
                'ie': '',
                'endereco': '',
                'telefone': '',
                'email': '',
                'data_entrega': '',
                'cond_pagamento': '',
                'comprador': '',
                'quant_itens': '',
                'total_geral': '',
                'observacoes': '',
                'produtos': []
            }

        # Atualiza totais de rodapé se presentes nesta página
        for k in ['quant_itens', 'total_geral', 'data_entrega', 'cond_pagamento', 'comprador', 'observacoes']:
            if info_pagina.get(k):
                pedidos_filiais[filial_chave_atual][k] = info_pagina[k]

        # 3. Extrai produtos cadastrados da página com o motor semântico
        novos_produtos = extrair_produtos_inteligente(linhas)
        if novos_produtos:
            pedidos_filiais[filial_chave_atual]['produtos'].extend(novos_produtos)
            log(f"   -> {len(novos_produtos)} produtos estruturados e validados no padrão!")
        else:
            log("   -> Nenhum produto detectado nesta página.")

    if total_caracteres_geral == 0:
        raise RuntimeError(
            "O arquivo PDF é uma imagem escaneada e nenhum motor de OCR conseguiu ler o conteúdo.\n\n"
            "Solução Simples no Terminal (CMD):\n"
            "pip install pymupdf rapidocr-onnxruntime pillow\n\n"
            "Isso ativa o OCR de alta precisão 100% automático no Python!"
        )

    # ==========================================================================
    # GERAÇÃO DA PLANILHA EXCEL FORMATADA (PADRÃO PRÉ-ESTABELECIDO)
    # ==========================================================================
    wb = Workbook()
    ws = wb.active
    ws.title = "Pedidos por Filial"

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    fonte_titulo = Font(name="Calibri", size=11, bold=True, color="1F497D")
    fonte_negrito = Font(name="Calibri", size=10, bold=True)
    fonte_normal = Font(name="Calibri", size=10)
    fonte_cabecalho_tab = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    fill_cabecalho_tab = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
    fill_divisor = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    fill_suspeito = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fonte_suspeito = Font(name="Calibri", size=10, bold=True, color="9C0006")
    avisos = 0

    borda_fina = Border(
        left=Side(style='thin', color="BFBFBF"),
        right=Side(style='thin', color="BFBFBF"),
        top=Side(style='thin', color="BFBFBF"),
        bottom=Side(style='thin', color="BFBFBF")
    )
    borda_divisor = Border(
        top=Side(style='medium', color="1F497D"),
        bottom=Side(style='medium', color="1F497D")
    )

    num_linha = 1
    total_filiais_processadas = 0

    for chave, filial in pedidos_filiais.items():
        if not filial['produtos']:
            continue

        total_filiais_processadas += 1

        # 1. TÍTULO DO BLOCO DA FILIAL
        rotulo_filial = f"PEDIDO DE COMPRA - FILIAL {filial['filial_cod']} ({filial['razao_social']})"
        ws.cell(row=num_linha, column=1, value=rotulo_filial).font = fonte_titulo
        num_linha += 1

        # 2. METADADOS DA EMPRESA E PEDIDO
        metadados = [
            ("Pedido de Compra:", filial['ped_compra'], "Data Pedido:", filial['data_pedido'], "Data Entrega:", filial['data_entrega']),
            ("Razão Social:", filial['razao_social'], "CNPJ:", filial['cnpj'], "Insc. Estadual:", filial['ie']),
            ("Endereço:", filial['endereco'], "Telefone:", filial['telefone'], "Email:", filial['email']),
            ("Comprador:", filial['comprador'], "Cond. Pagamento:", filial['cond_pagamento'], "", "")
        ]

        for meta in metadados:
            c1, v1, c2, v2, c3, v3 = meta
            ws.cell(row=num_linha, column=1, value=c1).font = fonte_negrito
            ws.cell(row=num_linha, column=2, value=v1).font = fonte_normal
            if c2:
                ws.cell(row=num_linha, column=3, value=c2).font = fonte_negrito
                ws.cell(row=num_linha, column=4, value=v2).font = fonte_normal
            if c3:
                ws.cell(row=num_linha, column=5, value=c3).font = fonte_negrito
                ws.cell(row=num_linha, column=6, value=v3).font = fonte_normal
            num_linha += 1

        num_linha += 1  # Espaço antes da tabela

        # 3. CABEÇALHO DA TABELA DE PRODUTOS (8 COLUNAS DO PADRÃO)
        colunas_tabela = [
            "Cód. Forn.",
            "Código ERP",
            "Código de Barras",
            "Descrição do Produto",
            "UM",
            "Quant.",
            "Preço Unit.",
            "Total"
        ]

        for col_idx, col_nome in enumerate(colunas_tabela, start=1):
            cell = ws.cell(row=num_linha, column=col_idx, value=col_nome)
            cell.font = fonte_cabecalho_tab
            cell.fill = fill_cabecalho_tab
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = borda_fina
        num_linha += 1

        # 4. LINHAS DE PRODUTOS
        soma_quant = 0.0
        soma_total = 0.0

        for p_idx, prod in enumerate(filial['produtos']):
            bg = fill_zebra if (p_idx % 2 == 1) else None

            # Validação da linha: quantidade x preço tem que bater com o total lido
            # ponytail: tolerância fixa de R$ 0,05; se o ERP aplicar desconto/frete por item, descontar aqui
            if abs(prod['quant_num'] * prod['preco_num'] - prod['total_num']) > 0.05:
                bg = fill_suspeito
                avisos += 1
                ws.cell(row=num_linha, column=9, value="CONFERIR: quant. x preço ≠ total").font = fonte_suspeito

            # Cód. Forn
            c = ws.cell(row=num_linha, column=1, value=prod['cod_forn'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = borda_fina
            if bg: c.fill = bg

            # Código ERP
            c = ws.cell(row=num_linha, column=2, value=prod['codigo_erp'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = borda_fina
            if bg: c.fill = bg

            # Código de Barras (EAN)
            c = ws.cell(row=num_linha, column=3, value=prod['cod_barras'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borda_fina
            if bg: c.fill = bg

            # Descrição do Produto
            c = ws.cell(row=num_linha, column=4, value=prod['descricao'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = borda_fina
            if bg: c.fill = bg

            # UM
            c = ws.cell(row=num_linha, column=5, value=prod['um'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borda_fina
            if bg: c.fill = bg

            # Quantidade
            c = ws.cell(row=num_linha, column=6, value=prod['quant_num'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = '#,##0.000'
            c.border = borda_fina
            if bg: c.fill = bg
            soma_quant += prod['quant_num']

            # Preço Unitário
            c = ws.cell(row=num_linha, column=7, value=prod['preco_num'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = 'R$ #,##0.00'
            c.border = borda_fina
            if bg: c.fill = bg

            # Total
            c = ws.cell(row=num_linha, column=8, value=prod['total_num'])
            c.font = fonte_normal
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.number_format = 'R$ #,##0.00'
            c.border = borda_fina
            if bg: c.fill = bg
            soma_total += prod['total_num']

            num_linha += 1

        # 5. LINHA DE SUBTOTAIS DA FILIAL
        c = ws.cell(row=num_linha, column=4, value="TOTAL DA FILIAL:")
        c.font = fonte_negrito
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = borda_fina

        c = ws.cell(row=num_linha, column=6, value=soma_quant)
        c.font = fonte_negrito
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.number_format = '#,##0.000'
        c.border = borda_fina

        c = ws.cell(row=num_linha, column=8, value=soma_total)
        c.font = fonte_negrito
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.number_format = 'R$ #,##0.00'
        c.border = borda_fina

        num_linha += 1

        # 6. RODAPÉ DE RESUMO
        num_linha += 1
        ws.cell(row=num_linha, column=1, value="Total de Itens Cadastrados:").font = fonte_negrito
        ws.cell(row=num_linha, column=2, value=len(filial['produtos'])).font = fonte_normal

        if filial['total_geral']:
            ws.cell(row=num_linha, column=4, value="Total Geral Informado no PDF:").font = fonte_negrito
            ws.cell(row=num_linha, column=5, value=f"R$ {filial['total_geral']}").font = fonte_normal
            # Validação da filial: soma das linhas tem que bater com o Total Geral do PDF
            if abs(soma_total - limpar_numero(filial['total_geral'])) > 0.05:
                avisos += 1
                c = ws.cell(row=num_linha, column=6, value="CONFERIR: soma das linhas ≠ Total Geral do PDF")
                c.font = fonte_suspeito
                c.fill = fill_suspeito
                log(f"ATENÇÃO: Filial {filial['filial_cod']}: soma das linhas {soma_total:.2f} ≠ Total Geral do PDF {filial['total_geral']}")
        else:
            # Sem o Total Geral não há como conferir a soma: avisa em vez de passar em silêncio
            avisos += 1
            c = ws.cell(row=num_linha, column=4, value="CONFERIR: Total Geral não encontrado no PDF")
            c.font = fonte_suspeito
            c.fill = fill_suspeito
            log(f"ATENÇÃO: Filial {filial['filial_cod']}: Total Geral não encontrado no PDF, soma não conferida.")
        num_linha += 1

        if filial['observacoes']:
            ws.cell(row=num_linha, column=1, value="Observações:").font = fonte_negrito
            ws.cell(row=num_linha, column=2, value=filial['observacoes']).font = fonte_normal
            num_linha += 1

        # 7. LINHA DIVISÓRIA ENTRE FILIAIS
        num_linha += 1
        for col_idx in range(1, 9):
            cell = ws.cell(row=num_linha, column=col_idx, value="")
            cell.fill = fill_divisor
            cell.border = borda_divisor
        num_linha += 2

    if total_filiais_processadas == 0:
        raise RuntimeError(
            "Nenhum produto ou filial pôde ser identificado no documento.\n"
            "Dica: Marque a opção 'Forçar OCR Inteligente em Todas as Páginas' se o documento foi escaneado."
        )

    # 8. AJUSTE AUTOMÁTICO DE LARGURA DAS COLUNAS
    larguras_minimas = {
        'A': 14,  # Cód. Forn.
        'B': 14,  # Código ERP
        'C': 18,  # Código de Barras
        'D': 42,  # Descrição do Produto
        'E': 8,   # UM
        'F': 14,  # Quant.
        'G': 14,  # Preço Unit.
        'H': 16   # Total
    }

    for col_letter, min_w in larguras_minimas.items():
        col_idx = ord(col_letter) - ord('A') + 1
        max_len = min_w
        for r in range(1, min(num_linha, 500)):
            val = ws.cell(row=r, column=col_idx).value
            if val is not None:
                max_len = max(max_len, min(len(str(val)) + 2, 60))
        ws.column_dimensions[col_letter].width = max_len

    # Salva planilha
    wb.save(caminho_excel)
    log(f"Sucesso! Planilha gerada com {total_filiais_processadas} filiais em: {caminho_excel}")
    if avisos:
        log(f"ATENÇÃO: {avisos} ponto(s) a conferir, marcados em vermelho na planilha.")
    return total_filiais_processadas, avisos


# ==============================================================================
# SEÇÃO 4B: CONVERSOR GENÉRICO (QUALQUER PDF -> EXCEL)
# ==============================================================================

def detectar_colunas(linhas):
    """
    Descobre as colunas da página pelas faixas verticais em branco: uma posição X pertence
    a uma coluna quando há texto nela em mais de 20% das linhas com mais de um trecho.
    Devolve a lista de faixas (x_inicio, x_fim), da esquerda para a direita.
    """
    # ponytail: uma grade de colunas por página; tabelas diferentes na mesma página dividem
    # a mesma grade. Se precisar, detectar colunas por bloco de linhas consecutivas.
    tabela = [l for l in linhas if len(l) > 1]
    if not tabela:
        return [(0, 1)]
    largura = int(max(b['x_max'] for l in tabela for b in l)) + 2
    ocupacao = [0] * largura
    for l in tabela:
        coberto = bytearray(largura)
        for b in l:
            x0 = max(int(b['x_min']), 0)
            coberto[x0:int(b['x_max']) + 1] = b'\x01' * (int(b['x_max']) + 1 - x0)
        for x in range(largura):
            ocupacao[x] += coberto[x]

    limite = 0.2 * len(tabela)
    alturas = sorted(b['height'] for l in tabela for b in l)
    vao_minimo = 0.6 * alturas[len(alturas) // 2]

    colunas, inicio, fim = [], None, 0
    for x in range(largura):
        if ocupacao[x] > limite:
            if inicio is not None and x - fim > vao_minimo:
                colunas.append((inicio, fim))
                inicio = None
            if inicio is None:
                inicio = x
            fim = x
    if inicio is not None:
        colunas.append((inicio, fim))
    return colunas or [(0, largura)]


def coluna_da_caixa(b, colunas):
    """Índice da coluna da caixa: a primeira que cobre 25% dela, senão a de maior sobreposição."""
    largura = max(b['x_max'] - b['x_min'], 1.0)
    sobre = [max(0.0, min(b['x_max'], c1) - max(b['x_min'], c0)) for c0, c1 in colunas]
    for i, s in enumerate(sobre):
        if s >= 0.25 * largura:
            return i
    if max(sobre) > 0:
        return sobre.index(max(sobre))
    # Caixa fora de todas as faixas: coluna mais próxima
    centro = (b['x_min'] + b['x_max']) / 2.0
    return min(range(len(colunas)), key=lambda i: abs((colunas[i][0] + colunas[i][1]) / 2.0 - centro))


def valor_celula(txt):
    """Número no formato brasileiro com vírgula (1.245,60) vira número; o resto fica texto."""
    # O OCR costuma soltar espaços dentro do número ("15, 10")
    limpo = re.sub(r'\s', '', txt) if re.fullmatch(r'-?[\d.,\s]+', txt) else txt
    if re.fullmatch(r'-?(\d{1,3}(\.\d{3})+|\d+),\d+', limpo):
        return float(limpo.replace('.', '').replace(',', '.'))
    return txt


def converter_pdf_generico(caminho_pdf, caminho_excel, forcar_ocr=False, aba_por_pagina=True, callback_log=None):
    """
    Converte qualquer PDF em Excel mantendo a disposição visual: cada linha do PDF vira uma
    linha da planilha e as colunas são detectadas pelos espaços em branco da página.
    Páginas com texto são lidas direto; páginas em imagem passam pelo OCR.
    """
    def log(msg):
        if callback_log:
            callback_log(msg)
        print(f"[PDF Genérico] {msg}")

    if not HAS_PYMUPDF:
        raise RuntimeError("A conversão genérica precisa do PyMuPDF:\npip install pymupdf")

    doc = fitz.open(caminho_pdf)
    total_paginas = len(doc)
    doc.close()
    log(f"Arquivo selecionado: '{os.path.basename(caminho_pdf)}' ({total_paginas} páginas)")
    if RAPIDOCR_ERRO:
        log(f"Aviso: o RapidOCR falhou ao iniciar: {RAPIDOCR_ERRO}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Pág 1" if aba_por_pagina else "PDF"
    fonte = Font(name="Calibri", size=10)
    num_linha = 1
    paginas_com_conteudo = 0

    for i in range(total_paginas):
        boxes = [] if forcar_ocr else caixas_texto_nativo(caminho_pdf, i)
        modo = "texto"
        if sum(len(b['text']) for b in boxes) < 40:
            modo = "OCR"
            if HAS_RAPIDOCR and rapid_ocr_engine:
                imagens = renderizar_pagina_para_imagens(caminho_pdf, i)
                boxes = caixas_rapidocr(imagens[0]) if imagens else []
            elif not boxes:
                log(f"Pág. {i+1}: sem texto e sem motor de OCR disponível.")

        if aba_por_pagina and i > 0:
            ws = wb.create_sheet(f"Pág {i+1}")
            num_linha = 1

        linhas = agrupar_caixas_em_linhas(boxes)
        if not linhas:
            log(f"Pág. {i+1}: nada encontrado.")
            continue
        paginas_com_conteudo += 1
        colunas = detectar_colunas(linhas)
        log(f"Pág. {i+1}: {len(linhas)} linhas e {len(colunas)} colunas via {modo}.")

        for linha in linhas:
            celulas = {}
            for b in linha:
                celulas.setdefault(coluna_da_caixa(b, colunas), []).append(b['text'])
            for col, textos in celulas.items():
                ws.cell(row=num_linha, column=col + 1, value=valor_celula(' '.join(textos))).font = fonte
            num_linha += 1
        num_linha += 1  # linha em branco entre páginas quando tudo vai na mesma aba

    if paginas_com_conteudo == 0:
        raise RuntimeError("Nenhum conteúdo pôde ser lido do PDF (sem texto e sem OCR disponível).")

    for aba in wb.worksheets:
        for col in aba.columns:
            maior = max((len(str(c.value)) for c in col if c.value is not None), default=0)
            aba.column_dimensions[get_column_letter(col[0].column)].width = min(max(maior + 2, 6), 60)

    wb.save(caminho_excel)
    log(f"Sucesso! {paginas_com_conteudo} página(s) convertida(s) em: {caminho_excel}")
    return paginas_com_conteudo


# ==============================================================================
# SEÇÃO 5: INTERFACE GRÁFICA TKINTER MODERNA
# ==============================================================================

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de Pedidos PDF -> Excel | Versão 5.0 com OCR Inteligente")
        self.root.geometry("820x640")
        self.root.minsize(700, 520)
        # Ícone da janela: dentro do .exe ele fica na pasta temporária do PyInstaller (_MEIPASS)
        pasta = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        try:
            self.root.iconbitmap(os.path.join(pasta, "icone.ico"))
        except tk.TclError:
            pass

        self.caminho_pdf = tk.StringVar()
        self.caminho_excel = tk.StringVar()
        self.forcar_ocr = tk.BooleanVar(value=False)
        self.is_processing = False

        self._criar_layout()

    def _criar_layout(self):
        # 1. CABEÇALHO SUPERIOR
        header_frame = tk.Frame(self.root, bg="#1F497D", padx=16, pady=12)
        header_frame.pack(fill=tk.X)

        lbl_titulo = tk.Label(
            header_frame,
            text="Conversor de Pedidos de Compra (PDF -> Excel)",
            font=("Segoe UI", 14, "bold"),
            fg="#FFFFFF",
            bg="#1F497D"
        )
        lbl_titulo.pack(anchor=tk.W)

        lbl_sub = tk.Label(
            header_frame,
            text="Versão 5.0: Reconstrução Espacial 2D de Dados Misturados & OCR com Costura de Produtos",
            font=("Segoe UI", 9),
            fg="#DCE6F1",
            bg="#1F497D"
        )
        lbl_sub.pack(anchor=tk.W)

        # 2. CONTAINER PRINCIPAL
        abas = ttk.Notebook(self.root)
        abas.pack(fill=tk.BOTH, expand=True)
        aba_generica = tk.Frame(abas, padx=16, pady=12)
        abas.add(aba_generica, text="  Qualquer PDF  ")
        self._criar_aba_generica(aba_generica)
        main_frame = tk.Frame(abas, padx=16, pady=12)
        abas.add(main_frame, text="  Pedido de Compra  ")

        # SELETOR DE ARQUIVOS
        file_frame = tk.LabelFrame(main_frame, text=" Arquivos ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # PDF Input
        lbl_pdf = tk.Label(file_frame, text="Arquivo PDF:", font=("Segoe UI", 9))
        lbl_pdf.grid(row=0, column=0, sticky=tk.W, pady=4)
        entry_pdf = tk.Entry(file_frame, textvariable=self.caminho_pdf, font=("Segoe UI", 9))
        entry_pdf.grid(row=0, column=1, sticky=tk.EW, padx=8, pady=4)
        btn_pdf = tk.Button(file_frame, text="Selecionar PDF...", command=self._selecionar_pdf, width=15)
        btn_pdf.grid(row=0, column=2, pady=4)

        # Excel Output
        lbl_excel = tk.Label(file_frame, text="Salvar Excel em:", font=("Segoe UI", 9))
        lbl_excel.grid(row=1, column=0, sticky=tk.W, pady=4)
        entry_excel = tk.Entry(file_frame, textvariable=self.caminho_excel, font=("Segoe UI", 9))
        entry_excel.grid(row=1, column=1, sticky=tk.EW, padx=8, pady=4)
        btn_excel = tk.Button(file_frame, text="Alterar Local...", command=self._selecionar_excel, width=15)
        btn_excel.grid(row=1, column=2, pady=4)

        file_frame.columnconfigure(1, weight=1)

        # OPÇÕES DO MOTOR OCR
        opts_frame = tk.LabelFrame(main_frame, text=" Opções do Motor OCR & Reconstrução ", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        opts_frame.pack(fill=tk.X, pady=(0, 10))

        chk_ocr = tk.Checkbutton(
            opts_frame,
            text="Forçar OCR Inteligente em Todas as Páginas (recomendado para documentos escaneados/fotos)",
            variable=self.forcar_ocr,
            font=("Segoe UI", 9)
        )
        chk_ocr.pack(anchor=tk.W)

        # BOTÕES DE AÇÃO
        act_frame = tk.Frame(main_frame)
        act_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_inspecionar = tk.Button(
            act_frame,
            text="🔍 Inspecionar Leitura / OCR da Pág 1",
            command=self._inspecionar_pagina_1,
            font=("Segoe UI", 9),
            padx=10,
            pady=6
        )
        self.btn_inspecionar.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_converter = tk.Button(
            act_frame,
            text="⚡ Exportar para Excel (.xlsx)",
            command=self._iniciar_conversao,
            font=("Segoe UI", 10, "bold"),
            bg="#1F497D",
            fg="#FFFFFF",
            activebackground="#16365C",
            activeforeground="#FFFFFF",
            padx=16,
            pady=6
        )
        self.btn_converter.pack(side=tk.RIGHT)

        # BARRA DE PROGRESSO
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        # LOG EM TEMPO REAL
        log_frame = tk.LabelFrame(main_frame, text=" Progresso e Diagnóstico em Tempo Real ", font=("Segoe UI", 9, "bold"), padx=8, pady=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.txt_log = scrolledtext.ScrolledText(log_frame, font=("Consolas", 9), state=tk.DISABLED, bg="#F9FAFB")
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        # Status inicial do sistema
        self._log("=== SISTEMA PRONTO ===")
        self._log(f"- Leitor Nativo (pypdf): {'Disponível' if HAS_PYPDF else 'Não instalado'}")
        self._log(f"- Renderizador (PyMuPDF): {'Disponível' if HAS_PYMUPDF else 'Não instalado (opcional)'}")
        self._log(f"- Motor OCR RapidOCR: {'Disponível (100% Python)' if HAS_RAPIDOCR else 'Não instalado'}")
        self._log(f"- Motor OCR Tesseract: {'Disponível no Windows' if HAS_TESSERACT else 'Não configurado'}")
        self._log("Selecione um arquivo PDF acima para iniciar.")

    def _criar_aba_generica(self, aba):
        """Segunda aba: converte qualquer PDF em Excel mantendo linhas e colunas da página."""
        self.gen_pdf = tk.StringVar()
        self.gen_excel = tk.StringVar()
        self.gen_forcar_ocr = tk.BooleanVar(value=False)
        self.gen_aba_por_pagina = tk.BooleanVar(value=True)

        tk.Label(
            aba,
            text="Converte qualquer PDF (relatórios, extratos, tabelas, PDFs escaneados) mantendo as linhas e colunas da página.",
            font=("Segoe UI", 9), anchor=tk.W, justify=tk.LEFT, wraplength=760
        ).pack(fill=tk.X, pady=(0, 8))

        file_frame = tk.LabelFrame(aba, text=" Arquivos ", font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Label(file_frame, text="Arquivo PDF:", font=("Segoe UI", 9)).grid(row=0, column=0, sticky=tk.W, pady=4)
        tk.Entry(file_frame, textvariable=self.gen_pdf, font=("Segoe UI", 9)).grid(row=0, column=1, sticky=tk.EW, padx=8, pady=4)
        tk.Button(file_frame, text="Selecionar PDF...", command=self._gen_selecionar_pdf, width=15).grid(row=0, column=2, pady=4)
        tk.Label(file_frame, text="Salvar Excel em:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky=tk.W, pady=4)
        tk.Entry(file_frame, textvariable=self.gen_excel, font=("Segoe UI", 9)).grid(row=1, column=1, sticky=tk.EW, padx=8, pady=4)
        tk.Button(file_frame, text="Alterar Local...", command=self._gen_selecionar_excel, width=15).grid(row=1, column=2, pady=4)
        file_frame.columnconfigure(1, weight=1)

        opts_frame = tk.LabelFrame(aba, text=" Opções ", font=("Segoe UI", 9, "bold"), padx=10, pady=6)
        opts_frame.pack(fill=tk.X, pady=(0, 10))
        tk.Checkbutton(opts_frame, text="Uma aba da planilha para cada página (desmarcado: tudo em uma aba só)",
                       variable=self.gen_aba_por_pagina, font=("Segoe UI", 9)).pack(anchor=tk.W)
        tk.Checkbutton(opts_frame, text="Forçar OCR em todas as páginas (use se o texto sair embaralhado)",
                       variable=self.gen_forcar_ocr, font=("Segoe UI", 9)).pack(anchor=tk.W)

        self.gen_btn = tk.Button(
            aba, text="⚡ Converter para Excel (.xlsx)", command=self._gen_converter,
            font=("Segoe UI", 10, "bold"), bg="#1F497D", fg="#FFFFFF",
            activebackground="#16365C", activeforeground="#FFFFFF", padx=16, pady=6
        )
        self.gen_btn.pack(anchor=tk.E, pady=(0, 10))

        self.gen_progress = ttk.Progressbar(aba, mode='indeterminate')
        self.gen_progress.pack(fill=tk.X, pady=(0, 8))

        log_frame = tk.LabelFrame(aba, text=" Progresso ", font=("Segoe UI", 9, "bold"), padx=8, pady=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.gen_txt_log = scrolledtext.ScrolledText(log_frame, font=("Consolas", 9), state=tk.DISABLED, bg="#F9FAFB")
        self.gen_txt_log.pack(fill=tk.BOTH, expand=True)

    def _gen_log(self, msg):
        self.gen_txt_log.config(state=tk.NORMAL)
        self.gen_txt_log.insert(tk.END, msg + "\n")
        self.gen_txt_log.see(tk.END)
        self.gen_txt_log.config(state=tk.DISABLED)

    def _gen_selecionar_pdf(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o PDF",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os Arquivos", "*.*")]
        )
        if caminho:
            self.gen_pdf.set(caminho)
            self.gen_excel.set(os.path.splitext(caminho)[0] + ".xlsx")
            self._gen_log(f"PDF carregado: {os.path.basename(caminho)}")

    def _gen_selecionar_excel(self):
        caminho = filedialog.asksaveasfilename(
            title="Salvar Planilha Excel",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel (*.xlsx)", "*.xlsx")]
        )
        if caminho:
            self.gen_excel.set(caminho)

    def _gen_converter(self):
        pdf_path = self.gen_pdf.get().strip()
        excel_path = self.gen_excel.get().strip()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo PDF existente!")
            return
        if not excel_path:
            messagebox.showwarning("Atenção", "Defina o local onde o Excel será salvo!")
            return

        self.gen_btn.config(state=tk.DISABLED)
        self.gen_progress.start(10)
        self._gen_log("\n--- INICIANDO CONVERSÃO ---")

        def task():
            try:
                paginas = converter_pdf_generico(
                    pdf_path, excel_path,
                    forcar_ocr=self.gen_forcar_ocr.get(),
                    aba_por_pagina=self.gen_aba_por_pagina.get(),
                    callback_log=self._gen_log
                )

                def sucesso():
                    if messagebox.askyesno("Sucesso!", f"{paginas} página(s) convertida(s).\n\nArquivo salvo em:\n{excel_path}\n\nDeseja abrir a planilha agora?"):
                        try:
                            os.startfile(excel_path)
                        except Exception as ex:
                            messagebox.showinfo("Aviso", f"Não foi possível abrir o arquivo automaticamente: {ex}")

                self.root.after(0, sucesso)
            except Exception as e:
                self._gen_log(f"[ERRO] {e}")
                self.root.after(0, lambda: messagebox.showerror("Erro na Conversão", str(e)))
            finally:
                self.gen_btn.config(state=tk.NORMAL)
                self.gen_progress.stop()

        threading.Thread(target=task, daemon=True).start()

    def _log(self, msg):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def _selecionar_pdf(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o Pedido de Compra em PDF",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os Arquivos", "*.*")]
        )
        if caminho:
            self.caminho_pdf.set(caminho)
            padrao_excel = os.path.splitext(caminho)[0] + "_Filiais.xlsx"
            self.caminho_excel.set(padrao_excel)
            self._log(f"PDF carregado: {os.path.basename(caminho)}")

    def _selecionar_excel(self):
        caminho = filedialog.asksaveasfilename(
            title="Salvar Planilha Excel",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel (*.xlsx)", "*.xlsx")]
        )
        if caminho:
            self.caminho_excel.set(caminho)

    def _inspecionar_pagina_1(self):
        pdf_path = self.caminho_pdf.get().strip()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showwarning("Atenção", "Selecione um arquivo PDF válido primeiro!")
            return

        def task():
            self.btn_inspecionar.config(state=tk.DISABLED)
            self.progress_bar.start(10)
            self._log("--- Inspecionando Página 1 com Reconstrução 2D ---")

            try:
                reader = PdfReader(pdf_path) if HAS_PYPDF else None
                page = reader.pages[0] if (reader and len(reader.pages) > 0) else None
                texto, modo = extrair_texto_pagina(
                    page, pdf_path, 0, forcar_ocr=self.forcar_ocr.get(), callback_log=self._log
                )

                linhas = [l.strip() for l in texto.splitlines() if l.strip()]
                meta = extrair_dados_cabecalho_rodape(linhas)
                prods = extrair_produtos_inteligente(linhas)

                def show_popup():
                    win = tk.Toplevel(self.root)
                    win.title(f"Diagnóstico da Página 1 [{modo}]")
                    win.geometry("780x540")

                    lbl = tk.Label(
                        win,
                        text=f"Filial: {meta['filial_cod']} | Pedido: {meta['ped_compra']} | {meta['razao_social']} | {len(prods)} Produtos Detectados",
                        font=("Segoe UI", 10, "bold"),
                        bg="#1F497D",
                        fg="#FFFFFF",
                        pady=6
                    )
                    lbl.pack(fill=tk.X)

                    txt = scrolledtext.ScrolledText(win, font=("Consolas", 9), padx=8, pady=8)
                    txt.pack(fill=tk.BOTH, expand=True)

                    txt.insert(tk.END, "=== METADADOS IDENTIFICADOS ===\n")
                    for k, v in meta.items():
                        txt.insert(tk.END, f"{k}: {v}\n")

                    txt.insert(tk.END, f"\n=== PRODUTOS IDENTIFICADOS ({len(prods)}) ===\n")
                    for idx, p in enumerate(prods, 1):
                        txt.insert(tk.END, f"{idx:02d}. ERP: {p['codigo_erp']:<8} | EAN: {p['cod_barras']:<14} | UM: {p['um']:<3} | Qtd: {p['quant']:<8} | Preço: {p['preco_unit']:<8} | {p['descricao']}\n")

                    txt.insert(tk.END, "\n=== TEXTO RECONSTRUÍDO PELA LINHA FÍSICA ===\n")
                    txt.insert(tk.END, texto)
                    txt.config(state=tk.DISABLED)

                self.root.after(0, show_popup)

            except Exception as e:
                self._log(f"Erro na inspeção: {e}")
                self.root.after(0, lambda: messagebox.showerror("Erro", str(e)))
            finally:
                self.progress_bar.stop()
                self.btn_inspecionar.config(state=tk.NORMAL)

        threading.Thread(target=task, daemon=True).start()

    def _iniciar_conversao(self):
        pdf_path = self.caminho_pdf.get().strip()
        excel_path = self.caminho_excel.get().strip()

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showwarning("Atenção", "Por favor, selecione um arquivo PDF existente!")
            return

        if not excel_path:
            messagebox.showwarning("Atenção", "Defina o local onde o Excel será salvo!")
            return

        if self.is_processing:
            return

        self.is_processing = True
        self.btn_converter.config(state=tk.DISABLED)
        self.progress_bar.start(10)
        self._log("\n--- INICIANDO CONVERSÃO BLINDADA ---")

        def task():
            try:
                total_f, avisos = extrair_pdf_para_excel(
                    pdf_path,
                    excel_path,
                    forcar_ocr=self.forcar_ocr.get(),
                    callback_log=self._log
                )

                def sucesso():
                    msg = (
                        f"Conversão concluída com sucesso!\n\n"
                        f"• Filiais geradas: {total_f}\n"
                        f"• Arquivo salvo em:\n{excel_path}\n\n"
                    )
                    if avisos:
                        msg += f"ATENÇÃO: {avisos} ponto(s) a conferir, marcados em vermelho na planilha.\n\n"
                    msg += "Deseja abrir a planilha Excel agora?"
                    resp = messagebox.askyesno("Concluído com avisos" if avisos else "Sucesso!", msg)
                    if resp:
                        try:
                            if os.name == 'nt':
                                os.startfile(excel_path)
                            elif os.uname().sysname == 'Darwin':
                                subprocess.Popen(['open', excel_path])
                            else:
                                subprocess.Popen(['xdg-open', excel_path])
                        except Exception as ex:
                            messagebox.showinfo("Aviso", f"Não foi possível abrir o arquivo automaticamente: {ex}")

                self.root.after(0, sucesso)

            except Exception as e:
                self._log(f"[ERRO CRÍTICO] {e}")
                self.root.after(0, lambda: messagebox.showerror("Erro na Conversão", str(e)))
            finally:
                self.is_processing = False
                self.btn_converter.config(state=tk.NORMAL)
                self.progress_bar.stop()

        threading.Thread(target=task, daemon=True).start()


def main():
    # Modo sem janela: Conversor_de_Pedidos.exe [--qualquer] arquivo.pdf [saida.xlsx]
    args = sys.argv[1:]
    generico = bool(args) and args[0] == "--qualquer"
    if generico:
        args = args[1:]
    if args:
        pdf = args[0]
        sufixo = ".xlsx" if generico else "_Filiais.xlsx"
        saida = args[1] if len(args) > 1 else os.path.splitext(pdf)[0] + sufixo
        converter = converter_pdf_generico if generico else extrair_pdf_para_excel
        # O .exe não tem console: o progresso vai para um .log ao lado da planilha
        with open(os.path.splitext(saida)[0] + ".log", "w", encoding="utf-8") as f:
            def log(msg):
                f.write(msg + "\n")
            try:
                converter(pdf, saida, callback_log=log)
            except Exception as e:
                log(f"[ERRO] {e}")
                sys.exit(1)
        return

    root = tk.Tk()
    app = ConversorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
