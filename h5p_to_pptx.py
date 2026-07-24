"""
h5p_to_pptx.py
 
Converte um arquivo .h5p do tipo "Course Presentation" (H5P.CoursePresentation)
para um arquivo .pptx editável no PowerPoint.
 
Uso:
    python h5p_to_pptx.py caminho/para/arquivo.h5p [saida.pptx]
 
Requisitos:
    pip install python-pptx beautifulsoup4
"""

import sys
import os
import json
import zipfile
import shutil
import tempfile
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
 
# Dimensões padrão de slide do H5P Course Presentation (em "unidades" internas 0-100%)
SLIDE_WIDTH_EMU = Emu(9144000)   # 10 polegadas
SLIDE_HEIGHT_EMU = Emu(5143500)  # ~5.625 polegadas (16:9)
 
 
def extrair_h5p(caminho_h5p, pasta_destino):
    """Descompacta o .h5p (é um .zip) na pasta_destino."""
    with zipfile.ZipFile(caminho_h5p, "r") as z:
        z.extractall(pasta_destino)
 
 
def carregar_content_json(pasta_extraida):
    caminho = os.path.join(pasta_extraida, "content", "content.json")
    if not os.path.exists(caminho):
        raise FileNotFoundError(
            "content/content.json não encontrado. "
            "Confira se o .h5p é realmente do tipo Course Presentation."
        )
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)
 
 
def html_para_texto(html):
    """Extrai texto simples de um trecho HTML (remove tags)."""
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return soup.get_text(separator="\n").strip()
 
 
def porcentagem_para_emu(valor_percentual, total_emu):
    return Emu(int(total_emu * (valor_percentual / 100.0)))
 
 
def adicionar_elemento(slide, elemento, pasta_extraida):
    """Adiciona um elemento do H5P (texto ou imagem) ao slide do pptx."""
    x = elemento.get("x", 0)
    y = elemento.get("y", 0)
    largura = elemento.get("width", 30)
    altura = elemento.get("height", 20)
 
    left = porcentagem_para_emu(x, SLIDE_WIDTH_EMU)
    top = porcentagem_para_emu(y, SLIDE_HEIGHT_EMU)
    width = porcentagem_para_emu(largura, SLIDE_WIDTH_EMU)
    height = porcentagem_para_emu(altura, SLIDE_HEIGHT_EMU)
 
    action = elemento.get("action", {})
    library = action.get("library", "")
 
    if library.startswith("H5P.AdvancedText") or library.startswith("H5P.Text"):
        html = action.get("params", {}).get("text", "")
        texto = html_para_texto(html)
        if texto:
            caixa = slide.shapes.add_textbox(left, top, width, height)
            tf = caixa.text_frame
            tf.word_wrap = True
            linhas = texto.split("\n")
            tf.text = linhas[0]
            for linha in linhas[1:]:
                p = tf.add_paragraph()
                p.text = linha
 
    elif library.startswith("H5P.Image"):
        params = action.get("params", {})
        arquivo_info = params.get("file", {})
        caminho_relativo = arquivo_info.get("path", "")
        if caminho_relativo:
            caminho_imagem = os.path.join(pasta_extraida, "content", caminho_relativo)
            if os.path.exists(caminho_imagem):
                slide.shapes.add_picture(caminho_imagem, left, top, width, height)
 
    else:
        # Outros tipos (vídeo, quiz, interações etc.) não têm equivalente
        # direto em slide estático — inserimos apenas uma nota de texto.
        caixa = slide.shapes.add_textbox(left, top, width, height)
        caixa.text_frame.text = f"[Conteúdo interativo não suportado: {library}]"
 
 
def converter(caminho_h5p, caminho_saida):
    pasta_temp = tempfile.mkdtemp(prefix="h5p_extract_")
    try:
        extrair_h5p(caminho_h5p, pasta_temp)
        dados = carregar_content_json(pasta_temp)
 
        slides_h5p = dados.get("presentation", {}).get("slides", [])
        if not slides_h5p:
            raise ValueError(
                "Nenhum slide encontrado em content.json — "
                "verifique se é mesmo um Course Presentation."
            )
 
        prs = Presentation()
        prs.slide_width = SLIDE_WIDTH_EMU
        prs.slide_height = SLIDE_HEIGHT_EMU
        layout_em_branco = prs.slide_layouts[6]  # layout em branco
 
        for slide_h5p in slides_h5p:
            slide = prs.slides.add_slide(layout_em_branco)
            for elemento in slide_h5p.get("elements", []):
                adicionar_elemento(slide, elemento, pasta_temp)
 
        prs.save(caminho_saida)
        print(f"OK: {len(slides_h5p)} slide(s) convertido(s) para: {caminho_saida}")
 
    finally:
        shutil.rmtree(pasta_temp, ignore_errors=True)
 
 
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python h5p_to_pptx.py arquivo.h5p [saida.pptx]")
        sys.exit(1)
 
    entrada = sys.argv[1]
    saida = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(entrada)[0] + ".pptx"
 
    converter(entrada, saida)
 
