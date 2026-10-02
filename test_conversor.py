"""
Checagem rápida: python test_conversor.py

Os PDFs de exemplo têm dados de clientes e ficam só na sua máquina, na pasta exemplos/
(ignorada pelo git). Sem eles, os testes que dependem dos PDFs são pulados.
  exemplos/pedido_texto.pdf  -> pedido com camada de texto (5 filiais, 160 itens)
  exemplos/pedido_imagem.pdf -> pedido em imagem, lido por OCR (2 filiais, 59 itens)
"""
import os
import tempfile

import fitz
from openpyxl import load_workbook

import app_gui

EXEMPLOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exemplos")
PEDIDO_TEXTO = os.path.join(EXEMPLOS, "pedido_texto.pdf")
PEDIDO_IMAGEM = os.path.join(EXEMPLOS, "pedido_imagem.pdf")


def produtos(caminho_xlsx):
    """Linhas de produto da planilha: as que têm código de barras de 13 dígitos na coluna C."""
    ws = load_workbook(caminho_xlsx).active
    return [r for r in ws.iter_rows(values_only=True) if r[2] and str(r[2]).isdigit() and len(str(r[2])) == 13]


def test_pedido_texto():
    saida = os.path.join(tempfile.gettempdir(), "test_pedido_texto.xlsx")
    filiais, avisos = app_gui.extrair_pdf_para_excel(PEDIDO_TEXTO, saida)
    assert (filiais, avisos) == (5, 0), (filiais, avisos)  # avisos 0 = cada filial bate com o Total Geral
    prods = produtos(saida)
    assert len(prods) == 160, len(prods)
    assert all(r[4] == "CX" for r in prods), {r[4] for r in prods}
    assert all(r[5] < 10000 and abs(r[5] * r[6] - r[7]) <= 0.05 for r in prods)
    os.remove(saida)


def test_pedido_imagem():
    saida = os.path.join(tempfile.gettempdir(), "test_pedido_imagem.xlsx")
    filiais, avisos = app_gui.extrair_pdf_para_excel(PEDIDO_IMAGEM, saida)
    assert (filiais, avisos) == (2, 0), (filiais, avisos)
    assert len(produtos(saida)) == 59
    os.remove(saida)


def test_generico_tabela():
    """Aba 'Qualquer PDF': cada campo da tabela vai para uma coluna e números viram números."""
    saida = os.path.join(tempfile.gettempdir(), "test_generico.xlsx")
    assert app_gui.converter_pdf_generico(PEDIDO_TEXTO, saida) == 10
    linhas = load_workbook(saida)["Pág 1"].iter_rows(values_only=True)
    item = [c for c in next(r for r in linhas if any(str(c).isdigit() and len(str(c)) == 13 for c in r)) if c is not None]
    assert item[3] == "CX  1" and all(isinstance(c, (int, float)) for c in item[4:]), item
    os.remove(saida)


def test_generico_texto_corrido():
    pdf = os.path.join(tempfile.gettempdir(), "test_texto.pdf")
    saida = os.path.join(tempfile.gettempdir(), "test_texto.xlsx")
    doc = fitz.open()
    doc.new_page().insert_text((72, 72), "Relatorio mensal\nLinha de texto corrido sem tabela")
    doc.save(pdf)
    app_gui.converter_pdf_generico(pdf, saida)
    linhas = list(load_workbook(saida).active.iter_rows(values_only=True))
    assert linhas[:2] == [("Relatorio mensal",), ("Linha de texto corrido sem tabela",)], linhas
    os.remove(saida)
    os.remove(pdf)


if __name__ == "__main__":
    test_generico_texto_corrido()
    for pdf, testes in [(PEDIDO_TEXTO, [test_pedido_texto, test_generico_tabela]), (PEDIDO_IMAGEM, [test_pedido_imagem])]:
        if not os.path.exists(pdf):
            print(f"Pulado (sem {os.path.relpath(pdf)}): {', '.join(t.__name__ for t in testes)}")
            continue
        for t in testes:
            t()
    print("OK")
