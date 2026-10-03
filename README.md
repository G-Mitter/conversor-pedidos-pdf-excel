# Conversor PDF → Excel

Aplicativo para Windows que transforma PDFs em planilhas Excel (`.xlsx`). Funciona com PDFs que têm texto e também com PDFs que são só imagem (escaneados ou exportados como foto por sistemas na nuvem), usando OCR.

O app tem duas abas:

| Aba | Para que serve |
| --- | --- |
| **Qualquer PDF** | Converte qualquer PDF mantendo a disposição da página: cada linha do PDF vira uma linha da planilha e as colunas são detectadas pelos espaços em branco. Ideal para relatórios, extratos e tabelas. |
| **Pedido de Compra** | Lê pedidos de compra emitidos por ERP e monta uma planilha organizada por filial, com cabeçalho do pedido e tabela de produtos (código, código de barras, descrição, UM, quantidade, preço e total). Confere os valores e marca em vermelho o que precisa de revisão. |

## Como usar

1. Baixe o `Conversor_de_Pedidos.exe` na página de **Releases** deste repositório (não precisa instalar Python).
2. Abra o programa, escolha a aba, selecione o PDF e clique em converter.
3. A planilha é salva ao lado do PDF (ou no local que você escolher).

> Como o `.exe` não tem assinatura digital, o Windows pode mostrar "O Windows protegeu o computador". Clique em **Mais informações** → **Executar assim mesmo**.

PDFs em imagem passam pelo OCR e demoram mais (cerca de 15 a 25 segundos por página).

### Sem abrir a janela

O mesmo `.exe` converte pela linha de comando, gravando um `.log` ao lado da planilha:

```powershell
Conversor_de_Pedidos.exe --qualquer arquivo.pdf [saida.xlsx]   # aba Qualquer PDF
Conversor_de_Pedidos.exe pedido.pdf [saida.xlsx]               # aba Pedido de Compra
```

## Conferência automática (aba Pedido de Compra)

- Em cada linha, **quantidade × preço** tem que bater com o total lido do PDF.
- Em cada filial, a **soma das linhas** tem que bater com o **Total Geral** impresso no pedido.
- O que não bate fica em vermelho com a indicação `CONFERIR`, e a janela final informa quantos pontos revisar.

## Rodando pelo código

Requer Python 3.10 ou mais recente.

```powershell
pip install pymupdf pdfplumber pypdf openpyxl pillow numpy rapidocr-onnxruntime
python app_gui.py
```

### Testes

```powershell
python test_conversor.py
```

Os testes que dependem de pedidos reais usam PDFs na pasta `exemplos/` (`pedido_texto.pdf` e `pedido_imagem.pdf`). Essa pasta não vai para o GitHub porque os pedidos têm dados de clientes; sem ela, esses testes são pulados.

### Gerando o `.exe`

```powershell
pip install pyinstaller
python -m PyInstaller Conversor_de_Pedidos.spec --noconfirm
```

O executável sai em `dist/`. O arquivo `.spec` já inclui o ícone, os modelos do OCR e as dependências que o PyInstaller não encontra sozinho.

## Como funciona

- **PDF com texto:** o texto é lido com PyMuPDF trecho a trecho, respeitando o recorte das células. Alguns ERPs escrevem a descrição inteira e escondem o excesso; outros leitores misturam esse texto escondido com a coluna seguinte.
- **PDF em imagem:** o OCR (RapidOCR) roda sobre a imagem embutida na **resolução original**. Ampliar a imagem antes do OCR borra as letras e piora a leitura.
- **Linhas e colunas:** os trechos são agrupados pela posição vertical (linhas) e, na aba genérica, distribuídos em colunas pelas faixas em branco da página.

## Limitações

- No OCR, descrições podem ter pequenos erros de letras (ex.: `O` no lugar de `0`), e palavras muito próximas às vezes vêm juntas na mesma célula.
- Na aba genérica, duas tabelas com colunas diferentes na mesma página compartilham a mesma grade de colunas.
- A aba de pedidos foi feita para o layout de pedido de compra com colunas de código, código de barras, descrição, UM, quantidade, preço e total.

## Privacidade

PDFs, planilhas e logs são ignorados pelo git (veja `.gitignore`). Nenhum dado de cliente faz parte do código.
