# h5p-to-pptx

Script Python para converter arquivos `.h5p` (Course Presentation) exportados do Moodle em apresentações `.pptx` editáveis.

## Status

- ✅ Conversão para **PPTX**
- 🔜 Conversão para **PDF** (planejado)

## Como funciona

Um arquivo `.h5p` é, na prática, um `.zip` contendo um `content/content.json` que descreve os slides, textos e imagens. O script extrai esse conteúdo e monta um `.pptx` reproduzindo a posição de cada elemento.

**Suporta atualmente:**
- Texto (`H5P.AdvancedText`, `H5P.Text`)
- Imagens (`H5P.Image`)

**Limitação:** elementos interativos (quiz, vídeo, drag-and-drop etc.) não têm equivalente em slide estático — aparecem como uma caixa de texto indicando "conteúdo interativo não suportado".

## Requisitos

```bash
pip install python-pptx beautifulsoup4 --break-system-packages
```

(ou use uma venv, se preferir isolar as dependências)

## Uso

```bash
python h5p_to_pptx.py arquivo.h5p
# gera arquivo.pptx na mesma pasta

python h5p_to_pptx.py arquivo.h5p saida.pptx
# gera saida.pptx com nome customizado
```

### Gerando PDF (via LibreOffice, por enquanto)

Até o suporte nativo a PDF ser implementado, converta o `.pptx` gerado:

```bash
soffice --headless --convert-to pdf saida.pptx
```

## Estrutura do projeto

```
h5p-to-pptx/
├── h5p_to_pptx.py
└── README.md
```
