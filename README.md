# Certificado de Participação Automático - Geração de DOCX + PDF

## What is it made for?
Script Python desenvolvido para automatizar a geração de certificados personalizados para participantes de campeonatos de Tiro Defensivo (TD). Ele processa uma planilha Excel com informações dos participantes e gera arquivos Word (.docx) e PDF prontos para distribuição.

## Structure
```
certificate-generator/
├── assets/
├── build_docx/              # Saída intermediária .docx dos certificados
├── certificados/            # Saída final .pdf dos certificados
├── data/                    # Local da planilha .xlsx com os dados
├── src/
│   └── gerar_certificados_docx.py  # Script principal
├── templates/
│   └── certificado.docx     # Template com placeholders {{...}}
├── .venv/                   # Ambiente virtual Python
```

## Prerequisites
- Python 3.9+
- LibreOffice instalado (recomendado: C:\\Arquivos de Programas\\LibreOffice)
- pacotes Python:
  - pandas
  - openpyxl
  - python-docx
  - docxtpl
  - python-slugify

> Para instalar as dependências:
```bash
pip install pandas openpyxl python-docx docxtpl python-slugify
```

## How to use

1. Coloque sua planilha Excel (.xlsx) com as informações dos participantes na pasta `data/`. A aba deve se chamar "Apuração - Geral".
2. O template Word `certificado.docx` deve estar na pasta `templates/` com os placeholders como `{{ nome }}`, `{{ cr_atleta }}` etc.
3. Ative o ambiente virtual Python e rode o script:

### Activate env ambient code
```bash
.\.venv\Scripts\activate
```

### Run
```bash
python src\gerar_certificados_docx.py
```

Certificados serão salvos em:
- `build_docx/` para os .docx gerados
- `certificados/` para os PDFs convertidos

Cada PDF é nomeado no formato `001-nome-participante.pdf` com base na classificação.

## PT-BR - README em Português Brasil

### Para que serve?
Script para automatizar a criação de certificados para atletas que participaram de etapas do campeonato de Tiro Defensivo.

### Estrutura
```
certificate-generator/
├── data/ -> sua planilha de resultados
├── templates/ -> certificado.docx com campos {{}}
├── src/gerar_certificados_docx.py
├── certificados/ -> saída .pdf
├── build_docx/ -> saída .docx
```

### Requisitos
- Python 3.9 ou superior
- LibreOffice instalado (modo headless)
- Biblioteca Python:
```bash
pip install pandas openpyxl docxtpl slugify
```

### Como usar
1. Ative o ambiente virtual:
```bash
.\.venv\Scripts\activate
```
2. Rode o script:
```bash
python src\gerar_certificados_docx.py
```
3. Os certificados serão gerados automaticamente nas pastas `certificados/` e `build_docx/`.

---

## Lógica do Sistema

1. **Leitura da planilha:** O script carrega o primeiro arquivo .xlsx na pasta `data/`, lendo a aba `Apuração - Geral`.
2. **Mapeamento de colunas:** A função `map_columns()` normaliza nomes como "CLASSIFICAÇÃO" → `classificacao` e associa ao nome de campo esperado no DOCX (ex: `{{ posicao }}`).
3. **Renderização do certificado:** Para cada linha (participante), cria-se um dicionário com os dados e gera um `.docx` via `docxtpl.render(ctx)`.
4. **Exporta para PDF:** Usa o LibreOffice via linha de comando (modo headless) para converter o `.docx` em `.pdf` automaticamente.

```python
subprocess.run([
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "--headless",
    "--convert-to", "pdf",
    str(docx_path),
    "--outdir", str(pdf_dest_dir)
])
```

### Observação importante:
A planilha de dados **não é enviada para o repositório Git** por conter dados pessoais. Apenas o template DOCX fica salvo.

Para usar:
- Crie uma planilha com os dados exigidos (nome, cr, arma, etc.) e coloque em `data/`
- Rode o script e os certificados são gerados com base nos campos do modelo .docx

### Final
Todo o sistema é 100% offline e automatizado. Ideal para competições e aplicações onde se deseja gerar em massa certificados personalizados de forma confiável e ordenada.

