# Certificate Generator - Automatic DOCX + PDF Creation

## What is it made for?

Python script designed to automate the generation of personalized certificates for participants of Defensive Shooting championships. It processes an Excel spreadsheet with participant data and generates ready-to-distribute Word (.docx) and PDF files.

## Structure

```
certificate-generator/
├── assets/
├── build_docx/              # Intermediate output (.docx) of certificates
├── certificados/            # Final output (.pdf) of certificates
├── data/                    # Excel (.xlsx) spreadsheet with input data
├── src/
│   └── gerar_certificados_docx.py  # Main script
├── templates/
│   └── certificado.docx     # Template with {{...}} placeholders
└── .venv/                   # Python virtual environment
```

## Prerequisites

- Python 3.9+
- LibreOffice installed (recommended: `C:\Program Files\LibreOffice`)
- Required Python packages:
  ```bash
  pip install pandas openpyxl python-docx docxtpl python-slugify
  ```

## How to Use

1. Place your `.xlsx` spreadsheet with participant data inside the `data/` folder. The sheet must be named `"Apuração - Geral"`.
2. Make sure the Word template `certificado.docx` is inside `templates/` and includes placeholders like `{{ nome }}`, `{{ cr_atleta }}`, etc.
3. Activate the virtual environment and run the script:

### Activate virtual environment

```bash
.\.venv\Scripts\activate
```

### Run the script

```bash
python src\gerar_certificados_docx.py
```

Certificates will be saved in:
- `build_docx/` for generated .docx files
- `certificados/` for the exported PDF files

Each PDF is named in the format `001-nome-participante.pdf`, based on participant ranking.

## Logic Behind the Script

1. **Excel Reading:** Loads the first `.xlsx` file inside `data/` and reads the `"Apuração - Geral"` sheet.
2. **Column Mapping:** `map_columns()` normalizes header names (e.g. "CLASSIFICAÇÃO" → `classificacao`) and maps them to the expected DOCX context keys (like `{{ posicao }}`).
3. **DOCX Rendering:** For each row (participant), a dictionary is built and passed to `docxtpl.render(ctx)` to generate a `.docx`.
4. **Export to PDF:** Uses LibreOffice in headless mode via subprocess to convert `.docx` into `.pdf`.

```python
subprocess.run([
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "--headless",
    "--convert-to", "pdf",
    str(docx_path),
    "--outdir", str(pdf_dest_dir)
])
```

### Important Note:
The spreadsheet with personal participant data is **not included in the Git repository**. Only the DOCX template is versioned.

To use:
- Prepare a spreadsheet with all required data (name, CR, weapon, etc.) and place it in the `data/` folder.
- Run the script, and the system will generate all certificates automatically from the `.docx` model.

---

# PT-BR — Gerador de Certificados - Geração Automática de DOCX + PDF

## Para que serve?

Script Python desenvolvido para automatizar a criação de certificados personalizados para participantes de campeonatos de Tiro Defensivo. Ele processa uma planilha Excel com os dados dos atletas e gera arquivos Word (.docx) e PDF prontos para distribuição.

## Estrutura

```
certificate-generator/
├── assets/
├── build_docx/              # Saída intermediária (.docx) dos certificados
├── certificados/            # Saída final (.pdf) dos certificados
├── data/                    # Planilha (.xlsx) com dados de entrada
├── src/
│   └── gerar_certificados_docx.py  # Script principal
├── templates/
│   └── certificado.docx     # Template com campos {{...}}
└── .venv/                   # Ambiente virtual Python
```

## Requisitos

- Python 3.9 ou superior
- LibreOffice instalado (preferencialmente em `C:\Program Files\LibreOffice`)
- Pacotes Python necessários:
  ```bash
  pip install pandas openpyxl python-docx docxtpl python-slugify
  ```

## Como Usar

1. Coloque sua planilha `.xlsx` com os dados dos participantes dentro da pasta `data/`. A aba deve se chamar `"Apuração - Geral"`.
2. O template `certificado.docx` precisa estar dentro da pasta `templates/` com os campos como `{{ nome }}`, `{{ cr_atleta }}`, etc.
3. Ative o ambiente virtual e execute o script:

### Ativar ambiente virtual

```bash
.\.venv\Scripts\activate
```

### Executar o script

```bash
python src\gerar_certificados_docx.py
```

Os certificados serão salvos em:
- `build_docx/` para os arquivos .docx gerados
- `certificados/` para os arquivos .pdf exportados

Cada certificado em PDF será nomeado no formato `001-nome-participante.pdf`, conforme a classificação.

## Lógica do Sistema

1. **Leitura da Planilha:** O script carrega o primeiro `.xlsx` dentro da pasta `data/`, lendo a aba `"Apuração - Geral"`.
2. **Mapeamento de Colunas:** A função `map_columns()` normaliza os nomes dos cabeçalhos (ex: "CLASSIFICAÇÃO" → `classificacao`) e associa aos campos esperados no DOCX.
3. **Renderização do DOCX:** Para cada linha (participante), cria-se um dicionário com os dados e gera-se um `.docx` via `docxtpl.render(ctx)`.
4. **Exportação para PDF:** Usa o LibreOffice em modo silencioso (headless) para converter o `.docx` em `.pdf`.

```python
subprocess.run([
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "--headless",
    "--convert-to", "pdf",
    str(docx_path),
    "--outdir", str(pdf_dest_dir)
])
```

### Observação:
A planilha com os dados dos participantes **não é incluída no Git** por conter informações sensíveis. Apenas o template `.docx` será versionado.

Para usar:
- Crie a planilha com os dados exigidos (nome, CR, arma, etc.) e coloque em `data/`
- Execute o script e os certificados serão gerados automaticamente a partir do template