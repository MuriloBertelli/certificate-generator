from datetime import datetime, timezone
from pathlib import Path
import pandas as pd
import unicodedata
import re
import shutil
import subprocess
from slugify import slugify
from docxtpl import DocxTemplate
import subprocess  

# Diretórios e arquivos base
BASE = Path(__file__).resolve().parent.parent
DATA_DIR = BASE / "data"
TPL_PATH = BASE / "templates" / "certificado.docx"
DOCX_TMP = BASE / "build_docx"
PDF_OUT = BASE / "certificados"

DEFAULT_MIN_DISPAROS = 16
SHEET_NAME = "Apuração - Geral"

EVENTO = {
    "liga": "LNTD – Liga Nacional de Tiro Defensivo",
    "cnpj": "27.347.774/0001-40",
    "cr_entidade": "150.113",
    "endereco": "Rua Henrique Correia, 1849, Bairro Alto, Curitiba – PR - CEP 82840-270",
    "titulo": "Certificado de Participação",
    "nome_prova": "4ª Etapa do Torneio de TD da LNTD 2025",
    "periodo": "11 a 25 de novembro de 2025",
    "organizador_exibido": "Liga Nacional de Tiro Defensivo (LNTD)",
}
def find_soffice():
    path = r"C:\Program Files\LibreOffice\program\soffice.exe"
    if not Path(path).exists():
        raise SystemExit(f"LibreOffice não encontrado em: {path}")
    return path



def normalize_header(header):
    h = header.strip().lower()
    h = ''.join(ch for ch in unicodedata.normalize('NFKD', h) if unicodedata.category(ch) != 'Mn')
    return h

def load_table():
    xlsxs = sorted(DATA_DIR.glob("*.xlsx"))
    if not xlsxs:
        raise SystemExit("Nenhum arquivo .xlsx encontrado em data/")

    df = pd.read_excel(xlsxs[0], sheet_name=SHEET_NAME, engine="openpyxl", header=0, dtype=str)
    df = df.dropna(how="all")
    df = df.loc[:, [c for c in df.columns if c and str(c).strip() != ""]].reset_index(drop=True)

    if not len(df.columns):
        raise SystemExit("Falha na detecção dos cabeçalhos.")

    return df

def map_columns(df):
    col_map = {}
    normalized_cols = {normalize_header(c): c for c in df.columns if c}

    mapping = {
        "nome": ["nome"],
        "cr_atleta": ["cr"],
        "arma_modelo": ["arma"],
        "calibre": ["calibre"],
        "sigma": ["sigma"],
        "divisao": ["divisao"],
        "categoria": ["categoria"],
        "posicao": ["classificacao", "posicao"],
        "data_identificador": ["data"]
    }

    for key, keywords in mapping.items():
        for k in keywords:
            k_norm = normalize_header(k)
            if k_norm in normalized_cols:
                col_map[key] = normalized_cols[k_norm]
                break

    required = ["nome", "cr_atleta", "arma_modelo", "calibre", "sigma", "divisao", "categoria", "posicao"]
    for r in required:
        if r not in col_map:
            raise SystemExit(f"Coluna obrigatória ausente: {r}")

    out = pd.DataFrame({k: df[col_map[k]] for k in required})
    out["min_disparos"] = DEFAULT_MIN_DISPAROS

    if "data_identificador" in col_map:
        d = pd.to_datetime(df[col_map["data_identificador"]], errors="coerce", utc=True)
        out["identificador"] = d.dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        mask = out["identificador"].isna()
        out.loc[mask, "identificador"] = (
            df.loc[mask, col_map["data_identificador"]].astype(str).str.strip()
        )
    else:
        out["identificador"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    return out.fillna("")

def export_pdf_via_libreoffice(docx_path, pdf_dest_dir):
    subprocess.run([
        find_soffice(),
        "--headless",
        "--convert-to", "pdf",
        str(docx_path),
        "--outdir", str(pdf_dest_dir)
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def main():
    if not TPL_PATH.exists():
        raise SystemExit(f"Template não encontrado: {TPL_PATH}")

    DOCX_TMP.mkdir(parents=True, exist_ok=True)
    PDF_OUT.mkdir(parents=True, exist_ok=True)

    df_raw = load_table()
    df = map_columns(df_raw)

    for _, row in df.iterrows():
        ctx = row.to_dict()
        ctx["evento"] = EVENTO

        try:
            pos = int(row["posicao"])
        except Exception:
            pos = 999
        base = f"{pos:03d}-{slugify(row['nome'])}"

        tpl = DocxTemplate(str(TPL_PATH))
        tpl.render(ctx)
        out_docx = DOCX_TMP / f"{base}.docx"
        tpl.save(out_docx)

        export_pdf_via_libreoffice(out_docx, PDF_OUT)
        print(f"OK: certificados/{base}.pdf")

if __name__ == "__main__":
    main()
