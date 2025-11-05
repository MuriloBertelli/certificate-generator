from datetime import datetime, timezone
import subprocess
import pandas as pd
from slugify import slugify
import unicodedata, re, shutil, sys
from docx2pdf import convert
import shutil
from pathlib import Path
# friendly import for template engine
try:
    from docxtpl import DocxTemplate
except ImportError as e:
    # make the error actionable for users running the script directly
    venv_python = None
    # try to guess the venv python used by this repository (configured environment)
    try:
        # known from environment configuration in this workspace
        venv_python = r"C:/dev/certificate-generator/certificate-generator/.venv/Scripts/python.exe"
        if not Path(venv_python).exists():
            venv_python = None
    except Exception:
        venv_python = None

    install_hint = (
        f"{venv_python} -m pip install python-docx-template\n" if venv_python else
        "python -m pip install python-docx-template"
    )
    raise SystemExit(
        "Missing dependency 'python-docx-template' (module name: docxtpl).\n"
        "Install it and try again. For example:\n"
        f"  {install_hint}\n"
        "Or activate your virtualenv and run: pip install python-docx-template"
    ) from e

# --- caminhos ---
from pathlib import Path
BASE = Path(__file__).resolve().parent.parent  # sobe de src/ para a raiz
DATA_DIR = BASE / "data"
TPL_PATH = BASE / "templates" / "certificado.docx"
DOCX_TMP = BASE / "build_docx"
PDF_OUT  = BASE / "certificados"

# --- caminho do LibreOffice (ajuste se necessário) ---
def find_soffice():
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in candidates:
        if Path(p).exists():
            return p
    return shutil.which("soffice") or shutil.which("soffice.exe")

SOFFICE = find_soffice()

# --- dados do evento (constantes) ---
DEFAULT_MIN_DISPAROS = 16  # valor padrão quando não vier na planilha
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

def norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_").lower()
    return s

def load_table() -> pd.DataFrame:
    xlsxs = sorted(DATA_DIR.glob("*.xlsx"))
    csvs  = sorted(DATA_DIR.glob("*.csv"))
    if xlsxs:
        return pd.read_excel(xlsxs[0])
    if csvs:
        return pd.read_csv(csvs[0])
    raise SystemExit("Nenhum .xlsx ou .csv encontrado em data/")

def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """mapeia colunas variadas da planilha para o padrão usado no template"""
    df = df.copy()
    df.columns = [norm(c) for c in df.columns]  # normaliza cabeçalhos

    # apelidos aceitos para cada campo
    opts = {
        "nome": ["nome", "atleta", "nome_atleta", "competidor", "participante"],
        "cr_atleta": ["cr", "cr_atleta", "cr_cac", "numero_cr", "inscrito_no_cr"],
        "arma_modelo": ["arma_modelo", "arma", "modelo", "modelo_arma"],
        "calibre": ["calibre", "caliber", "cal"],
        "sigma": ["sigma", "craf", "registro_arma", "numero_sigma", "n_sigma"],
        "min_disparos": ["minimo_de_disparos", "min_disparos", "minimo_disparos", "qtd_minima_disparos"],
        "divisao": ["divisao", "divisao_categoria", "divisao_modalidade"],
        "categoria": ["categoria", "classe", "classe_categoria"],
        # sua planilha tem "CLASSIFICAÇÃO" → normaliza para "classificacao"
        "posicao": ["posicao", "ranking", "colocacao", "classificacao"],
        # IDENTIFICADOR = DATA
        "data_identificador": ["data", "data_identificador", "data_da_prova", "periodo", "data_certificado"],
    }

    # >>> AQUI construímos o 'resolved' <<<
    resolved = {}
    for target, candidates in opts.items():
        for c in candidates:
            if c in df.columns:
                resolved[target] = c
                break

    # checagem mínima (agora 'min_disparos' NÃO é obrigatório)
    missing_min = [k for k in ["nome", "cr_atleta", "arma_modelo", "calibre", "sigma",
                               "divisao", "categoria", "posicao"] if k not in resolved]
    if missing_min:
        raise SystemExit(
            f"Faltam colunas na planilha: {missing_min}\n"
            f"Encontradas: {df.columns.tolist()}"
        )

    # monta DF com os campos obrigatórios
    out = pd.DataFrame({k: df[resolved[k]] for k in
                        ["nome", "cr_atleta", "arma_modelo", "calibre", "sigma",
                         "divisao", "categoria", "posicao"]})

    # min_disparos: usa a coluna se existir; senão, valor padrão
    if "min_disparos" in resolved:
        out["min_disparos"] = df[resolved["min_disparos"]]
    else:
        out["min_disparos"] = DEFAULT_MIN_DISPAROS

   # IDENTIFICADOR = DATA (regra)
    if "data_identificador" in resolved:
        # lida com valores com ou sem timezone sem estourar erro
        d = pd.to_datetime(df[resolved["data_identificador"]], errors="coerce", utc=True)
        out["identificador"] = d.dt.strftime("%Y-%m-%dT%H:%M:%SZ")

        # onde não deu parse, mantém texto original
        mask = out["identificador"].isna()
        out.loc[mask, "identificador"] = (
            df.loc[mask, resolved["data_identificador"]].astype(str).str.strip()
        )
    else:
        out["identificador"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    return out.fillna("")


from docx2pdf import convert

def export_pdf_via_libreoffice(docx_path, pdf_dest):
    pdf_dest.parent.mkdir(parents=True, exist_ok=True)
    # o docx2pdf cria o PDF dentro da pasta destino com o mesmo nome
    convert(str(docx_path), str(pdf_dest.parent))


def main():
    if not TPL_PATH.exists():
        raise SystemExit(f"Template não encontrado: {TPL_PATH}")
    DOCX_TMP.mkdir(parents=True, exist_ok=True)
    PDF_OUT.mkdir(parents=True, exist_ok=True)

    df = map_columns(load_table())
    for _, row in df.iterrows():
        ctx = row.to_dict()
        ctx["evento"] = EVENTO

        # nome do arquivo com posição + slug do nome
        try:
            pos = int(row["posicao"])
        except Exception:
            pos = 999
        base = f"{pos:03d}-{slugify(row['nome'])}"

        # 1) renderiza DOCX
        tpl = DocxTemplate(str(TPL_PATH))
        tpl.render(ctx)
        out_docx = DOCX_TMP / f"{base}.docx"
        tpl.save(out_docx)

        # 2) converte para PDF mantendo layout
        out_pdf = PDF_OUT / f"{base}.pdf"
        export_pdf_via_libreoffice(out_docx, out_pdf)
        print("OK:", out_pdf)

if __name__ == "__main__":
    main()
