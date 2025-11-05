from pathlib import Path
import unicodedata
import re
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from slugify import slugify
from datetime import datetime, timezone

# ----- caminhos -----
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path("data")
OUT_DIR = Path("certificados")
TPL_DIR = Path("templates")
TPL_NAME = "certificado.html"

# ----- constantes do evento (edite se quiser) -----
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

# ----- utilitários -----
def norm(s: str) -> str:
    """normaliza cabeçalhos: minúsculas, sem acento, sem pontuação/espaço extra"""
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-zA-Z0-9]+", "_", s).strip("_").lower()
    return s

def load_table_from_data_dir() -> pd.DataFrame:
    """carrega o primeiro .xlsx ou .csv encontrado em data/"""
    xlsxs = sorted(DATA_DIR.glob("*.xlsx"))
    csvs  = sorted(DATA_DIR.glob("*.csv"))
    if xlsxs:
        df = pd.read_excel(xlsxs[0])
    elif csvs:
        df = pd.read_csv(csvs[0])
    else:
        raise SystemExit("Nenhum arquivo encontrado em data/. Coloque seu .xlsx ou .csv lá.")
    return df

def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """mapeia colunas variadas da planilha para o padrão usado no template"""
    original_cols = list(df.columns)
    df = df.copy()
    df.columns = [norm(c) for c in df.columns]

    # dicionário de opções por campo (aceita variações comuns)
    options = {
        "nome": ["nome", "atleta", "nome_atleta", "competidor", "participante"],
        "cr_atleta": ["cr", "cr_atleta", "cr_cac", "numero_cr", "inscrito_no_cr"],
        "arma_modelo": ["arma_modelo", "arma", "modelo", "modelo_arma"],
        "calibre": ["calibre", "caliber", "cal"],
        "sigma": ["sigma", "craf", "registro_arma", "numero_sigma", "n_sigma"],
        "min_disparos": ["minimo_de_disparos", "min_disparos", "minimo_disparos", "qtd_minima_disparos"],
        "divisao": ["divisao", "divisao_categoria", "divisao_modalidade"],
        "categoria": ["categoria", "classe", "classe_categoria"],
        "posicao": ["posicao", "ranking", "colocacao", "classificacao"],
        # coluna para IDENTIFICADOR (= DATA)
        "data_identificador": ["data", "data_identificador", "data_da_prova", "periodo", "data_certificado"],
    }

    resolved = {}
    for target, candidates in options.items():
        found = None
        for c in candidates:
            if c in df.columns:
                found = c
                break
        if found is not None:
            resolved[target] = found

    # checagens mínimas
    missing_min = [k for k in ["nome", "cr_atleta", "arma_modelo", "calibre", "sigma",
                               "min_disparos", "divisao", "categoria", "posicao"] if k not in resolved]
    if missing_min:
        raise SystemExit(f"Faltam colunas na planilha (ou renomeie-as): {missing_min}\n"
                         f"Cabeçalhos normalizados encontrados: {df.columns.tolist()}")

    # monta DF só com o que precisamos
    out = pd.DataFrame()
    for k in ["nome","cr_atleta","arma_modelo","calibre","sigma","min_disparos","divisao","categoria","posicao"]:
        out[k] = df[resolved[k]]

    # IDENTIFICADOR = DATA (se existir); senão usa agora()
    if "data_identificador" in resolved:
        # tenta parsear; se falhar, usa string original
        d = pd.to_datetime(df[resolved["data_identificador"]], errors="coerce")
        # formata em ISO com Z
        out["identificador"] = d.dt.tz_localize("UTC", nonexistent="NaT", ambiguous="NaT").dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        # onde não deu parse, mantém original
        mask = out["identificador"].isna()
        out.loc[mask, "identificador"] = df.loc[mask, resolved["data_identificador"]].astype(str).str.strip()
    else:
        now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        out["identificador"] = now_iso

    return out

def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    df_raw = load_table_from_data_dir()
    df = map_columns(df_raw).fillna("")

    env = Environment(loader=FileSystemLoader(str(TPL_DIR)))
    tpl = env.get_template(TPL_NAME)

    for _, row in df.iterrows():
        p = row.to_dict()
        html = tpl.render(evento=EVENTO, p=p)
        # usa ranking/posição no nome do arquivo se for número, senão 999
        try:
            pos = int(row["posicao"])
        except Exception:
            pos = 999
        fname = f"{pos:03d}-{slugify(row['nome'])}.pdf"
        HTML(string=html).write_pdf(OUT_DIR / fname)
        print("OK:", fname)

if __name__ == "__main__":
    main()
