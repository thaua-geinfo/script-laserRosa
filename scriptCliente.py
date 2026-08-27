#!/usr/bin/env python3
"""Tratamento da planilha de Clientes - Laser Rosa.

Estrutura do projeto:
    raiz/
      entrada/   -> Clientes*.xlsx e modeloImportacaoCliente*.csv
      saida/     -> planilhaTratadaCliente.csv e validacoes
      scriptCliente.py

Regras importantes:
- o modelo serve somente para cabecalho/formato; nenhuma linha tratada e usada como fonte;
- colunas de codigo proprio com indicacao de maximo de 6 digitos usam 100000+;
- remove aspas simples/duplas, barra invertida, controles, quebras de linha e espacos excedentes;
- campos obrigatorios vazios sao excluidos da planilha final e listados em relatorio;
- Observacao recebe "Importacao dd/mm/aaaa" e o maximo de informacoes da extracao
  que nao foram importadas em outras colunas, sempre com o nome/indicativo do campo;
- a saida sempre e gravada em ./saida por padrao.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import os
import posixpath
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
from datetime import date, datetime, time, timedelta
from pathlib import Path, PurePosixPath
from typing import Any, Iterable
from collections import defaultdict
from zipfile import ZipFile

VERSION = "2026-08-26.10"
SCRIPT_DIR = Path(__file__).resolve().parent

def _existing_layout_dir(root: Path, logical_name: str) -> Path | None:
    wanted = unicodedata.normalize("NFKD", logical_name).encode("ascii", "ignore").decode().casefold()
    wanted = re.sub(r"[^a-z0-9]+", "", wanted)
    if not root.is_dir():
        return None
    try:
        children = list(root.iterdir())
    except OSError:
        return None
    for child in children:
        if not child.is_dir():
            continue
        key = unicodedata.normalize("NFKD", child.name).encode("ascii", "ignore").decode().casefold()
        key = re.sub(r"[^a-z0-9]+", "", key)
        if key == wanted:
            return child.resolve()
    return None

def _discover_project_layout() -> tuple[Path, Path, Path]:
    env_root = os.environ.get("LASER_ROSA_PROJECT_ROOT", "").strip()
    roots: list[Path] = []
    if env_root:
        roots.append(Path(env_root).expanduser())
    roots.extend((SCRIPT_DIR, *SCRIPT_DIR.parents))
    seen: set[Path] = set()
    for candidate in roots:
        try:
            root = candidate.resolve()
        except OSError:
            continue
        if root in seen:
            continue
        seen.add(root)
        entrada = _existing_layout_dir(root, "entrada")
        saida = _existing_layout_dir(root, "saida")
        if entrada is not None and saida is not None:
            return root, entrada, saida
    # Mantem caminhos previsiveis para produzir uma mensagem de erro clara.
    return SCRIPT_DIR, SCRIPT_DIR / "entrada", SCRIPT_DIR / "saida"

PROJECT_ROOT, INPUT_DIR, OUTPUT_DIR = _discover_project_layout()

FORBIDDEN_CHARS = str.maketrans({'"': '', "'": '', '\\': ''})
CODE_MAX6_RE = re.compile(r"c[oó]digo.*max\s*6\s*d[ií]gitos", re.I)
PHONE_OUTPUT_INDEXES = (6, 7, 18)
NO_WHITESPACE_OUTPUT_INDEXES = (6, 7, 8, 9, 10, 17, 18, 25, 26, 27)


VALID_UFS = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO",
}
INVALID_CLIENT_PLACEHOLDERS = {
    "buscando", "invalid date", "invaliddate", "undefined", "indefinido",
    "invalido", "null", "none", "nan", "nat", "n/a", "na",
    "nao informado", "sem informacao",
}


def _client_value_key(value: Any) -> str:
    text = clean_text(value)
    folded = unicodedata.normalize("NFKD", text)
    folded = "".join(ch for ch in folded if not unicodedata.combining(ch)).casefold()
    folded = re.sub(r"\s+", " ", folded).strip()
    # Placeholders podem chegar com pontuacao no final: Buscando..., NULL., etc.
    return folded.strip(" .,:;!?_-–—()[]{}")


def _is_punctuation_only_client_value(value: Any) -> bool:
    text = clean_text(value)
    return bool(text) and not any(ch.isalnum() for ch in text)


def _is_invalid_client_placeholder(value: Any) -> bool:
    text = clean_text(value)
    if not text:
        return False
    return _client_value_key(text) in INVALID_CLIENT_PLACEHOLDERS


def _is_garbage_client_value(value: Any) -> bool:
    # Mantida como auxiliar para validacoes especificas de conteudo real,
    # como a normalizacao estrutural do Nome quando ele nao e apenas pontuacao.
    return _is_punctuation_only_client_value(value) or _is_invalid_client_placeholder(value)


def normalize_client_name(value: Any) -> str:
    """Remove apenas lixo de prefixo sem destruir observacoes legitimas no nome."""
    text = clean_text(value)
    if not text or _is_garbage_client_value(text):
        return ""
    # Evita que Excel trate o conteudo como formula e corrige casos como
    # '-Camilla Laura...' e '...Maria'. Parenteses sao preservados porque
    # algumas extracoes os usam legitimamente no nome/observacao.
    text = re.sub(r"^[=+@]+", "", text).strip()
    text = re.sub(r"^[\-–—.,;:_!?#$%&*/|]+\s*", "", text).strip()
    if not text or not any(ch.isalpha() for ch in text):
        return ""
    return " ".join(text.split())


def normalize_client_gender(value: Any) -> str:
    key = _client_value_key(value)
    if key == "masculino":
        return "Masculino"
    if key == "feminino":
        return "Feminino"
    return ""


def normalize_client_state(value: Any) -> str:
    text = clean_text(value).upper().strip()
    return text if text in VALID_UFS else ""


def sanitize_client_output_row(
    row: list[Any],
    headers: list[str],
    required: list[int] | tuple[int, ...] | set[int] | None = None,
) -> list[str]:
    """Limpeza final de Cliente respeitando obrigatorios e opcionais."""
    result = [clean_cell(value) for value in row]
    keys = [_source_header_compact(header) for header in headers]
    required_set = set(required or ())
    for index, key in enumerate(keys):
        if index >= len(result):
            break
        # Regras especificas continuam tendo prioridade.
        if key.startswith("sexo"):
            result[index] = normalize_client_gender(result[index])
            continue
        if key == "ufsigla" or key.startswith("uf"):
            result[index] = normalize_client_state(result[index])
            continue
        # Regra de pontuacao: obrigatorio preserva exatamente o valor recebido;
        # opcional pode ser limpo quando contem somente pontuacao/simbolos.
        if _is_punctuation_only_client_value(result[index]):
            if index not in required_set:
                result[index] = ""
            continue
        if key.startswith("nome") and "origem" not in key:
            result[index] = normalize_client_name(result[index])
            continue
        # Placeholders conhecidos continuam sem valor real.
        if _is_invalid_client_placeholder(result[index]):
            result[index] = ""
            continue
    return result

# Colunas da extracao que alimentam diretamente alguma coluna de destino.
# As demais colunas preenchidas sao preservadas na Observacao com o respectivo rotulo.
IMPORTED_SOURCE_INDEXES = {
    0, 1, 2, 3, 4, 5, 6, 11, 12, 13, 14, 15, 16, 17, 23, 24, 25, 26, 27
}


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).translate(FORBIDDEN_CHARS)
    text = "".join(ch for ch in text if unicodedata.category(ch) not in {"Cc", "Cf"})
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    return clean_text(value)


def digits_only(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    result: list[str] = []
    for char in text:
        if char.isdigit():
            try:
                result.append(str(unicodedata.digit(char)))
            except (TypeError, ValueError):
                result.append(char)
    return "".join(result)


def has_any_whitespace(value: str) -> bool:
    return any(ch.isspace() or unicodedata.category(ch) in {"Zl", "Zp", "Zs"} for ch in value)


def is_fake_phone(digits: str) -> bool:
    if not digits:
        return False
    if digits in {"0000000000", "00000000000", "0011112222", "00111112222"}:
        return True
    return len(set(digits)) == 1


def normalize_phone(value: Any, *, mobile: bool) -> str:
    digits = digits_only(value)
    if not digits:
        return ""
    if digits.startswith("55") and len(digits) in {12, 13}:
        digits = digits[2:]
    if is_fake_phone(digits):
        return ""
    if mobile and len(digits) == 10:
        digits = digits[:2] + "9" + digits[2:]
    if len(digits) == 11:
        result = f"({digits[:2]}){digits[2:7]}-{digits[7:]}"
    elif len(digits) == 10:
        result = f"({digits[:2]}){digits[2:6]}-{digits[6:]}"
    else:
        return ""
    # Regra especifica ja validada no projeto: nenhum espaco pode permanecer no telefone.
    return "".join(ch for ch in result if not ch.isspace())


def normalize_cpf(value: Any) -> str:
    digits = digits_only(value)
    if not digits:
        return ""
    if len(digits) in {9, 10}:
        digits = digits.zfill(11)
    if len(digits) != 11 or digits == "12345678900" or len(set(digits)) == 1:
        return ""
    return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"


def normalize_cep(value: Any) -> str:
    digits = digits_only(value)
    if not digits:
        return ""
    if len(digits) in {6, 7}:
        digits = digits.zfill(8)
    if len(digits) != 8 or digits == "00000000":
        return ""
    return f"{digits[:5]}-{digits[5:]}"


def normalize_numeric(value: Any) -> str:
    return digits_only(value)


def _excel_serial_date(value: Any, date1904: bool = False) -> date | None:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    # Datas modernas do Excel normalmente possuem serial acima de 10000;
    # evita confundir ano, dia ou outros numeros curtos com uma data.
    if not (10000 <= number < 100000):
        return None
    base = date(1904, 1, 1) if date1904 else date(1899, 12, 30)
    try:
        return base + timedelta(days=int(number))
    except (OverflowError, ValueError):
        return None


def format_date(value: Any, date1904: bool = False) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    text = clean_text(value)
    if not text or text == "11/22/3333":
        return ""
    candidates = [text]
    if "T" in text:
        candidates.append(text.split("T", 1)[0])
    if " " in text:
        candidates.append(text.split(" ", 1)[0])
    for candidate in candidates:
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
            try:
                return datetime.strptime(candidate, fmt).strftime("%d/%m/%Y")
            except ValueError:
                pass
    serial = _excel_serial_date(value, date1904)
    return serial.strftime("%d/%m/%Y") if serial else ""


def birth_date_from_client(values: tuple[Any, ...], date1904: bool = False) -> str:
    if len(values) > 27 and values[27] not in (None, ""):
        return format_date(values[27], date1904)
    day = values[24] if len(values) > 24 else None
    month = values[25] if len(values) > 25 else None
    year = values[26] if len(values) > 26 else None
    if day in (None, "") or month in (None, ""):
        return ""
    try:
        d = int(float(day))
        m = int(float(month))
        y = int(float(year)) if year not in (None, "") else 1900
        if y <= 0:
            y = 1900
        return date(y, m, d).strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        return ""

def _source_header_key(value: Any) -> str:
    text = clean_text(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode().casefold()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def _source_header_compact(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", _source_header_key(value))


def _header_alias_score(value: Any) -> int:
    key = _source_header_compact(value)
    if not key:
        return 0
    exact = {
        "codigo", "id", "nome", "nomecliente", "nomecompleto", "genero", "sexo", "cpf",
        "email", "telefone", "fone", "celular", "midia", "campanha", "dataultimavenda",
        "logradouro", "rua", "endereco", "numero", "numeroendereco", "bairro", "complemento",
        "cidade", "municipio", "estado", "uf", "cep", "observacao", "observacaotransferencia",
        "observacaounidade", "datacadastro", "dtcadastro", "datanascimento", "dtnascimento",
        "dianascimento", "mesnascimento", "anonascimento", "rg", "estadocivil", "profissao",
    }
    if key in exact:
        return 3
    if any(token in key for token in (
        "nomecliente", "datanascimento", "datacadastro", "dataultimavenda",
        "observacao", "logradouro", "telefone", "celular", "email", "cpf", "cep",
    )):
        return 1
    return 0


def _detect_source_header(source_rows: list[tuple[int, tuple[Any, ...]]]) -> tuple[int, tuple[Any, ...]]:
    """Localiza a linha real de cabecalho, mesmo quando existem titulos/metadados antes dela."""
    best: tuple[int, int, tuple[Any, ...]] | None = None
    for row_number, values in source_rows[:40]:
        score = sum(_header_alias_score(value) for value in values if value not in (None, ""))
        compact = {_source_header_compact(value) for value in values if value not in (None, "")}
        # Nome e pelo menos dois identificadores/endereco tornam a linha muito provavel.
        if "nome" in compact or "nomecliente" in compact:
            score += 4
        score += 2 * sum(1 for token in ("cpf", "email", "celular", "telefone", "logradouro", "cidade", "cep") if token in compact)
        candidate = (score, -row_number, values)
        if best is None or candidate[:2] > best[:2]:
            best = candidate
            best_row = row_number
    if best is None or best[0] < 8:
        # Fallback conservador para layouts antigos que realmente comecam no primeiro registro nao vazio.
        return source_rows[0]
    return best_row, best[2]


def _find_source_index(headers: tuple[Any, ...], aliases: tuple[str, ...], *, contains: bool = False, exclude: tuple[str, ...] = ()) -> int | None:
    keys = [_source_header_compact(value) for value in headers]
    alias_keys = tuple(_source_header_compact(value) for value in aliases)
    exclude_keys = tuple(_source_header_compact(value) for value in exclude)
    for index, key in enumerate(keys):
        if not key or any(ex and ex in key for ex in exclude_keys):
            continue
        if key in alias_keys:
            return index
    if contains:
        for index, key in enumerate(keys):
            if not key or any(ex and ex in key for ex in exclude_keys):
                continue
            if any(alias and alias in key for alias in alias_keys):
                return index
    return None


def _source_mapping(headers: tuple[Any, ...]) -> dict[str, int | None]:
    return {
        "source_code": _find_source_index(headers, ("Codigo", "Codigo Cliente", "Id Cliente", "ID"), contains=False),
        "name": _find_source_index(headers, ("Nome", "Nome Cliente", "Nome Completo"), contains=True, exclude=("Mae", "Pai")),
        "gender": _find_source_index(headers, ("Genero", "Sexo"), contains=True),
        "cpf": _find_source_index(headers, ("CPF",), contains=True),
        "rg": _find_source_index(headers, ("RG", "Identidade"), contains=False),
        "email": _find_source_index(headers, ("E-mail", "Email"), contains=True),
        "phone": _find_source_index(headers, ("Telefone", "Fone", "Telefone Fixo"), contains=True, exclude=("Celular",)),
        "mobile": _find_source_index(headers, ("Celular", "Telefone Celular", "Whatsapp"), contains=True),
        "mobile2": _find_source_index(headers, ("Celular 2", "Segundo Celular", "Telefone 2"), contains=True),
        "street": _find_source_index(headers, ("Logradouro", "Rua", "Endereco"), contains=True, exclude=("Numero", "Complemento")),
        "number": _find_source_index(headers, ("Numero", "Numero Endereco", "Numero do Endereco"), contains=True),
        "neighborhood": _find_source_index(headers, ("Bairro",), contains=True),
        "complement": _find_source_index(headers, ("Complemento", "Complemento Endereco"), contains=True),
        "city": _find_source_index(headers, ("Cidade", "Municipio"), contains=True),
        "state": _find_source_index(headers, ("Estado", "UF"), contains=False),
        "cep": _find_source_index(headers, ("CEP",), contains=True),
        "birth": _find_source_index(headers, ("Data Nascimento", "Dt Nascimento", "Nascimento"), contains=True, exclude=("Dia", "Mes", "Ano")),
        "birth_day": _find_source_index(headers, ("Dia Nascimento", "Dia de Nascimento"), contains=True),
        "birth_month": _find_source_index(headers, ("Mes Nascimento", "Mes de Nascimento"), contains=True),
        "birth_year": _find_source_index(headers, ("Ano Nascimento", "Ano de Nascimento"), contains=True),
        "register_date": _find_source_index(headers, ("Data Cadastro", "Dt Cadastro", "Data de Cadastro"), contains=True),
        "civil": _find_source_index(headers, ("Estado Civil", "Est Civil"), contains=True),
        "profession": _find_source_index(headers, ("Profissao",), contains=True),
    }


def _source_value(values: tuple[Any, ...], index: int | None) -> Any:
    if index is None or index < 0 or index >= len(values):
        return None
    return values[index]


def _birth_date_from_mapping(values: tuple[Any, ...], mapping: dict[str, int | None], date1904: bool) -> str:
    direct = _source_value(values, mapping.get("birth"))
    if direct not in (None, ""):
        formatted = format_date(direct, date1904)
        if formatted:
            return formatted
    day = _source_value(values, mapping.get("birth_day"))
    month = _source_value(values, mapping.get("birth_month"))
    year = _source_value(values, mapping.get("birth_year"))
    if day in (None, "") or month in (None, ""):
        return ""
    try:
        d = int(float(day))
        m = int(float(month))
        y = int(float(year)) if year not in (None, "") else 1900
        if y <= 0:
            y = 1900
        return date(y, m, d).strftime("%d/%m/%Y")
    except (TypeError, ValueError):
        return ""


def build_client_observation(
    source_headers: tuple[Any, ...],
    values: tuple[Any, ...],
    import_date: str,
    date1904: bool = False,
    imported_indexes: set[int] | None = None,
) -> str:
    """Preserva somente dados sem campo proprio, sempre com o cabecalho real da extracao."""
    imported = imported_indexes or set()
    parts: list[str] = []

    def clean_observation_text(value: Any) -> str:
        text = clean_cell(value)
        text = re.sub(
            r"(?i)(?:^|\s*\|\s*)importa[cç][aã]o\s+\d{2}/\d{2}/\d{4}(?=\s*\||$)",
            " ", text,
        )
        return clean_text(text.strip(" |"))

    limit = min(max(len(source_headers), len(values)), len(source_headers))
    for index in range(limit):
        if index in imported:
            continue
        label = clean_text(source_headers[index])
        if not label:
            continue
        raw = values[index] if index < len(values) else None
        value = clean_observation_text(raw)
        if not value or _is_garbage_client_value(value):
            continue
        key = _source_header_compact(label)
        if "data" in key:
            formatted = format_date(raw, date1904)
            if formatted:
                value = formatted
        parts.append(f"{label}: {value}")

    parts.append(f"Importação {import_date}")
    return clean_text(" | ".join(parts))

def read_csv_template(path: Path) -> tuple[list[str], str, str]:
    raw = path.read_bytes()
    text: str | None = None
    encoding: str | None = None
    for candidate in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = raw.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue
    if text is None or encoding is None:
        raise ValueError(f"Nao foi possivel identificar a codificacao do modelo: {path}")

    first_line = next((line for line in text.splitlines() if line.strip()), "")
    if ";" in first_line:
        delimiter = ";"
    else:
        try:
            delimiter = csv.Sniffer().sniff(text[:10000], delimiters=",\t|").delimiter
        except csv.Error:
            delimiter = ","

    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    try:
        headers = next(reader)
    except StopIteration as exc:
        raise ValueError(f"Modelo vazio: {path}") from exc
    # Nenhuma linha posterior do modelo e lida/utilizada.
    return headers, delimiter, encoding


def required_indexes(headers: list[str]) -> list[int]:
    indexes: list[int] = []
    for index, header in enumerate(headers):
        normalized = clean_text(header).lower()
        mandatory = "obrigatório" in normalized or "obrigatorio" in normalized
        conditional = any(
            token in normalized
            for token in (
                "obrigatório se",
                "obrigatorio se",
                "obrigatório no",
                "obrigatorio no",
            )
        )
        if mandatory and not conditional:
            indexes.append(index)
    return indexes


def apply_global_rules(row: list[Any], headers: list[str], sequence_number: int) -> list[str]:
    result = [clean_cell(value) for value in row]
    for index, header in enumerate(headers):
        if CODE_MAX6_RE.search(clean_text(header)):
            result[index] = str(sequence_number)
    return result


def find_sheet(workbook, preferred: str):
    if preferred in workbook.sheetnames:
        return workbook[preferred]
    lowered = {name.lower(): name for name in workbook.sheetnames}
    if preferred.lower() in lowered:
        return workbook[lowered[preferred.lower()]]
    return workbook[workbook.sheetnames[0]]


def write_csv(
    path: Path,
    headers: list[str],
    rows: Iterable[list[str]],
    delimiter: str,
) -> str:
    encoding = "utf-8-sig"
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    try:
        with temporary.open("w", encoding=encoding, newline="") as handle:
            writer = csv.writer(handle, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
            writer.writerow(headers)
            writer.writerows(rows)
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()
    return encoding


def write_rejected_report(
    path: Path,
    rejected: list[tuple[int, list[str]]],
) -> None:
    if not rejected:
        if path.exists():
            path.unlink()
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=";")
        writer.writerow(["Linha na planilha de extração", "Coluna obrigatória em branco"])
        for source_line, missing in rejected:
            for column in missing:
                writer.writerow([source_line, clean_text(column)])


def validate_output_csv(path: Path, delimiter: str, encoding: str) -> int:
    violations: list[str] = []
    checked_rows = 0
    with path.open("r", encoding=encoding, newline="") as handle:
        reader = csv.reader(handle, delimiter=delimiter)
        headers = next(reader)
        for csv_line, row in enumerate(reader, start=2):
            checked_rows += 1
            if len(row) != 28:
                violations.append(f"linha {csv_line}: {len(row)} colunas")
                continue
            for index in NO_WHITESPACE_OUTPUT_INDEXES:
                if index < len(row) and has_any_whitespace(row[index]):
                    violations.append(
                        f"linha {csv_line}, coluna {index + 1} ({headers[index]}): {row[index]!r}"
                    )
                    break
            if len(violations) >= 10:
                break
    if violations:
        try:
            path.unlink()
        except OSError:
            pass
        raise ValueError(
            "O CSV foi bloqueado porque ainda continha espaco em campo numerico:\n"
            + "\n".join(violations)
        )
    return checked_rows


def _identity_text(value: Any) -> str:
    text = unicodedata.normalize("NFKD", clean_text(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", text.upper()).split())


def _find_header_index(headers: list[str], tokens: tuple[str, ...]) -> int | None:
    normalized = [_identity_text(header).lower() for header in headers]
    for index, header in enumerate(normalized):
        if all(token in header for token in tokens):
            return index
    return None


def _strip_import_marker(value: str) -> tuple[str, str]:
    text = clean_text(value)
    match = re.search(r"(?i)(?:^|\s*\|\s*)(Importa[cç][aã]o\s+\d{2}/\d{2}/\d{4})\s*$", text)
    marker = match.group(1) if match else ""
    if match:
        text = clean_text(text[:match.start()].strip(" |"))
    return text, marker


def _merge_observation(primary: str, secondary: str) -> str:
    p_text, p_marker = _strip_import_marker(primary)
    s_text, s_marker = _strip_import_marker(secondary)
    pieces: list[str] = []
    seen: set[str] = set()
    for source in (p_text, s_text):
        for piece in [clean_text(item) for item in source.split("|") if clean_text(item)]:
            key = piece.casefold()
            if key not in seen:
                seen.add(key)
                pieces.append(piece)
    marker = p_marker or s_marker
    if marker:
        pieces.append(marker)
    return clean_text(" | ".join(pieces))


def deduplicate_client_rows(rows: list[list[str]], headers: list[str]) -> tuple[list[list[str]], int]:
    """Remove cadastros duplicados por identidade.

    Criterios: mesmo CPF valido; mesmo Nome + telefone/celular; ou mesmo
    Nome + e-mail. Telefone sozinho nunca elimina cadastro.
    """
    if len(rows) < 2:
        return rows, 0

    code_i = _find_header_index(headers, ("codigo",)) or 0
    name_i = _find_header_index(headers, ("nome",))
    cpf_i = _find_header_index(headers, ("cpf",))
    email_indexes = [i for i, h in enumerate(headers) if "mail" in _identity_text(h).lower()]
    phone_indexes = [
        i for i, h in enumerate(headers)
        if any(token in _identity_text(h).lower() for token in ("telefone", "celular", "fone"))
        and "ddi" not in _identity_text(h).lower()
    ]
    obs_i = _find_header_index(headers, ("observacao",))

    parent = list(range(len(rows)))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    seen: dict[tuple[str, ...], int] = {}
    for i, row in enumerate(rows):
        name = _identity_text(row[name_i]) if name_i is not None and name_i < len(row) else ""
        cpf = digits_only(row[cpf_i]) if cpf_i is not None and cpf_i < len(row) else ""
        phones = {digits_only(row[j]) for j in phone_indexes if j < len(row) and digits_only(row[j])}
        emails = {clean_text(row[j]).casefold() for j in email_indexes if j < len(row) and clean_text(row[j])}
        keys: list[tuple[str, ...]] = []
        if len(cpf) == 11 and len(set(cpf)) > 1:
            keys.append(("cpf", cpf))
        if name:
            keys.extend(("nomefone", name, phone) for phone in phones)
            keys.extend(("nomeemail", name, email) for email in emails)
        for key in keys:
            if key in seen:
                union(i, seen[key])
            else:
                seen[key] = i

    groups: dict[int, list[int]] = defaultdict(list)
    for i in range(len(rows)):
        groups[find(i)].append(i)

    kept: list[tuple[int, list[str]]] = []
    removed = 0
    for indexes in groups.values():
        if len(indexes) == 1:
            kept.append((indexes[0], rows[indexes[0]]))
            continue

        def rank(i: int) -> tuple[int, int, int]:
            code = digits_only(rows[i][code_i]) if code_i < len(rows[i]) else ""
            numeric = int(code) if code else 10**12
            filled = sum(bool(clean_text(value)) for value in rows[i])
            return (numeric, -filled, i)

        primary_i = min(indexes, key=rank)
        primary = list(rows[primary_i])
        for other_i in indexes:
            if other_i == primary_i:
                continue
            other = rows[other_i]
            for col in range(min(len(primary), len(other))):
                if col == code_i:
                    continue
                if obs_i is not None and col == obs_i:
                    primary[col] = _merge_observation(primary[col], other[col])
                elif not clean_text(primary[col]) and clean_text(other[col]):
                    primary[col] = other[col]
            removed += 1
        kept.append((min(indexes), primary))

    kept.sort(key=lambda item: item[0])
    return [row for _, row in kept], removed


def _column_index(reference: str) -> int:
    match = re.match(r"([A-Z]+)", reference.upper())
    if not match:
        return -1
    value = 0
    for char in match.group(1):
        value = value * 26 + (ord(char) - 64)
    return value - 1


def _read_client_xlsx(path: Path, preferred_sheet: str = "Sheet1", max_col: int = 128) -> tuple[str, bool, list[tuple[int, tuple[Any, ...]]]]:
    """Le a extracao diretamente do XML, ignorando completamente styles.xml.

    Isso torna o processamento tolerante a arquivos gerados por sistemas que
    possuem estilos XLSX corrompidos/incompativeis com openpyxl.
    """
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with ZipFile(path) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{{{main_ns}}}si"):
                shared.append("".join(node.text or "" for node in item.iter(f"{{{main_ns}}}t")))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        workbook_pr = workbook.find(f"{{{main_ns}}}workbookPr")
        date1904 = bool(
            workbook_pr is not None
            and workbook_pr.attrib.get("date1904", "0").casefold() in {"1", "true"}
        )
        rel_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relations: dict[str, str] = {}
        for relation in rel_root.findall(f"{{{pkg_rel_ns}}}Relationship"):
            rid = relation.attrib.get("Id", "")
            target = relation.attrib.get("Target", "")
            if not rid or not target:
                continue
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(str(PurePosixPath("xl") / target))
            relations[rid] = resolved

        sheets: list[tuple[str, str]] = []
        sheets_node = workbook.find(f"{{{main_ns}}}sheets")
        if sheets_node is not None:
            for sheet in sheets_node.findall(f"{{{main_ns}}}sheet"):
                name = sheet.attrib.get("name", "")
                rid = sheet.attrib.get(f"{{{rel_ns}}}id", "")
                target = relations.get(rid, "")
                if name and target:
                    sheets.append((name, target))
        if not sheets:
            raise ValueError(f"Nenhuma aba encontrada em {path.name}")
        selected = next((item for item in sheets if item[0].casefold() == preferred_sheet.casefold()), sheets[0])
        sheet_name, sheet_path = selected

        rows: list[tuple[int, tuple[Any, ...]]] = []
        with archive.open(sheet_path) as handle:
            for _, element in ET.iterparse(handle, events=("end",)):
                if element.tag != f"{{{main_ns}}}row":
                    continue
                try:
                    row_number = int(element.attrib.get("r", "0"))
                except ValueError:
                    row_number = 0
                values: list[Any] = [None] * max_col
                for cell in element.findall(f"{{{main_ns}}}c"):
                    index = _column_index(cell.attrib.get("r", ""))
                    if index < 0 or index >= max_col:
                        continue
                    cell_type = cell.attrib.get("t", "")
                    value_node = cell.find(f"{{{main_ns}}}v")
                    inline_node = cell.find(f"{{{main_ns}}}is")
                    value: Any = None
                    if cell_type == "inlineStr" and inline_node is not None:
                        value = "".join(node.text or "" for node in inline_node.iter(f"{{{main_ns}}}t"))
                    elif value_node is not None and value_node.text is not None:
                        raw = value_node.text
                        if cell_type == "s":
                            try:
                                value = shared[int(raw)]
                            except (ValueError, IndexError):
                                value = ""
                        elif cell_type == "b":
                            value = raw == "1"
                        else:
                            value = raw
                    values[index] = value
                if any(value not in (None, "") for value in values):
                    rows.append((row_number, tuple(values)))
                element.clear()
        return sheet_name, date1904, rows


def process_cliente(
    input_xlsx: Path,
    template_csv: Path,
    output_csv: Path,
    rejected_report: Path,
) -> tuple[int, int, int, int]:
    headers, delimiter, _template_encoding = read_csv_template(template_csv)
    if len(headers) != 28:
        raise ValueError(f"O modelo de Cliente deve ter 28 colunas; possui {len(headers)}.")

    _sheet_name, date1904, source_rows = _read_client_xlsx(input_xlsx, "Sheet1", 128)
    if not source_rows:
        raise ValueError(f"Extracao de clientes vazia: {input_xlsx.name}")
    header_row_number, header_values = _detect_source_header(source_rows)
    source_headers = tuple(header_values)
    mapping = _source_mapping(source_headers)
    if mapping.get("name") is None:
        raise ValueError(
            f"Nao foi possivel identificar a coluna de Nome no cabecalho da extracao (linha {header_row_number})."
        )

    imported_indexes = {index for index in mapping.values() if isinstance(index, int)}
    rows_out: list[list[str]] = []
    rejected: list[tuple[int, list[str]]] = []
    required = required_indexes(headers)
    import_date = date.today().strftime("%d/%m/%Y")

    for source_line, values in source_rows:
        if source_line <= header_row_number or all(value in (None, "") for value in values):
            continue

        observation = build_client_observation(
            source_headers, values, import_date, date1904, imported_indexes
        )
        phone = normalize_phone(_source_value(values, mapping.get("phone")), mobile=False)
        mobile = normalize_phone(_source_value(values, mapping.get("mobile")), mobile=True)
        mobile2 = normalize_phone(_source_value(values, mapping.get("mobile2")), mobile=True)
        row: list[Any] = [
            "",  # Codigo proprio e sempre gerado abaixo.
            _source_value(values, mapping.get("name")),
            _source_value(values, mapping.get("street")),
            _source_value(values, mapping.get("neighborhood")),
            _source_value(values, mapping.get("city")),
            _source_value(values, mapping.get("state")),
            phone,
            mobile,
            _source_value(values, mapping.get("rg")),
            normalize_cpf(_source_value(values, mapping.get("cpf"))),
            normalize_cep(_source_value(values, mapping.get("cep"))),
            _birth_date_from_mapping(values, mapping, date1904),
            _source_value(values, mapping.get("civil")),
            _source_value(values, mapping.get("gender")),
            _source_value(values, mapping.get("email")),
            format_date(_source_value(values, mapping.get("register_date")), date1904),
            observation,
            normalize_numeric(_source_value(values, mapping.get("number"))),
            mobile2,
            _source_value(values, mapping.get("complement")),
            _source_value(values, mapping.get("profession")),
            "Leads",
            "Parcerias",
            "",
            "",
            "55" if phone else "",
            "55" if mobile else "",
            "55" if mobile2 else "",
        ]
        row = apply_global_rules(row, headers, 100000 + len(rows_out))
        row = sanitize_client_output_row(row, headers, required)

        for index in NO_WHITESPACE_OUTPUT_INDEXES:
            if index < len(row):
                row[index] = "".join(ch for ch in row[index] if not ch.isspace())

        missing = [headers[index] for index in required if not clean_text(row[index])]
        if missing:
            rejected.append((source_line, missing))
            continue
        rows_out.append(row)

    rows_out, duplicates_removed = deduplicate_client_rows(rows_out, headers)
    encoding = write_csv(output_csv, headers, rows_out, delimiter)
    write_rejected_report(rejected_report, rejected)
    validated_rows = validate_output_csv(output_csv, delimiter, encoding)
    return len(rows_out), len(rejected), validated_rows, duplicates_removed

def newest(paths: list[Path]) -> Path:
    return max(paths, key=lambda item: item.stat().st_mtime)


def ensure_project_dirs() -> tuple[Path, Path, Path]:
    if not INPUT_DIR.is_dir():
        raise FileNotFoundError(f"Pasta de entrada nao encontrada: {INPUT_DIR}")
    if not OUTPUT_DIR.is_dir():
        raise FileNotFoundError(
            f"Pasta de saida nao encontrada: {OUTPUT_DIR}. "
            "Use uma pasta existente chamada saida ou saída na raiz do projeto."
        )
    return PROJECT_ROOT, INPUT_DIR, OUTPUT_DIR


def find_default_client_files() -> tuple[Path, Path]:
    ensure_project_dirs()
    inputs = [
        path
        for path in INPUT_DIR.glob("Clientes*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    if not inputs:
        inputs = [
            path
            for path in INPUT_DIR.glob("*Cliente*.xlsx")
            if path.is_file() and not path.name.startswith("~$")
        ]

    exact_model = INPUT_DIR / "modeloImportacaoCliente.csv"
    models = (
        [exact_model]
        if exact_model.is_file()
        else [path for path in INPUT_DIR.glob("modeloImportacaoCliente*.csv") if path.is_file()]
    )

    if not inputs:
        raise FileNotFoundError(f"Nao encontrei a extracao Clientes*.xlsx em {INPUT_DIR}")
    if not models:
        raise FileNotFoundError(f"Nao encontrei modeloImportacaoCliente*.csv em {INPUT_DIR}")
    return newest(inputs), (exact_model if exact_model.is_file() else newest(models))


def resolve_input_path(value: Path | None) -> Path | None:
    if value is None:
        return None
    if value.is_absolute():
        return value.resolve()
    if value.parent == Path("."):
        return (INPUT_DIR / value).resolve()
    return (PROJECT_ROOT / value).resolve()


def resolve_output_path(value: Path | None, default_name: str) -> Path:
    if value is None:
        return (OUTPUT_DIR / default_name).resolve()
    if value.is_absolute():
        return value.resolve()
    if value.parent == Path("."):
        return (OUTPUT_DIR / value).resolve()
    return (PROJECT_ROOT / value).resolve()


def script_hash() -> str:
    try:
        return hashlib.sha256(Path(__file__).read_bytes()).hexdigest()[:12]
    except OSError:
        return "indisponivel"


def main() -> int:
    print(f"Laser Rosa - {Path(__file__).name} versao {VERSION}")
    print(f"Arquivo executado: {Path(__file__).resolve()}")
    print(f"Identificador do arquivo: {script_hash()}")
    print(f"Raiz do projeto: {PROJECT_ROOT}")
    print(f"Entrada: {INPUT_DIR}")
    print(f"Saida: {OUTPUT_DIR}")

    parser = argparse.ArgumentParser(description="Tratamento automatico de Clientes - Laser Rosa")
    parser.add_argument("--entrada", type=Path, default=None, help="Extracao dentro de entrada (opcional)")
    parser.add_argument("--modelo", type=Path, default=None, help="Modelo dentro de entrada (opcional)")
    parser.add_argument("--saida", type=Path, default=None, help="Saida dentro de saida (opcional)")
    args = parser.parse_args()

    try:
        ensure_project_dirs()
        if args.entrada is None and args.modelo is None:
            input_file, model_file = find_default_client_files()
        elif args.entrada is None or args.modelo is None:
            parser.error("Informe --entrada e --modelo juntos, ou nao informe nenhum.")
        else:
            input_file = resolve_input_path(args.entrada)
            model_file = resolve_input_path(args.modelo)
            assert input_file is not None and model_file is not None

        output_file = resolve_output_path(args.saida, "planilhaTratadaCliente.csv")
        rejected_file = output_file.with_name(output_file.stem + "_linhas_rejeitadas.csv")

        if not input_file.is_file():
            raise FileNotFoundError(f"Entrada nao encontrada: {input_file}")
        if not model_file.is_file():
            raise FileNotFoundError(f"Modelo nao encontrado: {model_file}")

        print(f"Entrada selecionada: {input_file}")
        print(f"Modelo selecionado: {model_file}")
        print(f"Saida selecionada: {output_file}")

        accepted, rejected, validated, duplicates_removed = process_cliente(
            input_file,
            model_file,
            output_file,
            rejected_file,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    print(f"OK: {accepted} linhas exportadas para {output_file}")
    if duplicates_removed:
        print(f"Clientes: {duplicates_removed} cadastro(s) duplicado(s) removido(s) por identidade.")
    if rejected:
        print(f"Clientes: {rejected} linhas nao importadas por campos obrigatorios vazios.")
        print(f"Relatorio: {rejected_file}")
    else:
        print("Clientes: nenhuma linha rejeitada por campo obrigatorio em branco.")
    print(f"VALIDACAO FINAL: {validated} linhas relidas; zero espacos em campos numericos/telefones controlados.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
