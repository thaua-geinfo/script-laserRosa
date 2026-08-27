#!/usr/bin/env python3
"""Importacao de Venda de Planos e Saldos - Laser Rosa.

Objetivo
--------
Cruzar as extracoes de Vendas e Sessoes com:
- planilhaTratadaCliente(s): localiza o codigo novo do cliente;
- DE-PARA: arquivo unico de relacionamentos; usa a aba de Servicos para codigo e valor do procedimento;
- modelos de Venda de Plano e Saldo de Venda de Plano, usados somente como layout/cabecalho de saida.

Principios de seguranca
-----------------------
- nenhum registro e descartado silenciosamente;
- conflitos sem confirmacao segura bloqueiam os CSVs finais;
- cada venda gera uma unica linha, mesmo que apareca em varios procedimentos;
- cada linha de saldo da extracao e preservada, inclusive saldo zero/vazio;
- os codigos de venda reutilizam a origem se todos couberem em 6 digitos; caso contrario usam sequencia 100000+;
- o mesmo de-para de venda e aplicado ao arquivo de saldos;
- a observacao recebe "Importacao dd/mm/aaaa" e os dados da extracao nao importados em outras colunas;
- Preco e Preco Final usam ValorVenda da extracao; ValorVendaPago nao altera o valor do plano;
- modelos e planilhas tratadas de exemplo sao usados somente como referencia de layout/formato;
- planilhaTratadaCliente e uma fonte de relacionamento valida para o codigo do cliente;
- clientes ausentes dessa planilha geram uma planilha complementar baseada em modeloImportacaoCliente;
- mapa, rastreabilidades e relatorio sao consolidados em saida/validacaoVendaPlanoSaldo.xlsx;
- remove aspas simples/duplas, barras invertidas, controles e quebras de linha.

Estrutura automatica:
    raiz/entrada/  -> extracoes, DE_PARA e modelos
    raiz/saida/    -> planilhas tratadas e validacao
    raiz/scriptVendaPlanoSaldo.py

Uso:
    py scriptVendaPlanoSaldo.py

Uso explicito:
    py scriptVendaPlanoSaldo.py ^
      --vendas "GV Shopping - Vendas.xls" ^
      --sessoes "GV Shopping - sessoes.xls" ^
      --clientes "planilhaTratadaCliente.csv" ^
      --de-para "DE-PARA.xlsx" ^
      --modelo-vendas "modeloImportacaoVendaPlano.csv" ^
      --modelo-saldos "modeloImportacaoSaldoVendaPlano.csv" ^
      --modelo-clientes "modeloImportacaoCliente.csv"

O script le .xls antigos sem depender de xlrd e le .xlsx sem depender de
LibreOffice. Somente bibliotecas da instalacao padrao do Python sao exigidas.
"""

from __future__ import annotations

import argparse
import csv
import difflib
import hashlib
import io
import math
import posixpath
import re
import struct
import sys
import unicodedata
import zipfile
from decimal import Decimal, InvalidOperation
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape
VERSION = "2026-08-26.8"
MODEL_PLACEHOLDER_DATE = "01/06/2026"
FORBIDDEN_TRANSLATION = str.maketrans({'"': "", "'": "", "\\": ""})
NULL_WORDS = {"", "null", "none", "nan", "nat"}
FREE_SECTOR = 0xFFFFFFFF
END_OF_CHAIN = 0xFFFFFFFE

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR
INPUT_DIR = PROJECT_ROOT / "entrada"
OUTPUT_DIR = PROJECT_ROOT / "saida"


def _layout_key(name: str) -> str:
    text = unicodedata.normalize("NFKD", str(name))
    return "".join(ch for ch in text if not unicodedata.combining(ch)).casefold()


def _existing_layout_dir(root: Path, logical_name: str) -> Path | None:
    if not root.is_dir():
        return None
    wanted = _layout_key(logical_name)
    try:
        children = list(root.iterdir())
    except OSError:
        return None
    for child in children:
        if child.is_dir() and _layout_key(child.name) == wanted:
            return child.resolve()
    return None


def _compact_layout_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", _layout_key(value))


def _project_root_score(root: Path, entrada: Path, saida: Path) -> int:
    """Pontua uma raiz candidata para evitar escolher pastas de teste antigas.

    Uma subpasta ``teste`` pode ter recebido ``entrada``/``saida`` por engano em
    execucoes anteriores. Quando mais de um ancestral possui o par de pastas,
    preferimos a raiz que realmente concentra os arquivos do projeto.
    """
    score = 0
    try:
        input_files = [item for item in entrada.iterdir() if item.is_file() and not item.name.startswith("~$")]
    except OSError:
        input_files = []
    try:
        output_files = [item for item in saida.iterdir() if item.is_file() and not item.name.startswith("~$")]
    except OSError:
        output_files = []

    input_names = [(_compact_layout_name(item.stem), item.suffix.lower()) for item in input_files]
    output_names = [_compact_layout_name(item.stem) for item in output_files]

    # Os nomes das EXTRACOES recebidas dos clientes nao fazem parte da
    # pontuacao: eles podem ser completamente arbitrarios. A raiz correta e
    # identificada somente por elementos padronizados do proprio projeto.
    if any(re.fullmatch(r"depara\d*", name) and suffix in {".xls", ".xlsx"} for name, suffix in input_names):
        score += 40
    if any(name.startswith("modeloimportacaovendaplano") and "saldo" not in name for name, _ in input_names):
        score += 15
    if any(name.startswith("modeloimportacaosaldovendaplano") for name, _ in input_names):
        score += 15
    if any(name.startswith("planilhatratadacliente") for name in output_names):
        score += 20

    # Uma pasta de entrada realmente populada recebe apenas um pequeno peso
    # adicional. Nao usamos os nomes desses arquivos, apenas a quantidade, para
    # continuar aceitando qualquer nomenclatura enviada pelo cliente.
    supported = {".xls", ".xlsx", ".csv", ".txt"}
    score += min(sum(suffix in supported for _, suffix in input_names), 10)
    return score


def configure_project_layout() -> tuple[Path, Path, Path]:
    """Localiza a raiz EXISTENTE do projeto a partir do proprio script.

    Regra do projeto Laser Rosa:
    - scripts podem estar na raiz ou em subpastas como ``teste``;
    - procura apenas nos ancestrais do proprio arquivo (__file__);
    - a raiz precisa possuir simultaneamente ``entrada`` e ``saida``/``saída``;
    - se houver mais de uma raiz candidata, escolhe a que contem mais arquivos
      caracteristicos do projeto (DE-PARA, modelos, extracoes e cliente tratado);
    - em empate, prefere o ancestral mais externo, evitando uma estrutura
      ``teste/entrada`` criada por execucoes antigas;
    - nenhuma pasta e criada automaticamente.
    """
    global PROJECT_ROOT, INPUT_DIR, OUTPUT_DIR

    checked: list[Path] = []
    candidates: list[tuple[int, int, Path, Path, Path]] = []
    for root in (SCRIPT_DIR, *SCRIPT_DIR.parents):
        root = root.resolve()
        checked.append(root)
        entrada = _existing_layout_dir(root, "entrada")
        saida = _existing_layout_dir(root, "saida")
        if entrada is None or saida is None:
            continue
        score = _project_root_score(root, entrada, saida)
        # maior score vence; em empate, menor profundidade (raiz mais externa)
        # vence. Usamos -len(parts) para manter max().
        candidates.append((score, -len(root.parts), root, entrada, saida))

    if candidates:
        _, _, root, entrada, saida = max(candidates, key=lambda item: (item[0], item[1]))
        PROJECT_ROOT = root
        INPUT_DIR = entrada
        OUTPUT_DIR = saida
        return PROJECT_ROOT, INPUT_DIR, OUTPUT_DIR

    locais = " -> ".join(str(path) for path in checked)
    raise FileNotFoundError(
        "Nao encontrei a raiz do projeto. O script procurou, a partir da pasta "
        "onde ele esta salvo e subindo pelos diretorios-pai, uma pasta que ja "
        "contenha simultaneamente 'entrada' e 'saida' (ou 'saída'). "
        f"Locais verificados: {locais}"
    )



# ---------------------------------------------------------------------------
# Limpeza, normalizacao e formatacao
# ---------------------------------------------------------------------------

def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip().lower() in NULL_WORDS
    return False


def clean_text(value: Any) -> str:
    """Limpa texto sem apagar os espacos normais entre palavras."""
    if is_blank(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).translate(FORBIDDEN_TRANSLATION)
    text = "".join(ch for ch in text if unicodedata.category(ch) not in {"Cc", "Cf"})
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def report_text(value: Any) -> str:
    """Limpa o relatorio sem destruir barras de caminhos do Windows."""
    if is_blank(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    text = "".join(ch for ch in text if unicodedata.category(ch) not in {"Cc", "Cf"})
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def null_text(value: Any) -> str:
    text = clean_text(value)
    return text if text else "null"


def ascii_fold(value: Any) -> str:
    text = unicodedata.normalize("NFKD", clean_text(value))
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize_header(value: Any) -> str:
    text = ascii_fold(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def normalize_name(value: Any) -> str:
    text = ascii_fold(value).upper()
    text = re.sub(r"\bRECORRENCIA\b", " ", text)
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return " ".join(text.split())


def normalize_service(value: Any) -> str:
    text = ascii_fold(value).upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    return " ".join(text.split())


def service_lookup_keys(value: Any) -> list[str]:
    """Gera chaves equivalentes conservadoras para descricoes de servico."""
    key = normalize_service(value)
    if not key:
        return []
    result = [key]
    stripped = re.sub(r"^(MIGRACAO|TRANSFERENCIA|CORTESIA)\s+", "", key).strip()
    if stripped and stripped not in result:
        result.append(stripped)
    gender_fixed = stripped.replace(" MASCULINA ", " MASCULINO ").replace(" FEMININA ", " FEMININO ")
    if gender_fixed and gender_fixed not in result:
        result.append(gender_fixed)
    return result


def normalize_status(value: Any) -> str:
    return normalize_service(value)


def digits_only(value: Any) -> str:
    if is_blank(value):
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


def normalize_id(value: Any) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return format(value, ".15g")
    text = clean_text(value)
    if re.fullmatch(r"[+-]?\d+[.,]0+", text):
        return re.split(r"[.,]", text, 1)[0]
    return text


def number_string(value: Any, *, blank_as_zero: bool = False) -> str:
    if is_blank(value):
        return "0" if blank_as_zero else ""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not math.isfinite(value):
            return "0" if blank_as_zero else ""
        if value.is_integer():
            return str(int(value))
        return format(value, ".15g")
    text = clean_text(value).replace(" ", "")
    if not text:
        return "0" if blank_as_zero else ""
    if re.fullmatch(r"[+-]?\d+[.,]0+", text):
        return re.split(r"[.,]", text, 1)[0]
    return text.replace(",", ".")


def numeric_value(value: Any) -> float | None:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        result = float(value)
        return result if math.isfinite(result) else None
    text = clean_text(value).replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", ".")
    try:
        result = float(text)
    except ValueError:
        return None
    return result if math.isfinite(result) else None


def values_equal_numeric(left: Any, right: Any, tolerance: float = 1e-9) -> bool:
    a = numeric_value(left)
    b = numeric_value(right)
    if a is None or b is None:
        return is_blank(left) and is_blank(right)
    return abs(a - b) <= tolerance


def excel_datetime(value: float, *, date1904: bool = False) -> datetime:
    origin = datetime(1904, 1, 1) if date1904 else datetime(1899, 12, 30)
    return origin + timedelta(days=float(value))


def parse_date_value(value: Any, *, date1904: bool = False) -> date | None:
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return excel_datetime(float(value), date1904=date1904).date()
        except (OverflowError, ValueError):
            return None
    text = clean_text(value)
    if not text:
        return None
    candidates = [text]
    if "T" in text:
        candidates.append(text.split("T", 1)[0])
    if " " in text:
        candidates.append(text.split(" ", 1)[0])
    for candidate in candidates:
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%m/%d/%Y"):
            try:
                return datetime.strptime(candidate, fmt).date()
            except ValueError:
                continue
    numeric = numeric_value(text)
    if numeric is not None and numeric > 1000:
        try:
            return excel_datetime(numeric, date1904=date1904).date()
        except (OverflowError, ValueError):
            return None
    return None


def format_date(value: Any, *, date1904: bool = False) -> str:
    parsed = parse_date_value(value, date1904=date1904)
    return parsed.strftime("%d/%m/%Y") if parsed else ""


def format_datetime(value: Any, *, date1904: bool = False) -> str:
    if is_blank(value):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            dt = excel_datetime(float(value), date1904=date1904)
            if dt.time() == time(0, 0):
                return dt.strftime("%d/%m/%Y")
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        except (OverflowError, ValueError):
            pass
    text = clean_text(value)
    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ):
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%d/%m/%Y %H:%M:%S") if "%H" in fmt else dt.strftime("%d/%m/%Y")
        except ValueError:
            continue
    numeric = numeric_value(text)
    if numeric is not None and numeric > 1000:
        try:
            dt = excel_datetime(numeric, date1904=date1904)
            return dt.strftime("%d/%m/%Y %H:%M:%S")
        except (OverflowError, ValueError):
            pass
    return text


def phone_keys(value: Any) -> set[str]:
    digits = digits_only(value)
    if not digits:
        return set()
    if digits.startswith("55") and len(digits) in {12, 13}:
        digits = digits[2:]
    keys = {digits}
    if len(digits) == 10:
        keys.add(digits[:2] + "9" + digits[2:])
    elif len(digits) == 11 and digits[2:3] == "9":
        keys.add(digits[:2] + digits[3:])
    return {key for key in keys if key}


def sort_identifier(value: str) -> tuple[int, int | str, str]:
    if re.fullmatch(r"\d+", value):
        return (0, int(value), value)
    return (1, value, value)


def distinct_clean(values: Iterable[Any], *, formatter: Any = clean_text) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        formatted = formatter(value)
        if not formatted or formatted in seen:
            continue
        seen.add(formatted)
        result.append(formatted)
    return result


# ---------------------------------------------------------------------------
# Leitor OLE/BIFF8 para arquivos .xls antigos
# ---------------------------------------------------------------------------

def _u16(data: bytes, offset: int = 0) -> int:
    return struct.unpack_from("<H", data, offset)[0]


def _u32(data: bytes, offset: int = 0) -> int:
    return struct.unpack_from("<I", data, offset)[0]


def _u64(data: bytes, offset: int = 0) -> int:
    return struct.unpack_from("<Q", data, offset)[0]


class CompoundFile:
    def __init__(self, path: Path):
        self.path = path
        self.data = path.read_bytes()
        header = self.data[:512]
        if header[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError(f"Arquivo nao e um OLE .xls valido: {path}")
        self.sector_size = 1 << _u16(header, 30)
        self.mini_sector_size = 1 << _u16(header, 32)
        self.num_fat = _u32(header, 44)
        self.first_dir = _u32(header, 48)
        self.mini_cutoff = _u32(header, 56)
        self.first_minifat = _u32(header, 60)
        self.num_minifat = _u32(header, 64)
        self.first_difat = _u32(header, 68)
        self.num_difat = _u32(header, 72)

        difat = [_u32(header, 76 + 4 * index) for index in range(109)]
        difat = [sector for sector in difat if sector != FREE_SECTOR]
        sector = self.first_difat
        for _ in range(self.num_difat):
            block = self.sector(sector)
            entries = self.sector_size // 4 - 1
            for index in range(entries):
                item = _u32(block, 4 * index)
                if item != FREE_SECTOR:
                    difat.append(item)
            sector = _u32(block, self.sector_size - 4)

        self.fat: list[int] = []
        for sector_id in difat[: self.num_fat]:
            block = self.sector(sector_id)
            self.fat.extend(struct.unpack("<" + "I" * (self.sector_size // 4), block))

        directory = self.read_chain(self.first_dir)
        self.entries: list[dict[str, Any]] = []
        for offset in range(0, len(directory), 128):
            entry = directory[offset : offset + 128]
            if len(entry) < 128:
                break
            name_length = _u16(entry, 64)
            name = entry[: max(0, name_length - 2)].decode("utf-16le", "replace") if name_length >= 2 else ""
            self.entries.append(
                {
                    "name": name,
                    "type": entry[66],
                    "start": _u32(entry, 116),
                    "size": _u64(entry, 120),
                }
            )
        self.root = next(entry for entry in self.entries if entry["type"] == 5)
        if self.num_minifat and self.first_minifat not in (FREE_SECTOR, END_OF_CHAIN):
            raw = self.read_chain(self.first_minifat)
            self.minifat = list(struct.unpack("<" + "I" * (len(raw) // 4), raw[: len(raw) // 4 * 4]))
        else:
            self.minifat = []
        self.ministream = self.read_chain(self.root["start"], self.root["size"]) if self.root["size"] else b""

    def sector(self, sector_id: int) -> bytes:
        offset = 512 + sector_id * self.sector_size
        return self.data[offset : offset + self.sector_size]

    def chain_ids(self, start: int, fat: Sequence[int] | None = None) -> list[int]:
        if start in (FREE_SECTOR, END_OF_CHAIN):
            return []
        chain = self.fat if fat is None else fat
        result: list[int] = []
        seen: set[int] = set()
        sector = start
        while sector not in (FREE_SECTOR, END_OF_CHAIN):
            if sector >= len(chain):
                raise ValueError(f"Cadeia OLE invalida em {self.path}: setor {sector}")
            if sector in seen:
                raise ValueError(f"Ciclo na cadeia OLE em {self.path}")
            seen.add(sector)
            result.append(sector)
            sector = chain[sector]
        return result

    def read_chain(self, start: int, size: int | None = None) -> bytes:
        result = b"".join(self.sector(sector) for sector in self.chain_ids(start))
        return result if size is None else result[:size]

    def read_stream(self, name: str) -> bytes:
        entry = next((item for item in self.entries if item["name"].lower() == name.lower()), None)
        if entry is None:
            raise KeyError(f"Fluxo {name!r} nao encontrado em {self.path}")
        if entry["size"] < self.mini_cutoff and entry["type"] == 2:
            result = b"".join(
                self.ministream[sector * self.mini_sector_size : (sector + 1) * self.mini_sector_size]
                for sector in self.chain_ids(entry["start"], self.minifat)
            )
            return result[: entry["size"]]
        return self.read_chain(entry["start"], entry["size"])


class SegmentedReader:
    def __init__(self, segments: Sequence[bytes]):
        self.segments = segments
        self.segment_index = 0
        self.offset = 0

    def _advance(self) -> None:
        while self.segment_index < len(self.segments) and self.offset >= len(self.segments[self.segment_index]):
            self.segment_index += 1
            self.offset = 0

    def read_plain(self, size: int) -> bytes:
        result = bytearray()
        while len(result) < size:
            self._advance()
            if self.segment_index >= len(self.segments):
                raise EOFError("Fim inesperado em SST")
            segment = self.segments[self.segment_index]
            take = min(size - len(result), len(segment) - self.offset)
            result += segment[self.offset : self.offset + take]
            self.offset += take
        return bytes(result)

    def read_chars(self, count: int, wide: bool) -> str:
        result: list[str] = []
        remaining = count
        current_wide = wide
        while remaining:
            self._advance()
            if self.segment_index >= len(self.segments):
                raise EOFError("Fim inesperado em caracteres SST")
            if self.offset == 0 and self.segment_index > 0:
                flag = self.read_plain(1)[0]
                current_wide = bool(flag & 1)
            segment = self.segments[self.segment_index]
            width = 2 if current_wide else 1
            available = (len(segment) - self.offset) // width
            if available <= 0:
                self.offset = len(segment)
                continue
            take = min(remaining, available)
            raw = segment[self.offset : self.offset + take * width]
            self.offset += take * width
            result.append(raw.decode("utf-16le" if current_wide else "latin1", "replace"))
            remaining -= take
        return "".join(result)


def _parse_sst(segments: Sequence[bytes]) -> list[str]:
    reader = SegmentedReader(segments)
    reader.read_plain(4)  # total de ocorrencias
    unique = _u32(reader.read_plain(4))
    strings: list[str] = []
    for index in range(unique):
        try:
            char_count = _u16(reader.read_plain(2))
            flags = reader.read_plain(1)[0]
            rich = bool(flags & 0x08)
            extended = bool(flags & 0x04)
            wide = bool(flags & 0x01)
            runs = _u16(reader.read_plain(2)) if rich else 0
            extended_size = _u32(reader.read_plain(4)) if extended else 0
            text = reader.read_chars(char_count, wide)
            if runs:
                reader.read_plain(runs * 4)
            if extended_size:
                reader.read_plain(extended_size)
            strings.append(text)
        except Exception as exc:
            raise RuntimeError(f"Falha ao ler SST {index + 1}/{unique}: {exc}") from exc
    return strings


def _iter_biff_records(data: bytes, start: int = 0, end: int | None = None) -> Iterator[tuple[int, int, bytes]]:
    limit = len(data) if end is None else min(end, len(data))
    position = start
    while position + 4 <= limit:
        record_id, length = struct.unpack_from("<HH", data, position)
        payload = data[position + 4 : position + 4 + length]
        if position + 4 + length > limit:
            break
        yield position, record_id, payload
        position += 4 + length


def _parse_biff_string(payload: bytes, offset: int = 0, char_count_bytes: int = 2) -> tuple[str, int]:
    if char_count_bytes == 2:
        char_count = _u16(payload, offset)
        offset += 2
    else:
        char_count = payload[offset]
        offset += 1
    flags = payload[offset]
    offset += 1
    rich = bool(flags & 8)
    extended = bool(flags & 4)
    wide = bool(flags & 1)
    runs = _u16(payload, offset) if rich else 0
    if rich:
        offset += 2
    extended_size = _u32(payload, offset) if extended else 0
    if extended:
        offset += 4
    byte_count = char_count * (2 if wide else 1)
    raw = payload[offset : offset + byte_count]
    text = raw.decode("utf-16le" if wide else "latin1", "replace")
    return text, offset + byte_count + runs * 4 + extended_size


def _decode_rk(rk: int) -> float:
    divide_100 = bool(rk & 1)
    is_integer = bool(rk & 2)
    if is_integer:
        value = rk >> 2
        if value & (1 << 29):
            value -= 1 << 30
        result = float(value)
    else:
        raw = struct.pack("<I", 0) + struct.pack("<I", rk & 0xFFFFFFFC)
        result = struct.unpack("<d", raw)[0]
    return result / 100 if divide_100 else result


def _compact_number(value: float) -> int | float:
    return int(value) if value.is_integer() else value


class XlsBook:
    def __init__(self, path: Path):
        self.path = path
        compound = CompoundFile(path)
        stream_name = "Workbook" if any(entry["name"] == "Workbook" for entry in compound.entries) else "Book"
        self.data = compound.read_stream(stream_name)
        self.sheets: list[tuple[str, int]] = []
        self.shared_strings: list[str] = []
        self.date1904 = False
        records = list(_iter_biff_records(self.data))
        index = 0
        while index < len(records):
            _, record_id, payload = records[index]
            if record_id == 0x0085:  # BOUNDSHEET
                offset = _u32(payload, 0)
                length = payload[6]
                flags = payload[7]
                raw = payload[8 : 8 + (2 * length if flags & 1 else length)]
                name = raw.decode("utf-16le" if flags & 1 else "latin1", "replace")
                self.sheets.append((name, offset))
            elif record_id == 0x0022 and len(payload) >= 2:
                self.date1904 = bool(_u16(payload))
            elif record_id == 0x00FC:  # SST + CONTINUE
                segments = [payload]
                next_index = index + 1
                while next_index < len(records) and records[next_index][1] == 0x003C:
                    segments.append(records[next_index][2])
                    next_index += 1
                self.shared_strings = _parse_sst(segments)
                index = next_index - 1
            index += 1

    @property
    def sheet_names(self) -> list[str]:
        return [name for name, _ in self.sheets]

    def sheet_cells(self, name: str) -> tuple[dict[tuple[int, int], Any], tuple[int, int]]:
        match = next((sheet for sheet in self.sheets if sheet[0].lower() == name.lower()), None)
        if match is None:
            raise KeyError(f"Aba {name!r} nao encontrada em {self.path}. Abas: {self.sheet_names}")
        start = match[1]
        cells: dict[tuple[int, int], Any] = {}
        pending_string: tuple[int, int] | None = None
        max_row = -1
        max_col = -1
        for _, record_id, payload in _iter_biff_records(self.data, start):
            if record_id == 0x000A:  # EOF
                break
            if record_id == 0x00FD:  # LABELSST
                row, col = _u16(payload, 0), _u16(payload, 2)
                string_index = _u32(payload, 6)
                cells[(row, col)] = self.shared_strings[string_index] if string_index < len(self.shared_strings) else ""
            elif record_id == 0x0203:  # NUMBER
                row, col = _u16(payload, 0), _u16(payload, 2)
                cells[(row, col)] = _compact_number(struct.unpack_from("<d", payload, 6)[0])
            elif record_id == 0x027E:  # RK
                row, col = _u16(payload, 0), _u16(payload, 2)
                cells[(row, col)] = _compact_number(_decode_rk(_u32(payload, 6)))
            elif record_id == 0x00BD:  # MULRK
                row = _u16(payload, 0)
                first_col = _u16(payload, 2)
                body = payload[4:-2]
                for body_offset in range(0, len(body), 6):
                    if body_offset + 6 > len(body):
                        break
                    col = first_col + body_offset // 6
                    cells[(row, col)] = _compact_number(_decode_rk(_u32(body, body_offset + 2)))
                    max_row = max(max_row, row)
                    max_col = max(max_col, col)
                continue
            elif record_id == 0x0006:  # FORMULA
                row, col = _u16(payload, 0), _u16(payload, 2)
                result = payload[6:14]
                if len(result) == 8 and result[6:8] == b"\xff\xff":
                    result_type = result[0]
                    if result_type == 0:
                        cells[(row, col)] = ""
                        pending_string = (row, col)
                    elif result_type == 1:
                        cells[(row, col)] = bool(result[2])
                        pending_string = None
                    elif result_type == 2:
                        cells[(row, col)] = f"#ERR{result[2]}"
                        pending_string = None
                    else:
                        cells[(row, col)] = ""
                        pending_string = None
                else:
                    cells[(row, col)] = _compact_number(struct.unpack("<d", result)[0])
                    pending_string = None
            elif record_id == 0x0207 and pending_string is not None:  # STRING
                try:
                    text, _ = _parse_biff_string(payload, 0, 2)
                except Exception:
                    text = ""
                cells[pending_string] = text
                pending_string = None
                continue
            elif record_id == 0x0205:  # BOOLERR
                row, col = _u16(payload, 0), _u16(payload, 2)
                value = payload[6]
                cells[(row, col)] = f"#ERR{value}" if payload[7] else bool(value)
            elif record_id == 0x0204:  # LABEL
                row, col = _u16(payload, 0), _u16(payload, 2)
                try:
                    text, _ = _parse_biff_string(payload, 6, 2)
                except Exception:
                    text = ""
                cells[(row, col)] = text
            else:
                continue
            max_row = max(max_row, row)
            max_col = max(max_col, col)
        return cells, (max_row, max_col)


# ---------------------------------------------------------------------------
# Leitor simples de .xlsx usando ZIP + XML
# ---------------------------------------------------------------------------

def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _xlsx_column_index(reference: str) -> int:
    letters = "".join(ch for ch in reference if ch.isalpha()).upper()
    result = 0
    for char in letters:
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result - 1


class XlsxBook:
    def __init__(self, path: Path):
        self.path = path
        self.archive = zipfile.ZipFile(path)
        self.shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in self.archive.namelist():
            root = ET.fromstring(self.archive.read("xl/sharedStrings.xml"))
            for item in root:
                text = "".join(node.text or "" for node in item.iter() if _local_name(node.tag) == "t")
                self.shared_strings.append(text)

        workbook = ET.fromstring(self.archive.read("xl/workbook.xml"))
        relations_root = ET.fromstring(self.archive.read("xl/_rels/workbook.xml.rels"))
        relations = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relations_root
            if "Id" in relation.attrib and "Target" in relation.attrib
        }
        self.sheets: list[tuple[str, str]] = []
        relation_attribute = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        for node in workbook.iter():
            if _local_name(node.tag) != "sheet":
                continue
            name = node.attrib.get("name", "")
            relation_id = node.attrib.get(relation_attribute, "")
            target = relations.get(relation_id, "")
            if not target:
                continue
            if target.startswith("/"):
                path_in_zip = target.lstrip("/")
            else:
                path_in_zip = posixpath.normpath(posixpath.join("xl", target))
            self.sheets.append((name, path_in_zip))
        self.date1904 = False
        for node in workbook.iter():
            if _local_name(node.tag) == "workbookPr":
                self.date1904 = node.attrib.get("date1904", "0").lower() in {"1", "true"}
                break

    @property
    def sheet_names(self) -> list[str]:
        return [name for name, _ in self.sheets]

    def sheet_cells(self, name: str) -> tuple[dict[tuple[int, int], Any], tuple[int, int]]:
        match = next((sheet for sheet in self.sheets if sheet[0].lower() == name.lower()), None)
        if match is None:
            raise KeyError(f"Aba {name!r} nao encontrada em {self.path}. Abas: {self.sheet_names}")
        cells: dict[tuple[int, int], Any] = {}
        max_row = -1
        max_col = -1
        with self.archive.open(match[1]) as handle:
            for _, element in ET.iterparse(handle, events=("end",)):
                if _local_name(element.tag) != "c":
                    continue
                reference = element.attrib.get("r", "")
                if not reference:
                    element.clear()
                    continue
                row_match = re.search(r"\d+", reference)
                if row_match is None:
                    element.clear()
                    continue
                row = int(row_match.group()) - 1
                col = _xlsx_column_index(reference)
                cell_type = element.attrib.get("t", "")
                value_node = next((child for child in element if _local_name(child.tag) == "v"), None)
                inline_node = next((child for child in element if _local_name(child.tag) == "is"), None)
                value: Any = ""
                if cell_type == "inlineStr" and inline_node is not None:
                    value = "".join(node.text or "" for node in inline_node.iter() if _local_name(node.tag) == "t")
                elif value_node is not None and value_node.text is not None:
                    raw = value_node.text
                    if cell_type == "s":
                        try:
                            index = int(raw)
                            value = self.shared_strings[index] if index < len(self.shared_strings) else ""
                        except ValueError:
                            value = ""
                    elif cell_type in {"str", "e"}:
                        value = raw
                    elif cell_type == "b":
                        value = raw == "1"
                    else:
                        try:
                            numeric = float(raw)
                            value = int(numeric) if numeric.is_integer() else numeric
                        except ValueError:
                            value = raw
                cells[(row, col)] = value
                max_row = max(max_row, row)
                max_col = max(max_col, col)
                element.clear()
        return cells, (max_row, max_col)


# ---------------------------------------------------------------------------
# Estruturas de tabela e relatorio
# ---------------------------------------------------------------------------
@dataclass
class Table:
    path: Path
    sheet: str
    headers: list[Any]
    rows: list[tuple[int, list[Any]]]
    date1904: bool = False


@dataclass
class Issue:
    level: str
    category: str
    file: str = ""
    line: str = ""
    sale_id: str = ""
    field: str = ""
    origin: str = ""
    compared: str = ""
    detail: str = ""


@dataclass
class ValidationReport:
    issues: list[Issue] = field(default_factory=list)

    def add(
        self,
        level: str,
        category: str,
        *,
        file: Path | str = "",
        line: int | str = "",
        sale_id: Any = "",
        field: str = "",
        origin: Any = "",
        compared: Any = "",
        detail: str = "",
    ) -> None:
        self.issues.append(
            Issue(
                level=level,
                category=category,
                file=Path(file).name if file else "",
                line=str(line) if line != "" else "",
                sale_id=normalize_id(sale_id),
                field=report_text(field),
                origin=report_text(origin),
                compared=report_text(compared),
                detail=report_text(detail),
            )
        )

    def info(self, category: str, **kwargs: Any) -> None:
        self.add("INFO", category, **kwargs)

    def warning(self, category: str, **kwargs: Any) -> None:
        self.add("AVISO", category, **kwargs)

    def error(self, category: str, **kwargs: Any) -> None:
        self.add("ERRO", category, **kwargs)

    @property
    def error_count(self) -> int:
        return sum(issue.level == "ERRO" for issue in self.issues)

    @property
    def warning_count(self) -> int:
        return sum(issue.level == "AVISO" for issue in self.issues)


def read_csv_table(path: Path) -> Table:
    raw = path.read_bytes()
    text: str | None = None
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"Nao foi possivel identificar a codificacao: {path}")
    # Os modelos da Laser Rosa podem conter somente uma linha de cabecalho.
    # Nessa situacao o csv.Sniffer costuma confundir virgulas existentes nos
    # titulos (ex.: "obrigatorio, max 6 digitos") com o delimitador do arquivo.
    # Os modelos oficiais usam ponto e virgula; por isso damos prioridade a ele
    # sempre que estiver presente no primeiro registro.
    first_line = next((line for line in text.splitlines() if line.strip()), "")
    if ";" in first_line:
        delimiter = ";"
    else:
        try:
            delimiter = csv.Sniffer().sniff(text[:10000], delimiters=",\t|").delimiter
        except csv.Error:
            delimiter = ","
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    all_rows = list(reader)
    if not all_rows:
        raise ValueError(f"Arquivo vazio: {path}")
    headers = list(all_rows[0])
    rows = [(line, list(row)) for line, row in enumerate(all_rows[1:], start=2) if any(clean_text(value) for value in row)]
    return Table(path=path, sheet=path.name, headers=headers, rows=rows)


def _select_sheet(
    sheet_names: Sequence[str],
    preferred: Sequence[str],
    *,
    require_preferred: bool = False,
) -> str:
    if not sheet_names:
        raise ValueError("Arquivo sem abas")

    normalized = [(normalize_header(name), name) for name in sheet_names]
    preferred_normalized = [normalize_header(name) for name in preferred if normalize_header(name)]

    # 1) correspondencia exata ignorando acentos, caixa e pontuacao.
    for wanted in preferred_normalized:
        for current, original in normalized:
            if current == wanted:
                return original

    # 2) nomes como "DE-PARA Servicos", "Servicos Laser", "Serv" etc.
    for wanted in preferred_normalized:
        for current, original in normalized:
            tokens = current.split()
            if wanted in tokens or any(token.startswith(wanted) or wanted.startswith(token) for token in tokens if len(token) >= 4):
                return original

    if require_preferred and preferred:
        raise KeyError(
            "Nenhuma aba correspondente foi encontrada. "
            f"Procurado: {', '.join(preferred)}. Abas disponiveis: {list(sheet_names)}"
        )
    return sheet_names[0]


def read_workbook_table(
    path: Path,
    preferred_sheets: Sequence[str] = (),
    *,
    require_preferred_sheet: bool = False,
) -> Table:
    suffix = path.suffix.lower()
    if suffix == ".xls":
        book: Any = XlsBook(path)
    elif suffix == ".xlsx":
        book = XlsxBook(path)
    else:
        raise ValueError(f"Formato de planilha nao suportado: {path}")
    sheet = _select_sheet(book.sheet_names, preferred_sheets, require_preferred=require_preferred_sheet)
    cells, shape = book.sheet_cells(sheet)
    max_row, max_col = shape
    if max_row < 0 or max_col < 0:
        raise ValueError(f"Aba vazia: {path} / {sheet}")
    headers = [cells.get((0, col), "") for col in range(max_col + 1)]
    rows: list[tuple[int, list[Any]]] = []
    for row in range(1, max_row + 1):
        values = [cells.get((row, col), "") for col in range(max_col + 1)]
        if any(not is_blank(value) for value in values):
            rows.append((row + 1, values))
    return Table(path=path, sheet=sheet, headers=headers, rows=rows, date1904=bool(book.date1904))


def read_table(path: Path, preferred_sheets: Sequence[str] = ()) -> Table:
    if path.suffix.lower() in {".csv", ".txt"}:
        return read_csv_table(path)
    return read_workbook_table(path, preferred_sheets)


def read_depara_sheet(path: Path, aliases: Sequence[str]) -> Table:
    """Abre o arquivo unico DE-PARA e exige uma aba da categoria solicitada."""
    if path.suffix.lower() not in {".xls", ".xlsx"}:
        raise ValueError(
            f"O arquivo DE-PARA precisa ser .xls ou .xlsx para conter abas: {path}"
        )
    return read_workbook_table(path, aliases, require_preferred_sheet=True)


def _header_exists(headers: Sequence[Any], aliases: Sequence[str]) -> bool:
    """Verifica uma coluna pelo layout, sem depender do nome do arquivo."""
    normalized_headers = [normalize_header(header) for header in headers]
    for alias in aliases:
        wanted = normalize_header(alias)
        if not wanted:
            continue
        wanted_tokens = set(wanted.split())
        for current in normalized_headers:
            if current == wanted:
                return True
            if wanted_tokens and wanted_tokens.issubset(set(current.split())):
                return True
    return False


def _is_sales_extraction_layout(headers: Sequence[Any]) -> tuple[bool, int]:
    """Reconhece a extracao principal/auxiliar de vendas pelos cabecalhos."""
    base = [
        ["VendaId", "Codigo Venda"],
        ["StatusVenda"],
        ["ClienteNome"],
        ["ClienteCpf"],
    ]
    if not all(_header_exists(headers, aliases) for aliases in base):
        return False, 0
    markers = [
        ["VendaDataFaturamento"],
        ["ValorVenda"],
        ["PacoteServico"],
        ["Observacao"],
        ["QtdTotal", "Quantidade"],
        ["QtdRealizado", "QuantidadeUtilizada"],
        ["Codigo Cliente", "Codigo do Cliente"],
    ]
    count = sum(_header_exists(headers, aliases) for aliases in markers)
    return count >= 3, count


def _is_sessions_extraction_layout(headers: Sequence[Any]) -> tuple[bool, int]:
    """Reconhece a extracao de sessoes/saldos pelos cabecalhos."""
    base = [
        ["VendaId"],
        ["StatusVenda"],
        ["ClienteNome"],
        ["ClienteCpf"],
    ]
    if not all(_header_exists(headers, aliases) for aliases in base):
        return False, 0
    markers = [
        ["ServicoId"],
        ["ServicoNome", "Nome Servico"],
        ["NomeItem", "Item"],
        ["Quantidade", "QtdTotal"],
        ["QuantidadeUtilizada", "QtdRealizado"],
        ["PacoteServicosId"],
    ]
    count = sum(_header_exists(headers, aliases) for aliases in markers)
    return count >= 4, count


def _candidate_spreadsheets(search_dir: Path) -> list[Path]:
    supported = {".xls", ".xlsx", ".csv", ".txt"}
    result: list[Path] = []
    for path in search_dir.iterdir():
        if not path.is_file() or path.name.startswith("~$") or path.suffix.lower() not in supported:
            continue
        normalized = normalized_filename(path)
        # Modelos e o DE-PARA possuem funcao propria e nunca devem concorrer
        # como planilhas de extracao do cliente.
        if normalized.startswith("modeloimportacao"):
            continue
        if re.fullmatch(r"depara\d*", normalized):
            continue
        result.append(path.resolve())
    return result


def discover_extraction_file(
    argument: Path | None,
    kind: str,
    *,
    search_dir: Path,
    exclude: Iterable[Path] = (),
) -> Path:
    """Localiza Vendas/Sessoes pelo CONTEUDO, nunca pelo nome do arquivo.

    Para Vendas podem existir arquivos auxiliares com o mesmo layout. Nesse
    caso a extracao principal e a candidata compativel com maior quantidade de
    linhas. Isso permite receber nomes completamente diferentes por cliente.
    """
    description = "Extracao de Vendas" if kind == "vendas" else "Extracao de Sessoes"
    excluded = {Path(path).resolve() for path in exclude}
    if argument is not None:
        path = _resolve_explicit_path(argument, default_dir=search_dir)
        if not path.is_file():
            raise FileNotFoundError(f"{description} nao encontrado: {path}")
        if path.suffix.lower() not in {".xls", ".xlsx", ".csv", ".txt"}:
            raise ValueError(f"Formato nao suportado para {description}: {path}")
        return path

    matches: list[tuple[int, int, float, Path]] = []
    inspected: list[str] = []
    for path in _candidate_spreadsheets(search_dir):
        if path in excluded:
            continue
        try:
            table = read_table(path, ("expdata 1", "Planilha1", "Sheet1", "Plan1"))
        except Exception as exc:
            inspected.append(f"{path.name}: nao lido ({type(exc).__name__})")
            continue
        if kind == "vendas":
            compatible, markers = _is_sales_extraction_layout(table.headers)
        else:
            compatible, markers = _is_sessions_extraction_layout(table.headers)
        inspected.append(
            f"{path.name}: linhas={len(table.rows)}, marcadores={markers}, compativel={'sim' if compatible else 'nao'}"
        )
        if compatible:
            matches.append((len(table.rows), markers, path.stat().st_mtime, path))

    if not matches:
        detail = "; ".join(inspected[:20]) or "nenhuma planilha suportada encontrada"
        raise FileNotFoundError(
            f"{description} nao identificado pelo layout em: {search_dir}. "
            f"Arquivos analisados: {detail}"
        )

    matches.sort(key=lambda item: (item[0], item[1], item[2]), reverse=True)
    return matches[0][3]


def discover_sales_auxiliary(
    argument: Path | None,
    kind: str,
    *,
    search_dir: Path,
    exclude: Iterable[Path] = (),
) -> Path | None:
    """Localiza auxiliares de vendas pelo conteudo quando o nome variar."""
    description = "Vendas pendentes de pagamento" if kind == "pendentes" else "Vendas sem data de faturamento"
    excluded = {Path(path).resolve() for path in exclude}
    if argument is not None:
        path = _resolve_explicit_path(argument, default_dir=search_dir)
        if not path.is_file():
            raise FileNotFoundError(f"{description} nao encontrado: {path}")
        return path

    matches: list[tuple[int, float, Path]] = []
    for path in _candidate_spreadsheets(search_dir):
        if path in excluded:
            continue
        try:
            table = read_table(path, ("expdata 1", "Planilha1", "Sheet1", "Plan1"))
        except Exception:
            continue
        compatible, _ = _is_sales_extraction_layout(table.headers)
        if not compatible or not table.rows:
            continue
        try:
            status_index = find_header(table.headers, ["StatusVenda"], required=True)
            billing_index = find_header(table.headers, ["VendaDataFaturamento"], required=True)
        except KeyError:
            continue
        statuses = [normalize_status(row_value(row, status_index)) for _, row in table.rows]
        dates = [format_date(row_value(row, billing_index), date1904=table.date1904) for _, row in table.rows]
        all_without_date = all(not value for value in dates)
        all_pending = all(status == "PENDENTE A PAGAMENTO" for status in statuses if status) and any(statuses)
        if kind == "pendentes":
            valid = all_without_date and all_pending
        else:
            valid = all_without_date
        if valid:
            matches.append((len(table.rows), path.stat().st_mtime, path))

    if not matches:
        return None
    matches.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return matches[0][2]


def ensure_width(row: Sequence[Any], width: int) -> list[Any]:
    result = list(row)
    if len(result) < width:
        result.extend([""] * (width - len(result)))
    return result


def header_positions(headers: Sequence[Any]) -> dict[str, list[int]]:
    result: dict[str, list[int]] = defaultdict(list)
    for index, header in enumerate(headers):
        normalized = normalize_header(header)
        if normalized:
            result[normalized].append(index)
    return result


def find_header(
    headers: Sequence[Any],
    aliases: Sequence[str],
    *,
    required: bool = True,
    occurrence: int = 0,
) -> int | None:
    normalized_headers = [normalize_header(header) for header in headers]
    normalized_aliases = [normalize_header(alias) for alias in aliases]
    exact: list[int] = []
    for alias in normalized_aliases:
        exact.extend(index for index, header in enumerate(normalized_headers) if header == alias)
    exact = sorted(set(exact))
    if len(exact) > occurrence:
        return exact[occurrence]
    contains: list[int] = []
    for alias in normalized_aliases:
        alias_tokens = set(alias.split())
        for index, header in enumerate(normalized_headers):
            if alias_tokens and alias_tokens.issubset(set(header.split())):
                contains.append(index)
    contains = sorted(set(contains))
    if len(contains) > occurrence:
        return contains[occurrence]
    if required:
        raise KeyError(f"Coluna nao encontrada. Procurado: {', '.join(aliases)}. Cabecalhos: {list(headers)}")
    return None


def row_value(row: Sequence[Any], index: int | None) -> Any:
    if index is None or index < 0 or index >= len(row):
        return ""
    return row[index]


# ---------------------------------------------------------------------------
# Descoberta de arquivos e gravacao atomica
# ---------------------------------------------------------------------------

def ensure_project_folders() -> tuple[Path, Path]:
    """Valida as pastas existentes do projeto. Nunca cria entrada/saida."""
    if not INPUT_DIR.is_dir():
        raise FileNotFoundError(f"Pasta de entrada nao encontrada: {INPUT_DIR}")
    if not OUTPUT_DIR.is_dir():
        raise FileNotFoundError(f"Pasta de saida nao encontrada: {OUTPUT_DIR}")
    return INPUT_DIR, OUTPUT_DIR


def require_existing_parent(path: Path) -> None:
    if not path.parent.is_dir():
        raise FileNotFoundError(
            f"Pasta de destino nao existe: {path.parent}. O script nao cria novas pastas automaticamente."
        )


def normalized_filename(path: Path) -> str:
    return re.sub(r"[^a-z0-9]+", "", ascii_fold(path.stem).lower())


def newest(paths: Sequence[Path]) -> Path:
    return max(paths, key=lambda item: item.stat().st_mtime)


def _resolve_explicit_path(argument: Path, *, default_dir: Path) -> Path:
    expanded = argument.expanduser()
    if expanded.is_absolute():
        return expanded.resolve()
    # Um nome simples e procurado na pasta padrao; caminhos com subpasta sao
    # interpretados a partir da raiz onde o script esta salvo.
    if expanded.parent == Path('.'):
        return (default_dir / expanded).resolve()
    return (PROJECT_ROOT / expanded).resolve()


def resolve_file(
    argument: Path | None,
    description: str,
    predicate: Any,
    *,
    search_dir: Path,
    required: bool = True,
) -> Path | None:
    if argument is not None:
        path = _resolve_explicit_path(argument, default_dir=search_dir)
        if not path.is_file():
            raise FileNotFoundError(f"{description} nao encontrado: {path}")
        return path

    candidates = [
        path for path in search_dir.iterdir()
        if path.is_file() and not path.name.startswith("~$") and predicate(path)
    ]
    if candidates:
        # Nome exato tem prioridade; em variacoes como (1), (2), usa o mais recente.
        return newest(candidates)
    if required:
        raise FileNotFoundError(f"{description} nao encontrado em: {search_dir}")
    return None


def resolve_exact_or_variant(
    argument: Path | None,
    description: str,
    *,
    search_dir: Path,
    exact_stem: str,
    suffixes: set[str],
    exclude_tokens: Sequence[str] = (),
    required: bool = True,
) -> Path | None:
    normalized_exact = re.sub(r"[^a-z0-9]+", "", ascii_fold(exact_stem).lower())

    def predicate(path: Path) -> bool:
        if path.suffix.lower() not in suffixes:
            return False
        normalized = normalized_filename(path)
        if not normalized.startswith(normalized_exact):
            return False
        return not any(token in normalized for token in exclude_tokens)

    if argument is not None:
        return resolve_file(argument, description, predicate, search_dir=search_dir, required=required)

    # Primeiro tenta nomes exatos em ordem de extensao preferida.
    for suffix in (".xls", ".xlsx", ".csv", ".txt"):
        if suffix not in suffixes:
            continue
        exact = search_dir / f"{exact_stem}{suffix}"
        if exact.is_file() and predicate(exact):
            return exact.resolve()
    return resolve_file(None, description, predicate, search_dir=search_dir, required=required)


def resolve_depara_file(argument: Path | None, *, search_dir: Path) -> Path:
    """Localiza exclusivamente o arquivo unico DE-PARA (.xls/.xlsx).

    Aceita copias como ``DE-PARA (1).xls``. Nao procura arquivos separados
    ``DE_PARA_SERVICOS`` porque os relacionamentos ficam em abas do mesmo arquivo.
    """
    if argument is not None:
        path = _resolve_explicit_path(argument, default_dir=search_dir)
        if not path.is_file():
            raise FileNotFoundError(f"DE-PARA nao encontrado: {path}")
        if path.suffix.lower() not in {".xls", ".xlsx"}:
            raise ValueError(f"DE-PARA precisa ser .xls ou .xlsx: {path}")
        return path

    candidates: list[Path] = []
    for path in search_dir.iterdir():
        if not path.is_file() or path.name.startswith("~$") or path.suffix.lower() not in {".xls", ".xlsx"}:
            continue
        normalized = normalized_filename(path)
        if re.fullmatch(r"depara\d*", normalized):
            candidates.append(path)
    if not candidates:
        raise FileNotFoundError(
            f"Arquivo DE-PARA (.xls/.xlsx) nao encontrado em: {search_dir}. "
            "O arquivo deve se chamar DE-PARA e conter abas como Servicos, Salas e FP."
        )

    # Nome exato antes de copias com (1), (2); entre equivalentes, o mais recente.
    exact = [path for path in candidates if normalized_filename(path) == "depara"]
    return newest(exact or candidates).resolve()


def write_csv_atomic(path: Path, headers: Sequence[Any], rows: Iterable[Sequence[Any]]) -> int:
    require_existing_parent(path)
    temporary = path.with_name(path.name + ".tmp")
    count = 0
    try:
        with temporary.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writerow(list(headers))
            for row in rows:
                writer.writerow([clean_text(value) for value in row])
                count += 1
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()
    return count


def validate_csv_output(path: Path, expected_headers: Sequence[str], expected_rows: int) -> None:
    table = read_csv_table(path)
    if list(table.headers) != list(expected_headers):
        raise ValueError(f"Cabecalho final divergiu do modelo em {path}")
    if len(table.rows) != expected_rows:
        raise ValueError(f"Quantidade final incorreta em {path}: {len(table.rows)}; esperado {expected_rows}")
    for line, row in table.rows:
        if len(row) != len(expected_headers):
            raise ValueError(f"Linha {line} de {path} possui {len(row)} colunas; esperado {len(expected_headers)}")
        for value in row:
            if any(char in value for char in ('"', "'", "\\", "\r", "\n", "\t")):
                raise ValueError(f"Linha {line} de {path} ainda contem caractere proibido")


def write_report(path: Path, report: ValidationReport) -> None:
    headers = ["Nivel", "Tipo", "Arquivo", "Linha", "Venda origem", "Campo", "Valor origem", "Valor comparado", "Detalhe"]
    rows = [
        [
            issue.level,
            issue.category,
            issue.file,
            issue.line,
            issue.sale_id,
            issue.field,
            issue.origin,
            issue.compared,
            issue.detail,
        ]
        for issue in report.issues
    ]
    require_existing_parent(path)
    temporary = path.with_name(path.name + ".tmp")
    try:
        with temporary.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writerow(headers)
            for row in rows:
                writer.writerow([report_text(value) for value in row])
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()




def _xlsx_col_letter(index: int) -> str:
    result = ""
    value = index + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _xlsx_cell_text(value: Any) -> str:
    text = report_text(value)
    # Limite oficial de texto por celula do Excel.
    return text[:32767]


def _sheet_xml(headers: Sequence[Any], rows: Iterable[Sequence[Any]]) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        '<sheetData>',
    ]

    def add_row(row_number: int, values: Sequence[Any], header: bool = False) -> None:
        cells: list[str] = []
        for col, value in enumerate(values):
            ref = f"{_xlsx_col_letter(col)}{row_number}"
            text = xml_escape(_xlsx_cell_text(value), {'"': '&quot;'})
            style = ' s="1"' if header else ''
            preserve = ' xml:space="preserve"' if text.startswith(' ') or text.endswith(' ') else ''
            cells.append(f'<c r="{ref}" t="inlineStr"{style}><is><t{preserve}>{text}</t></is></c>')
        parts.append(f'<row r="{row_number}">' + ''.join(cells) + '</row>')

    add_row(1, headers, True)
    row_number = 2
    for row in rows:
        add_row(row_number, row, False)
        row_number += 1
    parts.extend(['</sheetData>', '</worksheet>'])
    return ''.join(parts)


def write_validation_xlsx(
    path: Path,
    *,
    sale_codes: dict[str, str],
    sales_trace: list[list[str]],
    balance_trace: list[list[str]],
    report: ValidationReport,
) -> None:
    """Consolida mapa, rastreabilidades e validacao em um unico XLSX."""
    require_existing_parent(path)
    temporary = path.with_name(path.name + '.tmp')

    report_headers = [
        "Nivel", "Tipo", "Arquivo", "Linha", "Venda origem", "Campo",
        "Valor origem", "Valor comparado", "Detalhe",
    ]
    report_rows = [
        [
            issue.level, issue.category, issue.file, issue.line, issue.sale_id,
            issue.field, issue.origin, issue.compared, issue.detail,
        ]
        for issue in report.issues
    ]
    sheets = [
        ("Mapa Codigos Venda", ["VendaId origem", "Codigo venda importacao"], mapping_rows(sale_codes)),
        ("Rastreabilidade Vendas", SALES_TRACE_HEADERS, sales_trace),
        ("Rastreabilidade Saldos", BALANCE_TRACE_HEADERS, balance_trace),
        ("Relatorio Validacao", report_headers, report_rows),
    ]

    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index in range(1, len(sheets) + 1):
        content_types.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    content_types.append('</Types>')

    workbook_sheets = ''.join(
        f'<sheet name="{xml_escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, (name, _, _) in enumerate(sheets, start=1)
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{workbook_sheets}</sheets></workbook>'
    )
    rels = ''.join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, len(sheets) + 1)
    )
    rels += (
        f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels + '</Relationships>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/>'
        '</styleSheet>'
    )

    try:
        with zipfile.ZipFile(temporary, 'w', compression=zipfile.ZIP_DEFLATED, allowZip64=True) as archive:
            archive.writestr('[Content_Types].xml', ''.join(content_types))
            archive.writestr('_rels/.rels', root_rels)
            archive.writestr('xl/workbook.xml', workbook_xml)
            archive.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
            archive.writestr('xl/styles.xml', styles_xml)
            for index, (_, headers, rows) in enumerate(sheets, start=1):
                archive.writestr(f'xl/worksheets/sheet{index}.xml', _sheet_xml(headers, rows))
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()


# ---------------------------------------------------------------------------
# Modelos de dados do cruzamento
# ---------------------------------------------------------------------------
@dataclass
class SaleAggregate:
    source_id: str
    source_lines: list[int]
    raw_rows: list[dict[str, Any]]
    old_client_code: str = ""
    client_name: str = ""
    client_cpf: str = ""
    client_phone: str = ""
    raw_status: str = ""
    billing_date: str = ""
    sale_value: str = "0"
    paid_value: str = "0"
    observations: list[str] = field(default_factory=list)
    unit_observations: list[str] = field(default_factory=list)
    cancellation_reasons: list[str] = field(default_factory=list)
    units: list[str] = field(default_factory=list)
    categories: list[str] = field(default_factory=list)
    packages: list[str] = field(default_factory=list)
    item_types: list[str] = field(default_factory=list)
    deletion_dates: list[str] = field(default_factory=list)
    imported_sale_code: str = ""
    imported_client_code: str = ""
    client_match_method: str = ""
    mapped_status: str = ""
    used_default_date: bool = False
    sale_date_source: str = ""
    importable: bool = True
    rejection_reason: str = ""
    output_row: list[str] = field(default_factory=list)


@dataclass
class ClientRecord:
    code: str
    lines: list[int] = field(default_factory=list)
    names: set[str] = field(default_factory=set)
    cpfs: set[str] = field(default_factory=set)
    phones: set[str] = field(default_factory=set)


@dataclass
class ClientCatalog:
    records: dict[str, ClientRecord]
    by_cpf: dict[str, set[str]]
    by_phone: dict[str, set[str]]
    by_name: dict[str, set[str]]


@dataclass
class GeneratedClient:
    code: str
    identity_key: str
    name: str
    cpf: str
    phone: str
    old_codes: set[str] = field(default_factory=set)
    sale_ids: list[str] = field(default_factory=list)


@dataclass
class GeneratedClientRegistry:
    used_codes: set[int]
    by_key: dict[str, GeneratedClient] = field(default_factory=dict)

    @classmethod
    def from_catalog(cls, catalog: ClientCatalog) -> "GeneratedClientRegistry":
        used: set[int] = set()
        for code in catalog.records:
            if re.fullmatch(r"\d{1,6}", code):
                numeric = int(code)
                if 100000 <= numeric <= 999999:
                    used.add(numeric)
        return cls(used_codes=used)

    def _next_code(self) -> str:
        code = 100000
        while code in self.used_codes and code <= 999999:
            code += 1
        if code > 999999:
            raise ValueError("Nao ha codigo de cliente disponivel entre 100000 e 999999.")
        self.used_codes.add(code)
        return str(code)

    def get_or_create(self, sale: SaleAggregate, report: ValidationReport, source_file: Path) -> str:
        key = generated_client_identity_key(sale)
        if not key:
            report.error(
                "CLIENTE_COMPLEMENTAR_SEM_IDENTIDADE",
                file=source_file,
                line=",".join(str(line) for line in sale.source_lines[:20]),
                sale_id=sale.source_id,
                detail="Nao ha CPF, codigo de cliente de origem, celular ou nome suficiente para criar o cliente complementar.",
            )
            return ""

        current = self.by_key.get(key)
        if current is None:
            current = GeneratedClient(
                code=self._next_code(),
                identity_key=key,
                name=clean_text(sale.client_name),
                cpf=digits_only(sale.client_cpf),
                phone=clean_text(sale.client_phone),
            )
            self.by_key[key] = current
        else:
            divergences: list[str] = []
            name = clean_text(sale.client_name)
            cpf = digits_only(sale.client_cpf)
            phone = clean_text(sale.client_phone)
            if name and current.name and normalize_name(name) != normalize_name(current.name):
                divergences.append(f"Nome={current.name} / {name}")
            if cpf and current.cpf and cpf != current.cpf:
                divergences.append(f"CPF={current.cpf} / {cpf}")
            if phone and current.phone and phone_keys(phone).isdisjoint(phone_keys(current.phone)):
                divergences.append(f"Celular={current.phone} / {phone}")
            if divergences:
                report.warning(
                    "CLIENTE_COMPLEMENTAR_DADOS_DIVERGENTES",
                    file=source_file,
                    sale_id=sale.source_id,
                    origin=" | ".join(divergences),
                    compared=current.code,
                    detail="A mesma identidade de cliente apareceu em mais de uma venda com dados divergentes; revisar a planilha complementar.",
                )
            if not current.name and name:
                current.name = name
            if not current.cpf and cpf:
                current.cpf = cpf
            if not current.phone and phone:
                current.phone = phone

        old_code = normalize_id(sale.old_client_code)
        if old_code:
            current.old_codes.add(old_code)
        if sale.source_id not in current.sale_ids:
            current.sale_ids.append(sale.source_id)
        return current.code


def generated_client_identity_key(sale: SaleAggregate) -> str:
    cpf = digits_only(sale.client_cpf)
    if cpf:
        return f"CPF:{cpf}"
    old_code = normalize_id(sale.old_client_code)
    if old_code:
        return f"ORIGEM:{old_code}"
    phones = sorted(phone_keys(sale.client_phone), key=lambda value: (-len(value), value))
    name = normalize_name(sale.client_name)
    if phones and name:
        return f"CEL:{phones[0]}|NOME:{name}"
    if phones:
        return f"CEL:{phones[0]}"
    if name:
        return f"NOME:{name}"
    return ""


def _client_phone_output(value: Any) -> str:
    digits = digits_only(value)
    if not digits:
        return ""
    if digits.startswith("55") and len(digits) in {12, 13}:
        digits = digits[2:]
    if len(digits) == 10:
        digits = digits[:2] + "9" + digits[2:]
    if len(digits) == 11:
        return f"({digits[:2]}){digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}){digits[2:6]}-{digits[6:]}"
    return clean_text(value)


def _client_cpf_output(value: Any) -> str:
    digits = digits_only(value)
    if not digits:
        return ""
    if len(digits) in {9, 10}:
        digits = digits.zfill(11)
    if len(digits) != 11:
        return clean_text(value)
    return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"


def _client_model_indexes(headers: Sequence[Any]) -> dict[str, int | None]:
    normalized = [normalize_header(header) for header in headers]

    def first(predicate: Any) -> int | None:
        for index, header in enumerate(normalized):
            if predicate(header, index):
                return index
        return None

    return {
        "code": first(lambda h, i: "codigo" in h.split() and ("cliente" in h.split() or "max" in h.split() or i == 0)),
        "name": first(lambda h, i: "nome" in h.split() and "mae" not in h.split() and "pai" not in h.split()),
        "phone": first(lambda h, i: any(token in h.split() for token in ("fone", "telefone")) and "ddi" not in h.split()),
        "mobile": first(lambda h, i: "celular" in h.split() and "ddi" not in h.split() and "2" not in h.split()),
        "cpf": first(lambda h, i: "cpf" in h.split()),
        "observation": first(lambda h, i: "observacao" in h.split()),
        "status": first(lambda h, i: h == "status"),
        "origin_type": first(lambda h, i: "tipo" in h.split() and "origem" in h.split()),
        "origin_code": first(lambda h, i: "codigo" in h.split() and "origem" in h.split()),
        "ddi_mobile": first(lambda h, i: "ddi" in h.split() and "celular" in h.split() and "2" not in h.split()),
    }


def build_generated_client_rows(
    template: Table,
    registry: GeneratedClientRegistry,
    report: ValidationReport,
) -> tuple[list[list[str]], set[str]]:
    indexes = _client_model_indexes(template.headers)
    if indexes["code"] is None or indexes["name"] is None:
        raise ValueError("modeloImportacaoCliente precisa conter colunas de Codigo e Nome.")

    required = [
        index for index, header in enumerate(template.headers)
        if "obrigatorio" in normalize_header(header)
        and "obrigatorio se" not in normalize_header(header)
        and "obrigatorio no" not in normalize_header(header)
    ]
    import_date = date.today().strftime("%d/%m/%Y")
    rows: list[list[str]] = []
    invalid_codes: set[str] = set()

    for client in sorted(registry.by_key.values(), key=lambda item: int(item.code)):
        row = ["" for _ in template.headers]
        row[indexes["code"]] = client.code
        row[indexes["name"]] = client.name
        if indexes["mobile"] is not None:
            row[indexes["mobile"]] = _client_phone_output(client.phone)
        elif indexes["phone"] is not None:
            row[indexes["phone"]] = _client_phone_output(client.phone)
        if indexes["cpf"] is not None:
            row[indexes["cpf"]] = _client_cpf_output(client.cpf)
        if indexes["observation"] is not None:
            row[indexes["observation"]] = f"Cliente complementar gerado pela Venda de Planos | Importação {import_date}"
        if indexes["status"] is not None:
            row[indexes["status"]] = "Leads"
        if indexes["origin_type"] is not None:
            row[indexes["origin_type"]] = "Parcerias"
        if indexes["origin_code"] is not None and client.old_codes:
            row[indexes["origin_code"]] = sorted(client.old_codes, key=sort_identifier)[0]
        if indexes["ddi_mobile"] is not None and client.phone:
            row[indexes["ddi_mobile"]] = "55"

        row = [clean_text(value) for value in row]
        missing = [clean_text(template.headers[index]) for index in required if not clean_text(row[index])]
        if missing:
            invalid_codes.add(client.code)
            report.error(
                "CLIENTE_COMPLEMENTAR_CAMPO_OBRIGATORIO_AUSENTE",
                sale_id=",".join(client.sale_ids[:20]),
                field=" | ".join(missing),
                origin=client.code,
                detail="A extracao de Vendas nao possui dados suficientes para preencher todos os campos obrigatorios do modelo de Cliente.",
            )
        rows.append(row)
    return rows, invalid_codes


def invalidate_sales_for_generated_clients(
    sales: dict[str, SaleAggregate],
    registry: GeneratedClientRegistry,
    invalid_codes: set[str],
) -> None:
    if not invalid_codes:
        return
    invalid_sales = {sale_id for client in registry.by_key.values() if client.code in invalid_codes for sale_id in client.sale_ids}
    for sale_id in invalid_sales:
        sale = sales.get(sale_id)
        if sale is None:
            continue
        sale.importable = False
        sale.rejection_reason = "CLIENTE_COMPLEMENTAR_INCOMPLETO"
        sale.imported_sale_code = ""
        sale.output_row = []


@dataclass
class ServiceCatalog:
    by_name: dict[str, set[str]]
    display_names: dict[str, list[str]]
    value_by_code: dict[str, str]


# ---------------------------------------------------------------------------
# Mapeamento dos arquivos de origem
# ---------------------------------------------------------------------------
SALES_FIELD_ALIASES: dict[str, list[str]] = {
    "sale_id": ["VendaId", "Codigo Venda"],
    "observation": ["Observacao"],
    "unit_observation": ["ObservacaoUnidade"],
    "cancellation_type_id": ["TipoCancelamentoId"],
    "cancellation_reason": ["MotivoCancelamento"],
    "status": ["StatusVenda"],
    "billing_date": ["VendaDataFaturamento"],
    "client_name": ["ClienteNome"],
    "old_client_code": ["Codigo Cliente", "Codigo do Cliente"],
    "unit": ["Unidade"],
    "client_phone": ["ClienteCelular"],
    "client_cpf": ["ClienteCpf"],
    "package": ["PacoteServico"],
    "item_type": ["Tipo"],
    "category": ["Categoria"],
    "item_status": ["StatusItemComandaServicoId"],
    "qty_total": ["QtdTotal", "Quantidade"],
    "qty_used": ["QtdRealizado", "QuantidadeUtilizada"],
    "qty_remaining": ["QtdFaltante"],
    "item_value": ["ValorItem"],
    "sale_value": ["ValorVenda"],
    "paid_value": ["ValorVendaPago"],
    "last_session": ["UltimaDataSessaoRealizada"],
    "deletion_date": ["DataDelete", "Data Exclusao"],
    "days_without_booking": ["QtdDiasSemAgendamento"],
    "unit_session_value": ["ValorUnitarioSessao"],
}

SESSION_FIELD_ALIASES: dict[str, list[str]] = {
    "sale_id": ["VendaId"],
    "status": ["StatusVenda"],
    "client_name": ["ClienteNome"],
    "client_cpf": ["ClienteCpf"],
    "package_id": ["PacoteServicosId"],
    "old_service_id": ["ServicoId"],
    "item_name": ["NomeItem", "Item"],
    "service_name": ["ServicoNome", "Nome Servico"],
    "qty_total": ["Quantidade", "QtdTotal"],
    "qty_used": ["QuantidadeUtilizada", "QtdRealizado"],
    "qty_remaining": ["QtdFaltante"],
    "deletion_date": ["DataDelete"],
}


def map_columns(headers: Sequence[Any], aliases: dict[str, list[str]], optional: set[str] | None = None) -> dict[str, int | None]:
    optional = optional or set()
    result: dict[str, int | None] = {}
    for field_name, names in aliases.items():
        result[field_name] = find_header(headers, names, required=field_name not in optional)
    return result


def source_row_dict(row: Sequence[Any], columns: dict[str, int | None]) -> dict[str, Any]:
    return {field_name: row_value(row, index) for field_name, index in columns.items()}


def canonical_critical(value: Any, field_name: str, *, date1904: bool = False) -> str:
    if field_name == "client_cpf":
        return digits_only(value)
    if field_name == "client_phone":
        keys = sorted(phone_keys(value))
        return "|".join(keys)
    if field_name == "client_name":
        return normalize_name(value)
    if field_name in {"sale_id", "old_client_code"}:
        return normalize_id(value)
    if field_name == "status":
        return normalize_status(value)
    if field_name == "billing_date":
        return format_date(value, date1904=date1904)
    if field_name in {"sale_value", "paid_value"}:
        # Valores financeiros da venda devem refletir exatamente a extracao.
        # Vazio/null nao pode ser convertido silenciosamente em zero.
        return number_string(value, blank_as_zero=False)
    return clean_text(value)


def aggregate_sales(table: Table, report: ValidationReport) -> tuple[dict[str, SaleAggregate], list[dict[str, Any]]]:
    optional = {
        "paid_value",
        "cancellation_type_id",
        "last_session",
        "days_without_booking",
        "unit_session_value",
        "qty_remaining",
        "deletion_date",
        "old_client_code",
    }
    columns = map_columns(table.headers, SALES_FIELD_ALIASES, optional)
    grouped: dict[str, list[tuple[int, dict[str, Any]]]] = defaultdict(list)
    all_source_rows: list[dict[str, Any]] = []
    for line, raw_row in table.rows:
        row = ensure_width(raw_row, len(table.headers))
        values = source_row_dict(row, columns)
        values["_line"] = line
        all_source_rows.append(values)
        sale_id = normalize_id(values["sale_id"])
        if not sale_id:
            report.error(
                "VENDA_SEM_CODIGO",
                file=table.path,
                line=line,
                field="VendaId",
                detail="Linha preenchida sem identificador da venda.",
            )
            continue
        grouped[sale_id].append((line, values))

    aggregates: dict[str, SaleAggregate] = {}
    critical_fields = [
        "old_client_code",
        "client_name",
        "client_cpf",
        "client_phone",
        "status",
        "billing_date",
        "sale_value",
        "paid_value",
    ]
    for sale_id, entries in grouped.items():
        raw_rows = [values for _, values in entries]
        aggregate = SaleAggregate(source_id=sale_id, source_lines=[line for line, _ in entries], raw_rows=raw_rows)
        critical: dict[str, str] = {}
        for field_name in critical_fields:
            values = [
                canonical_critical(row[field_name], field_name, date1904=table.date1904)
                for row in raw_rows
            ]
            distinct = [value for value in dict.fromkeys(values) if value]
            if len(distinct) > 1:
                report.error(
                    "CONFLITO_DENTRO_DA_VENDA",
                    file=table.path,
                    line=",".join(str(line) for line in aggregate.source_lines[:20]),
                    sale_id=sale_id,
                    field=field_name,
                    origin=" | ".join(distinct[:10]),
                    detail="A mesma venda possui valores diferentes em um campo critico.",
                )
            critical[field_name] = distinct[0] if distinct else ""

        aggregate.old_client_code = critical["old_client_code"]
        aggregate.client_name = clean_text(next((row["client_name"] for row in raw_rows if not is_blank(row["client_name"])), ""))
        aggregate.client_cpf = critical["client_cpf"]
        aggregate.client_phone = clean_text(next((row["client_phone"] for row in raw_rows if not is_blank(row["client_phone"])), ""))
        aggregate.raw_status = clean_text(next((row["status"] for row in raw_rows if not is_blank(row["status"])), ""))
        aggregate.billing_date = critical["billing_date"]
        aggregate.sale_value = critical["sale_value"]
        aggregate.paid_value = critical["paid_value"]
        aggregate.observations = distinct_clean(row["observation"] for row in raw_rows)
        aggregate.unit_observations = distinct_clean(row["unit_observation"] for row in raw_rows)
        aggregate.cancellation_reasons = distinct_clean(row["cancellation_reason"] for row in raw_rows)
        aggregate.units = distinct_clean(row["unit"] for row in raw_rows)
        aggregate.categories = distinct_clean(row["category"] for row in raw_rows)
        aggregate.packages = distinct_clean(row["package"] for row in raw_rows)
        aggregate.item_types = distinct_clean(row["item_type"] for row in raw_rows)
        aggregate.deletion_dates = distinct_clean(
            (row["deletion_date"] for row in raw_rows),
            formatter=lambda value: format_datetime(value, date1904=table.date1904),
        )

        for label, values in (
            ("Observacao", aggregate.observations),
            ("ObservacaoUnidade", aggregate.unit_observations),
            ("MotivoCancelamento", aggregate.cancellation_reasons),
            ("Unidade", aggregate.units),
            ("Categoria", aggregate.categories),
            ("Tipo", aggregate.item_types),
            ("DataDelete", aggregate.deletion_dates),
        ):
            if len(values) > 1:
                report.info(
                    "MULTIPLOS_VALORES_CONSOLIDADOS",
                    file=table.path,
                    line=",".join(str(line) for line in aggregate.source_lines[:20]),
                    sale_id=sale_id,
                    field=label,
                    origin=" | ".join(values[:10]),
                    detail="Todos os valores distintos foram preservados na rastreabilidade/observacao.",
                )
        aggregates[sale_id] = aggregate

    report.info(
        "RESUMO_VENDAS_ORIGEM",
        file=table.path,
        origin=len(table.rows),
        compared=len(aggregates),
        detail="Linhas preenchidas da extracao e vendas unicas encontradas.",
    )
    return aggregates, all_source_rows


def build_client_catalog(table: Table, report: ValidationReport) -> ClientCatalog:
    headers_norm = [normalize_header(header) for header in table.headers]
    code_candidates = [
        index
        for index, header in enumerate(headers_norm)
        if "codigo" in header.split() and ("cliente" in header.split() or index == 0)
    ]
    if not code_candidates:
        code_candidates = [index for index, header in enumerate(headers_norm) if "codigo" in header.split()]
    if not code_candidates:
        raise KeyError(f"Nao encontrei a coluna de codigo do cliente em {table.path}")
    code_index = code_candidates[0]

    name_candidates = [
        index
        for index, header in enumerate(headers_norm)
        if "nome" in header.split() and "mae" not in header.split() and "pai" not in header.split()
    ]
    if not name_candidates:
        raise KeyError(f"Nao encontrei a coluna de nome do cliente em {table.path}")
    name_index = name_candidates[0]
    cpf_candidates = [index for index, header in enumerate(headers_norm) if "cpf" in header.split()]
    if not cpf_candidates:
        raise KeyError(f"Nao encontrei a coluna de CPF em {table.path}")
    cpf_index = cpf_candidates[0]
    phone_indexes = [
        index
        for index, header in enumerate(headers_norm)
        if any(token in header.split() for token in ("celular", "fone", "telefone"))
        and "ddi" not in header.split()
    ]

    records: dict[str, ClientRecord] = {}
    for line, raw_row in table.rows:
        row = ensure_width(raw_row, len(table.headers))
        code = normalize_id(row_value(row, code_index))
        if not code:
            report.error(
                "CLIENTE_SEM_CODIGO",
                file=table.path,
                line=line,
                field=clean_text(table.headers[code_index]),
                detail="Linha da planilha tratada de clientes sem codigo.",
            )
            continue
        record = records.setdefault(code, ClientRecord(code=code))
        record.lines.append(line)
        name = normalize_name(row_value(row, name_index))
        cpf = digits_only(row_value(row, cpf_index))
        if name:
            record.names.add(name)
        if cpf:
            record.cpfs.add(cpf)
        for index in phone_indexes:
            record.phones.update(phone_keys(row_value(row, index)))

    by_cpf: dict[str, set[str]] = defaultdict(set)
    by_phone: dict[str, set[str]] = defaultdict(set)
    by_name: dict[str, set[str]] = defaultdict(set)
    for code, record in records.items():
        for value in record.cpfs:
            by_cpf[value].add(code)
        for value in record.phones:
            by_phone[value].add(code)
        for value in record.names:
            by_name[value].add(code)

    report.info(
        "RESUMO_CLIENTES_DE_PARA",
        file=table.path,
        origin=len(table.rows),
        compared=len(records),
        detail="Linhas lidas e codigos de cliente distintos no arquivo tratado.",
    )
    return ClientCatalog(records=records, by_cpf=dict(by_cpf), by_phone=dict(by_phone), by_name=dict(by_name))


def _candidate_codes(index: dict[str, set[str]], keys: Iterable[str]) -> set[str]:
    result: set[str] = set()
    for key in keys:
        result.update(index.get(key, set()))
    return result


def match_client(
    sale: SaleAggregate,
    catalog: ClientCatalog,
    report: ValidationReport,
    source_file: Path,
    historical_code: Any = "",
) -> tuple[str, str]:
    """Localiza o codigo do cliente sem depender do modelo de Venda de Plano.

    Ordem de seguranca:
    1. O Codigo Cliente da extracao pode resolver diretamente somente quando
       ele existe na planilhaTratadaCliente e e confirmado por CPF/celular/nome.
       Isso resolve CPFs duplicados sem escolher um registro por acaso.
    2. Se o codigo de origem nao resolver, usa intersecao/consenso das chaves
       CPF, celular e nome na planilha tratada.
    3. Clientes ausentes da planilha tratada recebem um novo codigo 100000+,
       sem colisao com os codigos existentes, e sao enviados para uma planilha
       complementar baseada no modeloImportacaoCliente. O mesmo novo codigo e
       reutilizado na Venda de Planos e no Saldo.
    """
    cpf_keys = {sale.client_cpf} if sale.client_cpf else set()
    phone_key_set = phone_keys(sale.client_phone)
    normalized_client_name = normalize_name(sale.client_name)
    name_keys = {normalized_client_name} if normalized_client_name else set()

    candidates = {
        "CPF": _candidate_codes(catalog.by_cpf, cpf_keys),
        "CELULAR": _candidate_codes(catalog.by_phone, phone_key_set),
        "NOME": _candidate_codes(catalog.by_name, name_keys),
    }
    old_code = normalize_id(sale.old_client_code)
    old_exists = bool(old_code and old_code in catalog.records)


    # O mesmo CPF valido pode aparecer em dois cadastros tratados quando houve
    # recadastro/migracao do mesmo cliente. Isso nao e uma ambiguidade de pessoa:
    # e duplicidade de cadastro. Para que todas as vendas e saldos da identidade
    # apontem sempre para o mesmo codigo, usa um codigo canonico deterministico.
    if len(candidates["CPF"]) > 1:
        selected = sorted(candidates["CPF"], key=sort_identifier)[-1]
        report.warning(
            "CLIENTE_CPF_DUPLICADO_RESOLVIDO",
            file=source_file,
            line=",".join(str(line) for line in sale.source_lines[:20]),
            sale_id=sale.source_id,
            field="CPF",
            origin=sale.client_cpf,
            compared=f"{selected} (duplicados: {','.join(sorted(candidates['CPF'], key=sort_identifier))})",
            detail="CPF valido identifica a mesma pessoa em mais de um cadastro tratado; utilizado um unico codigo canonico para evitar divergencia entre vendas e saldos.",
        )
        return selected, "CPF_DUPLICADO_CANONICO"

    # O codigo presente na propria extracao e a melhor chave de desempate quando
    # ele tambem existe na planilha tratada. Exigimos confirmacao por ao menos
    # uma chave de identidade sempre que essas chaves estiverem disponiveis.
    if old_exists:
        supporting = [source for source, values in candidates.items() if old_code in values]
        any_identity_candidate = any(candidates.values())
        if supporting or not any_identity_candidate:
            conflicting = [
                (source, values) for source, values in candidates.items()
                if values and old_code not in values
            ]
            if conflicting:
                report.warning(
                    "CLIENTE_CODIGO_ORIGEM_COM_CHAVE_SECUNDARIA_DIVERGENTE",
                    file=source_file,
                    line=",".join(str(line) for line in sale.source_lines[:20]),
                    sale_id=sale.source_id,
                    field="CPF/Celular/Nome",
                    origin=old_code,
                    compared=" | ".join(
                        f"{source}={','.join(sorted(values, key=sort_identifier))}"
                        for source, values in conflicting
                    ),
                    detail="Codigo Cliente existe na planilha tratada e foi confirmado por outra chave; divergencias secundarias ficaram registradas para revisao.",
                )
            method = "CODIGO_ORIGEM"
            if supporting:
                method += "+" + "+".join(supporting)
            else:
                report.warning(
                    "CLIENTE_RESOLVIDO_APENAS_POR_CODIGO_ORIGEM",
                    file=source_file,
                    line=",".join(str(line) for line in sale.source_lines[:20]),
                    sale_id=sale.source_id,
                    field="Codigo Cliente",
                    origin=old_code,
                    detail="O codigo existe em planilhaTratadaCliente, mas a extracao nao trouxe uma chave de identidade utilizavel para confirmar.",
                )
            return old_code, method

    nonempty = [(source, values) for source, values in candidates.items() if values]
    union: set[str] = set().union(*(values for _, values in nonempty)) if nonempty else set()
    common: set[str] = set(nonempty[0][1]) if nonempty else set()
    for _, values in nonempty[1:]:
        common.intersection_update(values)

    # CPF unico tem prioridade. Para CPF duplicado, codigo de origem ja teria
    # resolvido acima se fosse um dos registros existentes na planilha tratada.
    if len(candidates["CPF"]) == 1:
        selected = next(iter(candidates["CPF"]))
        conflicts = [
            (source, values) for source, values in candidates.items()
            if source != "CPF" and values and selected not in values
        ]
        if conflicts:
            report.warning(
                "CLIENTE_CPF_UNICO_COM_CHAVE_SECUNDARIA_DIVERGENTE",
                file=source_file,
                line=",".join(str(line) for line in sale.source_lines[:20]),
                sale_id=sale.source_id,
                field="CPF/Celular/Nome",
                origin=selected,
                compared=" | ".join(
                    f"{source}={','.join(sorted(values, key=sort_identifier))}"
                    for source, values in conflicts
                ),
                detail="CPF identificou um unico cliente; as chaves secundarias divergentes foram mantidas no relatorio.",
            )
        return selected, "CPF"

    if len(common) == 1:
        selected = next(iter(common))
        return selected, "+".join(source for source, values in nonempty if selected in values)

    if len(union) == 1:
        selected = next(iter(union))
        return selected, "+".join(source for source, values in nonempty if selected in values)

    # A planilhaTratadaCliente e a fonte oficial do codigo de cliente para esta
    # importacao. O codigo antigo da extracao pode apenas desempatar/confirmar um
    # registro que exista nela; nunca e usado como substituto quando o cliente
    # nao foi localizado. Nesse caso a venda permanece registrada na validacao
    # e os CSVs finais sao bloqueados para evitar conflito de relacionamento.
    if not union and old_code:
        report.info(
            "CLIENTE_NAO_LOCALIZADO_NA_PLANILHA_TRATADA",
            file=source_file,
            line=",".join(str(line) for line in sale.source_lines[:20]),
            sale_id=sale.source_id,
            field="Codigo Cliente",
            origin=f"CPF={sale.client_cpf}; Celular={sale.client_phone}; Nome={sale.client_name}",
            compared=old_code,
            detail="O Codigo Cliente da extracao nao foi reutilizado porque nao existe correspondencia segura em planilhaTratadaCliente.",
        )
        return "", "NAO_LOCALIZADO"

    if union:
        report.error(
            "CLIENTE_AMBIGUO",
            file=source_file,
            line=",".join(str(line) for line in sale.source_lines[:20]),
            sale_id=sale.source_id,
            field="CPF/Celular/Nome",
            origin=" | ".join(
                f"{source}={','.join(sorted(values, key=sort_identifier))}"
                for source, values in candidates.items() if values
            ),
            compared=old_code,
            detail="As chaves apontam para mais de um cliente e o Codigo Cliente da extracao nao resolve a ambiguidade na planilha tratada.",
        )
        return "", "AMBIGUO"
    else:
        report.info(
            "CLIENTE_NAO_LOCALIZADO",
            file=source_file,
            line=",".join(str(line) for line in sale.source_lines[:20]),
            sale_id=sale.source_id,
            field="CPF/Celular/Nome",
            origin=f"CPF={sale.client_cpf}; Celular={sale.client_phone}; Nome={sale.client_name}",
            detail="Nenhuma chave e nenhum Codigo Cliente utilizavel foram encontrados.",
        )
    return "", "NAO_LOCALIZADO"

def build_service_catalog(table: Table, report: ValidationReport) -> ServiceCatalog:
    """Carrega codigo/descricao da aba de Servicos do DE-PARA.

    O arquivo do cliente pode usar cabecalhos compactos, como ``codServicos``.
    Quando a aba nao possuir coluna de valor, aplica 0, conforme a regra do
    projeto para Saldo de Venda de Planos.
    """
    normalized = [normalize_header(header) for header in table.headers]
    compact = [header.replace(" ", "") for header in normalized]

    code_candidates = [
        index for index, header in enumerate(compact)
        if (
            header in {
                "codservicos", "codservico", "codserv",
                "codprocedimentos", "codprocedimento", "codproced",
                "codigoservicos", "codigoservico",
                "codigoprocedimentos", "codigoprocedimento",
            }
            or (header.startswith("cod") and ("servic" in header or "proced" in header))
            or (header.startswith("codigo") and ("servic" in header or "proced" in header))
        )
    ]
    if not code_candidates:
        code_candidates = [
            index for index, header in enumerate(normalized)
            if "codigo" in header.split() and any(token in header.split() for token in ("servico", "servicos", "procedimento", "procedimentos"))
        ]
    if not code_candidates:
        raise KeyError(
            f"Nao encontrei a coluna de codigo do servico na aba {table.sheet!r} de {table.path.name}. "
            f"Cabecalhos encontrados: {', '.join(clean_text(h) or '<vazio>' for h in table.headers)}"
        )
    code_index = code_candidates[0]

    value_candidates = [
        index for index, header in enumerate(normalized)
        if index != code_index
        and ("valor" in header.split() or "preco" in header.split())
        and "total" not in header.split()
        and "desconto" not in header.split()
    ]
    value_index = value_candidates[0] if value_candidates else None
    if value_index is None:
        report.info(
            "DE_PARA_SERVICOS_SEM_COLUNA_VALOR",
            file=table.path,
            field="Valor Servico",
            origin=table.sheet,
            compared="0",
            detail="A aba de Servicos nao possui coluna de valor; aplicado valor 0 conforme regra de Saldo de Venda de Planos.",
        )

    excluded_indexes = {code_index}
    if value_index is not None:
        excluded_indexes.add(value_index)
    name_candidates = [
        index for index, header in enumerate(normalized)
        if index not in excluded_indexes
        and any(token in header.split() for token in ("procedimento", "procedimentos", "servico", "servicos", "nome", "item", "descricao"))
        and "codigo" not in header.split()
        and "valor" not in header.split()
        and "preco" not in header.split()
        and "id" not in header.split()
    ]
    if not name_candidates:
        name_candidates = [
            index for index in range(len(table.headers))
            if index not in excluded_indexes and compact[index] not in {"tempo", "duracao", "minutos"}
        ]
    if not name_candidates:
        raise KeyError(f"Nao encontrei coluna de descricao do servico na aba {table.sheet!r} de {table.path.name}")

    by_name: dict[str, set[str]] = defaultdict(set)
    display_names: dict[str, list[str]] = defaultdict(list)
    values_by_code: dict[str, set[str]] = defaultdict(set)

    for line, raw_row in table.rows:
        row = ensure_width(raw_row, len(table.headers))
        code = normalize_id(row_value(row, code_index))
        value = number_string(row_value(row, value_index)) if value_index is not None else "0"
        if not value:
            value = "0"
        names = distinct_clean(row_value(row, index) for index in name_candidates)

        if not code or not names:
            report.error(
                "DE_PARA_SERVICO_INCOMPLETO",
                file=table.path,
                line=line,
                origin=" | ".join(names),
                compared=code,
                detail="Descricao e codigo do servico sao obrigatorios na aba de Servicos do DE-PARA.",
            )
            continue

        values_by_code[code].add(value)
        for name in names:
            for key in service_lookup_keys(name):
                by_name[key].add(code)
                if name not in display_names[key]:
                    display_names[key].append(name)

    for key, codes in by_name.items():
        if len(codes) > 1:
            report.error(
                "SERVICO_AMBIGUO_NO_DE_PARA",
                file=table.path,
                field=key,
                origin=" | ".join(display_names[key]),
                compared=", ".join(sorted(codes, key=sort_identifier)),
                detail="O mesmo procedimento normalizado possui codigos diferentes na aba de Servicos do DE-PARA.",
            )

    value_by_code: dict[str, str] = {}
    for code, values in values_by_code.items():
        if len(values) > 1:
            report.error(
                "VALOR_SERVICO_AMBIGUO_NO_DE_PARA",
                file=table.path,
                field="Valor Servico",
                origin=code,
                compared=" | ".join(sorted(values)),
                detail="O mesmo codigo de servico possui mais de um valor na aba de Servicos do DE-PARA.",
            )
        elif values:
            value_by_code[code] = next(iter(values))

    report.info(
        "RESUMO_SERVICOS_DE_PARA",
        file=table.path,
        origin=len(table.rows),
        compared=len(by_name),
        detail="Codigo e descricao lidos da aba de Servicos do DE-PARA; valor ausente no arquivo e tratado como 0.",
    )
    return ServiceCatalog(by_name=dict(by_name), display_names=dict(display_names), value_by_code=value_by_code)


def build_existing_sale_reference(table: Table, report: ValidationReport) -> dict[str, dict[str, str]]:
    """Le o conteudo ja preenchido do modelo sem reutilizar codigos tratados.

    O arquivo enviado possui datas manuais para algumas vendas faturadas cuja
    extracao principal traz VendaDataFaturamento nulo. Essas datas sao
    preservadas como uma fonte auxiliar, sempre com validacao contra a origem.
    """
    if not table.rows:
        report.info(
            "MODELO_VENDAS_SEM_DADOS",
            file=table.path,
            detail="O modelo contem apenas cabecalho; nao ha referencia historica para cruzar.",
        )
        return {}

    groups: dict[str, list[tuple[int, list[Any]]]] = defaultdict(list)
    for line, raw in table.rows:
        row = ensure_width(raw, 12)
        source_id = normalize_id(row[0])
        if not source_id:
            report.error(
                "MODELO_VENDA_SEM_CODIGO",
                file=table.path,
                line=line,
                detail="Linha preenchida do modelo sem codigo da venda.",
            )
            continue
        groups[source_id].append((line, row))

    fields: list[tuple[str, int, Any]] = [
        ("old_client_code", 1, normalize_id),
        ("sale_date", 2, format_date),
        ("seller", 3, clean_text),
        ("validity", 4, number_string),
        ("sale_value", 5, lambda value: number_string(value, blank_as_zero=True)),
        ("discount", 6, clean_text),
        ("discount_type", 7, clean_text),
        ("paid_value", 8, lambda value: number_string(value, blank_as_zero=True)),
        ("status", 9, clean_text),
        ("suspended_date", 10, format_date),
    ]
    references: dict[str, dict[str, str]] = {}
    for source_id, entries in groups.items():
        reference: dict[str, str] = {}
        for field_name, column, formatter in fields:
            values = []
            for _, row in entries:
                value = formatter(row[column])
                if value not in values:
                    values.append(value)
            nonblank = [value for value in values if value]
            if len(nonblank) > 1:
                report.error(
                    "CONFLITO_NO_MODELO_VENDA_EXISTENTE",
                    file=table.path,
                    line=",".join(str(line) for line, _ in entries[:20]),
                    sale_id=source_id,
                    field=field_name,
                    origin=" | ".join(nonblank[:10]),
                    detail="O modelo existente possui valores criticos diferentes para a mesma venda.",
                )
            reference[field_name] = nonblank[0] if nonblank else ""
        references[source_id] = reference
        if len(entries) > 1:
            report.warning(
                "VENDA_DUPLICADA_NO_MODELO_EXISTENTE",
                file=table.path,
                line=",".join(str(line) for line, _ in entries[:20]),
                sale_id=source_id,
                origin=len(entries),
                detail="A referencia existente repete a venda. O novo arquivo gerara somente uma linha.",
            )

    report.info(
        "RESUMO_MODELO_VENDAS_EXISTENTE",
        file=table.path,
        origin=len(table.rows),
        compared=len(references),
        detail="Linhas preenchidas e VendaId distintos usados para cruzamento historico.",
    )
    return references


def validate_sale_reference_sets(
    sales: dict[str, SaleAggregate],
    references: dict[str, dict[str, str]],
    report: ValidationReport,
    reference_file: Path,
) -> None:
    if not references:
        return
    source_ids = set(sales)
    reference_ids = set(references)
    for source_id in sorted(source_ids - reference_ids, key=sort_identifier):
        report.warning(
            "VENDA_SEM_REFERENCIA_NO_MODELO_EXISTENTE",
            file=reference_file,
            sale_id=source_id,
            detail="A venda existe na extracao, mas nao no conteudo previamente preenchido do modelo.",
        )
    for source_id in sorted(reference_ids - source_ids, key=sort_identifier):
        report.error(
            "MODELO_EXISTENTE_COM_VENDA_FORA_DA_ORIGEM",
            file=reference_file,
            sale_id=source_id,
            detail="O modelo existente possui venda que nao existe na extracao principal atual.",
        )


def build_existing_balance_reference(table: Table, report: ValidationReport) -> list[tuple[int, list[Any]]]:
    if not table.rows:
        report.info(
            "MODELO_SALDOS_SEM_DADOS",
            file=table.path,
            detail="O modelo contem apenas cabecalho; nao ha referencia posicional para cruzar.",
        )
        return []
    rows: list[tuple[int, list[Any]]] = []
    for line, raw in table.rows:
        row = ensure_width(raw, 7)
        if not normalize_id(row[0]):
            report.error("MODELO_SALDO_SEM_VENDA", file=table.path, line=line)
        if not normalize_id(row[1]):
            report.error("MODELO_SALDO_SEM_SERVICO", file=table.path, line=line, sale_id=row[0])
        rows.append((line, row))
    report.info(
        "RESUMO_MODELO_SALDOS_EXISTENTE",
        file=table.path,
        origin=len(rows),
        detail="Linhas preenchidas usadas para confirmar VendaId, servico e quantidade.",
    )
    return rows

def assign_sale_codes(
    sales: dict[str, SaleAggregate],
    mode: str,
    report: ValidationReport,
) -> dict[str, str]:
    """Aplica a regra Laser Rosa para codigo da venda de plano.

    automatico: se TODOS os VendaId couberem em ate 6 digitos, reutiliza a origem;
    se ao menos um ultrapassar 6 digitos (ou nao for numerico), gera sequencia 100000+.
    """
    ordered = sorted(sales, key=sort_identifier)
    effective_mode = mode
    if mode == "automatico":
        effective_mode = "origem" if all(re.fullmatch(r"\d{1,6}", source_id) for source_id in ordered) else "sequencial"

    result: dict[str, str] = {}
    if effective_mode == "sequencial":
        last_code = 100000 + len(ordered) - 1
        if last_code > 999999:
            report.error(
                "LIMITE_CODIGO_VENDA",
                origin=len(ordered),
                compared=last_code,
                detail="A sequencia ultrapassaria os 6 digitos permitidos pelo modelo.",
            )
            return result
        for offset, source_id in enumerate(ordered):
            result[source_id] = str(100000 + offset)
    else:
        for source_id in ordered:
            if not re.fullmatch(r"\d{1,6}", source_id):
                report.error(
                    "CODIGO_VENDA_ORIGEM_INVALIDO",
                    sale_id=source_id,
                    origin=source_id,
                    detail="O codigo de origem precisa ser numerico e possuir no maximo 6 digitos.",
                )
                continue
            result[source_id] = source_id

    report.info(
        "RESUMO_CODIGOS_VENDA",
        origin=f"solicitado={mode}; aplicado={effective_mode}",
        compared=len(result),
        detail="De-para de codigos de venda preparado conforme regra de 6 digitos.",
    )
    return result


STATUS_MAP = {
    "CANCELADO": "Suspenso",
    "FATURADA": "Aprovado",
    "ABERTA": "Pendente",
    "PENDENTE A PAGAMENTO": "Pendente",
    "SUSPENSO": "Suspenso",
    "APROVADO": "Aprovado",
    "PENDENTE": "Pendente",
}


def build_sale_observation(sale: SaleAggregate, import_date: str) -> str:
    """Inclui somente informacoes relevantes sem destino em outra importacao.

    Dados de identificacao do cliente (nome, CPF, celular e codigo), data/status
    da venda, ValorVenda e saldo de sessoes nao sao repetidos aqui porque ja sao
    importados em Cliente, Venda de Planos ou Saldo de Venda de Planos.
    O modelo/planilha tratada nunca e usado como fonte.
    """

    def collected(field_name: str, formatter: Any = clean_text) -> list[str]:
        return distinct_clean(
            (raw.get(field_name, "") for raw in sale.raw_rows),
            formatter=formatter,
        )

    parts: list[str] = []

    # O codigo original so e relevante quando o codigo da venda foi alterado
    # para a sequencia de importacao e, portanto, nao existe outra coluna para ele.
    if sale.imported_sale_code and sale.imported_sale_code != sale.source_id:
        parts.append(f"Venda origem: {sale.source_id}")

    # Informacoes gerais que nao possuem coluna de destino nas importacoes.
    fields: list[tuple[str, str, Any]] = [
        ("Observação origem", "observation", clean_text),
        ("Observação unidade", "unit_observation", clean_text),
        ("Tipo cancelamento ID", "cancellation_type_id", normalize_id),
        ("Motivo cancelamento", "cancellation_reason", clean_text),
        ("Unidade", "unit", clean_text),
        ("Valor venda pago origem", "paid_value", number_string),
        ("Última sessão realizada", "last_session", lambda value: format_datetime(value)),
        ("Data exclusão", "deletion_date", lambda value: format_datetime(value)),
        ("Qtd dias sem agendamento", "days_without_booking", number_string),
    ]

    for label, field_name, formatter in fields:
        values = collected(field_name, formatter)
        if values:
            parts.append(f"{label}: {' / '.join(values)}")

    # Os detalhes dos itens sao preservados por linha para nao perder a relacao
    # entre pacote, quantidade e valores. O saldo importado em outra planilha
    # (QtdFaltante) nao e repetido aqui.
    item_details: list[str] = []
    seen_items: set[str] = set()
    for raw in sale.raw_rows:
        detail_parts: list[str] = []
        item_fields: list[tuple[str, str, Any]] = [
            ("Pacote", "package", clean_text),
            ("Tipo", "item_type", clean_text),
            ("Categoria", "category", clean_text),
            ("Status item ID", "item_status", normalize_id),
            ("Qtd total", "qty_total", number_string),
            ("Qtd realizada", "qty_used", number_string),
            ("Valor item origem", "item_value", number_string),
            ("Valor unitário sessão origem", "unit_session_value", number_string),
        ]
        for label, field_name, formatter in item_fields:
            value = formatter(raw.get(field_name, ""))
            if value:
                detail_parts.append(f"{label}: {value}")
        detail = "; ".join(detail_parts)
        if detail and detail not in seen_items:
            seen_items.add(detail)
            item_details.append(detail)
    if item_details:
        parts.append("Itens origem: " + " / ".join(item_details))

    # A anotacao da importacao deve ser sempre o ultimo trecho da observacao.
    parts.append(f"Importação {import_date}")

    return clean_text(" | ".join(parts))


def sale_price_from_extraction(value: Any) -> str | None:
    """Converte ValorVenda em centavos para valor monetario com virgula decimal.

    Exemplos: 280788 -> 2807,88; 71700 -> 717; vazio/null -> 0.
    Retorna None somente quando ha conteudo nao numerico.
    """
    if is_blank(value):
        return "0"
    numeric = numeric_value(value)
    if numeric is None:
        return None
    try:
        result = Decimal(str(numeric)) / Decimal("100")
    except (InvalidOperation, ValueError):
        return None
    if result == 0:
        return "0"
    # Sem separador de milhar; usa virgula como separador decimal do layout.
    text = format(result, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text.replace(".", ",")


def build_sales_output(
    sales: dict[str, SaleAggregate],
    client_catalog: ClientCatalog,
    sale_codes: dict[str, str],
    existing_references: dict[str, dict[str, str]],
    default_date: str,
    source_file: Path,
    reference_file: Path,
    report: ValidationReport,
    generated_clients: GeneratedClientRegistry | None = None,
) -> list[list[str]]:
    import_date = date.today().strftime("%d/%m/%Y")
    old_to_new: dict[str, set[str]] = defaultdict(set)
    rows: list[list[str]] = []
    for source_id in sorted(sales, key=sort_identifier):
        sale = sales[source_id]
        sale.imported_sale_code = sale_codes.get(source_id, "")
        reference = existing_references.get(source_id, {})
        client_code, method = match_client(
            sale,
            client_catalog,
            report,
            source_file,
            historical_code=reference.get("old_client_code", ""),
        )
        if not client_code and method == "NAO_LOCALIZADO" and generated_clients is not None:
            client_code = generated_clients.get_or_create(sale, report, source_file)
            if client_code:
                method = "CLIENTE_COMPLEMENTAR_GERADO"
                report.info(
                    "CLIENTE_COMPLEMENTAR_GERADO",
                    file=source_file,
                    sale_id=source_id,
                    origin=f"CPF={sale.client_cpf}; Celular={sale.client_phone}; Nome={sale.client_name}",
                    compared=client_code,
                    detail="Cliente ausente da planilhaTratadaCliente; novo codigo criado e reutilizado nesta venda. Importar a planilha complementar de clientes antes das vendas.",
                )
        sale.imported_client_code = client_code
        sale.client_match_method = method
        if not client_code:
            sale.importable = False
            sale.rejection_reason = "CLIENTE_NAO_LOCALIZADO_NA_PLANILHA_TRATADA"
            sale.client_match_method = "REJEITADO_CLIENTE"
            sale.imported_sale_code = ""
            continue
        if sale.old_client_code and client_code:
            old_to_new[sale.old_client_code].add(client_code)

        mapped_status = STATUS_MAP.get(normalize_status(sale.raw_status), "")
        if not mapped_status:
            report.error(
                "STATUS_VENDA_DESCONHECIDO",
                file=source_file,
                line=",".join(str(line) for line in sale.source_lines[:20]),
                sale_id=source_id,
                origin=sale.raw_status,
                detail="Nao foi possivel converter o status para Pendente, Aprovado ou Suspenso.",
            )
        sale.mapped_status = mapped_status

        # A data da venda vem exclusivamente da extracao. Quando estiver
        # ausente, aplica o primeiro dia do mes anterior ao da execucao.
        sale_date = sale.billing_date
        if sale_date:
            sale.sale_date_source = "VendaDataFaturamento"
        else:
            sale_date = default_date
            sale.used_default_date = True
            sale.sale_date_source = "primeiro dia do mes anterior a execucao"
            report.info(
                "DATA_FATURAMENTO_AUSENTE",
                file=source_file,
                line=",".join(str(line) for line in sale.source_lines[:20]),
                sale_id=source_id,
                field="VendaDataFaturamento",
                origin="vazio/null",
                compared=default_date,
                detail="Aplicado o primeiro dia do mes anterior ao da execucao do script.",
            )

        suspended_date = sale_date if mapped_status == "Suspenso" else ""

        if reference:
            comparisons = [
                ("Codigo Cliente origem", sale.old_client_code, reference.get("old_client_code", "")),
                ("Vendedor", "Administrador", reference.get("seller", "")),
                ("Validade", "60", reference.get("validity", "")),
                ("Preco", sale_price_from_extraction(sale.sale_value) or "", reference.get("sale_value", "")),
                ("Desconto", "", reference.get("discount", "")),
                ("Tipo desconto", "%", reference.get("discount_type", "")),
                ("Preco Final", sale_price_from_extraction(sale.sale_value) or "", reference.get("sale_value", "")),
                ("Status", mapped_status, reference.get("status", "")),
            ]
            comparisons.append(("Data Suspenso", suspended_date, reference.get("suspended_date", "")))
            for field_name, expected, historical in comparisons:
                if clean_text(expected) != clean_text(historical):
                    report.error(
                        "DIVERGENCIA_MODELO_VENDA_EXISTENTE",
                        file=reference_file,
                        sale_id=source_id,
                        field=field_name,
                        origin=expected,
                        compared=historical,
                        detail="A regra calculada pela origem diverge do conteudo previamente preenchido.",
                    )
        # ValorVenda e a unica fonte de Preco e Preco Final. A extracao
        # armazena esse valor em centavos; por isso divide por 100. Vazio/null
        # deve virar zero. ValorVendaPago nao substitui o valor do plano.
        sale_price = sale_price_from_extraction(sale.sale_value)
        if sale_price is None:
            report.error(
                "VALOR_VENDA_INVALIDO",
                file=source_file,
                line=",".join(str(line) for line in sale.source_lines[:20]),
                sale_id=source_id,
                field="ValorVenda",
                origin=sale.sale_value,
                detail="ValorVenda precisa ser numerico para ser dividido por 100 e preencher Preco/Preco Final.",
            )
            sale_price = "0"

        observation = build_sale_observation(sale, import_date)
        row = [
            sale.imported_sale_code,
            client_code,
            sale_date,
            "Administrador",
            "60",
            sale_price,
            "",
            "%",
            sale_price,
            mapped_status,
            suspended_date,
            observation,
        ]
        sale.output_row = [clean_text(value) for value in row]
        rows.append(sale.output_row)

    for old_code, new_codes in old_to_new.items():
        if len(new_codes) > 1:
            report.error(
                "CODIGO_CLIENTE_ORIGEM_DIVERGENTE",
                field="Codigo Cliente origem",
                origin=old_code,
                compared=", ".join(sorted(new_codes, key=sort_identifier)),
                detail="O mesmo codigo antigo foi relacionado a mais de um cliente novo.",
            )
    report.info(
        "RESUMO_VENDAS_TRATADAS",
        file=source_file,
        origin=len(sales),
        compared=len(rows),
        detail="Vendas unicas e linhas preparadas para o modelo.",
    )
    return rows


def _service_value_for_code(
    code: str,
    catalog: ServiceCatalog,
    *,
    line: int,
    sale_id: str,
    source_file: Path,
    report: ValidationReport,
) -> str:
    value = catalog.value_by_code.get(code, "0")
    return value if value else "0"


def match_service(
    service_name: Any,
    item_name: Any,
    catalog: ServiceCatalog,
    *,
    line: int,
    sale_id: str,
    source_file: Path,
    report: ValidationReport,
) -> tuple[str, str, str]:
    """Localiza servico no DE_PARA por exato/alias e, por ultimo, proximidade."""

    def direct_candidates(value: Any) -> tuple[set[str], str]:
        keys = service_lookup_keys(value)
        if not keys:
            return set(), ""
        exact = catalog.by_name.get(keys[0], set())
        if exact:
            return set(exact), "EXATO"
        aliases: set[str] = set()
        for key in keys[1:]:
            aliases.update(catalog.by_name.get(key, set()))
        return aliases, "ALIAS" if aliases else ""

    primary, primary_kind = direct_candidates(service_name)
    fallback, fallback_kind = direct_candidates(item_name)
    all_candidates = primary | fallback

    if len(all_candidates) > 1:
        report.error(
            "SERVICO_AMBIGUO",
            file=source_file,
            line=line,
            sale_id=sale_id,
            field="ServicoNome/NomeItem",
            origin=f"ServicoNome={clean_text(service_name)}; NomeItem={clean_text(item_name)}",
            compared=", ".join(sorted(all_candidates, key=sort_identifier)),
            detail="A descricao aponta para mais de um codigo na aba de Servicos do DE-PARA.",
        )
        return "", "", ""

    if len(all_candidates) == 1:
        selected = next(iter(all_candidates))
        methods: list[str] = []
        if selected in primary:
            methods.append("ServicoNome" + (" alias" if primary_kind == "ALIAS" else ""))
        if selected in fallback and fallback != primary:
            methods.append("NomeItem" + (" alias" if fallback_kind == "ALIAS" else ""))
        value = _service_value_for_code(selected, catalog, line=line, sale_id=sale_id, source_file=source_file, report=report)
        return selected, value, "+".join(methods) or "DE_PARA"

    # Busca aproximada somente sobre os nomes existentes na aba de Servicos do DE-PARA.
    targets = []
    for raw in (service_name, item_name):
        for key in service_lookup_keys(raw):
            if key and key not in targets:
                targets.append(key)
    if targets and catalog.by_name:
        scored: list[tuple[float, str, str]] = []
        for target in targets:
            for candidate_key, codes in catalog.by_name.items():
                score = difflib.SequenceMatcher(None, target, candidate_key).ratio()
                for code in codes:
                    scored.append((score, candidate_key, code))
        if scored:
            best_score = max(item[0] for item in scored)
            best = [item for item in scored if abs(item[0] - best_score) < 1e-12]
            best_codes = {item[2] for item in best}
            if len(best_codes) == 1:
                selected = next(iter(best_codes))
                candidate_names = sorted({name for _, key, _ in best for name in catalog.display_names.get(key, [key])})
                report.warning(
                    "SERVICO_LOCALIZADO_APROXIMADO",
                    file=source_file,
                    line=line,
                    sale_id=sale_id,
                    field="ServicoNome/NomeItem",
                    origin=f"ServicoNome={clean_text(service_name)}; NomeItem={clean_text(item_name)}",
                    compared=f"{selected} - {' / '.join(candidate_names[:3])} - similaridade={best_score:.4f}",
                    detail="Nao houve correspondencia exata; foi escolhido o nome mais proximo dentre os servicos da aba de Servicos do DE-PARA.",
                )
                value = _service_value_for_code(selected, catalog, line=line, sale_id=sale_id, source_file=source_file, report=report)
                return selected, value, "APROXIMADO"

    report.error(
        "SERVICO_NAO_LOCALIZADO",
        file=source_file,
        line=line,
        sale_id=sale_id,
        field="ServicoNome/NomeItem",
        origin=f"ServicoNome={clean_text(service_name)}; NomeItem={clean_text(item_name)}",
        detail="Nenhum codigo foi encontrado na aba de Servicos do DE-PARA.",
    )
    return "", "", ""


def _decimal_from_text(value: Any) -> Decimal | None:
    if is_blank(value):
        return None
    text = clean_text(value).replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _decimal_output(value: Decimal) -> str:
    # Formato sem notacao cientifica e sem zeros decimais desnecessarios.
    if value == value.to_integral_value():
        return str(value.quantize(Decimal("1")))
    text = format(value.normalize(), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def process_balances(
    table: Table,
    sales: dict[str, SaleAggregate],
    sale_codes: dict[str, str],
    services: ServiceCatalog,
    report: ValidationReport,
) -> tuple[list[list[str]], list[list[str]], int]:
    optional = {"qty_remaining", "deletion_date", "old_service_id", "package_id"}
    columns = map_columns(table.headers, SESSION_FIELD_ALIASES, optional)
    if columns["qty_remaining"] is None and len(table.headers) > 10:
        columns["qty_remaining"] = 10

    output_rows: list[list[str]] = []
    trace_rows: list[list[str]] = []
    source_sale_ids: set[str] = set()
    old_service_to_new: dict[str, set[str]] = defaultdict(set)
    rejected_balance_counts: Counter[str] = Counter()

    for line, raw_row in table.rows:
        row = ensure_width(raw_row, len(table.headers))
        values = source_row_dict(row, columns)
        sale_id = normalize_id(values["sale_id"])
        if not sale_id:
            report.error("SALDO_SEM_VENDA", file=table.path, line=line, field="VendaId")
            continue
        source_sale_ids.add(sale_id)
        if sale_id not in sales:
            report.error(
                "SALDO_SEM_VENDA_CORRESPONDENTE",
                file=table.path,
                line=line,
                sale_id=sale_id,
                detail="O VendaId existe em Sessoes, mas nao na extracao de Vendas.",
            )
            continue
        sale = sales[sale_id]

        if normalize_status(values["status"]) != normalize_status(sale.raw_status):
            report.error(
                "DIVERGENCIA_STATUS_VENDAS_SESSOES",
                file=table.path,
                line=line,
                sale_id=sale_id,
                origin=values["status"],
                compared=sale.raw_status,
            )
        session_cpf = digits_only(values["client_cpf"])
        if session_cpf and sale.client_cpf and session_cpf != sale.client_cpf:
            report.error(
                "DIVERGENCIA_CPF_VENDAS_SESSOES",
                file=table.path,
                line=line,
                sale_id=sale_id,
                origin=session_cpf,
                compared=sale.client_cpf,
            )
        session_name = normalize_name(values["client_name"])
        if session_name and normalize_name(sale.client_name) and session_name != normalize_name(sale.client_name):
            report.error(
                "DIVERGENCIA_NOME_VENDAS_SESSOES",
                file=table.path,
                line=line,
                sale_id=sale_id,
                origin=values["client_name"],
                compared=sale.client_name,
            )

        service_code, service_value, match_method = match_service(
            values["service_name"],
            values["item_name"],
            services,
            line=line,
            sale_id=sale_id,
            source_file=table.path,
            report=report,
        )
        old_service_id = normalize_id(values["old_service_id"])
        if old_service_id and service_code:
            old_service_to_new[old_service_id].add(service_code)

        # SALDO_CALCULADO_POR_QUANTIDADE_V20260826_2
        # Regra do projeto: Quantidade faltante = Quantidade total - Quantidade utilizada.
        # QtdFaltante nao e usada como fonte obrigatoria porque algumas unidades
        # nao possuem esse campo e outros layouts podem fazer esse indice coincidir
        # com uma coluna de data (por exemplo DataDelete).
        total = numeric_value(values["qty_total"])
        used = numeric_value(values["qty_used"])
        if total is None or used is None:
            report.error(
                "QUANTIDADE_INVALIDA",
                file=table.path,
                line=line,
                sale_id=sale_id,
                origin=f"Quantidade={clean_text(values['qty_total'])}; Utilizada={clean_text(values['qty_used'])}",
                detail="Quantidade e QuantidadeUtilizada precisam ser numericas.",
            )
            calculated = None
            remaining_output = ""
        else:
            calculated = total - used
            if calculated < -1e-9:
                report.error(
                    "SALDO_NEGATIVO",
                    file=table.path,
                    line=line,
                    sale_id=sale_id,
                    origin=number_string(calculated),
                    detail="Quantidade utilizada e maior que a quantidade total.",
                )
            remaining_output = number_string(calculated)

        # Valor de servico zero e valido. O valor total sera zero quando aplicavel.
        service_decimal = _decimal_from_text(service_value)
        remaining_decimal = _decimal_from_text(remaining_output)
        total_value = ""
        if service_decimal is not None and remaining_decimal is not None:
            total_value = _decimal_output(service_decimal * remaining_decimal)
        elif service_code and calculated is not None:
            report.error(
                "VALOR_SERVICO_INVALIDO",
                file=table.path,
                line=line,
                sale_id=sale_id,
                origin=f"Valor Servico={service_value}",
                detail="O valor do servico precisa ser numerico; zero e permitido.",
            )

        output = [
            sale_codes.get(sale_id, ""),
            service_code,
            service_value,
            remaining_output,
            "",
            "%",
            total_value,
        ]
        if sale.importable:
            output_rows.append([clean_text(value) for value in output])
        else:
            rejected_balance_counts[sale_id] += 1
        trace_rows.append(
            [
                line,
                sale_id,
                sale_codes.get(sale_id, ""),
                sale.old_client_code,
                sale.imported_client_code,
                sale.client_match_method,
                clean_text(values["client_name"]),
                digits_only(values["client_cpf"]),
                old_service_id,
                clean_text(values["service_name"]),
                clean_text(values["item_name"]),
                service_code,
                service_value,
                remaining_output,
                total_value,
                match_method,
                number_string(values["qty_total"]),
                number_string(values["qty_used"]),
                clean_text(values["status"]),
                clean_text(values["package_id"]),
                format_datetime(values["deletion_date"], date1904=table.date1904),
            ]
        )

    sale_ids = set(sales)
    for missing in sorted(sale_ids - source_sale_ids, key=sort_identifier):
        report.error(
            "VENDA_SEM_LINHA_DE_SALDO",
            file=table.path,
            sale_id=missing,
            detail="Venda existe na extracao de Vendas, mas nao em Sessoes.",
        )
    for extra in sorted(source_sale_ids - sale_ids, key=sort_identifier):
        report.error("SALDO_DE_VENDA_INEXISTENTE", file=table.path, sale_id=extra)
    for old_id, new_codes in old_service_to_new.items():
        if len(new_codes) > 1:
            report.error(
                "SERVICO_ORIGEM_DIVERGENTE",
                file=table.path,
                field="ServicoId",
                origin=old_id,
                compared=", ".join(sorted(new_codes, key=sort_identifier)),
                detail="O mesmo ServicoId antigo foi relacionado a mais de um codigo novo.",
            )

    rejected_balance_total = sum(rejected_balance_counts.values())
    if rejected_balance_total:
        report.warning(
            "SALDOS_NAO_EXPORTADOS_POR_VENDA_REJEITADA",
            file=table.path,
            origin=rejected_balance_total,
            compared="; ".join(
                f"Venda {sale_id}: {count} saldo(s)"
                for sale_id, count in sorted(rejected_balance_counts.items(), key=lambda item: sort_identifier(item[0]))
            ),
            detail="As linhas permanecem na rastreabilidade, mas nao foram gravadas no CSV de saldo porque a venda correspondente nao possui Codigo do Cliente valido.",
        )
    report.info(
        "RESUMO_SALDOS_TRATADOS",
        file=table.path,
        origin=len(table.rows),
        compared=len(output_rows),
        detail="Linhas validas de Sessoes exportadas; linhas de vendas rejeitadas permanecem somente na rastreabilidade/validacao.",
    )
    return output_rows, trace_rows, rejected_balance_total


# ---------------------------------------------------------------------------
# Validacao dos arquivos auxiliares
# ---------------------------------------------------------------------------
AUX_SIGNATURE_FIELDS = [
    "sale_id",
    "observation",
    "unit_observation",
    "cancellation_reason",
    "status",
    "billing_date",
    "client_name",
    "unit",
    "client_phone",
    "client_cpf",
    "package",
    "item_type",
    "category",
    "item_status",
    "qty_total",
    "qty_used",
    "qty_remaining",
    "item_value",
    "sale_value",
]


def signature_value(value: Any, field_name: str, *, date1904: bool = False) -> str:
    if field_name == "sale_id":
        return normalize_id(value)
    if field_name == "client_cpf":
        return digits_only(value)
    if field_name == "client_phone":
        return digits_only(value)
    if field_name == "client_name":
        return normalize_name(value)
    if field_name == "status":
        return normalize_status(value)
    if field_name == "billing_date":
        return format_date(value, date1904=date1904)
    if field_name in {"qty_total", "qty_used", "qty_remaining", "item_value", "sale_value"}:
        return number_string(value)
    return clean_text(value)


def validate_auxiliary(
    aux_table: Table | None,
    main_source_rows: list[dict[str, Any]],
    report: ValidationReport,
    *,
    kind: str,
) -> None:
    if aux_table is None:
        report.warning(
            "ARQUIVO_AUXILIAR_NAO_ENCONTRADO",
            field=kind,
            detail="A validacao principal continua, mas este cruzamento adicional nao foi executado.",
        )
        return
    optional = set(SALES_FIELD_ALIASES) - {"sale_id", "status", "billing_date"}
    columns = map_columns(aux_table.headers, SALES_FIELD_ALIASES, optional)
    available_fields = [field_name for field_name in AUX_SIGNATURE_FIELDS if columns.get(field_name) is not None]
    if "sale_id" not in available_fields or "status" not in available_fields or "billing_date" not in available_fields:
        report.error(
            "AUXILIAR_LAYOUT_INVALIDO",
            file=aux_table.path,
            detail="VendaId, StatusVenda e VendaDataFaturamento sao necessarios.",
        )
        return

    def make_signature(values: dict[str, Any], date1904: bool) -> tuple[str, ...]:
        return tuple(signature_value(values.get(field_name, ""), field_name, date1904=date1904) for field_name in available_fields)

    main_filtered: list[dict[str, Any]] = []
    for values in main_source_rows:
        status = normalize_status(values.get("status", ""))
        date_blank = not format_date(values.get("billing_date", ""))
        if kind == "pendentes":
            if status == "PENDENTE A PAGAMENTO" and date_blank:
                main_filtered.append(values)
        elif kind == "sem_data":
            if date_blank:
                main_filtered.append(values)

    aux_values: list[dict[str, Any]] = []
    for line, raw_row in aux_table.rows:
        row = ensure_width(raw_row, len(aux_table.headers))
        values = source_row_dict(row, columns)
        values["_line"] = line
        aux_values.append(values)
        status = normalize_status(values.get("status", ""))
        date_blank = not format_date(values.get("billing_date", ""), date1904=aux_table.date1904)
        if kind == "pendentes" and status != "PENDENTE A PAGAMENTO":
            report.error(
                "AUXILIAR_PENDENTE_STATUS_INVALIDO",
                file=aux_table.path,
                line=line,
                sale_id=values.get("sale_id", ""),
                origin=values.get("status", ""),
            )
        if not date_blank:
            report.error(
                "AUXILIAR_DATA_NAO_VAZIA",
                file=aux_table.path,
                line=line,
                sale_id=values.get("sale_id", ""),
                origin=values.get("billing_date", ""),
            )

    main_counter = Counter(make_signature(values, False) for values in main_filtered)
    aux_counter = Counter(make_signature(values, aux_table.date1904) for values in aux_values)
    missing = main_counter - aux_counter
    extra = aux_counter - main_counter
    if missing or extra:
        report.error(
            "AUXILIAR_DIVERGENTE_DA_EXTRACAO",
            file=aux_table.path,
            field=kind,
            origin=sum(missing.values()),
            compared=sum(extra.values()),
            detail="Linhas ausentes no auxiliar e linhas excedentes em relacao a extracao principal.",
        )
        for signature, count in list(missing.items())[:20]:
            sale_id = signature[available_fields.index("sale_id")]
            report.error(
                "LINHA_AUSENTE_NO_AUXILIAR",
                file=aux_table.path,
                sale_id=sale_id,
                origin=count,
                detail="Assinatura: " + " | ".join(signature),
            )
        for signature, count in list(extra.items())[:20]:
            sale_id = signature[available_fields.index("sale_id")]
            report.error(
                "LINHA_EXCEDENTE_NO_AUXILIAR",
                file=aux_table.path,
                sale_id=sale_id,
                origin=count,
                detail="Assinatura: " + " | ".join(signature),
            )
    report.info(
        "RESUMO_AUXILIAR",
        file=aux_table.path,
        field=kind,
        origin=len(aux_values),
        compared=len(main_filtered),
        detail="Linhas do arquivo auxiliar e linhas equivalentes da extracao principal.",
    )


# ---------------------------------------------------------------------------
# Rastreabilidade e verificacoes finais
# ---------------------------------------------------------------------------
SALES_TRACE_HEADERS = [
    "Linha na extracao de Vendas",
    "VendaId origem",
    "Codigo venda importacao",
    "Codigo cliente origem",
    "Codigo cliente importacao",
    "Metodo busca cliente",
    "Observacao origem",
    "ObservacaoUnidade origem",
    "TipoCancelamentoId origem",
    "MotivoCancelamento origem",
    "StatusVenda origem",
    "Status importacao",
    "VendaDataFaturamento origem",
    "Data venda importacao",
    "Data padrao aplicada",
    "Origem da data de venda",
    "ClienteNome origem",
    "Unidade origem",
    "ClienteCelular origem",
    "ClienteCpf origem",
    "PacoteServico origem",
    "Tipo origem",
    "Categoria origem",
    "StatusItemComandaServicoId origem",
    "QtdTotal origem",
    "QtdRealizado origem",
    "QtdFaltante origem",
    "ValorItem origem",
    "ValorVenda origem",
    "ValorVendaPago origem",
    "UltimaDataSessaoRealizada origem",
    "DataDelete origem",
    "QtdDiasSemAgendamento origem",
    "ValorUnitarioSessao origem",
    "Observacao final importacao",
]

BALANCE_TRACE_HEADERS = [
    "Linha na extracao de Sessoes",
    "VendaId origem",
    "Codigo venda importacao",
    "Codigo cliente origem",
    "Codigo cliente importacao",
    "Metodo busca cliente",
    "ClienteNome origem",
    "ClienteCpf origem",
    "ServicoId origem",
    "ServicoNome origem",
    "NomeItem origem",
    "Codigo servico importacao",
    "Valor servico importacao",
    "Quantidade faltante",
    "Valor total",
    "Metodo busca servico",
    "Quantidade total",
    "Quantidade utilizada",
    "Status origem",
    "PacoteServicosId origem",
    "DataDelete origem",
]


def _trace_number(value: Any) -> str:
    result = number_string(value)
    return result if result else "null"


def _trace_date(value: Any, *, date1904: bool = False, with_time: bool = False) -> str:
    result = format_datetime(value, date1904=date1904) if with_time else format_date(value, date1904=date1904)
    return result if result else "null"


def sales_trace_rows(sales: dict[str, SaleAggregate], *, date1904: bool = False) -> list[list[str]]:
    """Gera uma linha de rastreabilidade para cada linha original de Vendas."""
    rows: list[list[str]] = []
    for source_id in sorted(sales, key=sort_identifier):
        sale = sales[source_id]
        output_date = sale.output_row[2] if len(sale.output_row) > 2 else ""
        output_observation = sale.output_row[11] if len(sale.output_row) > 11 else ""
        for raw in sorted(sale.raw_rows, key=lambda item: int(item.get("_line", 0))):
            rows.append(
                [
                    raw.get("_line", ""),
                    source_id,
                    sale.imported_sale_code,
                    normalize_id(raw.get("old_client_code", "")),
                    sale.imported_client_code,
                    sale.client_match_method,
                    null_text(raw.get("observation", "")),
                    null_text(raw.get("unit_observation", "")),
                    null_text(raw.get("cancellation_type_id", "")),
                    null_text(raw.get("cancellation_reason", "")),
                    null_text(raw.get("status", "")),
                    sale.mapped_status,
                    _trace_date(raw.get("billing_date", ""), date1904=date1904),
                    output_date,
                    "sim" if sale.used_default_date else "",
                    sale.sale_date_source,
                    null_text(raw.get("client_name", "")),
                    null_text(raw.get("unit", "")),
                    null_text(raw.get("client_phone", "")),
                    null_text(raw.get("client_cpf", "")),
                    null_text(raw.get("package", "")),
                    null_text(raw.get("item_type", "")),
                    null_text(raw.get("category", "")),
                    _trace_number(raw.get("item_status", "")),
                    _trace_number(raw.get("qty_total", "")),
                    _trace_number(raw.get("qty_used", "")),
                    _trace_number(raw.get("qty_remaining", "")),
                    _trace_number(raw.get("item_value", "")),
                    _trace_number(raw.get("sale_value", "")),
                    _trace_number(raw.get("paid_value", "")),
                    _trace_date(raw.get("last_session", ""), date1904=date1904, with_time=True),
                    _trace_date(raw.get("deletion_date", ""), date1904=date1904, with_time=True),
                    _trace_number(raw.get("days_without_booking", "")),
                    _trace_number(raw.get("unit_session_value", "")),
                    output_observation,
                ]
            )
    return rows


def mapping_rows(sale_codes: dict[str, str]) -> list[list[str]]:
    return [[source_id, sale_codes[source_id]] for source_id in sorted(sale_codes, key=sort_identifier)]


def check_output_integrity(
    sales: dict[str, SaleAggregate],
    sales_rows: list[list[str]],
    balance_rows: list[list[str]],
    source_balance_count: int,
    rejected_balance_count: int,
    report: ValidationReport,
) -> None:
    expected_sales = sum(1 for sale in sales.values() if sale.importable)
    if len(sales_rows) != expected_sales:
        report.error(
            "CONTAGEM_VENDAS_FINAL",
            origin=len(sales_rows),
            compared=expected_sales,
            detail="Deve existir exatamente uma linha por VendaId apto para importacao.",
        )
    expected_balances = source_balance_count - rejected_balance_count
    if len(balance_rows) != expected_balances:
        report.error(
            "CONTAGEM_SALDOS_FINAL",
            origin=len(balance_rows),
            compared=expected_balances,
            detail="Cada linha de Sessoes de uma venda apta deve gerar uma linha no saldo.",
        )
    sale_codes = [row[0] for row in sales_rows if row]
    if len(sale_codes) != len(set(sale_codes)):
        report.error("CODIGO_VENDA_FINAL_DUPLICADO", detail="Ha codigos repetidos no arquivo de venda tratado.")
    sale_code_set = set(sale_codes)
    missing_references = Counter(row[0] for row in balance_rows if row and row[0] not in sale_code_set)
    for code, count in missing_references.items():
        report.error(
            "SALDO_REFERENCIA_VENDA_INEXISTENTE",
            origin=code,
            compared=count,
            detail="Linhas de saldo apontam para codigo que nao existe no arquivo de vendas.",
        )
    for index, row in enumerate(sales_rows, start=2):
        if len(row) != 12:
            report.error("LAYOUT_VENDA_FINAL", line=index, origin=len(row), compared=12)
        if any(not clean_text(row[position]) for position in (0, 1, 2, 9)):
            report.error("OBRIGATORIO_VENDA_VAZIO", line=index, origin=" | ".join(row))
        if row[9] == "Suspenso" and not clean_text(row[10]):
            report.error("DATA_SUSPENSO_VAZIA", line=index, origin=row[0])
    for index, row in enumerate(balance_rows, start=2):
        if len(row) != 7:
            report.error("LAYOUT_SALDO_FINAL", line=index, origin=len(row), compared=7)
        if any(not clean_text(row[position]) for position in (0, 1)):
            report.error("OBRIGATORIO_SALDO_VAZIO", line=index, origin=" | ".join(row))


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------
def script_hash() -> str:
    try:
        return hashlib.sha256(Path(__file__).read_bytes()).hexdigest()[:12]
    except OSError:
        return "indisponivel"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Tratamento de Venda de Planos e Saldos - Laser Rosa")
    parser.add_argument("--vendas", type=Path, help="Extracao principal de Vendas; se omitida, e identificada pelos cabecalhos em entrada")
    parser.add_argument("--sessoes", type=Path, help="Extracao principal de Sessoes; se omitida, e identificada pelos cabecalhos em entrada")
    parser.add_argument("--clientes", type=Path, help="planilhaTratadaCliente; padrao: pasta saida")
    parser.add_argument("--de-para", "--servicos", dest="de_para", type=Path, help="Arquivo unico DE-PARA .xls/.xlsx; o script seleciona a aba de Servicos")
    parser.add_argument("--modelo-vendas", type=Path, help="modeloImportacaoVendaPlano.csv; padrao: pasta entrada")
    parser.add_argument("--modelo-saldos", type=Path, help="modeloImportacaoSaldoVendaPlano.csv; padrao: pasta entrada")
    parser.add_argument("--modelo-clientes", type=Path, help="modeloImportacaoCliente.csv; usado somente se houver cliente ausente")
    parser.add_argument("--pendentes", type=Path, help="vendasPendentesPagamento.xlsx; padrao: pasta entrada")
    parser.add_argument("--sem-data", type=Path, help="vendasSemDataFaturamento.xlsx; padrao: pasta entrada")
    parser.add_argument("--saida-vendas", type=Path, help="CSV final; padrao: saida/planilhaTratadaVendaPlano.csv")
    parser.add_argument("--saida-saldos", type=Path, help="CSV final; padrao: saida/planilhaTratadaSaldoVendaPlano.csv")
    parser.add_argument("--saida-clientes-complementares", type=Path, help="CSV complementar; padrao: saida/planilhaTratadaClienteVendaPlano.csv")
    parser.add_argument("--validacao", type=Path, help="XLSX unico de auditoria; padrao: saida/validacaoVendaPlanoSaldo.xlsx")
    parser.add_argument(
        "--data-padrao-sem-faturamento",
        default=None,
        help="Opcional. Se omitida, usa automaticamente o primeiro dia do mes anterior ao da execucao.",
    )
    parser.add_argument(
        "--codigo-venda",
        choices=("automatico", "sequencial", "origem"),
        default="automatico",
        help="automatico reutiliza VendaId se todos couberem em 6 digitos; caso contrario gera 100000+",
    )
    return parser


def _output_path(argument: Path | None, default_name: str) -> Path:
    if argument is None:
        return (OUTPUT_DIR / default_name).resolve()
    return _resolve_explicit_path(argument, default_dir=OUTPUT_DIR)


def main() -> int:
    args = build_parser().parse_args()
    try:
        configure_project_layout()
        ensure_project_folders()
    except Exception as exc:
        print(f"Laser Rosa - Venda de Planos/Saldos - versao {VERSION}")
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    print(f"Laser Rosa - Venda de Planos/Saldos - versao {VERSION}")
    print(f"Raiz: {PROJECT_ROOT}")
    print(f"Entrada: {INPUT_DIR} | Saida: {OUTPUT_DIR}")

    report = ValidationReport()
    sales: dict[str, SaleAggregate] = {}
    sale_codes: dict[str, str] = {}
    sales_trace: list[list[str]] = []
    balance_trace: list[list[str]] = []
    generated_client_rows: list[list[str]] = []
    generated_client_template: Table | None = None
    generated_client_registry: GeneratedClientRegistry | None = None
    validation_path = _output_path(args.validacao, "validacaoVendaPlanoSaldo.xlsx")

    try:
        if args.data_padrao_sem_faturamento:
            default_date_parsed = parse_date_value(args.data_padrao_sem_faturamento)
            if default_date_parsed is None:
                raise ValueError("--data-padrao-sem-faturamento deve estar no formato dd/mm/aaaa")
        else:
            first_this_month = date.today().replace(day=1)
            last_previous_month = first_this_month - timedelta(days=1)
            default_date_parsed = last_previous_month.replace(day=1)
        default_date = default_date_parsed.strftime("%d/%m/%Y")

        # As planilhas recebidas dos clientes NAO possuem padrao de nome.
        # Vendas e Sessoes sao identificadas exclusivamente pelo layout/cabecalhos.
        vendas = discover_extraction_file(
            args.vendas,
            "vendas",
            search_dir=INPUT_DIR,
        )
        sessoes = discover_extraction_file(
            args.sessoes,
            "sessoes",
            search_dir=INPUT_DIR,
            exclude=(vendas,),
        )
        clientes = resolve_file(
            args.clientes,
            "planilhaTratadaCliente(s)",
            lambda path: path.suffix.lower() in {".csv", ".xlsx", ".xls"} and normalized_filename(path).startswith("planilhatratadacliente"),
            search_dir=OUTPUT_DIR,
        )
        de_para = resolve_depara_file(args.de_para, search_dir=INPUT_DIR)
        modelo_vendas = resolve_exact_or_variant(
            args.modelo_vendas,
            "modeloImportacaoVendaPlano",
            search_dir=INPUT_DIR,
            exact_stem="modeloImportacaoVendaPlano",
            suffixes={".csv"},
            exclude_tokens=("saldo",),
        )
        modelo_saldos = resolve_exact_or_variant(
            args.modelo_saldos,
            "modeloImportacaoSaldoVendaPlano",
            search_dir=INPUT_DIR,
            exact_stem="modeloImportacaoSaldoVendaPlano",
            suffixes={".csv"},
        )
        # Os auxiliares tambem podem chegar com nomes diferentes. Primeiro tenta
        # identifica-los pelo conteudo; se nao existirem, continuam opcionais.
        pendentes = discover_sales_auxiliary(
            args.pendentes,
            "pendentes",
            search_dir=INPUT_DIR,
            exclude=(vendas, sessoes, de_para, modelo_vendas, modelo_saldos),
        )
        sem_data = discover_sales_auxiliary(
            args.sem_data,
            "sem_data",
            search_dir=INPUT_DIR,
            exclude=tuple(path for path in (vendas, sessoes, de_para, modelo_vendas, modelo_saldos, pendentes) if path),
        )

        assert vendas and sessoes and clientes and de_para and modelo_vendas and modelo_saldos
        selected = {
            "Vendas": vendas,
            "Sessoes": sessoes,
            "Clientes tratados": clientes,
            "DE-PARA": de_para,
            "Modelo Vendas": modelo_vendas,
            "Modelo Saldos": modelo_saldos,
            "Pendentes": pendentes,
            "Sem data": sem_data,
        }
        for label, path in selected.items():
            report.info("ARQUIVO_SELECIONADO", file=path or "", field=label, origin=str(path or "nao encontrado"))
        print(f"Vendas: {vendas.name} | Sessoes: {sessoes.name} | DE-PARA: {de_para.name}")

        sales_table = read_table(vendas, ("expdata 1", "Planilha1", "Sheet1"))
        sessions_table = read_table(sessoes, ("expdata 1", "Planilha1", "Sheet1"))
        client_table = read_table(clientes, ("Sheet1", "Planilha1", "Plan1"))
        service_table = read_depara_sheet(de_para, ("serviços", "servicos", "serviço", "servico", "serv"))
        print(f"DE-PARA/Servicos: aba {service_table.sheet}")
        report.info("ABA_DE_PARA_SELECIONADA", file=de_para, field="Servicos", origin=service_table.sheet)
        sale_template = read_csv_table(modelo_vendas)
        balance_template = read_csv_table(modelo_saldos)
        if len(sale_template.headers) != 12:
            raise ValueError(f"Modelo de Venda de Plano possui {len(sale_template.headers)} colunas; esperado: 12")
        if len(balance_template.headers) != 7:
            raise ValueError(f"Modelo de Saldo de Venda de Plano possui {len(balance_template.headers)} colunas; esperado: 7")

        # IMPORTANTE: os modelos e planilhas tratadas enviados como exemplo servem apenas como estrutura.
        # Nenhuma linha dos modelos e usada para preencher ou cruzar dados. planilhaTratadaCliente
        # e excecao porque foi definida como fonte oficial de relacionamento de clientes.
        if sale_template.rows:
            report.info(
                "DADOS_NO_MODELO_VENDAS_IGNORADOS",
                file=modelo_vendas,
                origin=len(sale_template.rows),
                detail="Somente o cabecalho/formato do modelo foi considerado.",
            )
        if balance_template.rows:
            report.info(
                "DADOS_NO_MODELO_SALDOS_IGNORADOS",
                file=modelo_saldos,
                origin=len(balance_template.rows),
                detail="Somente o cabecalho/formato do modelo foi considerado.",
            )

        aux_pending = read_table(pendentes, ("Planilha1", "Plan1", "Sheet1")) if pendentes else None
        aux_no_date = read_table(sem_data, ("Plan1", "Planilha1", "Sheet1")) if sem_data else None

        sales, source_sales_rows = aggregate_sales(sales_table, report)
        client_catalog = build_client_catalog(client_table, report)
        generated_client_registry = GeneratedClientRegistry.from_catalog(client_catalog)
        service_catalog = build_service_catalog(service_table, report)
        sale_codes = assign_sale_codes(sales, args.codigo_venda, report)
        sale_rows = build_sales_output(
            sales,
            client_catalog,
            sale_codes,
            {},
            default_date,
            vendas,
            modelo_vendas,
            report,
            generated_clients=generated_client_registry,
        )

        # Se algum cliente nao existia na planilhaTratadaCliente, cria uma
        # planilha complementar com o mesmo layout do modeloImportacaoCliente.
        # Esses novos codigos ja foram usados nas vendas acima.
        if generated_client_registry.by_key:
            modelo_clientes = resolve_exact_or_variant(
                args.modelo_clientes,
                "modeloImportacaoCliente",
                search_dir=INPUT_DIR,
                exact_stem="modeloImportacaoCliente",
                suffixes={".csv"},
            )
            generated_client_template = read_csv_table(modelo_clientes)
            generated_client_rows, invalid_generated_codes = build_generated_client_rows(
                generated_client_template,
                generated_client_registry,
                report,
            )
            invalidate_sales_for_generated_clients(sales, generated_client_registry, invalid_generated_codes)
            if invalid_generated_codes:
                sale_rows = [
                    sales[source_id].output_row
                    for source_id in sorted(sales, key=sort_identifier)
                    if sales[source_id].importable and sales[source_id].output_row
                ]

        balance_rows, balance_trace, rejected_balance_count = process_balances(
            sessions_table,
            sales,
            sale_codes,
            service_catalog,
            report,
        )
        validate_auxiliary(aux_pending, source_sales_rows, report, kind="pendentes")
        validate_auxiliary(aux_no_date, source_sales_rows, report, kind="sem_data")
        check_output_integrity(
            sales,
            sale_rows,
            balance_rows,
            len(sessions_table.rows),
            rejected_balance_count,
            report,
        )
        sales_trace = sales_trace_rows(sales, date1904=sales_table.date1904)

        report.info(
            "RESUMO_FINAL",
            origin=f"vendas={len(sale_rows)}; saldos={len(balance_rows)}",
            compared=f"erros={report.error_count}; avisos={report.warning_count}",
            detail="Entradas lidas de entrada; saidas geradas em saida; modelos usados somente como formato.",
        )

        # A validacao consolidada so existe quando ha algo que exige revisao.
        # INFOs de rotina nao justificam criar um arquivo de validacao.
        has_validation_findings = bool(report.error_count or report.warning_count)
        if has_validation_findings:
            write_validation_xlsx(
                validation_path,
                sale_codes={sale_id: code for sale_id, code in sale_codes.items() if sales.get(sale_id) and sales[sale_id].importable},
                sales_trace=sales_trace,
                balance_trace=balance_trace,
                report=report,
            )
        elif validation_path.exists():
            # Remove uma validacao antiga para nao parecer que ela pertence a
            # uma execucao atual que terminou sem erros/avisos.
            validation_path.unlink()

        if report.error_count:
            print(f"BLOQUEADO: {report.error_count} erro(s) e {report.warning_count} aviso(s).")
            print(f"Validacao: {validation_path}")
            print("Nenhum CSV final de importacao foi criado ou sobrescrito nesta execucao.")
            return 2

        output_sales = _output_path(args.saida_vendas, "planilhaTratadaVendaPlano.csv")
        output_balances = _output_path(args.saida_saldos, "planilhaTratadaSaldoVendaPlano.csv")
        output_generated_clients = _output_path(
            getattr(args, "saida_clientes_complementares", None),
            "planilhaTratadaClienteVendaPlano.csv",
        )
        if generated_client_rows and generated_client_template is not None:
            write_csv_atomic(output_generated_clients, generated_client_template.headers, generated_client_rows)
            validate_csv_output(
                output_generated_clients,
                [str(value) for value in generated_client_template.headers],
                len(generated_client_rows),
            )
        elif output_generated_clients.exists():
            output_generated_clients.unlink()

        write_csv_atomic(output_sales, sale_template.headers, sale_rows)
        write_csv_atomic(output_balances, balance_template.headers, balance_rows)
        validate_csv_output(output_sales, [str(value) for value in sale_template.headers], len(sale_rows))
        validate_csv_output(output_balances, [str(value) for value in balance_template.headers], len(balance_rows))

        if generated_client_rows:
            print(f"OK: {len(generated_client_rows)} cliente(s) complementar(es) -> {output_generated_clients.name}")
            print("Ordem de importacao: clientes complementares primeiro; depois Venda de Planos e Saldo.")
        # RESUMO_CONTAGEM_ORIGEM_V20260826_8
        print(f"OK: {len(sale_rows)} vendas unicas de {len(source_sales_rows)} linha(s) da extracao -> {output_sales.name}")
        print(f"OK: {len(balance_rows)} saldos de {len(sessions_table.rows)} linha(s) da extracao -> {output_balances.name}")
        rejected_sales_count = sum(1 for sale in sales.values() if not sale.importable)
        if rejected_sales_count:
            print(f"Revisao: {rejected_sales_count} venda(s) ainda nao puderam ser importadas e {rejected_balance_count} saldo(s) ficaram fora dos CSVs.")
        if has_validation_findings:
            print(f"Validacao consolidada: {validation_path}")
            print(f"Validacao: {report.warning_count} aviso(s), zero erros.")
        else:
            print("Validacao: nenhum erro ou aviso; arquivo de validacao nao foi criado.")
        return 0

    except Exception as exc:
        report.error("FALHA_FATAL", detail=f"{type(exc).__name__}: {exc}")
        try:
            write_validation_xlsx(
                validation_path,
                sale_codes=sale_codes,
                sales_trace=sales_trace,
                balance_trace=balance_trace,
                report=report,
            )
        except Exception as xlsx_exc:
            print(f"ERRO ao gravar validacao XLSX: {xlsx_exc}", file=sys.stderr)
        print(f"ERRO: {exc}", file=sys.stderr)
        print(f"Validacao: {validation_path}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
