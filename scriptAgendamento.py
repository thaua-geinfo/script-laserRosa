#!/usr/bin/env python3
"""Tratamento de Agendamentos - Laser Rosa.

Estrutura do projeto:
    raiz/
      entrada/   -> extracao de agendamentos/sessoes, modeloImportacaoAgendamento*.csv e DE-PARA
      saida/     -> planilhaTratadaCliente.csv, planilhaTratadaAgendamento.csv e validacoes
      scriptAgendamento.py

Fontes de dados:
- a extracao de agendamentos fornece os dados do agendamento;
- planilhaTratadaCliente, em ./saida, fornece o codigo de cliente;
- planilhaTratadaClienteVendaPlano, quando existir, e carregada apenas como complemento;
- o arquivo DE-PARA, em ./entrada, fornece codigos de sala/agenda e servicos;
- o modelo de importacao serve somente como cabecalho/formato. Nenhuma linha
  preenchida do modelo participa do tratamento.

Regras:
- remove aspas simples/duplas, barra invertida, controles, quebras de linha e espacos excedentes;
- cliente e relacionado pela base principal + complementar; ausentes com identidade suficiente geram cliente complementar;
- sala e servico sao relacionados pelo nome no DE-PARA;
- o profissional e opcional, pois a importacao utiliza a Sala/Agenda como parametro;
- nomes de servico aceitam busca aproximada quando nao ha correspondencia exata;
- Observacao recebe "Importacao dd/mm/aaaa" e os dados da extracao nao importados
  nas demais colunas, com rotulo/indicativo;
- linhas com campo obrigatorio ou relacionamento obrigatorio ausente sao retiradas
  da planilha final e registradas em saida/planilhaTratadaAgendamento_linhas_rejeitadas.csv.

O leitor principal de XLSX ignora styles.xml para tolerar estilos corrompidos na extracao.
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
import xml.etree.ElementTree as ET
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path, PurePosixPath
from typing import Any, Iterator, Sequence
from zipfile import ZipFile

VERSION = "2026-08-26.7"
SCRIPT_DIR = Path(__file__).resolve().parent

def _discover_project_root(script_dir: Path) -> Path:
    """Localiza a raiz procurando entrada + saida/saída a partir de __file__."""
    for candidate in (script_dir, *script_dir.parents):
        if not (candidate / "entrada").is_dir():
            continue
        if (candidate / "saida").is_dir() or (candidate / "saída").is_dir():
            return candidate
    return script_dir


PROJECT_ROOT = _discover_project_root(SCRIPT_DIR)
INPUT_DIR = PROJECT_ROOT / "entrada"
if (PROJECT_ROOT / "saida").is_dir():
    OUTPUT_DIR = PROJECT_ROOT / "saida"
elif (PROJECT_ROOT / "saída").is_dir():
    OUTPUT_DIR = PROJECT_ROOT / "saída"
else:
    OUTPUT_DIR = PROJECT_ROOT / "saida"

FORBIDDEN_CHARS = str.maketrans({'"': '', "'": '', '\\': ''})
CELL_REF_RE = re.compile(r"([A-Z]+)(\d+)")
FREE_SECTOR = 0xFFFFFFFF
END_OF_CHAIN = 0xFFFFFFFE


# ---------------------------------------------------------------------------
# Normalizacao
# ---------------------------------------------------------------------------

def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).translate(FORBIDDEN_CHARS)
    text = "".join(ch for ch in text if unicodedata.category(ch) not in {"Cc", "Cf"})
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).strip()


def ascii_fold(value: Any) -> str:
    text = unicodedata.normalize("NFKD", clean_text(value))
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize_header(value: Any) -> str:
    text = ascii_fold(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def digits_only(value: Any) -> str:
    return re.sub(r"\D", "", clean_text(value))


def normalize_phone_keys(value: Any) -> set[str]:
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


def normalize_cpf(value: Any) -> str:
    digits = digits_only(value)
    return digits if len(digits) == 11 else ""


def service_alias(value: Any) -> str:
    text = normalize_header(value)
    text = re.sub(r"^(migracao|transferencia|cortesia)\s+", "", text).strip()
    text = re.sub(r"\bmasculina\b", "masculino", text)
    text = re.sub(r"\bfeminina\b", "feminino", text)
    return text


def room_alias(value: Any) -> str:
    text = normalize_header(value)
    text = re.sub(r"^(agenda|sala)\s+", "", text).strip()
    text = re.sub(r"\s+(volante|agenda)\b.*$", "", text).strip()
    return text


def map_status(value: Any) -> str:
    raw = clean_text(value)
    key = normalize_header(raw)
    mapping = {
        "realizado": "Atendido",
        "comparecido": "Atendido",
        "atendido": "Atendido",
        "agendado": "Marcado",
        "marcado": "Marcado",
        "faltou": "Falhou",
        "falhou": "Falhou",
        "confirmado": "Confirmado",
        "em andamento": "Em Andamento",
        "aguardando": "Aguardando",
        "cancelado": "Cancelado",
        "desmarcado": "Desmarcado",
        "antecipado": "Antecipado",
    }
    return mapping.get(key, raw)


def column_index(reference: str) -> int:
    match = CELL_REF_RE.match(reference)
    if not match:
        return -1
    value = 0
    for char in match.group(1):
        value = value * 26 + (ord(char) - 64)
    return value - 1


def excel_serial_to_datetime(value: Any, date1904: bool = False) -> datetime | None:
    try:
        serial = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(serial):
        return None
    if date1904:
        return datetime(1904, 1, 1) + timedelta(days=serial)
    return datetime(1899, 12, 30) + timedelta(days=serial)


def format_date(value: Any, date1904: bool = False) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    text = clean_text(value)
    candidates = [text]
    if "T" in text:
        candidates.append(text.split("T", 1)[0])
    if " " in text:
        candidates.append(text.split(" ", 1)[0])
    for candidate in candidates:
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(candidate, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
    parsed = excel_serial_to_datetime(value, date1904)
    return parsed.strftime("%d/%m/%Y") if parsed else text


def format_time(value: Any, date1904: bool = False) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    text = clean_text(value)
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).strftime("%H:%M")
        except ValueError:
            continue
    parsed = excel_serial_to_datetime(value, date1904)
    if parsed:
        return parsed.strftime("%H:%M")
    return text[:5] if len(text) >= 5 else text


def find_header_index(headers: Sequence[Any], *names: str) -> int | None:
    normalized = [normalize_header(header) for header in headers]
    targets = [normalize_header(name) for name in names]
    for target in targets:
        for index, header in enumerate(normalized):
            if header == target:
                return index
    for target in targets:
        target_tokens = set(target.split())
        for index, header in enumerate(normalized):
            if target_tokens and target_tokens.issubset(set(header.split())):
                return index
    return None


# ---------------------------------------------------------------------------
# Leitura de XLSX sem styles.xml
# ---------------------------------------------------------------------------

class XlsxReader:
    MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

    def __init__(self, path: Path):
        self.path = path
        self.archive = ZipFile(path)
        self.shared_strings = self._load_shared_strings()
        self.sheets, self.date1904 = self._load_workbook_info()

    def close(self) -> None:
        self.archive.close()

    def __enter__(self) -> "XlsxReader":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def _load_shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self.archive.namelist():
            return []
        root = ET.fromstring(self.archive.read("xl/sharedStrings.xml"))
        result: list[str] = []
        for item in root.findall(f"{{{self.MAIN_NS}}}si"):
            parts = [node.text or "" for node in item.iter(f"{{{self.MAIN_NS}}}t")]
            result.append("".join(parts))
        return result

    def _load_workbook_info(self) -> tuple[dict[str, str], bool]:
        workbook = ET.fromstring(self.archive.read("xl/workbook.xml"))
        relations_root = ET.fromstring(self.archive.read("xl/_rels/workbook.xml.rels"))
        relations: dict[str, str] = {}
        for relation in relations_root.findall(f"{{{self.PKG_REL_NS}}}Relationship"):
            relation_id = relation.attrib.get("Id", "")
            target = relation.attrib.get("Target", "")
            if not relation_id or not target:
                continue
            if target.startswith("/"):
                resolved = target.lstrip("/")
            else:
                resolved = posixpath.normpath(str(PurePosixPath("xl") / target))
            relations[relation_id] = resolved

        date1904 = False
        workbook_properties = workbook.find(f"{{{self.MAIN_NS}}}workbookPr")
        if workbook_properties is not None:
            date1904 = workbook_properties.attrib.get("date1904", "0").lower() in {"1", "true"}

        sheets: dict[str, str] = {}
        sheets_node = workbook.find(f"{{{self.MAIN_NS}}}sheets")
        if sheets_node is not None:
            for sheet in sheets_node.findall(f"{{{self.MAIN_NS}}}sheet"):
                name = sheet.attrib.get("name", "")
                relation_id = sheet.attrib.get(f"{{{self.REL_NS}}}id", "")
                if name and relation_id in relations:
                    sheets[name] = relations[relation_id]
        return sheets, date1904

    def iter_rows(
        self,
        sheet_name: str,
        min_row: int = 1,
        max_col: int = 100,
        max_row: int | None = None,
    ) -> Iterator[tuple[int, list[Any]]]:
        if sheet_name in self.sheets:
            sheet_path = self.sheets[sheet_name]
        else:
            lowered = {name.lower(): path for name, path in self.sheets.items()}
            if sheet_name.lower() in lowered:
                sheet_path = lowered[sheet_name.lower()]
            elif self.sheets:
                sheet_path = next(iter(self.sheets.values()))
            else:
                raise KeyError(f"Nenhuma aba encontrada em {self.path}")

        with self.archive.open(sheet_path) as handle:
            for _, element in ET.iterparse(handle, events=("end",)):
                if element.tag != f"{{{self.MAIN_NS}}}row":
                    continue
                try:
                    row_number = int(element.attrib.get("r", "0"))
                except ValueError:
                    row_number = 0
                if row_number < min_row:
                    element.clear()
                    continue
                if max_row is not None and row_number > max_row:
                    element.clear()
                    break

                values: list[Any] = [None] * max_col
                for cell in element.findall(f"{{{self.MAIN_NS}}}c"):
                    reference = cell.attrib.get("r", "")
                    index = column_index(reference)
                    if index < 0 or index >= max_col:
                        continue
                    cell_type = cell.attrib.get("t", "")
                    value_node = cell.find(f"{{{self.MAIN_NS}}}v")
                    inline_node = cell.find(f"{{{self.MAIN_NS}}}is")
                    value: Any = None
                    if cell_type == "inlineStr" and inline_node is not None:
                        value = "".join(
                            node.text or "" for node in inline_node.iter(f"{{{self.MAIN_NS}}}t")
                        )
                    elif value_node is not None and value_node.text is not None:
                        raw = value_node.text
                        if cell_type == "s":
                            try:
                                value = self.shared_strings[int(raw)]
                            except (ValueError, IndexError):
                                value = ""
                        elif cell_type == "b":
                            value = raw == "1"
                        else:
                            value = raw
                    values[index] = value
                yield row_number, values
                element.clear()


# ---------------------------------------------------------------------------
# Leitura de XLS legado para DE-PARA
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
            self.entries.append({"name": name, "type": entry[66], "start": _u32(entry, 116), "size": _u64(entry, 120)})
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
                current_wide = bool(self.read_plain(1)[0] & 1)
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
    reader.read_plain(4)
    unique = _u32(reader.read_plain(4))
    strings: list[str] = []
    for _ in range(unique):
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
    return strings


def _iter_biff_records(data: bytes, start: int = 0) -> Iterator[tuple[int, int, bytes]]:
    position = start
    while position + 4 <= len(data):
        record_id, length = struct.unpack_from("<HH", data, position)
        payload = data[position + 4 : position + 4 + length]
        if position + 4 + length > len(data):
            break
        yield position, record_id, payload
        position += 4 + length


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
        records = list(_iter_biff_records(self.data))
        index = 0
        while index < len(records):
            _, record_id, payload = records[index]
            if record_id == 0x0085:
                offset = _u32(payload, 0)
                length = payload[6]
                flags = payload[7]
                raw = payload[8 : 8 + (2 * length if flags & 1 else length)]
                name = raw.decode("utf-16le" if flags & 1 else "latin1", "replace")
                self.sheets.append((name, offset))
            elif record_id == 0x00FC:
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

    def sheet_rows(self, name: str) -> list[tuple[int, list[Any]]]:
        match = next((sheet for sheet in self.sheets if sheet[0].lower() == name.lower()), None)
        if match is None:
            raise KeyError(name)
        cells: dict[tuple[int, int], Any] = {}
        max_row = -1
        max_col = -1
        for _, record_id, payload in _iter_biff_records(self.data, match[1]):
            if record_id == 0x000A:
                break
            row = col = None
            value: Any = None
            if record_id == 0x00FD:
                row, col = _u16(payload, 0), _u16(payload, 2)
                string_index = _u32(payload, 6)
                value = self.shared_strings[string_index] if string_index < len(self.shared_strings) else ""
            elif record_id == 0x0203:
                row, col = _u16(payload, 0), _u16(payload, 2)
                value = _compact_number(struct.unpack_from("<d", payload, 6)[0])
            elif record_id == 0x027E:
                row, col = _u16(payload, 0), _u16(payload, 2)
                value = _compact_number(_decode_rk(_u32(payload, 6)))
            elif record_id == 0x00BD:
                row = _u16(payload, 0)
                first_col = _u16(payload, 2)
                body = payload[4:-2]
                for body_offset in range(0, len(body), 6):
                    if body_offset + 6 > len(body):
                        break
                    current_col = first_col + body_offset // 6
                    cells[(row, current_col)] = _compact_number(_decode_rk(_u32(body, body_offset + 2)))
                    max_row = max(max_row, row)
                    max_col = max(max_col, current_col)
                continue
            if row is not None and col is not None:
                cells[(row, col)] = value
                max_row = max(max_row, row)
                max_col = max(max_col, col)
        rows: list[tuple[int, list[Any]]] = []
        for row_index in range(max_row + 1):
            values = [cells.get((row_index, col_index), "") for col_index in range(max_col + 1)]
            if any(clean_text(value) for value in values):
                rows.append((row_index + 1, values))
        return rows


# ---------------------------------------------------------------------------
# Tabelas auxiliares
# ---------------------------------------------------------------------------

def read_csv_rows(path: Path) -> list[tuple[int, list[Any]]]:
    raw = path.read_bytes()
    text = None
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"Nao foi possivel ler {path}")
    first_line = next((line for line in text.splitlines() if line.strip()), "")
    if ";" in first_line:
        delimiter = ";"
    else:
        try:
            delimiter = csv.Sniffer().sniff(text[:10000], delimiters=",\t|").delimiter
        except csv.Error:
            delimiter = ","
    return [
        (line_number, list(row))
        for line_number, row in enumerate(csv.reader(io.StringIO(text), delimiter=delimiter), start=1)
        if any(clean_text(value) for value in row)
    ]


def read_all_sheets(path: Path) -> dict[str, list[tuple[int, list[Any]]]]:
    suffix = path.suffix.lower()
    if suffix in {".csv", ".txt"}:
        return {path.stem: read_csv_rows(path)}
    if suffix == ".xlsx":
        with XlsxReader(path) as book:
            return {
                name: list(book.iter_rows(name, min_row=1, max_col=120))
                for name in book.sheets
            }
    if suffix == ".xls":
        book = XlsBook(path)
        return {name: book.sheet_rows(name) for name in book.sheet_names}
    raise ValueError(f"Formato nao suportado: {path}")


def read_csv_template(path: Path) -> tuple[list[str], str, str]:
    rows = read_csv_rows(path)
    if not rows:
        raise ValueError(f"Modelo vazio: {path}")
    raw = path.read_bytes()
    encoding = "utf-8-sig"
    for candidate in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = raw.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue
    first_line = next((line for line in text.splitlines() if line.strip()), "")
    if ";" in first_line:
        delimiter = ";"
    else:
        try:
            delimiter = csv.Sniffer().sniff(text[:10000], delimiters=",\t|").delimiter
        except csv.Error:
            delimiter = ","
    # Somente a primeira linha (cabecalho) do modelo e utilizada.
    return [clean_text(value) for value in rows[0][1]], delimiter, encoding


@dataclass
class ClientCatalog:
    codes: set[str] = field(default_factory=set)
    by_name: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    by_phone: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    by_cpf: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    by_origin: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))


def add_client_to_catalog(
    catalog: ClientCatalog,
    *,
    code: Any,
    name: Any = "",
    phone: Any = "",
    cpf: Any = "",
    origin_code: Any = "",
) -> None:
    code_text = clean_text(code)
    if not code_text:
        return
    catalog.codes.add(code_text)
    name_key = normalize_header(name)
    if name_key:
        catalog.by_name[name_key].add(code_text)
    cpf_key = normalize_cpf(cpf)
    if cpf_key:
        catalog.by_cpf[cpf_key].add(code_text)
    for phone_key in normalize_phone_keys(phone):
        catalog.by_phone[phone_key].add(code_text)
    origin = clean_text(origin_code)
    if origin:
        catalog.by_origin[origin].add(code_text)


@dataclass
class DeParaCatalog:
    service_codes: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    service_names: dict[str, list[str]] = field(default_factory=lambda: defaultdict(list))
    duration_by_code: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    room_codes: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    room_names: dict[str, list[str]] = field(default_factory=lambda: defaultdict(list))


def _header_code_index(headers: Sequence[Any], context: str) -> int | None:
    """Localiza a coluna de codigo, inclusive cabecalhos compactos como codServicos."""
    normalized = [normalize_header(value) for value in headers]
    compact = [re.sub(r"[^a-z0-9]+", "", header) for header in normalized]
    context_fragments = ("servic", "proced", "item") if context == "service" else ("sala", "agenda")

    # Prioriza cabecalhos que trazem codigo + contexto, mesmo quando vierem
    # concatenados/camelCase na origem (ex.: codServicos, codigoSala, idAgenda).
    for index, header in enumerate(compact):
        is_code = header.startswith(("codigo", "cod", "id"))
        has_context = any(fragment in header for fragment in context_fragments)
        if is_code and has_context:
            return index

    # Depois aceita o formato separado por espacos, por compatibilidade.
    context_tokens = {"servico", "servicos", "procedimento", "procedimentos"} if context == "service" else {"sala", "salas", "agenda", "agendas"}
    for index, header in enumerate(normalized):
        tokens = set(header.split())
        if tokens & {"codigo", "cod", "id"} and tokens & context_tokens:
            return index

    # Por fim aceita uma coluna generica Codigo/Cod/ID quando a aba ja fornece
    # o contexto (como ocorre na aba sala do DE-PARA atual).
    for index, header in enumerate(compact):
        if header in {"codigo", "cod", "id"}:
            return index
    return None


def _header_name_indexes(headers: Sequence[Any], context: str, code_index: int) -> list[int]:
    normalized = [normalize_header(value) for value in headers]
    if context == "service":
        tokens = {"servico", "procedimento", "item", "descricao", "nome"}
    else:
        tokens = {"sala", "agenda", "descricao", "nome"}
    result = []
    for index, header in enumerate(normalized):
        if index == code_index:
            continue
        words = set(header.split())
        if words & tokens and "codigo" not in words and "id" not in words:
            result.append(index)
    return result


def _header_duration_index(headers: Sequence[Any]) -> int | None:
    normalized = [normalize_header(value) for value in headers]
    for index, header in enumerate(normalized):
        words = set(header.split())
        if words & {"duracao", "tempo", "minutos"}:
            return index
    return None


def build_client_catalog(paths: Path | Sequence[Path]) -> ClientCatalog:
    """Combina a planilha principal de clientes com eventuais complementares."""
    client_paths = [paths] if isinstance(paths, Path) else list(paths)
    if not client_paths:
        raise ValueError("Nenhuma planilha de clientes foi informada.")

    catalog = ClientCatalog()
    for path in client_paths:
        sheets = read_all_sheets(path)
        best: tuple[int, str, int, list[Any]] | None = None
        for sheet_name, rows in sheets.items():
            for line, headers in rows[:20]:
                normalized = [normalize_header(value) for value in headers]
                code_index = next(
                    (
                        index
                        for index, header in enumerate(normalized)
                        if "codigo" in header.split() and ("cliente" in header.split() or "max" in header.split() or index == 0)
                    ),
                    None,
                )
                name_index = next(
                    (
                        index
                        for index, header in enumerate(normalized)
                        if "nome" in header.split() and "mae" not in header.split() and "pai" not in header.split()
                    ),
                    None,
                )
                if code_index is None or name_index is None:
                    continue
                score = 2 + sum(
                    any(token in header.split() for token in ("cpf", "celular", "fone", "telefone", "origem"))
                    for header in normalized
                )
                if best is None or score > best[0]:
                    best = (score, sheet_name, line, headers)
        if best is None:
            raise ValueError(f"Nao consegui localizar Codigo/Nome em {path}")

        _, sheet_name, header_line, headers = best
        normalized = [normalize_header(value) for value in headers]
        code_index = next(
            index for index, header in enumerate(normalized)
            if "codigo" in header.split() and ("cliente" in header.split() or "max" in header.split() or index == 0)
        )
        name_index = next(
            index for index, header in enumerate(normalized)
            if "nome" in header.split() and "mae" not in header.split() and "pai" not in header.split()
        )
        cpf_indexes = [index for index, header in enumerate(normalized) if "cpf" in header.split()]
        phone_indexes = [
            index for index, header in enumerate(normalized)
            if any(token in header.split() for token in ("celular", "fone", "telefone")) and "ddi" not in header.split()
        ]
        origin_indexes = [
            index for index, header in enumerate(normalized)
            if "codigo" in header.split() and "origem" in header.split()
        ]

        for line, row in sheets[sheet_name]:
            if line <= header_line:
                continue
            code = clean_text(row[code_index] if code_index < len(row) else "")
            if not code:
                continue
            name = row[name_index] if name_index < len(row) else ""
            cpf = next((row[index] for index in cpf_indexes if index < len(row) and clean_text(row[index])), "")
            phone = next((row[index] for index in phone_indexes if index < len(row) and clean_text(row[index])), "")
            origin = next((row[index] for index in origin_indexes if index < len(row) and clean_text(row[index])), "")
            add_client_to_catalog(
                catalog,
                code=code,
                name=name,
                phone=phone,
                cpf=cpf,
                origin_code=origin,
            )
    return catalog



def build_depara_catalog(path: Path) -> DeParaCatalog:
    sheets = read_all_sheets(path)
    catalog = DeParaCatalog()

    for context in ("service", "room"):
        for sheet_name, rows in sheets.items():
            sheet_norm = normalize_header(sheet_name)
            for header_line, headers in rows[:30]:
                code_index = _header_code_index(headers, context)
                if code_index is None:
                    continue
                name_indexes = _header_name_indexes(headers, context, code_index)
                if not name_indexes:
                    continue

                header_norm = " ".join(normalize_header(value) for value in headers)
                header_compact = re.sub(r"[^a-z0-9]+", "", header_norm)
                if context == "service":
                    hinted = (
                        "serv" in sheet_norm
                        or any(token in header_norm.split() for token in ("servico", "procedimento", "item"))
                        or "servic" in header_compact
                        or "proced" in header_compact
                    )
                else:
                    hinted = (
                        any(token in sheet_norm.split() for token in ("sala", "agenda"))
                        or any(token in header_norm.split() for token in ("sala", "agenda"))
                        or "sala" in header_compact
                        or "agenda" in header_compact
                    )
                if not hinted:
                    continue

                duration_index = _header_duration_index(headers) if context == "service" else None
                for line, row in rows:
                    if line <= header_line:
                        continue
                    code = clean_text(row[code_index] if code_index < len(row) else "")
                    if not code:
                        continue
                    names = [clean_text(row[index] if index < len(row) else "") for index in name_indexes]
                    names = [name for name in names if name]
                    if not names:
                        continue
                    if context == "service":
                        for name in names:
                            key = service_alias(name)
                            if key:
                                catalog.service_codes[key].add(code)
                                if name not in catalog.service_names[key]:
                                    catalog.service_names[key].append(name)
                        if duration_index is not None and duration_index < len(row):
                            duration = clean_text(row[duration_index])
                            if duration:
                                catalog.duration_by_code[code].add(duration)
                    else:
                        for name in names:
                            key = normalize_header(name)
                            alias = room_alias(name)
                            for candidate in {key, alias}:
                                if candidate:
                                    catalog.room_codes[candidate].add(code)
                                    if name not in catalog.room_names[candidate]:
                                        catalog.room_names[candidate].append(name)
                break

    if not catalog.service_codes:
        raise ValueError(
            f"Nao encontrei tabela de servicos com nome/codigo no DE-PARA: {path}"
        )
    if not catalog.room_codes:
        raise ValueError(
            f"Nao encontrei tabela de salas/agendas com nome/codigo no DE-PARA: {path}"
        )
    return catalog


def _unique_from_sets(sets: list[set[str]]) -> str:
    nonempty = [values for values in sets if values]
    if not nonempty:
        return ""
    common = set(nonempty[0])
    for values in nonempty[1:]:
        common.intersection_update(values)
    if len(common) == 1:
        return next(iter(common))
    union = set().union(*nonempty)
    return next(iter(union)) if len(union) == 1 else ""


def _canonical_client_code(codes: set[str]) -> str:
    """Escolhe o menor codigo tratado para a mesma identidade."""
    if not codes:
        return ""
    numeric = [clean_text(code) for code in codes if re.fullmatch(r"\d+", clean_text(code))]
    if numeric:
        return min(numeric, key=int)
    return sorted((clean_text(code) for code in codes), key=str.casefold)[0]


def map_appointment_type(type_value: Any, item_value: Any, agenda_value: Any) -> str:
    """Converte/inferre o Tipo do modelo de Agendamento.

    Quando a extracao possui Tipo, ele prevalece. Nas extracoes antigas, em que essa
    coluna nao existe, Item preenchido continua significando Servico. Agenda contendo
    AVALIACAO com Item vazio e tratada como Avaliacao.
    """
    key = normalize_header(type_value)
    aliases = {
        "servico": "Serviço",
        "servicos": "Serviço",
        "sessao": "Serviço",
        "sessao servico": "Serviço",
        "avaliacao": "Avaliação",
        "consulta": "Consulta",
        "retorno": "Retorno",
    }
    if key in aliases:
        return aliases[key]

    if clean_text(item_value):
        return "Serviço"

    agenda_key = normalize_header(agenda_value)
    if "avaliacao" in agenda_key.split() or agenda_key == "avaliacao":
        return "Avaliação"
    if "consulta" in agenda_key.split() or agenda_key == "consulta":
        return "Consulta"
    if "retorno" in agenda_key.split() or agenda_key == "retorno":
        return "Retorno"
    return ""


def match_client(
    catalog: ClientCatalog,
    *,
    source_code: Any,
    name: Any,
    phone: Any,
    cpf: Any,
) -> tuple[str, str]:
    name_key = normalize_header(name)
    phone_keys = normalize_phone_keys(phone)
    cpf_key = normalize_cpf(cpf)
    candidates_name = set(catalog.by_name.get(name_key, set())) if name_key else set()
    candidates_phone: set[str] = set()
    for key in phone_keys:
        candidates_phone.update(catalog.by_phone.get(key, set()))
    candidates_cpf = set(catalog.by_cpf.get(cpf_key, set())) if cpf_key else set()

    old_code = clean_text(source_code)
    origin_candidates = set(catalog.by_origin.get(old_code, set())) if old_code else set()

    # Se a extracao ja traz um codigo que tambem e codigo tratado, ele e aceito.
    if old_code and old_code in catalog.codes:
        return old_code, "CODIGO_TRATADO"

    # Quando um cliente complementar anterior preservou Codigo Origem, esse e o
    # vinculo mais direto para reutilizar o mesmo cadastro em Agendamentos.
    if origin_candidates:
        selected = _canonical_client_code(origin_candidates)
        return selected, "CODIGO_ORIGEM" if len(origin_candidates) == 1 else "CODIGO_ORIGEM_DUPLICADO_CANONICO"

    # CPF valido e a identidade mais forte. Duplicidades historicas do mesmo CPF
    # usam um codigo canonico, evitando rejeitar todos os agendamentos da pessoa.
    if candidates_cpf:
        selected = _canonical_client_code(candidates_cpf)
        return selected, "CPF" if len(candidates_cpf) == 1 else "CPF_DUPLICADO_CANONICO"

    # Celular + nome em conjunto pode resolver casos em que cada chave, isolada,
    # aparece em mais de um cadastro.
    if candidates_phone and candidates_name:
        common = candidates_phone & candidates_name
        if common:
            selected = _canonical_client_code(common)
            return selected, "CELULAR+NOME" if len(common) == 1 else "CELULAR+NOME_DUPLICADO_CANONICO"

    if len(candidates_phone) == 1:
        return next(iter(candidates_phone)), "CELULAR"
    if len(candidates_name) == 1:
        return next(iter(candidates_name)), "NOME"

    if candidates_phone or candidates_name:
        return "", "AMBIGUO"
    return "", "NAO_LOCALIZADO"





@dataclass
class GeneratedClient:
    code: str
    name: str
    cpf: str
    phone: str
    origin_code: str
    identity_key: str


@dataclass
class GeneratedClientRegistry:
    used_codes: set[int]
    by_identity: dict[str, GeneratedClient] = field(default_factory=dict)

    @classmethod
    def from_catalog(cls, catalog: ClientCatalog) -> "GeneratedClientRegistry":
        used = {
            int(code) for code in catalog.codes
            if re.fullmatch(r"\d{1,6}", clean_text(code)) and 100000 <= int(code) <= 999999
        }
        return cls(used_codes=used)

    def _next_code(self) -> str:
        code = 100000
        while code in self.used_codes and code <= 999999:
            code += 1
        if code > 999999:
            raise ValueError("Nao ha codigo de cliente disponivel entre 100000 e 999999.")
        self.used_codes.add(code)
        return str(code)

    def get_or_create(self, *, name: Any, cpf: Any, phone: Any, source_code: Any) -> GeneratedClient | None:
        clean_name = clean_text(name)
        cpf_key = normalize_cpf(cpf)
        origin = clean_text(source_code)
        phones = sorted(normalize_phone_keys(phone), key=lambda value: (-len(value), value))

        # O modelo de cliente exige nome. Para criar automaticamente tambem
        # exigimos pelo menos uma segunda chave de identidade confiavel.
        if not clean_name or not (cpf_key or origin or phones):
            return None
        if cpf_key:
            identity = f"CPF:{cpf_key}"
        elif origin:
            identity = f"ORIGEM:{origin}"
        else:
            identity = f"CEL:{phones[0]}|NOME:{normalize_header(clean_name)}"

        current = self.by_identity.get(identity)
        if current is not None:
            return current
        current = GeneratedClient(
            code=self._next_code(),
            name=clean_name,
            cpf=cpf_key,
            phone=clean_text(phone),
            origin_code=origin,
            identity_key=identity,
        )
        self.by_identity[identity] = current
        return current


def _format_client_phone(value: Any) -> str:
    digits = digits_only(value)
    if digits.startswith("55") and len(digits) in {12, 13}:
        digits = digits[2:]
    if len(digits) == 10:
        digits = digits[:2] + "9" + digits[2:]
    if len(digits) == 11:
        return f"({digits[:2]}){digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}){digits[2:6]}-{digits[6:]}"
    return ""


def _format_client_cpf(value: Any) -> str:
    digits = digits_only(value)
    if len(digits) in {9, 10}:
        digits = digits.zfill(11)
    if len(digits) != 11:
        return ""
    return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"


def _client_model_indexes(headers: Sequence[Any]) -> dict[str, int | None]:
    normalized = [normalize_header(header) for header in headers]

    def first(predicate: Any) -> int | None:
        return next((index for index, header in enumerate(normalized) if predicate(header, index)), None)

    return {
        "code": first(lambda h, i: "codigo" in h.split() and ("cliente" in h.split() or "max" in h.split() or i == 0)),
        "name": first(lambda h, i: "nome" in h.split() and "mae" not in h.split() and "pai" not in h.split()),
        "phone": first(lambda h, i: any(token in h.split() for token in ("fone", "telefone")) and "ddi" not in h.split()),
        "mobile": first(lambda h, i: "celular" in h.split() and "ddi" not in h.split() and "2" not in h.split()),
        "cpf": first(lambda h, i: "cpf" in h.split()),
        "observation": first(lambda h, i: "observacao" in h.split()),
        "status": first(lambda h, i: h.startswith("status")),
        "origin_type": first(lambda h, i: "tipo" in h.split() and "origem" in h.split()),
        "origin_code": first(lambda h, i: "codigo" in h.split() and "origem" in h.split()),
        "ddi_phone": first(lambda h, i: "ddi" in h.split() and any(token in h.split() for token in ("fone", "telefone"))),
        "ddi_mobile": first(lambda h, i: "ddi" in h.split() and "celular" in h.split() and "2" not in h.split()),
    }


def normalize_complementary_client_file(path: Path) -> int:
    """Remove Codigo Origem exportado anteriormente e garante constantes."""
    if not path.is_file() or path.suffix.lower() not in {".csv", ".txt"}:
        return 0
    raw = path.read_bytes()
    text = None
    encoding = "utf-8-sig"
    for candidate in ("utf-8-sig", "cp1252", "latin1"):
        try:
            text = raw.decode(candidate)
            encoding = candidate
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return 0
    first = next((line for line in text.splitlines() if line.strip()), "")
    delimiter = ";" if ";" in first else ","
    parsed = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    if not parsed:
        return 0
    headers = parsed[0]
    indexes = _client_model_indexes(headers)
    cleared = 0
    for row in parsed[1:]:
        if len(row) < len(headers):
            row.extend([""] * (len(headers) - len(row)))
        origin_i = indexes.get("origin_code")
        if origin_i is not None and clean_text(row[origin_i]):
            row[origin_i] = ""
            cleared += 1
        status_i = indexes.get("status")
        if status_i is not None and not clean_text(row[status_i]):
            row[status_i] = "Leads"
        type_i = indexes.get("origin_type")
        if type_i is not None and not clean_text(row[type_i]):
            row[type_i] = "Parcerias"
    temp = path.with_name(path.name + ".tmp")
    with temp.open("w", encoding=encoding, newline="") as handle:
        csv.writer(handle, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL).writerows(parsed)
    temp.replace(path)
    return cleared


def persist_generated_clients(
    target: Path,
    model_path: Path,
    generated: Sequence[GeneratedClient],
    import_date: str,
) -> int:
    if not generated:
        return 0
    model_headers, model_delimiter, model_encoding = read_csv_template(model_path)
    indexes = _client_model_indexes(model_headers)
    if indexes["code"] is None or indexes["name"] is None:
        raise ValueError("modeloImportacaoCliente precisa conter colunas de Codigo e Nome.")

    existing_rows: list[list[str]] = []
    delimiter = model_delimiter
    output_encoding = "cp1252" if model_encoding in ("cp1252", "latin1") else "utf-8-sig"
    if target.exists():
        if target.suffix.lower() not in {".csv", ".txt"}:
            raise ValueError(f"Clientes complementares precisam ser gravados em CSV: {target.name}")
        rows = read_csv_rows(target)
        if rows:
            existing_headers = [clean_text(value) for value in rows[0][1]]
            if existing_headers != [clean_text(value) for value in model_headers]:
                raise ValueError("Cabecalho de planilhaTratadaClienteVendaPlano diverge do modeloImportacaoCliente.")
            existing_rows = [[clean_text(value) for value in row] for _, row in rows[1:]]
            for existing in existing_rows:
                if len(existing) < len(model_headers):
                    existing.extend([""] * (len(model_headers) - len(existing)))
                if indexes["origin_code"] is not None:
                    existing[indexes["origin_code"]] = ""
                if indexes["status"] is not None and not clean_text(existing[indexes["status"]]):
                    existing[indexes["status"]] = "Leads"
                if indexes["origin_type"] is not None and not clean_text(existing[indexes["origin_type"]]):
                    existing[indexes["origin_type"]] = "Parcerias"
        raw = target.read_bytes()
        for encoding in ("utf-8-sig", "cp1252", "latin1"):
            try:
                decoded = raw.decode(encoding)
                output_encoding = encoding
                break
            except UnicodeDecodeError:
                continue
        first_line = next((line for line in decoded.splitlines() if line.strip()), "") if decoded else ""
        delimiter = ";" if ";" in first_line else model_delimiter

    new_rows: list[list[str]] = []
    for client in generated:
        row = ["" for _ in model_headers]
        row[indexes["code"]] = client.code
        row[indexes["name"]] = client.name
        phone_out = _format_client_phone(client.phone)
        if indexes["mobile"] is not None and phone_out:
            row[indexes["mobile"]] = phone_out
        elif indexes["phone"] is not None and phone_out:
            row[indexes["phone"]] = phone_out
        if indexes["cpf"] is not None:
            row[indexes["cpf"]] = _format_client_cpf(client.cpf)
        if indexes["observation"] is not None:
            row[indexes["observation"]] = f"Cliente gerado por Agendamentos | Importação {import_date}"
        if indexes["status"] is not None:
            row[indexes["status"]] = "Leads"
        if indexes["origin_type"] is not None:
            row[indexes["origin_type"]] = "Parcerias"
        if indexes["ddi_mobile"] is not None and phone_out:
            row[indexes["ddi_mobile"]] = "55"
        if indexes["ddi_phone"] is not None and phone_out and indexes["ddi_mobile"] is None:
            row[indexes["ddi_phone"]] = "55"
        new_rows.append([clean_text(value) for value in row])

    temporary = target.with_name(target.name + ".tmp")
    try:
        with temporary.open("w", encoding=output_encoding, newline="") as handle:
            writer = csv.writer(handle, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
            writer.writerow(model_headers)
            writer.writerows(existing_rows)
            writer.writerows(new_rows)
        temporary.replace(target)
    finally:
        if temporary.exists():
            temporary.unlink()
    return len(new_rows)

def match_service(catalog: DeParaCatalog, value: Any) -> tuple[str, str, str]:
    key = service_alias(value)
    direct = catalog.service_codes.get(key, set()) if key else set()
    if len(direct) == 1:
        code = next(iter(direct))
        durations = catalog.duration_by_code.get(code, set())
        duration = next(iter(durations)) if len(durations) == 1 else ""
        return code, duration, "EXATO"
    if len(direct) > 1:
        return "", "", "AMBIGUO"

    if not key:
        return "", "", "NAO_LOCALIZADO"
    scored: list[tuple[float, str, str]] = []
    for candidate_key, codes in catalog.service_codes.items():
        score = difflib.SequenceMatcher(None, key, candidate_key).ratio()
        for code in codes:
            scored.append((score, candidate_key, code))
    if not scored:
        return "", "", "NAO_LOCALIZADO"
    best_score = max(score for score, _, _ in scored)
    if best_score < 0.75:
        return "", "", f"NAO_LOCALIZADO_APROXIMACAO {best_score:.4f}"
    best = [item for item in scored if abs(item[0] - best_score) < 1e-12]
    codes = {code for _, _, code in best}
    if len(codes) != 1:
        return "", "", "AMBIGUO_APROXIMADO"
    code = next(iter(codes))
    durations = catalog.duration_by_code.get(code, set())
    duration = next(iter(durations)) if len(durations) == 1 else ""
    return code, duration, f"APROXIMADO {best_score:.4f}"


def match_room(catalog: DeParaCatalog, value: Any) -> tuple[str, str]:
    keys = [normalize_header(value), room_alias(value)]
    candidates: set[str] = set()
    for key in keys:
        if key:
            candidates.update(catalog.room_codes.get(key, set()))
    if len(candidates) == 1:
        return next(iter(candidates)), "EXATO/ALIAS"
    if len(candidates) > 1:
        return "", "AMBIGUO"

    target = room_alias(value) or normalize_header(value)
    if not target:
        return "", "NAO_LOCALIZADO"
    scored: list[tuple[float, str]] = []
    for candidate_key, codes in catalog.room_codes.items():
        score = difflib.SequenceMatcher(None, target, candidate_key).ratio()
        for code in codes:
            scored.append((score, code))
    if not scored:
        return "", "NAO_LOCALIZADO"
    best_score = max(score for score, _ in scored)
    best_codes = {code for score, code in scored if abs(score - best_score) < 1e-12}
    if len(best_codes) == 1 and best_score >= 0.80:
        return next(iter(best_codes)), f"APROXIMADO {best_score:.4f}"
    return "", "NAO_LOCALIZADO"


def build_observation(
    source_headers: Sequence[Any],
    values: Sequence[Any],
    used_indexes: set[int],
    import_date: str,
) -> str:
    parts: list[str] = []
    for index, raw in enumerate(values):
        if index in used_indexes:
            continue
        value = clean_text(raw)
        if not value:
            continue
        label = clean_text(source_headers[index] if index < len(source_headers) else "")
        if not label:
            continue
        value = re.sub(
            r"(?i)(?:^|\s*\|\s*)importa[cç][aã]o\s+\d{2}/\d{2}/\d{4}(?=\s*\||$)",
            " ",
            value,
        ).strip(" |")
        if value:
            parts.append(f"{label}: {clean_text(value)}")
    parts.append(f"Importação {import_date}")
    return clean_text(" | ".join(parts))


# ---------------------------------------------------------------------------
# Tratamento
# ---------------------------------------------------------------------------

def process_agendamento(
    input_xlsx: Path,
    template_csv: Path,
    client_files: Sequence[Path],
    depara_file: Path,
    output_csv: Path,
    report_csv: Path,
    client_model_file: Path,
    complementary_client_file: Path,
) -> tuple[int, int, int]:
    headers, delimiter, template_encoding = read_csv_template(template_csv)
    if len(headers) != 14:
        raise ValueError(f"O modelo de Agendamento deve ter 14 colunas; possui {len(headers)}.")

    clients = build_client_catalog(client_files)
    generated_registry = GeneratedClientRegistry.from_catalog(clients)
    depara = build_depara_catalog(depara_file)
    import_date = date.today().strftime("%d/%m/%Y")
    rows_out: list[list[str]] = []
    rejected: list[tuple[int, list[str]]] = []

    with XlsxReader(input_xlsx) as xlsx:
        required_groups = [
            ("Cliente", "Nome Cliente"),
            ("Item", "Servico", "Serviço"),
            ("Dt Agendamento", "Data Agendamento", "Data"),
            ("Hora Inicio", "Hora Inicial", "Hora Início"),
            ("Status",),
        ]
        best: tuple[int, str, int, list[Any]] | None = None
        for sheet_name in xlsx.sheets:
            for row_number, candidate in xlsx.iter_rows(sheet_name, min_row=1, max_col=120, max_row=50):
                score = sum(find_header_index(candidate, *group) is not None for group in required_groups)
                if best is None or score > best[0]:
                    best = (score, sheet_name, row_number, candidate)
        if best is None or best[0] < 5:
            raise ValueError("Nao consegui localizar automaticamente o cabecalho da extracao de Agendamento.")

        _, source_sheet, header_line, source_headers = best
        index_client_name = find_header_index(source_headers, "Cliente", "Nome Cliente")
        index_source_client_code = find_header_index(
            source_headers,
            "Código do Cliente", "Codigo do Cliente", "Código Cliente", "Codigo Cliente", "Cliente ID", "ID Cliente",
        )
        index_phone = find_header_index(source_headers, "Telefone", "Celular", "Fone")
        index_cpf = find_header_index(source_headers, "CPF", "Cliente CPF")
        index_agenda = find_header_index(source_headers, "Agenda", "Sala")
        index_item = find_header_index(source_headers, "Item", "Servico", "Serviço")
        index_duration = find_header_index(source_headers, "Tempo", "Duracao", "Duração", "Tempo Atendimento")
        index_type = find_header_index(source_headers, "Tipo de Agendamento", "Tipo Agendamento", "Tipo Atendimento", "Tipo")
        index_date = find_header_index(source_headers, "Dt Agendamento", "Data Agendamento", "Data")
        index_time = find_header_index(source_headers, "Hora Inicio", "Hora Inicial", "Hora Início")
        index_status = find_header_index(source_headers, "Status")
        index_executor = find_header_index(source_headers, "Executado por", "Profissional", "Executor")

        mandatory_source = [
            ("Cliente", index_client_name),
            ("Item/Servico", index_item),
            ("Data", index_date),
            ("Hora", index_time),
            ("Status", index_status),
        ]
        missing_source = [label for label, index in mandatory_source if index is None]
        if missing_source:
            raise ValueError(
                f"Na aba {source_sheet}, cabecalho na linha {header_line}, faltam: " + ", ".join(missing_source)
            )

        indexes = [
            index_client_name, index_source_client_code, index_phone, index_cpf,
            index_agenda, index_item, index_duration, index_type,
            index_date, index_time, index_status, index_executor,
        ]
        max_needed = max(index for index in indexes if index is not None) + 1
        used_indexes = {index for index in indexes if index is not None}

        for source_line, values in xlsx.iter_rows(
            source_sheet,
            min_row=header_line + 1,
            max_col=max(120, max_needed),
        ):
            def get(index: int | None) -> Any:
                return values[index] if index is not None and index < len(values) else None

            client_name = get(index_client_name)
            source_client_code = get(index_source_client_code)
            phone = get(index_phone)
            cpf = get(index_cpf)
            agenda_raw = get(index_agenda)
            item_raw = get(index_item)
            duration_raw = get(index_duration)
            type_raw = get(index_type)
            appointment_date = get(index_date)
            appointment_time = get(index_time)
            status_raw = get(index_status)
            executor = get(index_executor)

            if all(
                value in (None, "")
                for value in (
                    client_name, source_client_code, phone, cpf, agenda_raw, item_raw,
                    duration_raw, type_raw, appointment_date, appointment_time, status_raw, executor,
                )
            ):
                continue

            client_code, client_method = match_client(
                clients,
                source_code=source_client_code,
                name=client_name,
                phone=phone,
                cpf=cpf,
            )
            if not client_code:
                generated = generated_registry.get_or_create(
                    name=client_name,
                    cpf=cpf,
                    phone=phone,
                    source_code=source_client_code,
                )
                if generated is not None:
                    client_code = generated.code
                    client_method = "CLIENTE_COMPLEMENTAR_GERADO"
                    add_client_to_catalog(
                        clients,
                        code=generated.code,
                        name=generated.name,
                        phone=generated.phone,
                        cpf=generated.cpf,
                        origin_code=generated.origin_code,
                    )

            appointment_type = map_appointment_type(type_raw, item_raw, agenda_raw)

            if clean_text(agenda_raw):
                room_code, room_method = match_room(depara, agenda_raw)
            else:
                room_code, room_method = match_room(depara, "Laser 1")
                room_method = "PADRAO_LASER_1" if room_code else "PADRAO_LASER_1_NAO_LOCALIZADO"

            if appointment_type == "Serviço":
                service_code, depara_duration, service_method = match_service(depara, item_raw)
            else:
                service_code, depara_duration, service_method = "", "", "NAO_APLICAVEL"

            duration = clean_text(duration_raw) or clean_text(depara_duration)
            status = map_status(status_raw)
            observation = build_observation(source_headers, values, used_indexes, import_date)

            row = [
                client_code,
                clean_text(executor),
                room_code,
                format_date(appointment_date, xlsx.date1904),
                format_time(appointment_time, xlsx.date1904),
                duration,
                appointment_type,
                status,
                service_code,
                "", "", "",
                observation,
                "",
            ]
            row = [clean_text(value) for value in row]

            missing: list[str] = []
            if not client_code:
                detail = "ambíguo" if client_method == "AMBIGUO" else "não localizado e sem dados suficientes para cliente complementar"
                missing.append(f"Código do Cliente: cliente {detail}")

            if not room_code:
                if clean_text(agenda_raw):
                    missing.append(f"Sala: '{clean_text(agenda_raw)}' não localizada/ambígua no DE-PARA ({room_method})")
                else:
                    missing.append("Sala: campo vazio e sala padrão Laser 1 não localizada no DE-PARA")

            if not appointment_type:
                missing.append("Tipo: não identificado como Serviço/Avaliação/Consulta/Retorno")
            elif appointment_type == "Serviço" and not service_code:
                missing.append(f"Código do Serviço: '{clean_text(item_raw)}' não localizado/ambíguo no DE-PARA ({service_method})")

            if not row[3]:
                missing.append("Data (dd/mm/aaaa) / Obrigatório")
            if not row[4]:
                missing.append("Hora Inicial (hh:mm) / Obrigatório")
            if not row[5]:
                missing.append("Tempo de Atendimento em Minutos / Obrigatório")
            else:
                try:
                    if float(row[5].replace(",", ".")) < 1:
                        missing.append("Tempo de Atendimento em Minutos / mínimo 1")
                except ValueError:
                    missing.append("Tempo de Atendimento em Minutos / inválido")
            if not row[7]:
                missing.append("Status / Obrigatório")

            # Profissional e opcional: a importacao utiliza Sala/Agenda.
            if missing:
                rejected.append((source_line, missing))
                continue
            rows_out.append(row)

    generated_clients = list(generated_registry.by_identity.values())
    generated_count = persist_generated_clients(
        complementary_client_file,
        client_model_file,
        generated_clients,
        import_date,
    ) if generated_clients else 0

    output_encoding = "cp1252" if template_encoding in ("cp1252", "latin1") else "utf-8-sig"
    if not output_csv.parent.is_dir():
        raise FileNotFoundError(f"Pasta de saida nao existe: {output_csv.parent}")
    temporary = output_csv.with_name(output_csv.name + ".tmp")
    try:
        with temporary.open("w", encoding=output_encoding, newline="") as handle:
            writer = csv.writer(handle, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
            writer.writerow(headers)
            writer.writerows(rows_out)
        temporary.replace(output_csv)
    finally:
        if temporary.exists():
            temporary.unlink()

    if rejected:
        if not report_csv.parent.is_dir():
            raise FileNotFoundError(f"Pasta de validacao nao existe: {report_csv.parent}")
        with report_csv.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.writer(handle, delimiter=";")
            writer.writerow(["Linha na planilha de extração", "Motivo da rejeição"])
            for source_line, reasons in rejected:
                for reason in reasons:
                    writer.writerow([source_line, clean_text(reason)])
    elif report_csv.exists():
        report_csv.unlink()

    return len(rows_out), len(rejected), generated_count




# ---------------------------------------------------------------------------
# Pastas e descoberta de arquivos
# ---------------------------------------------------------------------------

def ensure_project_dirs() -> None:
    if not INPUT_DIR.is_dir():
        raise FileNotFoundError(
            f"Pasta de entrada nao encontrada: {INPUT_DIR}. A raiz deve possuir 'entrada'."
        )
    if not OUTPUT_DIR.is_dir():
        raise FileNotFoundError(
            f"Pasta de saida nao encontrada na raiz {PROJECT_ROOT}. Use 'saida' ou 'saída'."
        )


def newest(paths: list[Path]) -> Path:
    return max(paths, key=lambda item: item.stat().st_mtime)


def find_newest(directory: Path, patterns: Sequence[str]) -> Path | None:
    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(
            path for path in directory.glob(pattern)
            if path.is_file() and not path.name.startswith("~$")
        )
    unique = {path.resolve(): path for path in candidates}
    return newest(list(unique.values())) if unique else None


def discover_agendamento_extraction(directory: Path) -> Path | None:
    """Encontra a extracao pelo cabecalho, sem depender do nome do arquivo."""
    required_groups = [
        ("Cliente", "Nome Cliente"),
        ("Item", "Servico", "Serviço"),
        ("Dt Agendamento", "Data Agendamento", "Data"),
        ("Hora Inicio", "Hora Inicial", "Hora Início"),
        ("Status",),
    ]
    matches: list[tuple[int, int, float, Path]] = []
    for path in directory.iterdir():
        if not path.is_file() or path.suffix.lower() != ".xlsx" or path.name.startswith("~$"):
            continue
        compact_stem = re.sub(r"[^a-z0-9]+", "", ascii_fold(path.stem).casefold())
        if compact_stem.startswith("modeloimportacao") or compact_stem.startswith("depara"):
            continue
        try:
            best_score = 0
            with XlsxReader(path) as xlsx:
                for sheet_name in xlsx.sheets:
                    for _, row in xlsx.iter_rows(sheet_name, min_row=1, max_col=120, max_row=50):
                        score = 0
                        for aliases in required_groups:
                            if find_header_index(row, *aliases) is not None:
                                score += 1
                        best_score = max(best_score, score)
            if best_score < 5:
                continue
        except Exception:
            continue
        name_hint = 1 if any(token in compact_stem for token in ("agendamento", "sessao", "sessoes")) else 0
        matches.append((best_score, name_hint, path.stat().st_mtime, path))
    return max(matches, key=lambda item: (item[0], item[1], item[2]))[3].resolve() if matches else None


def find_default_depara() -> Path:
    candidates = [
        path for path in INPUT_DIR.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xls", ".xlsx", ".csv"}
        and ("depara" in normalize_header(path.stem).replace(" ", "") or "de para" in normalize_header(path.stem))
    ]
    if not candidates:
        # aceita nomes como DE_PARA..., pois normalize_header vira "de para ..."
        candidates = [
            path for path in INPUT_DIR.iterdir()
            if path.is_file()
            and path.suffix.lower() in {".xls", ".xlsx", ".csv"}
            and normalize_header(path.stem).startswith("de para")
        ]
    if not candidates:
        raise FileNotFoundError(f"Nao encontrei arquivo DE-PARA em {INPUT_DIR}")

    def score(path: Path) -> tuple[int, float]:
        stem = normalize_header(path.stem)
        generic = 3 if stem in {"de para", "depara"} else 0
        service_only = -2 if "servico" in stem and not any(token in stem for token in ("sala", "agenda")) else 0
        return generic + service_only, path.stat().st_mtime

    return max(candidates, key=score)


def _client_stem_compact(path: Path) -> str:
    return re.sub(r"[^a-z0-9]+", "", ascii_fold(path.stem).lower())


def find_default_client_files() -> list[Path]:
    """Prioriza planilhaTratadaCliente e acrescenta VendaPlano apenas como complemento."""
    candidates = [
        path for path in OUTPUT_DIR.iterdir()
        if path.is_file()
        and not path.name.startswith("~$")
        and path.suffix.lower() in {".csv", ".xlsx", ".xls"}
        and _client_stem_compact(path).startswith("planilhatratadacliente")
    ]

    primary = [
        path for path in candidates
        if not _client_stem_compact(path).startswith("planilhatratadaclientevendaplano")
        and not any(
            marker in _client_stem_compact(path)
            for marker in ("contaspagar", "contasreceber", "fornecedor", "validacao", "rejeitada")
        )
    ]
    if not primary:
        raise FileNotFoundError(
            f"Nao encontrei a planilha principal planilhaTratadaCliente em {OUTPUT_DIR}. "
            "Execute scriptCliente.py antes."
        )

    def primary_score(path: Path) -> tuple[int, float]:
        compact = _client_stem_compact(path)
        exact = 10 if compact == "planilhatratadacliente" else 0
        return exact, path.stat().st_mtime

    primary_file = max(primary, key=primary_score)

    complementary = [
        path for path in candidates
        if _client_stem_compact(path).startswith("planilhatratadaclientevendaplano")
    ]
    result = [primary_file]
    if complementary:
        result.append(newest(complementary))
    return result



def find_default_client_model() -> Path:
    model = find_newest(INPUT_DIR, ("modeloImportacaoCliente.csv", "modeloImportacaoCliente*.csv"))
    if model is None:
        raise FileNotFoundError(f"Nao encontrei modeloImportacaoCliente*.csv em {INPUT_DIR}")
    return model


def complementary_client_target(client_files: Sequence[Path]) -> Path:
    for path in client_files[1:]:
        if _client_stem_compact(path).startswith("planilhatratadaclientevendaplano") and path.suffix.lower() in {".csv", ".txt"}:
            return path.resolve()
    return (OUTPUT_DIR / "planilhaTratadaClienteVendaPlano.csv").resolve()

def find_default_files() -> tuple[Path, Path, list[Path], Path]:
    ensure_project_dirs()
    input_file = discover_agendamento_extraction(INPUT_DIR)
    model_file = find_newest(INPUT_DIR, ("modeloImportacaoAgendamento.csv", "modeloImportacaoAgendamento*.csv"))
    if input_file is None:
        raise FileNotFoundError(f"Nao encontrei extracao de Agendamentos/Sessoes pelos cabecalhos em {INPUT_DIR}")
    if model_file is None:
        raise FileNotFoundError(f"Nao encontrei modeloImportacaoAgendamento*.csv em {INPUT_DIR}")
    return input_file, model_file, find_default_client_files(), find_default_depara()


def resolve_path(value: Path | None, default_dir: Path) -> Path | None:
    if value is None:
        return None
    if value.is_absolute():
        return value.resolve()
    if value.parent == Path("."):
        return (default_dir / value).resolve()
    return (PROJECT_ROOT / value).resolve()


def output_path(value: Path | None, default_name: str) -> Path:
    resolved = resolve_path(value, OUTPUT_DIR)
    return resolved if resolved is not None else (OUTPUT_DIR / default_name).resolve()


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

    parser = argparse.ArgumentParser(description="Tratamento automatico de Agendamentos - Laser Rosa")
    parser.add_argument("--entrada", type=Path, help="Extracao; relativo a entrada")
    parser.add_argument("--modelo", type=Path, help="Modelo; relativo a entrada")
    parser.add_argument("--clientes", type=Path, help="planilhaTratadaCliente; relativo a saida")
    parser.add_argument("--depara", type=Path, help="DE-PARA; relativo a entrada")
    parser.add_argument("--saida", type=Path, help="CSV final; relativo a saida")
    args = parser.parse_args()

    try:
        ensure_project_dirs()
        if all(value is None for value in (args.entrada, args.modelo, args.clientes, args.depara)):
            input_file, model_file, client_files, depara_file = find_default_files()
        else:
            defaults = find_default_files()
            input_file = resolve_path(args.entrada, INPUT_DIR) or defaults[0]
            model_file = resolve_path(args.modelo, INPUT_DIR) or defaults[1]
            explicit_client = resolve_path(args.clientes, OUTPUT_DIR)
            if explicit_client is not None:
                client_files = [explicit_client]
                for complementary in find_default_client_files()[1:]:
                    if complementary.resolve() != explicit_client.resolve():
                        client_files.append(complementary)
            else:
                client_files = defaults[2]
            depara_file = resolve_path(args.depara, INPUT_DIR) or defaults[3]

        output_file = output_path(args.saida, "planilhaTratadaAgendamento.csv")
        report_file = output_file.with_name(output_file.stem + "_linhas_rejeitadas.csv")
        client_model_file = find_default_client_model()
        complementary_file = complementary_client_target(client_files)
        if complementary_file.is_file():
            cleared = normalize_complementary_client_file(complementary_file)
            if cleared:
                print(f"Clientes complementares: {cleared} Codigo Origem removido(s).")

        for label, path in (("Extracao", input_file), ("Modelo", model_file), ("DE-PARA", depara_file), ("Modelo Clientes", client_model_file)):
            if not path.is_file():
                raise FileNotFoundError(f"{label} nao encontrado: {path}")
            print(f"{label}: {path}")
        for index, client_path in enumerate(client_files):
            if not client_path.is_file():
                raise FileNotFoundError(f"Clientes nao encontrado: {client_path}")
            label = "Clientes" if index == 0 else "Clientes complementares"
            print(f"{label}: {client_path}")
        print(f"Saida: {output_file}")

        accepted, rejected, generated = process_agendamento(
            input_file,
            model_file,
            client_files,
            depara_file,
            output_file,
            report_file,
            client_model_file,
            complementary_file,
        )
    except Exception as exc:
        print(f"ERRO: {exc}", file=sys.stderr)
        return 1

    if generated:
        print(f"OK: {generated} cliente(s) complementar(es) gerado(s) em {complementary_file.name}")
        print("Ordem de importacao: clientes complementares primeiro; depois Agendamentos.")
    print(f"OK: {accepted} linhas exportadas para {output_file}")
    if rejected:
        print(f"Validacao: {rejected} linhas rejeitadas; detalhes em {report_file}")
    else:
        print("Validacao: nenhuma linha rejeitada. Nenhum relatorio de rejeicoes foi criado.")
    return 0



if __name__ == "__main__":
    raise SystemExit(main())
