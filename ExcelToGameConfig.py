import ast
import hashlib
import json
import os
import struct
import sys
from datetime import datetime
from typing import Any, Dict, List

import openpyxl

TYPE_INT = "int"
TYPE_STRING = "string"
TYPE_FLOAT = "float"
TYPE_FLOAT_ARR = "floatArr"
TYPE_INT_ARR = "intArr"
TYPE_STRING_ARR = "stringArr"
TYPE_FLOAT_ARR2 = "floatArr2"
TYPE_INT_ARR2 = "intArr2"
TYPE_STRING_ARR2 = "stringArr2"
SUPPORTED_TYPES = {
    TYPE_INT,
    TYPE_STRING,
    TYPE_FLOAT,
    TYPE_FLOAT_ARR,
    TYPE_INT_ARR,
    TYPE_STRING_ARR,
    TYPE_FLOAT_ARR2,
    TYPE_INT_ARR2,
    TYPE_STRING_ARR2,
}
BACKUP_NAME_MARKER = "备份"


def _default_base_dir() -> str:
    return os.path.dirname(os.path.abspath(sys.argv[0]))


def _default_excel_dir(base_dir: str) -> str:
    return base_dir

def _default_output_dir(base_dir: str) -> str:
    return os.path.join(base_dir, "Config")


def _safe_mkdir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def _is_excel_file(name: str) -> bool:
    lower = name.lower()
    return not lower.startswith("~$") and (lower.endswith(".xlsx") or lower.endswith(".xlsm"))


def _is_backup_name(name: str) -> bool:
    return BACKUP_NAME_MARKER in name


def _parse_json_text(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        text = value
        if text == "":
            return ""
        text = text.strip()
        if not text:
            return value
        if (text.startswith("[") and text.endswith("]")) or (text.startswith("{") and text.endswith("}")):
            try:
                return json.loads(text)
            except json.JSONDecodeError:
                try:
                    return ast.literal_eval(text)
                except (ValueError, SyntaxError):
                    return value
        return value
    return value


def _cast_scalar(value: Any, typ: str) -> Any:
    if value is None:
        return None
    if typ == TYPE_STRING:
        if isinstance(value, str):
            if value.strip() == '""':
                return ""
            return value
        return str(value)
    if typ == TYPE_INT:
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, int):
            return value
        if isinstance(value, float):
            return int(value)
        if isinstance(value, str):
            text = value.strip()
            if text == "":
                return None
            return int(float(text))
        return int(value)
    if typ == TYPE_FLOAT:
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            text = value.strip()
            if text == "":
                return None
            return float(text)
        return float(value)
    return value


def _ensure_array_value(value: Any, typ: str) -> List[Any]:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return list(value)
    return [_cast_scalar(value, typ)]


def _ensure_array2_value(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if not isinstance(value, (list, tuple)):
        raise ValueError(f"Expected 2D array value, got {type(value).__name__}: {value!r}")
    rows = list(value)
    for row in rows:
        if not isinstance(row, (list, tuple)):
            raise ValueError(f"Expected each 2D array row to be a list, got {type(row).__name__}: {row!r}")
    return [list(row) for row in rows]


def _cast_by_type(value: Any, typ: str) -> Any:
    if value is None:
        return None
    if typ in (TYPE_STRING, TYPE_INT, TYPE_FLOAT):
        return _cast_scalar(value, typ)
    parsed = _parse_json_text(value)
    if typ == TYPE_INT_ARR:
        return [int(float(x)) for x in _ensure_array_value(parsed, TYPE_INT)]
    if typ == TYPE_FLOAT_ARR:
        return [float(x) for x in _ensure_array_value(parsed, TYPE_FLOAT)]
    if typ == TYPE_STRING_ARR:
        return [str(x) for x in _ensure_array_value(parsed, TYPE_STRING)]
    if typ == TYPE_INT_ARR2:
        return [[int(float(x)) for x in row] for row in _ensure_array2_value(parsed)]
    if typ == TYPE_FLOAT_ARR2:
        return [[float(x) for x in row] for row in _ensure_array2_value(parsed)]
    if typ == TYPE_STRING_ARR2:
        return [[str(x) for x in row] for row in _ensure_array2_value(parsed)]
    return parsed


def _sheet_to_table(ws) -> Dict[str, Any]:
    cols: List[int] = []
    field_names: List[str] = []
    field_types: List[str] = []

    # Row 1: comment only, ignored by exporter.
    # Row 2: field names.
    # Row 3: field types.
    for c in range(1, ws.max_column + 1):
        key = ws.cell(row=2, column=c).value
        typ = ws.cell(row=3, column=c).value
        if key is None or str(key).strip() == "":
            continue
        if typ is None:
            continue
        typ_text = str(typ).strip()
        if typ_text == "" or typ_text.lower() == "null":
            # null type means comment-only/export-ignore column.
            continue
        if typ_text not in SUPPORTED_TYPES:
            raise ValueError(f"Unsupported type '{typ_text}' in sheet '{ws.title}', column '{key}'")
        cols.append(c)
        field_names.append(str(key).strip())
        field_types.append(typ_text)

    rows: List[Dict[str, Any]] = []
    for r in range(4, ws.max_row + 1):
        row_obj: Dict[str, Any] = {}
        empty = True
        for idx, key in enumerate(field_names):
            raw = ws.cell(row=r, column=cols[idx]).value
            raw_blank = raw is None or (isinstance(raw, str) and raw.strip() == "")
            if not raw_blank:
                empty = False
            if raw_blank:
                continue
            row_obj[key] = _cast_by_type(raw, field_types[idx])
        if not empty:
            rows.append(row_obj)

    return {
        "name": ws.title,
        "field_names": field_names,
        "field_types": field_types,
        "rows": rows,
    }


def build_tables(input_dir: str) -> List[Dict[str, Any]]:
    tables: List[Dict[str, Any]] = []
    existing_sheet_names = set()
    excel_files = [
        os.path.join(input_dir, n)
        for n in os.listdir(input_dir)
        if _is_excel_file(n) and not _is_backup_name(n)
    ]
    excel_files.sort()
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in: {input_dir}")

    for excel_path in excel_files:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        for sheet_name in wb.sheetnames:
            if _is_backup_name(sheet_name):
                # Backup sheets are intentionally excluded from game_config.bin.
                continue
            if sheet_name.lower().endswith("_map"):
                # Mapping tables are not exported into game_config.bin.
                continue
            if sheet_name in existing_sheet_names:
                raise ValueError(f"Duplicate sheet name detected across excels: {sheet_name}")
            existing_sheet_names.add(sheet_name)
            tables.append(_sheet_to_table(wb[sheet_name]))
    return tables


def _write_int32_be(f, value: int) -> None:
    f.write(struct.pack(">i", int(value)))


def _write_uint8(f, value: int) -> None:
    f.write(struct.pack(">B", int(value)))


def _write_float_be(f, value: float) -> None:
    f.write(struct.pack(">f", float(value)))


def _write_string(f, value: str) -> None:
    text = "" if value is None else str(value)
    data = text.encode("utf-8")
    _write_int32_be(f, len(data))
    f.write(data)


def _write_int_array(f, values: List[Any]) -> None:
    _write_int32_be(f, len(values))
    for value in values:
        _write_int32_be(f, int(value))


def _write_float_array(f, values: List[Any]) -> None:
    _write_int32_be(f, len(values))
    for value in values:
        _write_float_be(f, float(value))


def _write_string_array(f, values: List[Any]) -> None:
    _write_int32_be(f, len(values))
    for value in values:
        _write_string(f, str(value))


def _write_int_array2(f, values: List[List[Any]]) -> None:
    _write_int32_be(f, len(values))
    for row in values:
        _write_int_array(f, row)


def _write_float_array2(f, values: List[List[Any]]) -> None:
    _write_int32_be(f, len(values))
    for row in values:
        _write_float_array(f, row)


def _write_string_array2(f, values: List[List[Any]]) -> None:
    _write_int32_be(f, len(values))
    for row in values:
        _write_string_array(f, row)


def _write_value(f, typ: str, value: Any) -> None:
    if typ == TYPE_INT:
        _write_int32_be(f, 0 if value is None else int(value))
        return
    if typ == TYPE_FLOAT:
        _write_float_be(f, 0.0 if value is None else float(value))
        return
    if typ == TYPE_STRING:
        _write_string(f, "" if value is None else str(value))
        return
    if typ == TYPE_INT_ARR:
        _write_int_array(f, [] if value is None else value)
        return
    if typ == TYPE_FLOAT_ARR:
        _write_float_array(f, [] if value is None else value)
        return
    if typ == TYPE_STRING_ARR:
        _write_string_array(f, [] if value is None else value)
        return
    if typ == TYPE_INT_ARR2:
        _write_int_array2(f, [] if value is None else value)
        return
    if typ == TYPE_FLOAT_ARR2:
        _write_float_array2(f, [] if value is None else value)
        return
    if typ == TYPE_STRING_ARR2:
        _write_string_array2(f, [] if value is None else value)
        return
    raise ValueError(f"Unsupported field type while writing bin: {typ}")


def write_game_config_bin(output_dir: str, tables: List[Dict[str, Any]]) -> str:
    _safe_mkdir(output_dir)
    path = os.path.join(output_dir, "game_config.bin")
    with open(path, "wb") as f:
        _write_int32_be(f, len(tables))
        for table in tables:
            field_names = table["field_names"]
            field_types = table["field_types"]
            rows = table["rows"]

            _write_string(f, table["name"])
            _write_int32_be(f, len(rows))
            _write_int32_be(f, len(field_names))

            for i in range(len(field_names)):
                _write_string(f, field_names[i])
                _write_string(f, field_types[i])

            for row in rows:
                for i in range(len(field_names)):
                    field_name = field_names[i]
                    field_type = field_types[i]
                    _write_value(f, field_type, row.get(field_name))
    return path


def write_version_file(output_dir: str, version: str, config_path: str) -> str:
    with open(config_path, "rb") as f:
        raw = f.read()
    sha256 = hashlib.sha256(raw).hexdigest()
    payload = {"version": version, "hash": sha256}
    version_path = os.path.join(output_dir, "config_version.json")
    with open(version_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=4)
    return version_path


def _resolve_version() -> str:
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        return sys.argv[1].strip()
    return datetime.now().strftime("%Y.%m.%d.%H%M%S")


def main() -> int:
    base_dir = _default_base_dir()
    input_dir = _default_excel_dir(base_dir)
    output_dir = _default_output_dir(base_dir)
    version = _resolve_version()

    tables = build_tables(input_dir)
    config_path = write_game_config_bin(output_dir, tables)
    version_path = write_version_file(output_dir, version, config_path)

    print("Build game_config.bin success")
    print(f"InputDir: {input_dir}")
    print(f"OutputDir: {output_dir}")
    print(f"Version: {version}")
    print(f"Config: {config_path}")
    print(f"VersionFile: {version_path}")
    print(f"TableCount: {len(tables)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
