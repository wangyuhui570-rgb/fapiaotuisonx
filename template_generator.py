import csv
import os
import re
import zipfile
from datetime import datetime
import xml.etree.ElementTree as ET


SUPPORTED_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "gbk")
SUPPORTED_TABULAR_EXTENSIONS = (".csv", ".xlsx", ".xlsm")
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
DEFAULT_FILENAME_TEMPLATE = "{{CSV文件名不带后缀}}_{{序号}}_申请模板"
XLSX_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def sanitize_filename(name):
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", str(name or "").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" ._")
    return cleaned or "未命名"


def is_supported_tabular_file(file_path):
    return os.path.isfile(file_path) and str(file_path).lower().endswith(SUPPORTED_TABULAR_EXTENSIONS)


def list_csv_files(csv_dir):
    if not os.path.isdir(csv_dir):
        raise FileNotFoundError(f"CSV 目录不存在: {csv_dir}")

    csv_paths = [
        os.path.join(csv_dir, entry)
        for entry in sorted(os.listdir(csv_dir))
        if entry.lower().endswith(SUPPORTED_TABULAR_EXTENSIONS) and os.path.isfile(os.path.join(csv_dir, entry))
    ]
    if not csv_paths:
        raise FileNotFoundError(f"目录中没有找到 CSV 文件: {csv_dir}")
    return csv_paths


def _normalize_headers(fieldnames):
    headers = []
    seen = {}
    for index, header in enumerate(fieldnames or [], start=1):
        name = (header or "").replace("\ufeff", "").strip() or f"列{index}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        headers.append(name)
    return headers


def _column_index_from_ref(cell_ref):
    letters = "".join(char for char in str(cell_ref or "") if char.isalpha()).upper()
    if not letters:
        return -1
    value = 0
    for char in letters:
        value = value * 26 + (ord(char) - 64)
    return value - 1


def _read_xlsx_rows(csv_path):
    with zipfile.ZipFile(csv_path, "r") as archive:
        sheet_name = "xl/worksheets/sheet1.xml"
        if sheet_name not in archive.namelist():
            raise ValueError(f"Excel 文件里缺少工作表: {csv_path}")

        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in shared_root.findall("a:si", XLSX_NS):
                texts = [node.text or "" for node in item.findall(".//a:t", XLSX_NS)]
                shared_strings.append("".join(texts))

        sheet_root = ET.fromstring(archive.read(sheet_name))
        worksheet_rows = []
        for row_node in sheet_root.findall(".//a:sheetData/a:row", XLSX_NS):
            row_values = []
            for cell in row_node.findall("a:c", XLSX_NS):
                column_index = _column_index_from_ref(cell.get("r"))
                if column_index < 0:
                    continue
                while len(row_values) <= column_index:
                    row_values.append("")

                cell_type = cell.get("t")
                value = ""
                if cell_type == "inlineStr":
                    text_nodes = cell.findall("a:is//a:t", XLSX_NS)
                    value = "".join(node.text or "" for node in text_nodes)
                else:
                    value_node = cell.find("a:v", XLSX_NS)
                    if value_node is not None:
                        value = value_node.text or ""
                        if cell_type == "s":
                            try:
                                value = shared_strings[int(value)]
                            except (ValueError, IndexError):
                                pass
                row_values[column_index] = str(value).strip()
            if any(value for value in row_values):
                worksheet_rows.append(row_values)

    if not worksheet_rows:
        return [], [], "xlsx"

    headers = _normalize_headers(worksheet_rows[0])
    result_rows = []
    for raw_row in worksheet_rows[1:]:
        values = list(raw_row) + [""] * max(0, len(headers) - len(raw_row))
        row = {header: (values[idx].strip() if idx < len(values) else "") for idx, header in enumerate(headers)}
        if any(value for value in row.values()):
            result_rows.append(row)
    return headers, result_rows, "xlsx"


def read_csv_rows(csv_path):
    try:
        if zipfile.is_zipfile(csv_path):
            return _read_xlsx_rows(csv_path)
    except OSError:
        pass

    last_error = None

    for encoding in SUPPORTED_ENCODINGS:
        try:
            with open(csv_path, "r", encoding=encoding, newline="") as file_obj:
                reader = csv.reader(file_obj)
                rows = list(reader)
            if not rows:
                return [], [], encoding

            headers = _normalize_headers(rows[0])
            result_rows = []
            for raw_row in rows[1:]:
                values = list(raw_row) + [""] * max(0, len(headers) - len(raw_row))
                row = {header: (values[idx].strip() if idx < len(values) else "") for idx, header in enumerate(headers)}
                if any(value for value in row.values()):
                    result_rows.append(row)
            return headers, result_rows, encoding
        except UnicodeDecodeError as exc:
            last_error = exc
            continue

    raise UnicodeDecodeError(
        last_error.encoding if last_error else "unknown",
        last_error.object if last_error else b"",
        last_error.start if last_error else 0,
        last_error.end if last_error else 0,
        f"无法识别 CSV 编码: {csv_path}",
    )


def build_row_context(row, csv_path, row_index, global_index):
    file_name = os.path.basename(csv_path)
    file_stem, _ = os.path.splitext(file_name)
    context = {str(key).strip(): str(value).strip() for key, value in row.items()}
    context.update(
        {
            "序号": str(row_index),
            "全局序号": str(global_index),
            "CSV文件名": file_name,
            "CSV文件名不带后缀": file_stem,
            "CSV路径": os.path.abspath(csv_path),
            "当前日期": datetime.now().strftime("%Y-%m-%d"),
            "当前时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
    )
    return context


def render_template(template_text, context):
    missing_keys = []

    def replace(match):
        key = match.group(1).strip()
        if key in context:
            return str(context[key])
        missing_keys.append(key)
        return ""

    rendered = PLACEHOLDER_RE.sub(replace, template_text or "")
    return rendered, sorted(set(missing_keys))


def ensure_unique_path(file_path):
    if not os.path.exists(file_path):
        return file_path

    directory, file_name = os.path.split(file_path)
    stem, ext = os.path.splitext(file_name)
    counter = 2
    while True:
        candidate = os.path.join(directory, f"{stem}_{counter}{ext}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def generate_templates(csv_dir, output_dir, template_text, filename_template=None, logger=None):
    if not (template_text or "").strip():
        raise ValueError("申请文字模板不能为空。")

    csv_paths = list_csv_files(csv_dir)
    os.makedirs(output_dir, exist_ok=True)

    filename_pattern = (filename_template or DEFAULT_FILENAME_TEMPLATE).strip()
    if not filename_pattern:
        filename_pattern = DEFAULT_FILENAME_TEMPLATE

    summary = {
        "csv_files": len(csv_paths),
        "rows": 0,
        "generated": 0,
        "missing_placeholders": {},
        "output_files": [],
    }

    global_index = 0
    for csv_path in csv_paths:
        headers, rows, encoding = read_csv_rows(csv_path)
        if logger:
            logger(f"已读取 CSV：{os.path.basename(csv_path)}，编码 {encoding}，数据 {len(rows)} 行。")
        if not rows:
            continue

        for row_index, row in enumerate(rows, start=1):
            global_index += 1
            summary["rows"] += 1
            context = build_row_context(row, csv_path, row_index, global_index)

            rendered_text, text_missing = render_template(template_text, context)
            rendered_name, name_missing = render_template(filename_pattern, context)

            if text_missing or name_missing:
                summary["missing_placeholders"].setdefault(os.path.basename(csv_path), set()).update(
                    text_missing + name_missing
                )

            safe_name = sanitize_filename(rendered_name)
            output_path = ensure_unique_path(os.path.join(output_dir, f"{safe_name}.txt"))
            with open(output_path, "w", encoding="utf-8") as file_obj:
                file_obj.write(rendered_text.rstrip())
                file_obj.write("\n")

            summary["generated"] += 1
            summary["output_files"].append(output_path)
            if logger:
                logger(f"已生成：{os.path.basename(output_path)}")

    summary["missing_placeholders"] = {
        key: sorted(value) for key, value in summary["missing_placeholders"].items()
    }
    return summary
