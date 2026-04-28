import os
from datetime import datetime

import template_generator


DEFAULT_STORE_NAME = "拼多多功率房"
DEFAULT_OWNED_STORE = "执行派"
DEFAULT_GOODS_NAME = "自行车配件"
DEFAULT_APPLICANT = "王宇辉"
DEFAULT_UNIT = "件"
DEFAULT_INVOICE_TYPE = "电子普票"
ENTERPRISE_TAX_FIELD = "企业税号"
OPTIONAL_ENTERPRISE_DETAIL_FIELDS = ("开户银行", "账号", "地址", "电话")
PERSONAL_TITLE_HINTS = ("个人", "自然人")
FIELD_ALIASES = {
    "订单号": ("订单号", "*订单编号", "订单编号"),
    "抬头类型": ("抬头类型", "*抬头类型"),
    "发票抬头": ("发票抬头", "*发票抬头"),
    "商品数量": ("商品数量", "数量"),
    "发票金额": ("发票金额", "*开票总金额"),
    "商品金额": ("商品金额",),
    "企业税号": ("企业税号", "购方税号"),
    "开户银行": ("开户银行", "开户行"),
    "账号": ("账号", "开户账号"),
    "地址": ("地址", "企业地址"),
    "电话": ("电话", "企业电话"),
}


FIELD_ALIASES["订单号"] = FIELD_ALIASES["订单号"] + ("订单号/iD", "订单号/ID")
FIELD_ALIASES["抬头类型"] = FIELD_ALIASES["抬头类型"] + ("发票类型",)
FIELD_ALIASES["发票抬头"] = FIELD_ALIASES["发票抬头"] + ("开票信息",)
FIELD_ALIASES["发票金额"] = FIELD_ALIASES["发票金额"] + ("发票金额",)


def format_today_text(now=None):
    current = now or datetime.now()
    return f"{current.year}-{current.month}-{current.day}"


def clean_value(value):
    return str(value or "").strip()


def get_row_value(row, field_name):
    for candidate in FIELD_ALIASES.get(field_name, (field_name,)):
        value = clean_value(row.get(candidate))
        if value:
            return value
    return ""


def get_row_or_setting_value(row, field_names, settings, setting_key, default_value=""):
    for field_name in field_names:
        value = clean_value(row.get(field_name))
        if value:
            return value
    return clean_value(settings.get(setting_key)) or clean_value(default_value)


def format_quantity(value):
    raw = clean_value(value)
    if not raw:
        return ""
    try:
        number = float(raw)
        if number.is_integer():
            return str(int(number))
        return f"{number:.2f}".rstrip("0").rstrip(".")
    except ValueError:
        return raw


def format_amount(value):
    raw = clean_value(value)
    if not raw:
        return ""
    try:
        return f"{float(raw):.2f}"
    except ValueError:
        return raw


def amount_to_float(value):
    raw = clean_value(value).replace(",", "")
    if not raw:
        return 0.0
    try:
        return float(raw)
    except ValueError:
        return 0.0


def is_taobao_rows(headers):
    header_set = set(headers or [])
    return "*订单编号" in header_set and "*开票总金额" in header_set


def normalize_source_rows(headers, rows):
    if is_taobao_rows(headers):
        return merge_taobao_orders(rows)
    return rows


def merge_taobao_orders(rows):
    merged_orders = []
    groups = {}

    for row in rows:
        order_no = get_row_value(row, "订单号")
        title_type = get_row_value(row, "抬头类型")
        invoice_title = get_row_value(row, "发票抬头")
        group_key = (order_no, title_type, invoice_title)

        current = groups.get(group_key)
        if current is None:
            current = dict(row)
            current["商品数量"] = ""
            current["发票金额"] = get_row_value(row, "发票金额")
            current["_quantity_total"] = 0.0
            current["_quantity_seen"] = False
            groups[group_key] = current
            merged_orders.append(current)

        quantity_raw = get_row_value(row, "商品数量")
        if quantity_raw:
            current["_quantity_total"] += amount_to_float(quantity_raw)
            current["_quantity_seen"] = True

        invoice_amount = get_row_value(row, "发票金额")
        if invoice_amount and not clean_value(current.get("发票金额")):
            current["发票金额"] = invoice_amount

        for field in (ENTERPRISE_TAX_FIELD,) + OPTIONAL_ENTERPRISE_DETAIL_FIELDS:
            if not get_row_value(current, field):
                value = get_row_value(row, field)
                if value:
                    current[field] = value

    normalized = []
    for row in merged_orders:
        quantity_total = row.pop("_quantity_total", 0.0)
        quantity_seen = row.pop("_quantity_seen", False)
        invoice_amount = amount_to_float(get_row_value(row, "发票金额"))
        if abs(invoice_amount) < 1e-9:
            continue
        row["发票金额"] = f"{invoice_amount:.2f}"
        if quantity_seen:
            row["商品数量"] = format_quantity(str(quantity_total))
        normalized.append(row)
    return normalized


def has_enterprise_invoice_details(row):
    return bool(get_row_value(row, ENTERPRISE_TAX_FIELD))


def is_enterprise_title(row):
    title_type = get_row_value(row, "抬头类型")
    if "企业" in title_type:
        return True
    if any(hint in title_type for hint in PERSONAL_TITLE_HINTS):
        return False
    return has_enterprise_invoice_details(row)


def build_output_text(row, settings):
    store_name = get_row_or_setting_value(row, ("店铺名", "店铺名称"), settings, "store_name", DEFAULT_STORE_NAME)
    owned_store = get_row_or_setting_value(row, ("所属公司", "所属店铺"), settings, "owned_store", DEFAULT_OWNED_STORE)
    goods_name = get_row_or_setting_value(row, ("货物名称", "商品名称"), settings, "goods_name", DEFAULT_GOODS_NAME)
    unit_name = get_row_or_setting_value(row, ("单位",), settings, "unit", DEFAULT_UNIT)
    applicant = get_row_or_setting_value(row, ("申请人",), settings, "applicant", DEFAULT_APPLICANT)
    lines = [
        f"订单号：{get_row_value(row, '订单号')}",
        f"店铺名：{store_name}",
        f"所属店铺：{owned_store}",
        f"抬头类型：{DEFAULT_INVOICE_TYPE}",
        f"发票抬头：{get_row_value(row, '发票抬头')}",
    ]

    if is_enterprise_title(row):
        lines.append(f"企业税号：{get_row_value(row, ENTERPRISE_TAX_FIELD)}")
        for field in OPTIONAL_ENTERPRISE_DETAIL_FIELDS:
            value = get_row_value(row, field)
            if value:
                lines.append(f"{field}：{value}")

    lines.extend(
        [
            (
                f"货物名称：{goods_name} "
                f"单位：{unit_name} "
                f"数量：{format_quantity(get_row_value(row, '商品数量'))} "
                f"发票金额：{format_amount(get_row_value(row, '发票金额'))}"
            ),
            f"申请人：{applicant}",
            f"日期：{format_today_text()}",
        ]
    )
    return "\n".join(lines).rstrip() + "\n"


def resolve_csv_paths(csv_dir, csv_paths=None):
    if csv_paths:
        normalized = []
        seen = set()
        for path in csv_paths:
            absolute = str(path or "").strip()
            if not absolute:
                continue
            absolute = os.path.abspath(absolute)
            if template_generator.is_supported_tabular_file(absolute) and absolute not in seen:
                normalized.append(absolute)
                seen.add(absolute)
        if not normalized:
            raise FileNotFoundError("未找到可用的 CSV 文件。")
        return normalized
    return template_generator.list_csv_files(csv_dir)


def generate_request_texts(csv_dir, settings, logger=None, csv_paths=None):
    resolved_csv_paths = resolve_csv_paths(csv_dir, csv_paths=csv_paths)
    summary = {
        "csv_files": len(resolved_csv_paths),
        "rows": 0,
        "generated": 0,
        "texts": [],
        "combined_text": "",
    }

    for csv_path in resolved_csv_paths:
        headers, rows, encoding = template_generator.read_csv_rows(csv_path)
        source_rows = normalize_source_rows(headers, rows)
        if logger:
            logger(
                f"已读取 CSV：{csv_path.split('\\')[-1]}，编码 {encoding}，字段 {len(headers)} 个，数据 {len(rows)} 行。"
            )
            if len(source_rows) != len(rows):
                logger(f"已按订单和货物合并淘宝明细：{len(rows)} 行 -> {len(source_rows)} 条申请模板。")

        for row in source_rows:
            summary["rows"] += 1
            content = build_output_text(row, settings).rstrip()
            summary["generated"] += 1
            summary["texts"].append(content)
            if logger:
                logger(f"已生成：订单 {get_row_value(row, '订单号') or '未知订单'}")

    summary["combined_text"] = "\n\n--------------------\n\n".join(summary["texts"]).strip()
    return summary
