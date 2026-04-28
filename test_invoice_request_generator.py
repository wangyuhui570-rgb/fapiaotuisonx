import os
import tempfile
import unittest
import zipfile

import invoice_request_generator as gen


def create_test_xlsx(path, headers, rows):
    def col_ref(index):
        result = ""
        index += 1
        while index:
            index, remainder = divmod(index - 1, 26)
            result = chr(65 + remainder) + result
        return result

    shared_strings = []
    string_index = {}

    def get_shared_index(value):
        value = str(value)
        if value not in string_index:
            string_index[value] = len(shared_strings)
            shared_strings.append(value)
        return string_index[value]

    all_rows = [headers] + rows
    row_xml_parts = []
    for row_number, row in enumerate(all_rows, start=1):
        cell_parts = []
        for col_number, value in enumerate(row):
            ref = f"{col_ref(col_number)}{row_number}"
            index = get_shared_index(value)
            cell_parts.append(f'<c r="{ref}" t="s"><v>{index}</v></c>')
        row_xml_parts.append(f'<row r="{row_number}">{"".join(cell_parts)}</row>')

    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{len(shared_strings)}" uniqueCount="{len(shared_strings)}">'
        + "".join(f"<si><t>{value}</t></si>" for value in shared_strings)
        + "</sst>"
    )
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(row_xml_parts)}</sheetData>'
        "</worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>'
        "</Relationships>"
    )
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        "</Types>"
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml)
        archive.writestr("_rels/.rels", root_rels_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        archive.writestr("xl/sharedStrings.xml", shared_xml)


class InvoiceRequestGeneratorTests(unittest.TestCase):
    def test_build_output_text_for_personal_title(self):
        row = {
            "订单号": "260001",
            "抬头类型": "个人",
            "发票抬头": "张三",
            "商品数量": "1",
            "发票金额": "794.0",
        }
        settings = {
            "store_name": "拼多多功率房",
            "owned_store": "执行派",
            "goods_name": "自行车配件",
            "applicant": "王宇辉",
        }
        text = gen.build_output_text(row, settings)
        self.assertIn("订单号：260001", text)
        self.assertIn("店铺名：拼多多功率房", text)
        self.assertIn("所属店铺：执行派", text)
        self.assertIn("抬头类型：电子普票", text)
        self.assertIn("发票抬头：张三", text)
        self.assertNotIn("企业税号：", text)
        self.assertIn("数量：1", text)
        self.assertIn("发票金额：794.00", text)

    def test_build_output_text_for_enterprise_title(self):
        row = {
            "订单号": "260002",
            "抬头类型": "企业",
            "发票抬头": "示例科技有限公司",
            "企业税号": "91320000123456789X",
            "开户银行": "中国银行南京分行",
            "账号": "6222000000000000",
            "地址": "南京市雨花台区",
            "电话": "13800000000",
            "商品数量": "2.0",
            "发票金额": "1280.5",
        }
        settings = {
            "store_name": "店铺A",
            "owned_store": "执行派",
            "goods_name": "自行车配件",
            "applicant": "王宇辉",
        }
        text = gen.build_output_text(row, settings)
        self.assertIn("企业税号：91320000123456789X", text)
        self.assertIn("开户银行：中国银行南京分行", text)
        self.assertIn("账号：6222000000000000", text)
        self.assertIn("地址：南京市雨花台区", text)
        self.assertIn("电话：13800000000", text)
        self.assertIn("数量：2", text)
        self.assertIn("发票金额：1280.50", text)

    def test_build_output_text_falls_back_to_enterprise_details_when_title_type_missing(self):
        row = {
            "订单号": "260003",
            "抬头类型": "",
            "发票抬头": "示例贸易有限公司",
            "企业税号": "91320000999999999X",
            "开户银行": "招商银行上海分行",
            "账号": "1234567890123456",
            "地址": "上海市闵行区",
            "电话": "021-12345678",
            "商品数量": "3",
            "发票金额": "300",
        }
        settings = {
            "store_name": "店铺A",
            "owned_store": "执行派",
            "goods_name": "自行车配件",
            "applicant": "王宇辉",
        }
        text = gen.build_output_text(row, settings)
        self.assertIn("企业税号：91320000999999999X", text)
        self.assertIn("开户银行：招商银行上海分行", text)
        self.assertIn("账号：1234567890123456", text)

    def test_build_output_text_omits_empty_optional_enterprise_fields(self):
        row = {
            "订单号": "260003A",
            "抬头类型": "企业",
            "发票抬头": "示例贸易有限公司",
            "企业税号": "91320000999999999X",
            "开户银行": "",
            "账号": "",
            "地址": "",
            "电话": "",
            "商品数量": "3",
            "发票金额": "300",
        }
        settings = {
            "store_name": "店铺A",
            "owned_store": "执行派",
            "goods_name": "自行车配件",
            "applicant": "王宇辉",
        }
        text = gen.build_output_text(row, settings)
        self.assertIn("企业税号：91320000999999999X", text)
        self.assertNotIn("开户银行：", text)
        self.assertNotIn("账号：", text)
        self.assertNotIn("地址：", text)
        self.assertNotIn("电话：", text)

    def test_build_output_text_keeps_personal_when_title_type_explicitly_personal(self):
        row = {
            "订单号": "260004",
            "抬头类型": "个人",
            "发票抬头": "张三",
            "企业税号": "91320000123456789X",
            "开户银行": "中国银行南京分行",
            "账号": "6222000000000000",
            "地址": "南京市雨花台区",
            "电话": "13800000000",
            "商品数量": "1",
            "发票金额": "88",
        }
        settings = {
            "store_name": "店铺A",
            "owned_store": "执行派",
            "goods_name": "自行车配件",
            "applicant": "王宇辉",
        }
        text = gen.build_output_text(row, settings)
        self.assertNotIn("企业税号：", text)
        self.assertNotIn("开户银行：", text)

    def test_generate_request_texts_uses_csv_rows(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_dir = os.path.join(tmp_dir, "csv")
            os.makedirs(csv_dir, exist_ok=True)
            with open(os.path.join(csv_dir, "demo.csv"), "w", encoding="utf-8", newline="") as f:
                f.write("订单号,抬头类型,发票抬头,商品数量,发票金额\n")
                f.write("260001,个人,张三,1,88\n")
                f.write("260002,个人,李四,2,99.5\n")

            summary = gen.generate_request_texts(
                csv_dir,
                {
                    "store_name": gen.DEFAULT_STORE_NAME,
                    "owned_store": gen.DEFAULT_OWNED_STORE,
                    "goods_name": gen.DEFAULT_GOODS_NAME,
                    "applicant": gen.DEFAULT_APPLICANT,
                },
            )

            self.assertEqual(summary["csv_files"], 1)
            self.assertEqual(summary["rows"], 2)
            self.assertEqual(summary["generated"], 2)
            self.assertEqual(len(summary["texts"]), 2)
            self.assertIn("订单号：260001", summary["texts"][0])
            self.assertIn("订单号：260002", summary["combined_text"])

    def test_generate_request_texts_supports_selected_csv_paths(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_dir = os.path.join(tmp_dir, "csv")
            os.makedirs(csv_dir, exist_ok=True)
            csv_a = os.path.join(csv_dir, "a.csv")
            csv_b = os.path.join(csv_dir, "b.csv")

            with open(csv_a, "w", encoding="utf-8", newline="") as f:
                f.write("订单号,抬头类型,发票抬头,商品数量,发票金额\n")
                f.write("100001,个人,张三,1,88\n")
            with open(csv_b, "w", encoding="utf-8", newline="") as f:
                f.write("订单号,抬头类型,发票抬头,商品数量,发票金额\n")
                f.write("100002,个人,李四,1,99\n")

            summary = gen.generate_request_texts(
                csv_dir,
                {
                    "store_name": gen.DEFAULT_STORE_NAME,
                    "owned_store": gen.DEFAULT_OWNED_STORE,
                    "goods_name": gen.DEFAULT_GOODS_NAME,
                    "applicant": gen.DEFAULT_APPLICANT,
                },
                csv_paths=[csv_b],
            )

            self.assertEqual(summary["csv_files"], 1)
            self.assertEqual(summary["rows"], 1)
            self.assertIn("订单号：100002", summary["combined_text"])
            self.assertNotIn("订单号：100001", summary["combined_text"])

    def test_generate_request_texts_supports_selected_xlsx_paths(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            xlsx_path = os.path.join(tmp_dir, "登记表.xlsx")
            create_test_xlsx(
                xlsx_path,
                ["订单号/iD", "店铺名", "所属公司", "发票类型", "开票信息", "货物名称", "单位", "数量", "发票金额", "申请人"],
                [["JD10001", "八戒家", "八戒家", "个人", "张三", "自行车配件", "件", "2", "305.88", "王宇辉"]],
            )

            summary = gen.generate_request_texts(
                tmp_dir,
                {
                    "store_name": "默认店铺",
                    "owned_store": "默认公司",
                    "goods_name": "默认货物",
                    "applicant": "默认申请人",
                },
                csv_paths=[xlsx_path],
            )

            self.assertEqual(summary["csv_files"], 1)
            self.assertEqual(summary["rows"], 1)
            self.assertIn("订单号：JD10001", summary["combined_text"])
            self.assertIn("店铺名：八戒家", summary["combined_text"])
            self.assertIn("所属店铺：八戒家", summary["combined_text"])
            self.assertIn("发票抬头：张三", summary["combined_text"])
            self.assertIn("货物名称：自行车配件", summary["combined_text"])
            self.assertIn("单位：件", summary["combined_text"])
            self.assertIn("申请人：王宇辉", summary["combined_text"])

    def test_build_output_text_supports_taobao_header_aliases(self):
        row = {
            "*订单编号": "TB10001",
            "*抬头类型": "企业",
            "*发票抬头": "淘宝测试公司",
            "*开票总金额": "256.8",
            "数量": "2",
            "购方税号": "91320000123456789X",
            "企业地址": "无锡市测试路1号",
            "企业电话": "0510-12345678",
            "开户行": "中国银行无锡分行",
            "开户账号": "6222000000000000",
        }
        settings = {
            "store_name": "拼多多功率房",
            "owned_store": "执行派",
            "goods_name": "自行车配件",
            "applicant": "王宇辉",
        }
        text = gen.build_output_text(row, settings)
        self.assertIn("订单号：TB10001", text)
        self.assertIn("发票抬头：淘宝测试公司", text)
        self.assertIn("企业税号：91320000123456789X", text)
        self.assertIn("开户银行：中国银行无锡分行", text)
        self.assertIn("账号：6222000000000000", text)
        self.assertIn("地址：无锡市测试路1号", text)
        self.assertIn("电话：0510-12345678", text)
        self.assertIn("数量：2", text)
        self.assertIn("发票金额：256.80", text)

    def test_merge_taobao_orders_outputs_one_invoice_per_order(self):
        rows = [
            {
                "*订单编号": "TB10001",
                "货物名称": "骑行台",
                "*抬头类型": "个人",
                "*发票抬头": "张三",
                "数量": "1",
                "商品金额": "1239.00",
                "*开票总金额": "1276.34",
            },
            {
                "*订单编号": "TB10001",
                "货物名称": "骑行台",
                "*抬头类型": "个人",
                "*发票抬头": "张三",
                "数量": "",
                "商品金额": "-57.10",
                "*开票总金额": "1276.34",
            },
            {
                "*订单编号": "TB10001",
                "货物名称": "地垫",
                "*抬头类型": "个人",
                "*发票抬头": "张三",
                "数量": "1",
                "商品金额": "99.00",
                "*开票总金额": "1276.34",
            },
            {
                "*订单编号": "TB10001",
                "货物名称": "地垫",
                "*抬头类型": "个人",
                "*发票抬头": "张三",
                "数量": "",
                "商品金额": "-4.56",
                "*开票总金额": "1276.34",
            },
        ]

        merged = gen.merge_taobao_orders(rows)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0]["商品数量"], "2")
        self.assertEqual(merged[0]["发票金额"], "1276.34")


if __name__ == "__main__":
    unittest.main()
