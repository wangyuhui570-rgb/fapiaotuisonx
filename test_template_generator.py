import os
import tempfile
import unittest
import zipfile

import template_generator


class TemplateGeneratorTests(unittest.TestCase):
    def test_render_template_replaces_row_and_builtin_fields(self):
        context = template_generator.build_row_context(
            {"公司名称": "示例科技", "金额": "128.50"},
            csv_path=r"C:\demo\orders.csv",
            row_index=3,
            global_index=5,
        )
        rendered, missing = template_generator.render_template(
            "申请人：{{公司名称}}\n金额：{{金额}}\n序号：{{序号}}\n来源：{{CSV文件名}}",
            context,
        )
        self.assertEqual(missing, [])
        self.assertIn("申请人：示例科技", rendered)
        self.assertIn("金额：128.50", rendered)
        self.assertIn("序号：3", rendered)
        self.assertIn("来源：orders.csv", rendered)

    def test_read_csv_rows_supports_utf8_sig_and_strips_blank_rows(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_path = os.path.join(tmp_dir, "demo.csv")
            with open(csv_path, "w", encoding="utf-8-sig", newline="") as file_obj:
                file_obj.write("公司名称,金额\n")
                file_obj.write("示例科技,88.00\n")
                file_obj.write(",\n")

            headers, rows, encoding = template_generator.read_csv_rows(csv_path)
            self.assertEqual(headers, ["公司名称", "金额"])
            self.assertEqual(encoding, "utf-8-sig")
            self.assertEqual(rows, [{"公司名称": "示例科技", "金额": "88.00"}])

    def test_generate_templates_creates_unique_output_files(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_dir = os.path.join(tmp_dir, "csv")
            output_dir = os.path.join(tmp_dir, "out")
            os.makedirs(csv_dir, exist_ok=True)

            csv_path = os.path.join(csv_dir, "demo.csv")
            with open(csv_path, "w", encoding="utf-8", newline="") as file_obj:
                file_obj.write("公司名称,金额\n")
                file_obj.write("示例科技,88.00\n")
                file_obj.write("示例科技,99.00\n")

            summary = template_generator.generate_templates(
                csv_dir=csv_dir,
                output_dir=output_dir,
                template_text="公司：{{公司名称}}\n金额：{{金额}}",
                filename_template="{{公司名称}}",
            )

            self.assertEqual(summary["csv_files"], 1)
            self.assertEqual(summary["rows"], 2)
            self.assertEqual(summary["generated"], 2)
            self.assertEqual(summary["missing_placeholders"], {})

            output_names = sorted(os.listdir(output_dir))
            self.assertEqual(output_names, ["示例科技.txt", "示例科技_2.txt"])

    def test_generate_templates_collects_missing_placeholders(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_dir = os.path.join(tmp_dir, "csv")
            output_dir = os.path.join(tmp_dir, "out")
            os.makedirs(csv_dir, exist_ok=True)

            csv_path = os.path.join(csv_dir, "demo.csv")
            with open(csv_path, "w", encoding="utf-8", newline="") as file_obj:
                file_obj.write("公司名称\n")
                file_obj.write("示例科技\n")

            summary = template_generator.generate_templates(
                csv_dir=csv_dir,
                output_dir=output_dir,
                template_text="公司：{{公司名称}}\n税号：{{税号}}",
                filename_template="{{公司名称}}_{{税号}}",
            )

            self.assertEqual(summary["generated"], 1)
            self.assertEqual(summary["missing_placeholders"], {"demo.csv": ["税号"]})

    def test_read_csv_rows_supports_xlsx_content_with_csv_extension(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            csv_path = os.path.join(tmp_dir, "taobao_export.csv")
            with zipfile.ZipFile(csv_path, "w") as archive:
                archive.writestr(
                    "[Content_Types].xml",
                    """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
</Types>""",
                )
                archive.writestr(
                    "_rels/.rels",
                    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""",
                )
                archive.writestr(
                    "xl/workbook.xml",
                    """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></workbook>""",
                )
                archive.writestr(
                    "xl/_rels/workbook.xml.rels",
                    """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>""",
                )
                archive.writestr(
                    "xl/worksheets/sheet1.xml",
                    """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>*订单编号</t></is></c>
      <c r="B1" t="inlineStr"><is><t>*开票总金额</t></is></c>
      <c r="C1" t="inlineStr"><is><t>数量</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>TB10001</t></is></c>
      <c r="B2" t="inlineStr"><is><t>256.80</t></is></c>
      <c r="C2" t="inlineStr"><is><t>2</t></is></c>
    </row>
  </sheetData>
</worksheet>""",
                )

            headers, rows, encoding = template_generator.read_csv_rows(csv_path)
            self.assertEqual(encoding, "xlsx")
            self.assertEqual(headers, ["*订单编号", "*开票总金额", "数量"])
            self.assertEqual(rows, [{"*订单编号": "TB10001", "*开票总金额": "256.80", "数量": "2"}])


if __name__ == "__main__":
    unittest.main()
