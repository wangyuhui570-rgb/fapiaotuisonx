import os
import tempfile
import unittest

import store_mapping


class StoreMappingTests(unittest.TestCase):
    def test_parse_mapping_text_supports_aligned_spaces(self):
        text = (
            "八戒家 正品老店               八戒家\n"
            "度谷运动专营店                无锡度谷商贸有限公司\n"
        )
        bindings = store_mapping.parse_mapping_text(text)
        self.assertEqual(
            bindings,
            [
                ("八戒家 正品老店", "八戒家"),
                ("度谷运动专营店", "无锡度谷商贸有限公司"),
            ],
        )

    def test_owned_stores_for_keeps_unique_order(self):
        bindings = [
            ("thinkrider旗舰店", "八戒家"),
            ("thinkrider旗舰店", "执行派"),
            ("thinkrider旗舰店", "执行派"),
        ]
        self.assertEqual(
            store_mapping.owned_stores_for(bindings, "thinkrider旗舰店"),
            ["八戒家", "执行派"],
        )

    def test_upsert_binding_only_appends_new_pair(self):
        bindings = [("店铺A", "公司A")]
        updated = store_mapping.upsert_binding(bindings, "店铺A", "公司A")
        self.assertEqual(updated, [("店铺A", "公司A")])
        updated = store_mapping.upsert_binding(bindings, "店铺A", "公司B")
        self.assertEqual(updated, [("店铺A", "公司A"), ("店铺A", "公司B")])

    def test_save_and_load_bindings(self):
        bindings = [("店铺A", "公司A"), ("店铺B", "公司B")]
        with tempfile.TemporaryDirectory() as tmp_dir:
            path = os.path.join(tmp_dir, "店铺发票对应公司.txt")
            store_mapping.save_bindings(path, bindings)
            self.assertEqual(store_mapping.load_bindings(path), bindings)


if __name__ == "__main__":
    unittest.main()
