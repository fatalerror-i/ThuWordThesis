"""官方文本、学位目录和已停用功能的发布契约。"""
import unittest
from zipfile import ZipFile
from lxml import etree
from src.build.degree_catalog import load_degree_config, degree_display_options, degree_english, discipline_options
from tests.support import ROOT, product_paths


class OfficialContentTest(unittest.TestCase):
    def test_disciplines_follow_guide_appendix_categories(self):
        config = load_degree_config()
        academic = {item["name_cn"]: item["disciplines"] for item in config["academic"]}
        self.assertEqual(len(academic), 11)
        self.assertEqual(academic["经济学"], ["理论经济学", "应用经济学"])
        self.assertEqual(academic["历史学"], ["中国史", "世界史"])
        self.assertIn("统计学", academic["理学"])
        self.assertNotIn("统计学", academic["经济学"])
        self.assertIn("心理学", academic["教育学"])
        self.assertIn("力学", academic["工学"])
        self.assertIn("生物医学工程", academic["工学"])
        self.assertNotIn("交叉学科", academic)
        names = [name for group in academic.values() for name in group] + config["cross_disciplinary"]
        self.assertEqual(len(names), 59)
        self.assertEqual(len(set(names)), 59)
        self.assertIn("统计学", discipline_options(config, "经济学"))
        for category in ("理学", "工学"):
            self.assertIn("集成电路科学与工程", discipline_options(config, category))
            self.assertIn("计算机科学与技术", discipline_options(config, category))
        for category in ("工学", "艺术学"):
            self.assertIn("设计学", discipline_options(config, category))
        self.assertNotIn("交叉学科硕士学位", degree_display_options(config))
        for item in config["professional"]:
            self.assertEqual(discipline_options(config, item["name_cn"]), [])

    def test_english_discipline_uses_two_paragraphs_without_soft_break(self):
        with ZipFile(product_paths()["docx"]) as archive:
            doc = etree.fromstring(archive.read("word/document.xml"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        cc = doc.xpath('//w:sdt[w:sdtPr/w:tag[@w:val="英文学科"]]', namespaces=ns)[0]
        self.assertEqual("".join(cc.xpath('.//w:t/text()', namespaces=ns)), "Computer Science and Technology")
        self.assertEqual(cc.xpath('.//w:br', namespaces=ns), [])
        prefix = cc.getparent().getprevious()
        self.assertEqual("".join(prefix.xpath('.//w:t/text()', namespaces=ns)), "in")
        self.assertEqual(prefix.xpath('.//w:sdt', namespaces=ns), [])

    def test_academic_degrees(self):
        config = load_degree_config()
        self.assertEqual(degree_english(config, "工学"), "Master of Science")
        for item in config["academic"]:
            self.assertIn(item["master_en"], ("Master of Arts", "Master of Science"))
            self.assertEqual(item["doctor_en"], "Doctor of Philosophy")

    def test_professional_degrees_follow_guide(self):
        config = load_degree_config()
        for category, expected in (("法律", "Juris Master"), ("体育", "Master of Science in Physical Education"),
                                   ("电子信息", "Master of Electronic and Information Engineering"),
                                   ("资源与环境", "Master of Resources and Environmental Engineering")):
            self.assertEqual(degree_english(config, category), expected)
        options = degree_display_options(config)
        self.assertNotIn("金融博士专业学位", options)
        self.assertIn("教育博士专业学位", options)
        self.assertNotIn("教育硕士专业学位", options)
        with self.assertRaises(ValueError):
            degree_english(config, "金融", "博士")

    def test_official_graduate_text(self):
        with ZipFile(product_paths()["docx"]) as archive:
            doc = etree.fromstring(archive.read("word/document.xml"))
        text = "".join(doc.xpath('//*[local-name()="t"]/text()'))
        self.assertIn("其中包括但不限于：（1）", text)
        self.assertIn("所取得的成果，不包含涉及国家秘密的内容。", text)

    def test_continuation_entry_is_disabled(self):
        for variant in ("graduate", "undergraduate"):
            ribbon = etree.parse(str(product_paths(variant)["ribbon"]))
            self.assertEqual(ribbon.xpath('//*[@id="btnContinuedTable"]'), [])
        source = (ROOT / "src/dotm/vba/modCaption.bas").read_text(encoding="utf-8")
        self.assertNotRegex(source, r"(?m)^Public Sub InsertContinuedTableCaption\b")


if __name__ == "__main__":
    unittest.main()
