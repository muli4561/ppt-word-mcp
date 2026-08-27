import unittest
import tempfile
from pathlib import Path

from docx import Document
from docx.shared import Pt

from ppt_word_gen.config import WORD_REPORT_TEMPLATE_PATH
from ppt_word_gen.report_models import WordFormatOverrides
from ppt_word_gen.word_format import issue_word_format_confirmation, verify_word_format_confirmation


class WordFormatTests(unittest.TestCase):
    def test_fixed_template_and_confirmation_token(self):
        self.assertTrue(WORD_REPORT_TEMPLATE_PATH.is_file())
        confirmation = issue_word_format_confirmation()
        verified = verify_word_format_confirmation(confirmation.confirmation_token)
        self.assertEqual(confirmation.profile_id, verified.profile_id)
        self.assertIn("正文", confirmation.confirmation_text)

    def test_uploaded_template_defaults_can_be_overridden_and_are_bound_to_token(self):
        with tempfile.TemporaryDirectory(prefix="word-template-") as directory:
            path = Path(directory) / "custom.docx"
            document = Document()
            document.styles["Normal"].font.name = "SimSun"
            document.styles["Normal"].font.size = Pt(11)
            document.save(path)
            confirmation = issue_word_format_confirmation(
                template_path=path,
                template_upload_id="0123456789abcdef",
                overrides=WordFormatOverrides(body_font="宋体", body_size_pt=12.5),
            )
            verified = verify_word_format_confirmation(
                confirmation.confirmation_token,
                template_resolver=lambda _upload_id: path,
            )
        self.assertEqual("uploaded", confirmation.template_source)
        self.assertEqual("宋体", verified.format.body_font)
        self.assertEqual(12.5, verified.format.body_size_pt)
        self.assertEqual("0123456789abcdef", verified.template_upload_id)

    def test_tampered_confirmation_is_rejected(self):
        confirmation = issue_word_format_confirmation()
        with self.assertRaisesRegex(ValueError, "确认已失效"):
            verify_word_format_confirmation(confirmation.confirmation_token + "x")


if __name__ == "__main__":
    unittest.main()
