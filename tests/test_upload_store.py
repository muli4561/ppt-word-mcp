import tempfile
import unittest
from pathlib import Path

from ppt_word_gen.upload_store import UploadNotFound, UploadPurposeMismatch, UploadStore


class UploadStoreTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory(prefix="upload-store-")
        self.store = UploadStore(Path(self.temp_dir.name), expire_hours=1)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_round_trip_and_delete(self):
        metadata = self.store.put(b"# source", "业务资料.md", "source")
        data, filename, loaded = self.store.get(metadata["upload_id"], "source")

        self.assertEqual(b"# source", data)
        self.assertEqual(".md", Path(filename).suffix)
        self.assertEqual(metadata["upload_id"], loaded["upload_id"])
        self.assertTrue(self.store.delete(metadata["upload_id"]))
        with self.assertRaises(UploadNotFound):
            self.store.get(metadata["upload_id"])

    def test_purpose_is_enforced(self):
        metadata = self.store.put(b"plain", "source.txt", "source")
        with self.assertRaises(UploadPurposeMismatch):
            self.store.get(metadata["upload_id"], "reference_template")

    def test_template_extension_is_checked(self):
        with self.assertRaises(ValueError):
            self.store.put(b"plain", "template.pdf", "reference_template")


if __name__ == "__main__":
    unittest.main()
