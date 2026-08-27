import unittest

from ppt_word_gen.report_agent import ReportAgent, _json_from_tool_arguments
from ppt_word_gen.report_models import ReportCreate


class ToolArgumentsCompatibilityTests(unittest.TestCase):
    def test_parses_normal_object(self):
        self.assertEqual({"value": "OK"}, _json_from_tool_arguments('{"value":"OK"}'))

    def test_accepts_relay_prefixed_empty_object(self):
        self.assertEqual({"value": "OK"}, _json_from_tool_arguments('{}{"value":"OK"}'))

    def test_rejects_non_object_or_trailing_garbage(self):
        self.assertIsNone(_json_from_tool_arguments('[1, 2]'))
        self.assertIsNone(_json_from_tool_arguments('{}{"value":"OK"}x'))

    def test_runapi_claude_uses_json_content_compatibility_mode(self):
        agent = ReportAgent(ReportCreate(
            language_base_url="https://runapi.co/v1",
            language_api_key="test-key",
            language_model="claude-opus-5",
        ))
        self.assertTrue(agent._json_content_mode)


if __name__ == "__main__":
    unittest.main()
