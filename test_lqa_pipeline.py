import asyncio
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

import lqa_pipeline as pipeline


class _NeverReturningModels:
    async def generate_content(self, **kwargs):
        await asyncio.Event().wait()


class _NeverReturningClient:
    def __init__(self):
        self.aio = type("Aio", (), {"models": _NeverReturningModels()})()


class TicketPipelineTests(unittest.TestCase):
    def test_language_filter_is_case_insensitive(self):
        self.assertEqual(
            pipeline._normalize_requested_language("gr", {"CZ": "cs-CZ", "GR": "el-GR"}),
            "GR",
        )

    def test_language_filter_rejects_unknown_code(self):
        with self.assertRaisesRegex(ValueError, "Unknown language code"):
            pipeline._normalize_requested_language("XX", {"CZ": "cs-CZ", "GR": "el-GR"})

    def test_three_parallel_runs_are_capped_at_thirty_combined_calls(self):
        config = pipeline.load_config()
        self.assertLessEqual(config["batching"]["max_batches_per_window"] * 3, 30)

    def test_all_ticket_docx_files_map_to_their_runtime_locale(self):
        config = pipeline.load_config()
        archive = Path("03_Car Hire in City-20260831T145126Z-1-001.zip")

        with ZipFile(archive) as zip_file:
            docx_paths = [Path(name) for name in zip_file.namelist() if name.endswith(".docx")]

        counts = Counter(
            pipeline._infer_lang_code_for_input_file(path, config["lang_map"])
            for path in docx_paths
        )

        self.assertEqual(
            counts,
            {"CZ": 5, "GR": 5, "HU": 5, "PL": 5, "PT": 5, "SA": 5, "TH": 5, "TR": 5},
        )

    def test_detailed_findings_report_shows_source_filename(self):
        issue = {
            "source_file": "/ticket/input/example.docx",
            "candidate": "https://example.test/page",
            "text": "Original text",
            "issue": "Objective error",
            "type_of_issue": "Grammar",
            "severity": "Minor",
            "score_deduction": 1,
            "solution": "Apply the correction.",
            "updated_sentence": "Corrected text",
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir)
            pipeline.write_language_report(
                lang_code="PT",
                evaluated_scope_label="test",
                evaluated_text="Original text",
                issues_with_url=[issue],
                output_dir=output_dir,
                save_raw_json=False,
            )
            report_path = next(output_dir.glob("*.xlsx"))
            workbook = load_workbook(report_path, read_only=True, data_only=True)
            rows = list(workbook["Detailed Findings"].iter_rows(values_only=True))
            workbook.close()

        self.assertEqual(rows[0][0], "Source File")
        self.assertEqual(rows[1][0], "example.docx")


class GeminiTimeoutTests(unittest.IsolatedAsyncioTestCase):
    async def test_hung_gemini_request_is_bounded_by_pipeline_timeout(self):
        started = asyncio.get_running_loop().time()
        original_timeout = getattr(pipeline, "GEMINI_REQUEST_TIMEOUT_SECONDS", None)
        pipeline.GEMINI_REQUEST_TIMEOUT_SECONDS = 0.01
        try:
            with self.assertRaises(asyncio.TimeoutError):
                await asyncio.wait_for(
                    pipeline._call_gemini_once.__wrapped__(
                        _NeverReturningClient(),
                        "gemini-2.5-pro",
                        1,
                        {"system": "test", "prompt": "test"},
                    ),
                    timeout=0.2,
                )
        finally:
            if original_timeout is None:
                del pipeline.GEMINI_REQUEST_TIMEOUT_SECONDS
            else:
                pipeline.GEMINI_REQUEST_TIMEOUT_SECONDS = original_timeout

        self.assertLess(asyncio.get_running_loop().time() - started, 0.1)


if __name__ == "__main__":
    unittest.main()
