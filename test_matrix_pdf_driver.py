import json
import os
import re
import socket
import tempfile
import threading
import time
import unittest
from datetime import datetime
from pathlib import Path
from unittest import mock

import matrix_pdf_driver as driver

# 아래 시험들은 한글 문구를 기준으로 쓰여 있다. 본체가 설정을 읽으면서 화면 언어를 바꾸지 못하게 고정하고,
# 영어 화면은 LanguageTests에서 따로 본다.
driver.LANGUAGE = "ko"
driver.set_language = lambda code: None

SAMPLE_PS = (
    b"%!PS-Adobe-3.0\n%%Title: (Sample Report)\n%%For: (tester)\n%%Pages: 1\n%%EndComments\n"
    b"/Helvetica findfont 14 scalefont setfont 72 720 moveto (Result 12.34) show showpage\n%%EOF\n"
)
BROKEN_PS = b"%!PS-Adobe-3.0\n%%Title: (Broken)\n%%EndComments\nthis-is-not-an-operator\n%%EOF\n"
WHEN = datetime(2026, 10, 8, 15, 4, 5)


def _free_port() -> int:
    probe = socket.socket()
    probe.bind(("127.0.0.1", 0))
    port = probe.getsockname()[1]
    probe.close()
    return port


def _values(title="결과", user="hong", counter=7):
    return driver.token_values(title, user, WHEN, counter, "Matrix PDF-Driver")


class HeaderTests(unittest.TestCase):
    def test_hex_title_is_decoded_as_korean(self):
        encoded = "시험성적서 HPLC-001".encode("cp949").hex().upper().encode("ascii")
        head = b"%!PS-Adobe-3.0\r\n%%Title: <" + encoded + b">\r\n%%Creator: PScript5.dll\r\n"
        self.assertEqual(driver.parse_title(head), "시험성적서 HPLC-001")

    def test_parenthesized_and_plain_values(self):
        self.assertEqual(driver.parse_title(b"%!PS\n%%Title: (Report 1)\n"), "Report 1")
        self.assertEqual(driver.parse_title(b"%!PS\n%%Title: Report 2\n"), "Report 2")
        self.assertEqual(driver.parse_dsc(b"%!PS\n%%For: JanghoSeo\r\n%%Title: x\n", "For"), "JanghoSeo")

    def test_octal_escaped_korean_title_is_decoded(self):
        head = b"%!PS\n%%Title: (Layout_03.txt - \\270\\336\\270\\360\\300\\345)\n"
        self.assertEqual(driver.parse_title(head), "Layout_03.txt - 메모장")
        self.assertEqual(driver.parse_title(b"%!PS\n%%Title: (a \\(b\\) c\\\\d)\n"), "a (b) c\\d")

    def test_notices_of_one_job_are_handled_in_the_order_they_happened(self):
        names = ["20261008_1_0002.saved.json", "20261008_1_0002.start.json", "20261008_1_0001.failed.json",
                 "20261008_1_0001.start.json", "20261008_1_0003.ask.json"]
        ordered = sorted((Path(name) for name in names), key=driver.notice_order)
        self.assertEqual([p.name for p in ordered], [
            "20261008_1_0001.start.json", "20261008_1_0001.failed.json",
            "20261008_1_0002.start.json", "20261008_1_0002.saved.json", "20261008_1_0003.ask.json"])

    def test_missing_or_broken_value(self):
        self.assertEqual(driver.parse_title(b"%!PS\n%%Pages: 1\n"), "")
        self.assertEqual(driver.parse_title(b"%!PS\n%%Title: <ZZ>\n"), "")
        self.assertEqual(driver.parse_dsc(b"%!PS\n%%Title: x\n", "For"), "")


class NamingTests(unittest.TestCase):
    def test_forbidden_characters_are_removed(self):
        self.assertEqual(driver.sanitize_filename('a/b\\c:d*e?f"g<h>i|j'), "a b c d e f g h i j")
        self.assertEqual(driver.sanitize_filename("  끝에 점...  "), "끝에 점")

    def test_every_listed_token_can_be_used(self):
        pattern = "_".join(token for token, _ in driver.TOKENS)
        self.assertEqual(
            driver.build_filename(pattern, _values()),
            "결과_20261008_150405_20261008_150405_2026_10_08_hong_%s_Matrix PDF-Driver_7" % socket.gethostname(),
        )
        driver.check_template(pattern)

    def test_empty_title_and_user_get_placeholders(self):
        values = _values(title="", user="")
        self.assertEqual(driver.build_filename("{title}-{user}", values), "print-unknown")

    def test_counter_accepts_padding(self):
        self.assertEqual(driver.build_filename("{counter:05d}", _values()), "00007")

    def test_broken_pattern_falls_back_and_is_reported(self):
        self.assertEqual(driver.build_filename("{unknown}", _values(title="x")), "Cube_20261008_150405")
        for bad in ("{nope}", "{title", "{0}"):
            with self.assertRaises(ValueError):
                driver.check_template(bad)

    def test_folder_rule_creates_dated_subfolders(self):
        template = r"C:\MatrixPDF\{year}\{month}\{user}"
        self.assertEqual(driver.render_folder(template, _values()), Path(r"C:\MatrixPDF\2026\10\hong"))
        self.assertEqual(driver.static_folder_root(template), r"C:\MatrixPDF")
        self.assertEqual(driver.static_folder_root(r"D:\out"), r"D:\out")

    def test_title_cannot_escape_the_folder(self):
        values = _values(title=r"..\..\Windows\evil")
        rendered = driver.render_folder(r"C:\MatrixPDF\{title}", values)
        self.assertEqual(rendered.parent, Path(r"C:\MatrixPDF"))

    def test_unique_path_skips_existing_and_in_progress(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            self.assertEqual(driver.unique_pdf_path(root, "a").name, "a.pdf")
            (root / "a.pdf").write_bytes(b"x")
            (root / "a_2.pdf.tmp").write_bytes(b"x")
            self.assertEqual(driver.unique_pdf_path(root, "a").name, "a_3.pdf")


class WrapperTests(unittest.TestCase):
    def test_leading_printer_commands_are_skipped(self):
        head = b"\x1b%-12345X@PJL JOB\r\n@PJL ENTER LANGUAGE=POSTSCRIPT\r\n\x04%!PS-Adobe-3.0\n"
        self.assertEqual(head[driver.find_postscript_start(head):], b"%!PS-Adobe-3.0\n")
        self.assertEqual(driver.find_postscript_start(b"plain text job"), -1)

    def test_trailer_after_eof_is_removed(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "job"
            path.write_bytes(b"%!PS\nshowpage\n%%EOF\r\n\x04\x1b%-12345X@PJL EOJ\r\n\x1b%-12345X")
            driver.trim_trailer(path)
            self.assertEqual(path.read_bytes(), b"%!PS\nshowpage\n%%EOF\n")

    def test_file_without_eof_is_untouched(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "job"
            path.write_bytes(b"%!PS\nshowpage\n")
            driver.trim_trailer(path)
            self.assertEqual(path.read_bytes(), b"%!PS\nshowpage\n")


class ConfigTests(unittest.TestCase):
    def test_defaults(self):
        config = driver.Config({})
        self.assertEqual(config.port, 9100)
        self.assertEqual(config.save_mode, "auto")
        self.assertEqual(config.image_quality, "lossless")
        self.assertEqual(config.printer_name, "Matrix PDF-Driver")
        self.assertEqual((config.filename_prefix, config.filename_pattern), ("Cube_", "{datetime}"))
        self.assertEqual(config.name_rule, "Cube_{datetime}")
        self.assertEqual((config.page_rotation, config.color_mode), ("keep", "keep"))
        self.assertFalse(config.notifies_helper)
        self.assertEqual(config.output_folder, r"C:\MatrixPDF")
        self.assertEqual(driver.Config({"image_quality": "weird", "save_mode": "weird"}).save_mode, "auto")

    def test_load_reads_bom_file(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "config.json"
            path.write_text(json.dumps({"output_folder": folder, "port": 9105, "save_mode": "ask"}), encoding="utf-8-sig")
            config = driver.load_config(path)
            self.assertEqual(config.port, 9105)
            self.assertEqual(config.save_mode, "ask")
            self.assertEqual(config.output_folder, folder)

    def test_name_rule_is_split_into_fixed_start_and_items(self):
        split = driver.split_name_rule
        self.assertEqual(split({"filename_pattern": "Cube_{datetime}"}), ("Cube_", "{datetime}"))
        self.assertEqual(split({"filename_pattern": "{datetime}_{title}"}), ("", "{datetime}_{title}"))
        self.assertEqual(split({"filename_pattern": "fixed"}), ("fixed", ""))
        self.assertEqual(split({"filename_prefix": "", "filename_pattern": "{title}"}), ("", "{title}"))
        self.assertEqual(split({"filename_prefix": "A{1}_", "filename_pattern": ""}), ("A{1}_", ""))
        config = driver.Config({"filename_prefix": "A{1}_", "filename_pattern": "{date}"})
        self.assertEqual(driver.build_filename(config.name_rule, _values()), "A{1}_20261008")

    def test_rotation_and_color_arguments(self):
        plain = driver.ghostscript_args(Path("gs"), "lossless", Path("o.pdf"), Path("i.ps"))
        tuned = driver.ghostscript_args(Path("gs"), "lossless", Path("o.pdf"), Path("i.ps"), "auto", "gray")
        self.assertIn("-dAutoRotatePages=/None", plain)
        self.assertNotIn("-sColorConversionStrategy=Gray", plain)
        self.assertIn("-dAutoRotatePages=/PageByPage", tuned)
        self.assertIn("-sColorConversionStrategy=Gray", tuned)

    def test_pdfa_arguments(self):
        plain = driver.ghostscript_args(Path("gs"), "lossless", Path("o.pdf"), Path("i.ps"))
        self.assertFalse([a for a in plain if "PDFA" in a or "ColorConversion" in a])
        self.assertEqual(plain[-1], "i.ps")
        archived = driver.ghostscript_args(Path("gs"), "lossless", Path("o.pdf"), Path("i.ps"), "keep", "keep", 2, Path("d.ps"))
        self.assertIn("-dPDFA=2", archived)
        self.assertIn("-sColorConversionStrategy=RGB", archived)
        self.assertEqual(archived[-2:], ["d.ps", "i.ps"])
        gray = driver.ghostscript_args(Path("gs"), "lossless", Path("o.pdf"), Path("i.ps"), "keep", "gray", 1, Path("d.ps"))
        self.assertIn("-sColorConversionStrategy=Gray", gray)
        self.assertNotIn("-sColorConversionStrategy=RGB", gray)
        self.assertEqual(driver.Config({"pdf_format": "PDFA3"}).pdf_format, "pdfa3")
        self.assertEqual(driver.Config({"pdf_format": "pdfa4"}).pdf_format, "pdf")

    def test_korean_font_names_are_rewritten_as_utf8(self):
        korean = "맑은 고딕".encode("cp949").hex().upper().encode("ascii")
        utf8 = "맑은 고딕".encode("utf-8").hex().upper().encode("ascii")
        with tempfile.TemporaryDirectory() as folder:
            source, target = Path(folder) / "a.job", Path(folder) / "a.utf8"
            source.write_bytes(b"%!PS\n/OrigFontName <" + korean + b"> def\n/OrigFontName <417269616C> def\n"
                               b"/Other <" + korean + b"> def\n")
            self.assertEqual(driver.rewrite_font_names_as_utf8(source, target), 1)
            self.assertEqual(target.read_bytes(), b"%!PS\n/OrigFontName <" + utf8 + b"> def\n"
                             b"/OrigFontName <417269616C> def\n/Other <" + korean + b"> def\n")

    def test_lossless_and_compact_arguments(self):
        lossless = driver.ghostscript_args(Path("gs"), "lossless", Path("o.pdf"), Path("i.ps"))
        compact = driver.ghostscript_args(Path("gs"), "compact", Path("o.pdf"), Path("i.ps"))
        self.assertIn("-dColorImageFilter=/FlateEncode", lossless)
        self.assertIn("-dPDFSETTINGS=/printer", compact)
        self.assertIn("-dAutoRotatePages=/None", compact)


class SettingsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.config_path = self.root / "config.json"

    def _settings(self, **changes):
        settings = {
            "save_mode": "auto", "output_folder": str(self.root), "filename_prefix": "Cube_",
            "filename_pattern": "{datetime}", "image_quality": "lossless", "page_rotation": "keep",
            "color_mode": "keep", "show_progress": False, "open_after_save": False,
            "also_print": False, "also_print_printer": "",
        }
        settings.update(changes)
        return settings

    def test_save_keeps_other_keys_and_creates_fixed_part_of_folder(self):
        self.config_path.write_text(json.dumps({"port": 9103, "printer_name": "P"}), encoding="utf-8")
        template = str(self.root / "새 폴더") + "\\{date}"
        driver.save_settings(self.config_path, self._settings(
            output_folder=template, filename_prefix="R-", filename_pattern="{date}_{title}", image_quality="compact",
            page_rotation="auto", color_mode="gray", show_progress=True, open_after_save=True,
            also_print=True, also_print_printer="Lab Printer"))
        raw = json.loads(self.config_path.read_text(encoding="utf-8"))
        self.assertEqual((raw["port"], raw["printer_name"]), (9103, "P"))
        self.assertEqual(raw["output_folder"], template)
        self.assertEqual((raw["filename_prefix"], raw["filename_pattern"]), ("R-", "{date}_{title}"))
        self.assertEqual((raw["image_quality"], raw["page_rotation"], raw["color_mode"]), ("compact", "auto", "gray"))
        self.assertEqual((raw["show_progress"], raw["open_after_save"]), (True, True))
        self.assertEqual((raw["also_print"], raw["also_print_printer"]), (True, "Lab Printer"))
        self.assertTrue((self.root / "새 폴더").is_dir())
        self.assertFalse((self.root / "새 폴더" / "{date}").exists())
        config = driver.load_config(self.config_path)
        self.assertTrue(config.notifies_helper)
        self.assertEqual(config.name_rule, "R-{date}_{title}")

    def test_ask_mode_does_not_need_a_folder(self):
        driver.save_settings(self.config_path, self._settings(save_mode="ask", output_folder=""))
        raw = json.loads(self.config_path.read_text(encoding="utf-8"))
        self.assertEqual(raw["save_mode"], "ask")
        self.assertNotIn("output_folder", raw)

    def test_invalid_values_are_rejected_without_writing(self):
        good = str(self.root)
        cases = (
            {"output_folder": ""}, {"output_folder": "relative"}, {"output_folder": r"\\server\share"},
            {"output_folder": good + r"\{nope}"}, {"filename_pattern": "{nope}"}, {"filename_pattern": "{title"},
            {"save_mode": "weird"}, {"save_mode": "ask", "output_folder": "relative"},
            {"filename_prefix": "a/b"}, {"filename_prefix": "", "filename_pattern": ""},
            {"also_print": True, "also_print_printer": ""},
        )
        for changes in cases:
            with self.assertRaises(ValueError, msg=changes):
                driver.save_settings(self.config_path, self._settings(**changes))
        self.assertFalse(self.config_path.exists())

    def test_subfolder_item_is_appended_without_separators(self):
        self.assertEqual(driver.with_trailing_separator(r"C:\MatrixPDF"), "C:\\MatrixPDF\\")
        self.assertEqual(driver.append_subfolder("C:\\MatrixPDF\\", "{time}"), "C:\\MatrixPDF\\{time}")
        self.assertEqual(driver.append_subfolder("C:\\MatrixPDF\\{year}", "{month}"), "C:\\MatrixPDF\\{year}{month}")
        self.assertEqual(driver.append_subfolder("", "{user}"), "C:\\MatrixPDF\\{user}")

    def _window(self, **kwargs):
        try:
            return driver.build_settings_window(self.config_path, **kwargs)
        except Exception as exc:  # 화면이 없는 환경
            self.skipTest(str(exc))

    def test_freshly_opened_window_adds_items_at_the_end(self):
        window = self._window()
        try:
            variables = window.variables
            self.assertEqual(variables["folder"].get(), "C:\\MatrixPDF\\")
            self.assertEqual((variables["prefix"].get(), variables["pattern"].get()), ("Cube_", "{datetime}"))
            window.add_subfolder("{time}")
            self.assertEqual(variables["folder"].get(), "C:\\MatrixPDF\\{time}")
            window.add_name_token("{title}")
            self.assertEqual(variables["pattern"].get(), "{datetime}{title}")
        finally:
            window.destroy()

    def test_full_rule_is_split_into_folder_start_and_items(self):
        split = driver.split_full_rule
        self.assertEqual(split("C:\\MatrixPDF\\Cube_{datetime}.pdf"), ("C:\\MatrixPDF\\", "Cube_", "{datetime}"))
        self.assertEqual(split("C:\\A\\{year}\\R-{title}_{date}.PDF"), ("C:\\A\\{year}\\", "R-", "{title}_{date}"))
        self.assertEqual(split("Cube_{datetime}"), ("", "Cube_", "{datetime}"))
        self.assertEqual(split("fixed.pdf"), ("", "fixed", ""))
        self.assertEqual(split("{title}.pdf"), ("", "", "{title}"))

    def test_full_field_and_the_three_fields_follow_each_other(self):
        window = self._window()
        try:
            variables = window.variables
            self.assertEqual(variables["full"].get(), "C:\\MatrixPDF\\Cube_{datetime}.pdf")
            variables["prefix"].set("Lab_")
            window.add_name_token("{title}")
            self.assertEqual(variables["full"].get(), "C:\\MatrixPDF\\Lab_{datetime}{title}.pdf")
            variables["full"].set("D:\\out\\{year}\\R-{date}.pdf")
            self.assertEqual(
                (variables["folder"].get(), variables["prefix"].get(), variables["pattern"].get()),
                ("D:\\out\\{year}\\", "R-", "{date}"))
            self.assertEqual(variables["full"].get(), "D:\\out\\{year}\\R-{date}.pdf")
            variables["mode"].set("ask")
            self.assertEqual(variables["full"].get(), "D:\\out\\{year}\\R-{date}.pdf")
            variables["full"].set("E:\\ask\\X_{time}.pdf")
            self.assertEqual(
                (variables["folder"].get(), variables["prefix"].get(), variables["pattern"].get()),
                ("E:\\ask\\", "X_", "{time}"))
        finally:
            window.destroy()

    def test_folder_with_items_is_shown_as_saved(self):
        self.config_path.write_text(json.dumps({"output_folder": "C:\\MatrixPDF\\{year}"}), encoding="utf-8")
        window = self._window()
        try:
            self.assertEqual(window.variables["folder"].get(), "C:\\MatrixPDF\\{year}")
        finally:
            window.destroy()

    def test_window_saves_and_reports(self):
        window = self._window()
        try:
            variables = window.variables
            self.assertEqual(variables["mode"].get(), "auto")
            variables["folder"].set(str(self.root / "out") + "\\{year}")
            variables["prefix"].set("R-")
            variables["pattern"].set("{title}_{date}")
            variables["color"].set("gray")
            variables["open"].set(True)
            for page in ("파일 이름", "품질", "출력", "정보", "저장"):
                window.show_page(page)
            window.update()
            expected = str(self.root / "out" / datetime.now().strftime("%Y") / "R-시험성적서_")
            self.assertTrue(variables["preview"].get().startswith(expected), variables["preview"].get())
            self.assertTrue(variables["preview"].get().endswith(".pdf"))
            self.assertTrue(window.save_settings())
            self.assertIn("저장했습니다", variables["message"].get())
            variables["folder"].set("relative")
            self.assertFalse(window.save_settings())
            self.assertIn("전체 경로", variables["message"].get())
            variables["folder"].set(str(self.root / "out") + "\\{year}")
            variables["mode"].set("ask")
            window.update()
            self.assertIn("저장 창", variables["preview_title"].get())
            self.assertTrue(variables["preview"].get().startswith(expected), variables["preview"].get())
            window.add_subfolder("{month}")
            self.assertTrue(variables["folder"].get().endswith("{year}{month}"))
        finally:
            window.destroy()
        saved = driver.load_config(self.config_path)
        self.assertEqual(saved.name_rule, "R-{title}_{date}")
        self.assertEqual((saved.color_mode, saved.open_after_save), ("gray", True))
        self.assertEqual(driver.render_folder(saved.output_folder, _values()), self.root / "out" / "2026")

    def test_install_window_starts_with_ask_unless_already_configured(self):
        window = self._window(install_mode=True)
        try:
            self.assertEqual(window.variables["mode"].get(), "ask")
        finally:
            window.destroy()
        self.config_path.write_text(json.dumps({"save_mode": "auto"}), encoding="utf-8")
        window = self._window(install_mode=True)
        try:
            self.assertEqual(window.variables["mode"].get(), "auto")
        finally:
            window.destroy()


class SetupWizardTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.config_path = Path(self.temp.name) / "config.json"

    def test_install_folder_rules(self):
        self.assertEqual(driver.normalize_install_folder("D:/Tools"), "D:\\Tools\\Matrix PDF-Driver")
        self.assertEqual(driver.normalize_install_folder("D:\\Tools\\matrix pdf-driver"), "D:\\Tools\\matrix pdf-driver")
        self.assertEqual(driver.check_install_folder(" D:\\Tools\\App "), Path("D:\\Tools\\App"))
        for bad in ("", "C:\\", "relative\\folder", "\\\\server\\share\\app", "C:\\a*b"):
            with self.assertRaises(ValueError, msg=bad):
                driver.check_install_folder(bad)

    def test_only_our_own_files_are_listed_for_removal(self):
        files = driver.program_files_in(Path("D:\\Tools\\App"))
        self.assertEqual(sorted(str(f) for f in files), sorted([
            "D:\\Tools\\App\\MatrixPdfDriver.exe", "D:\\Tools\\App\\gs\\gswin32c.exe", "D:\\Tools\\App\\gs\\gsdll32.dll",
            "D:\\Tools\\App\\LICENSE.txt", "D:\\Tools\\App\\THIRD-PARTY.txt"]))

    def _pump(self, window, until, seconds=10):
        deadline = time.time() + seconds
        while time.time() < deadline and not until():
            window.update()
            time.sleep(0.05)
        return until()

    def test_wizard_walks_through_license_folder_install_and_finish(self):
        calls = []

        def fake_install(config_path, output_folder=None, save_mode=None, filename_pattern=None, target_dir=None, language=None):
            calls.append((save_mode, target_dir))
            return driver.Config({})

        with mock.patch.object(driver, "install", side_effect=fake_install), \
                mock.patch.object(driver.subprocess, "Popen") as launched:
            try:
                window = driver.build_setup_window(self.config_path)
            except Exception as exc:  # 화면이 없는 환경
                self.skipTest(str(exc))
            try:
                self.assertEqual(window.step, "welcome")
                window.go_next()
                self.assertEqual(window.step, "license")
                window.go_next()
                self.assertEqual(window.step, "license")
                self.assertIn("동의", window.variables["message"].get())
                window.variables["agree"].set(True)
                window.go_next()
                self.assertEqual(window.step, "folder")
                window.go_back()
                self.assertEqual(window.step, "license")
                window.go_next()
                window.variables["folder"].set("C:\\")
                window.go_next()
                self.assertEqual(window.step, "folder")
                self.assertIn("폴더 이름", window.variables["message"].get())
                window.variables["folder"].set("D:\\Tools\\Matrix PDF-Driver")
                window.go_next()
                self.assertTrue(self._pump(window, lambda: window.step == "done"))
                self.assertTrue(window.installed)
                self.assertEqual(calls, [("ask", "D:\\Tools\\Matrix PDF-Driver")])
                window.variables["open_settings"].set(False)
                window.go_next()
            finally:
                try:
                    window.destroy()
                except Exception:
                    pass
            launched.assert_not_called()

    def test_failed_install_returns_to_folder_step_with_reason(self):
        self.config_path.write_text(json.dumps({"save_mode": "auto"}), encoding="utf-8")
        calls = []

        def failing_install(config_path, output_folder=None, save_mode=None, filename_pattern=None, target_dir=None, language=None):
            calls.append(save_mode)
            raise driver.SetupError("프린터를 등록하지 못했습니다.")

        with mock.patch.object(driver, "install", side_effect=failing_install):
            try:
                window = driver.build_setup_window(self.config_path)
            except Exception as exc:
                self.skipTest(str(exc))
            try:
                window.go_next()
                window.variables["agree"].set(True)
                window.go_next()
                window.variables["folder"].set("D:\\Tools\\Matrix PDF-Driver")
                window.go_next()
                self.assertTrue(self._pump(window, lambda: window.step == "folder"))
                self.assertIn("프린터를 등록하지 못했습니다", window.variables["message"].get())
                self.assertFalse(window.installed)
                self.assertEqual(calls, [None])
            finally:
                window.destroy()


class LanguageTests(unittest.TestCase):
    HANGUL = re.compile("[\uac00-\ud7a3]")

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.config_path = Path(self.temp.name) / "config.json"
        self.addCleanup(setattr, driver, "LANGUAGE", "ko")

    def test_english_is_the_default_and_korean_is_kept_as_written(self):
        self.assertEqual(driver.Config({}).language, "en")
        self.assertEqual(driver.Config({"language": "KO"}).language, "ko")
        self.assertEqual(driver.Config({"language": "fr"}).language, "en")
        driver.LANGUAGE = "en"
        self.assertEqual(driver.tr("저장"), "Save")
        self.assertEqual(driver.tr("표에 없는 말"), "표에 없는 말")
        driver.LANGUAGE = "ko"
        self.assertEqual(driver.tr("저장"), "저장")

    def test_every_translated_phrase_in_the_source_has_english(self):
        import ast
        source = Path(driver.__file__).read_text(encoding="utf-8")
        phrases = set()
        for node in ast.walk(ast.parse(source)):
            if isinstance(node, ast.Call) and getattr(node.func, "id", "") == "tr" and node.args \
                    and isinstance(node.args[0], ast.Constant):
                phrases.add(node.args[0].value)
        phrases |= {label for _token, label in driver.TOKENS} | set(driver.SETTINGS_PAGES)
        self.assertGreater(len(phrases), 100)
        self.assertEqual(sorted(phrases - set(driver.TRANSLATIONS)), [])
        for korean, english in driver.TRANSLATIONS.items():
            self.assertFalse(self.HANGUL.search(english), english)
            self.assertEqual(korean.count("%"), english.count("%"), korean)

    def _visible_texts(self, window):
        import tkinter as tk
        texts, stack = [], [window]
        while stack:
            widget = stack.pop()
            stack.extend(widget.winfo_children())
            for option in ("text", "values"):
                try:
                    value = widget.cget(option)
                except tk.TclError:
                    continue
                texts.extend(str(v) for v in value) if isinstance(value, (tuple, list)) else texts.append(str(value))
            if isinstance(widget, tk.Menu) and widget.index("end") is not None:
                texts.extend(str(widget.entrycget(index, "label")) for index in range(widget.index("end") + 1))
        return [t for t in texts if t and t != "한국어"]

    def test_settings_window_in_english_shows_no_korean(self):
        driver.LANGUAGE = "en"
        try:
            window = driver.build_settings_window(self.config_path)
        except Exception as exc:  # 화면이 없는 환경
            self.skipTest(str(exc))
        try:
            for mode in ("ask", "auto"):
                window.variables["mode"].set(mode)
                window.update()
                texts = self._visible_texts(window) + [window.title()]
                self.assertEqual([t for t in texts if self.HANGUL.search(t)], [])
            self.assertIn("Save mode", texts)
            self.assertIn("Language", texts)
            self.assertIn("About", texts)
            window.variables["folder"].set("relative")
            self.assertFalse(window.save_settings())
            self.assertIn("full path", window.variables["message"].get())
        finally:
            window.destroy()

    def test_setup_window_in_english_shows_no_korean(self):
        with mock.patch.object(driver, "install", return_value=driver.Config({})):
            try:
                window = driver.build_setup_window(self.config_path, "en")
            except Exception as exc:
                self.skipTest(str(exc))
            driver.LANGUAGE = "en"  # 시험에서는 set_language를 막아 두었으므로 직접 맞춘다
            window.destroy()
            window = driver.build_setup_window(self.config_path, "en")
            try:
                seen = []
                for _ in range(3):
                    seen += self._visible_texts(window)
                    window.variables["agree"].set(True)
                    window.go_next()
                deadline = time.time() + 10
                while time.time() < deadline and window.step != "done":
                    window.update()
                    time.sleep(0.05)
                seen += self._visible_texts(window) + [window.variables["status"].get(), window.title()]
                self.assertEqual([t for t in seen if self.HANGUL.search(t)], [])
                self.assertIn("I accept the license", seen)
                self.assertIn("Finish", seen)
            finally:
                window.destroy()

    def test_changing_language_in_settings_saves_it_and_asks_to_reopen(self):
        try:
            window = driver.build_settings_window(self.config_path)
        except Exception as exc:
            self.skipTest(str(exc))
        self.assertEqual(window.variables["language"].get(), "en")
        self.assertFalse(window.reopen)
        window.variables["language"].set("ko")
        self.assertTrue(window.save_settings())
        self.assertTrue(window.reopen)
        self.assertEqual(driver.load_config(self.config_path).language, "ko")

    def test_picking_language_on_first_setup_screen_asks_to_reopen_in_it(self):
        try:
            window = driver.build_setup_window(self.config_path)
        except Exception as exc:
            self.skipTest(str(exc))
        self.assertEqual(window.language, "en")
        window.language_box.current(1)
        window.change_language()
        self.assertEqual((window.language, window.reopen), ("ko", True))


class AfterSaveTests(unittest.TestCase):
    def _config(self, **raw):
        return driver.Config(dict({"printer_name": "Matrix PDF-Driver"}, **raw))

    def test_nothing_happens_when_nothing_is_ticked(self):
        with mock.patch.object(driver.os, "startfile", create=True) as opened, \
                mock.patch.object(driver, "print_pdf") as printed:
            driver.after_save(self._config(), Path("x.pdf"))
            time.sleep(0.2)
        opened.assert_not_called()
        printed.assert_not_called()

    def test_opens_and_prints_when_ticked(self):
        with mock.patch.object(driver.os, "startfile", create=True) as opened, \
                mock.patch.object(driver, "find_ghostscript", return_value=Path("gs.exe")), \
                mock.patch.object(driver, "print_pdf") as printed:
            driver.after_save(self._config(open_after_save=True, also_print=True, also_print_printer="Lab"), Path("x.pdf"))
            time.sleep(0.3)
        opened.assert_called_once_with("x.pdf")
        printed.assert_called_once_with(Path("gs.exe"), Path("x.pdf"), "Lab")

    def test_never_prints_back_into_itself(self):
        with mock.patch.object(driver, "print_pdf") as printed:
            driver.after_save(self._config(also_print=True, also_print_printer="Matrix PDF-Driver"), Path("x.pdf"))
            time.sleep(0.2)
        printed.assert_not_called()

    def test_print_failure_is_reported(self):
        reports = []
        with mock.patch.object(driver, "find_ghostscript", return_value=Path("gs.exe")), \
                mock.patch.object(driver, "print_pdf", side_effect=OSError("오프라인")):
            driver.after_save(self._config(also_print=True, also_print_printer="Lab"), Path("x.pdf"),
                              lambda headline, detail: reports.append((headline, detail)))
            time.sleep(0.3)
        self.assertEqual(reports, [("Lab 프린터로 출력하지 못했습니다", "오프라인")])

    def test_printer_list_leaves_out_this_printer(self):
        printers = driver.list_printers()
        self.assertIsInstance(printers, list)
        self.assertNotIn("Matrix PDF-Driver", printers)


class PendingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.pending = Path(self.temp.name) / "pending"
        self.pending.mkdir()

    def _add(self, stem, user="hong", age=0.0):
        pdf = self.pending / (stem + ".pdf")
        pdf.write_bytes(b"%PDF-1.7 test")
        meta = {"type": "ask", "name": "결과", "title": "결과", "user": user,
                "received_at": time.time() - age, "pdf": str(pdf)}
        path = self.pending / (stem + ".ask.json")
        path.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
        return path

    def test_only_one_helper_gets_the_job(self):
        path = self._add("a")
        first = driver.claim_pending(path, "hong")
        self.assertIsNotNone(first)
        self.assertIsNone(driver.claim_pending(path, "hong"))
        self.assertEqual(first["name"], "결과")

    def test_other_user_waits_then_may_take_it(self):
        fresh = self._add("fresh", user="kim")
        self.assertIsNone(driver.claim_pending(fresh, "hong"))
        self.assertIsNotNone(driver.claim_pending(fresh, "KIM"))
        old = self._add("old", user="kim", age=driver.OTHER_USER_GRACE_SECONDS + 5)
        self.assertIsNotNone(driver.claim_pending(old, "hong"))
        unknown = self._add("unknown", user="")
        self.assertIsNotNone(driver.claim_pending(unknown, "hong"))

    def test_saving_copies_and_cleans_up(self):
        meta = driver.claim_pending(self._add("a"), "hong")
        target = Path(self.temp.name) / "새 폴더" / "결과.pdf"
        driver.finish_pending(meta, target)
        self.assertEqual(target.read_bytes(), b"%PDF-1.7 test")
        self.assertEqual(list(self.pending.iterdir()), [])

    def test_discarding_cleans_up(self):
        meta = driver.claim_pending(self._add("a"), "hong")
        driver.finish_pending(meta, None)
        self.assertEqual(list(self.pending.iterdir()), [])

    def test_failed_save_keeps_the_job(self):
        meta = driver.claim_pending(self._add("a"), "hong")
        blocker = Path(self.temp.name) / "file"
        blocker.write_bytes(b"x")
        with self.assertRaises(OSError):
            driver.finish_pending(meta, blocker / "sub" / "x.pdf")
        self.assertTrue(Path(meta["pdf"]).exists())

    def test_abandoned_claim_returns_to_queue(self):
        meta = driver.claim_pending(self._add("a"), "hong")
        old = time.time() - 3600
        os.utime(meta["claimed"], (old, old))
        driver.release_stale_claims(self.pending)
        self.assertTrue((self.pending / "a.ask.json").exists())

    def test_notice_without_pdf_is_just_cleared(self):
        path = self.pending / "b.saved.json"
        path.write_text(json.dumps({"type": "saved", "user": "", "received_at": time.time(), "path": "x.pdf"}), encoding="utf-8")
        notice = driver.claim_pending(path, "hong")
        self.assertEqual(notice["type"], "saved")
        driver.finish_pending(notice)
        self.assertEqual(list(self.pending.iterdir()), [])


@unittest.skipUnless(driver.find_ghostscript(driver.Config({})), "Ghostscript가 vendor 폴더에 없음")
class EndToEndTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.output = root / "출력 폴더"
        self.data = root / "data"
        patcher = mock.patch.object(driver, "data_dir", return_value=self.data)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self.temp.cleanup)
        self.config_path = root / "config.json"
        self._start({"output_folder": str(self.output)})
        self.addCleanup(self._stop)

    def _start(self, raw):
        raw = dict(raw, port=_free_port(), log_dir=str(Path(self.temp.name) / "logs"))
        raw.setdefault("filename_pattern", "{datetime}_{title}")
        self.config_path.write_text(json.dumps(raw, ensure_ascii=False), encoding="utf-8")
        self.config = driver.load_config(self.config_path)
        self.server = driver.PrintServer(self.config, self.config_path)
        self.thread = threading.Thread(target=self.server.serve, daemon=True)
        self.thread.start()
        self._wait(lambda: driver.is_running(self.config.port))

    def _stop(self):
        self.server.stop()
        self.thread.join(timeout=30)

    def _wait(self, check, seconds=60):
        deadline = time.time() + seconds
        while time.time() < deadline:
            if check():
                return True
            time.sleep(0.1)
        return False

    def _send(self, payload: bytes):
        with socket.create_connection(("127.0.0.1", self.config.port), timeout=10) as conn:
            conn.sendall(payload)

    def _pdfs(self, folder=None):
        folder = folder or self.output
        return sorted(folder.glob("*.pdf")) if folder.exists() else []

    def test_print_job_becomes_pdf(self):
        self._send(b"\x1b%-12345X@PJL JOB\r\n" + SAMPLE_PS + b"\x04\x1b%-12345X")
        self.assertTrue(self._wait(lambda: len(self._pdfs()) == 1))
        pdf = self._pdfs()[0]
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF"))
        self.assertTrue(pdf.name.endswith("_Sample Report.pdf"))
        self.assertEqual(list(self.output.glob("*.tmp")), [])
        self.assertEqual(list(self.config.spool_dir.glob("*")), [])

    def test_ten_consecutive_jobs_all_arrive(self):
        for _ in range(10):
            self._send(SAMPLE_PS)
        self.assertTrue(self._wait(lambda: len(self._pdfs()) == 10))

    def test_empty_connection_is_ignored(self):
        self._send(b"")
        time.sleep(1)
        self.assertEqual(self._pdfs(), [])
        self.assertFalse(self.config.failed_dir.exists())

    def test_broken_and_non_postscript_jobs_go_to_failed(self):
        self._send(BROKEN_PS)
        self._send(b"just some text, not PostScript")
        self.assertTrue(self._wait(
            lambda: self.config.failed_dir.exists() and len(list(self.config.failed_dir.glob("*.ps"))) == 2
        ))
        self.assertEqual(len(list(self.config.failed_dir.glob("*.reason.txt"))), 2)
        self.assertEqual(self._pdfs(), [])

    def test_folder_and_name_rules_use_job_details(self):
        self._stop()
        self._start({"output_folder": str(self.output / "{user}" / "{year}"), "filename_pattern": "{counter:03d}_{title}"})
        self._send(SAMPLE_PS)
        self._send(SAMPLE_PS)
        folder = self.output / "tester" / datetime.now().strftime("%Y")
        self.assertTrue(self._wait(lambda: len(self._pdfs(folder)) == 2))
        self.assertEqual([p.name for p in self._pdfs(folder)], ["001_Sample Report.pdf", "002_Sample Report.pdf"])

    def test_ask_mode_hands_the_pdf_to_the_helper(self):
        self._stop()
        self._start({"output_folder": str(self.output / "{user}"), "save_mode": "ask"})
        self._send(SAMPLE_PS)
        pending = self.config.pending_dir
        self.assertTrue(self._wait(lambda: pending.exists() and len(list(pending.glob("*.json"))) == 1))
        self.assertEqual(self._pdfs(), [])
        meta = driver.claim_pending(next(pending.glob("*.json")), "tester")
        self.assertEqual(meta["user"], "tester")
        self.assertEqual(Path(meta["folder"]), self.output / "tester")
        self.assertTrue((self.output / "tester").is_dir())
        self.assertTrue(meta["name"].endswith("_Sample Report"))
        target = Path(self.temp.name) / "고른 곳" / "x.pdf"
        driver.finish_pending(meta, target)
        self.assertTrue(target.read_bytes().startswith(b"%PDF"))
        self.assertEqual(list(pending.iterdir()), [])

    def test_ticked_options_leave_notices_for_the_helper(self):
        self._stop()
        self._start({"output_folder": str(self.output), "show_progress": True, "color_mode": "gray"})
        self._send(SAMPLE_PS)
        self._send(BROKEN_PS)
        pending = self.config.pending_dir
        self.assertTrue(self._wait(lambda: pending.exists() and len(list(pending.glob("*.json"))) == 4))
        kinds = sorted(json.loads(p.read_text(encoding="utf-8"))["type"] for p in pending.glob("*.json"))
        self.assertEqual(kinds, ["failed", "saved", "start", "start"])
        saved = json.loads(next(pending.glob("*.saved.json")).read_text(encoding="utf-8"))
        self.assertEqual(Path(saved["path"]), self._pdfs()[0])
        self.assertEqual(saved["user"], "tester")

    def test_each_archive_format_is_marked_inside_the_file(self):
        korean_font = b"/OrigFontName <" + "맑은 고딕".encode("cp949").hex().encode("ascii") + b"> def\n"
        job = SAMPLE_PS.replace(b"%%EndComments\n", b"%%EndComments\n" + korean_font)
        for pdf_format, part in (("pdfa1", b"1"), ("pdfa2", b"2"), ("pdfa3", b"3"), ("pdf", None)):
            self._stop()
            folder = self.output / pdf_format
            self._start({"output_folder": str(folder), "pdf_format": pdf_format})
            self._send(job)
            self.assertTrue(self._wait(lambda: len(self._pdfs(folder)) == 1), pdf_format)
            content = self._pdfs(folder)[0].read_bytes()
            if part is None:
                self.assertNotIn(b"pdfaid:part", content)
            else:
                self.assertRegex(content, rb"pdfaid:part[^0-9]{0,4}" + part)
                self.assertRegex(content, rb"pdfaid:conformance[^A-Za-z]{0,4}B")
            self.assertEqual([p.name for p in self.config.spool_dir.iterdir()], [])

    def test_no_notices_when_nothing_is_ticked(self):
        self._send(SAMPLE_PS)
        self.assertTrue(self._wait(lambda: len(self._pdfs()) == 1))
        self.assertFalse(self.config.pending_dir.exists())

    def test_changed_settings_apply_to_next_job_without_restart(self):
        self._send(SAMPLE_PS)
        self.assertTrue(self._wait(lambda: len(self._pdfs()) == 1))
        moved = Path(self.temp.name) / "바뀐 폴더"
        time.sleep(1.1)
        driver.save_settings(self.config_path, {
            "save_mode": "auto", "output_folder": str(moved), "filename_prefix": "", "filename_pattern": "{title}"})
        self._send(SAMPLE_PS)
        self.assertTrue(self._wait(lambda: (moved / "Sample Report.pdf").exists()))
        self.assertEqual(len(self._pdfs()), 1)

    def test_leftover_job_is_converted_on_start(self):
        self._stop()
        self.config.spool_dir.mkdir(parents=True, exist_ok=True)
        (self.config.spool_dir / "20261008_000000_0001.job").write_bytes(SAMPLE_PS)
        (self.config.spool_dir / "half.part").write_bytes(b"%!PS")
        self._start({"output_folder": str(self.output)})
        self.assertTrue(self._wait(lambda: len(self._pdfs()) == 1))
        # 시작 확인용 빈 연결이 잠깐 만드는 .part가 사라질 때까지 기다린 뒤 본다.
        self.assertTrue(self._wait(lambda: not (self.config.spool_dir / "half.part").exists()))


if __name__ == "__main__":
    unittest.main()
