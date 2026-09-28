#!/usr/bin/env python3
"""
table-merger 自检脚本

跑法（在项目目录里）：
    python3 selftest.py

只用标准库 unittest，不需要 pytest，也不需要联网。
所有临时文件都建在系统临时目录里，跑完自己删掉，不会碰你的数据。
处理 .xlsx 的用例需要 openpyxl；没装的话那部分会自动跳过（不会算失败）。

这里测的都是「约定」，不是「实现细节」——
比如「永远不覆盖源文件」「输出已存在必须 --force」「含公式的 Excel 一律拒绝」，
这些是工具的安全底线，改代码时如果被改坏了，这里必须报警。

其中两条是小环明确要求的硬规则，专门有一组用例守着，不许改松：
  1. 字段别名只认显式配置 —— 不配别名表时，「姓名」和「名字」必须保持成两个独立的列，
     工具不许猜它们是不是同一个字段。
  2. 同一张表里两列归到同一个列名 —— 必须报错停机，且不产出任何文件；
     不许自动改名（曾经有过自动改名成「姓名(2)」的行为，已按用户要求删除），
     不许静默合并、不许丢掉其中一列。
     （跨文件归到同一列是允许的，那是别名表的用途，不同行不会互相覆盖。）
"""

import contextlib
import csv
import datetime
import io
import shutil
import tempfile
import unittest
from pathlib import Path

import merge

try:
    import openpyxl

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# --------------------------------------------------------------------------
# 小工具
# --------------------------------------------------------------------------

def run_main(argv):
    """跑一次 merge.main，把打印内容抓下来，返回 (退出码, 输出文本)。"""
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
        code = merge.main(argv)
    return code, buffer.getvalue()


def read_csv_rows(path):
    """读工具写出来的结果文件（去 BOM 后按 CSV 解析）。"""
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        return [row for row in csv.reader(f)]


def make_xlsx(path, sheets):
    """用 openpyxl 造一个测试用 Excel。sheets 是 [(表名, 二维列表), ...]。"""
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets:
        sheet = workbook.create_sheet(title=title)
        for row in rows:
            sheet.append(row)
    workbook.save(path)


class TempCase(unittest.TestCase):
    """每个用例一个临时目录，收尾自动删干净。"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="table-merger-selftest-"))
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)
        self.data = self.tmp / "data"
        self.data.mkdir()

    def write(self, name, text, encoding="utf-8"):
        path = self.data / name
        path.write_text(text, encoding=encoding)
        return path


# --------------------------------------------------------------------------
# 文本标准化
# --------------------------------------------------------------------------

class TestTextHelpers(unittest.TestCase):
    def test_fullwidth_alnum_converted(self):
        self.assertEqual(merge.to_halfwidth("１２３"), "123")
        self.assertEqual(merge.to_halfwidth("ＡＢｃ"), "ABc")
        self.assertEqual(merge.to_halfwidth("１００１号"), "1001号")

    def test_fullwidth_space_converted(self):
        self.assertEqual(merge.to_halfwidth("张　三"), "张 三")

    def test_fullwidth_punctuation_untouched(self):
        # 中文标点是数据的一部分，转了就是改数据，这里守死这条线
        self.assertEqual(merge.to_halfwidth("北京，上海。"), "北京，上海。")
        self.assertEqual(merge.to_halfwidth("（备注）"), "（备注）")

    def test_clean_text_value_strips_and_converts(self):
        self.assertEqual(merge.clean_text_value(" １２３ \n"), "123")
        self.assertEqual(merge.clean_text_value(None), None)
        self.assertEqual(merge.clean_text_value(7), 7)

    def test_normalize_key_ignores_case_and_width(self):
        self.assertEqual(merge.normalize_key(" Name "), "name")
        self.assertEqual(merge.normalize_key("ＮＡＭＥ"), "name")
        self.assertEqual(merge.normalize_key("姓名"), "姓名")

    def test_cell_to_text(self):
        self.assertEqual(merge.cell_to_text(None), "")
        self.assertEqual(merge.cell_to_text(" 值 "), "值")
        self.assertEqual(merge.cell_to_text(True), "TRUE")
        self.assertEqual(merge.cell_to_text(False), "FALSE")
        self.assertEqual(merge.cell_to_text(1001), "1001")
        self.assertEqual(merge.cell_to_text(3.5), "3.5")
        self.assertEqual(
            merge.cell_to_text(datetime.datetime(2026, 9, 28, 0, 0)), "2026-09-28"
        )
        self.assertEqual(
            merge.cell_to_text(datetime.datetime(2026, 9, 28, 13, 5, 6)),
            "2026-09-28 13:05:06",
        )
        self.assertEqual(merge.cell_to_text(datetime.date(2026, 9, 28)), "2026-09-28")

    def test_is_blank_row(self):
        self.assertTrue(merge.is_blank_row([""]))
        self.assertTrue(merge.is_blank_row([]))
        self.assertFalse(merge.is_blank_row(["0"]))
        self.assertFalse(merge.is_blank_row(["", "x"]))


# --------------------------------------------------------------------------
# 别名表
# --------------------------------------------------------------------------

class TestAliases(unittest.TestCase):
    def test_parse_mixed_separators_and_comments(self):
        text = (
            "# 注释\n"
            "\n"
            "姓名 = 名字，员工姓名、人员姓名\n"
            "部门 = 所属部门\n"
        )
        mapping = merge.parse_aliases_text(text, "测试")
        self.assertEqual(mapping[merge.normalize_key("名字")], "姓名")
        self.assertEqual(mapping[merge.normalize_key("人员姓名")], "姓名")
        # 标准名自己也算自己的别名
        self.assertEqual(mapping[merge.normalize_key("姓名")], "姓名")
        self.assertEqual(mapping[merge.normalize_key("所属部门")], "部门")
        self.assertEqual(mapping[merge.normalize_key("部门")], "部门")

    def test_case_insensitive_alias(self):
        mapping = merge.parse_aliases_text("Date = DATE, date\n", "测试")
        self.assertEqual(mapping["date"], "Date")

    def test_conflict_raises(self):
        with self.assertRaises(merge.MergeError):
            merge.parse_aliases_text("姓名 = 名字\n部门 = 名字\n", "测试")

    def test_missing_equals_raises(self):
        with self.assertRaises(merge.MergeError):
            merge.parse_aliases_text("姓名 名字\n", "测试")

    def test_empty_canonical_raises(self):
        with self.assertRaises(merge.MergeError):
            merge.parse_aliases_text("= 名字\n", "测试")

    def test_load_explicit_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.txt"
            path.write_text("姓名 = 名字\n", encoding="utf-8")
            mapping, note = merge.load_aliases(str(path), False)
            self.assertEqual(mapping[merge.normalize_key("名字")], "姓名")
            self.assertIn("1 个标准名", note)

    def test_load_missing_explicit_file_raises(self):
        with self.assertRaises(merge.MergeError):
            merge.load_aliases("/nonexistent/aliases.txt", False)

    def test_disabled(self):
        mapping, note = merge.load_aliases(None, True)
        self.assertEqual(mapping, {})
        self.assertIn("--no-aliases", note)


# --------------------------------------------------------------------------
# 列规格与并集
# --------------------------------------------------------------------------

class TestHeaderSpecs(unittest.TestCase):
    def specs(self, rows, aliases=None):
        return merge.build_header_specs(rows, aliases or {}, "测试表")

    def test_basic(self):
        specs = self.specs([["姓名", "部门"], ["张三", "技术部"]])
        self.assertEqual([s.display for s in specs], ["姓名", "部门"])
        self.assertEqual([s.key for s in specs], ["姓名", "部门"])

    def test_display_keeps_first_spelling_but_key_ignores_case(self):
        specs = self.specs([["Name"], ["x"]])
        self.assertEqual(specs[0].display, "Name")
        self.assertEqual(specs[0].key, "name")

    def test_unnamed_column(self):
        specs = self.specs([["姓名", None], ["张三", "x"]])
        self.assertEqual(specs[1].display, merge.UNNAMED_COLUMN_TEMPLATE.format(2))

    def test_blank_header_cell_is_unnamed(self):
        specs = self.specs([["", "部门"], ["a", "b"]])
        self.assertEqual(specs[0].display, "未命名列1")

    def test_duplicate_literal_header_is_refused(self):
        # 同一张表里两列都叫「姓名」→ 同一行两个值抢一个格子，必须报错停机
        with self.assertRaises(merge.MergeError) as ctx:
            self.specs([["姓名", "姓名"], ["a", "b"]])
        message = str(ctx.exception)
        self.assertIn("无法安全合并", message)
        self.assertIn("第 1 列", message)
        self.assertIn("第 2 列", message)

    def test_alias_collision_inside_one_table_is_refused(self):
        aliases = merge.parse_aliases_text("姓名 = 名字\n", "测试")
        with self.assertRaises(merge.MergeError) as ctx:
            self.specs([["姓名", "名字"], ["a", "b"]], aliases)
        message = str(ctx.exception)
        self.assertIn("无法安全合并", message)
        self.assertIn("别名表", message)

    def test_case_only_difference_is_also_refused(self):
        with self.assertRaises(merge.MergeError):
            self.specs([["Name", "name"], ["a", "b"]])

    def test_no_file_is_produced_when_columns_collide(self):
        # 报错必须发生在写盘之前：跑完磁盘上不能多出任何文件
        with tempfile.TemporaryDirectory() as tmp:
            data = Path(tmp) / "data"
            data.mkdir()
            (data / "a.csv").write_text("姓名,名字\n张三,小三\n", encoding="utf-8")
            alias_file = Path(tmp) / "al.txt"
            alias_file.write_text("姓名 = 名字\n", encoding="utf-8")
            output = Path(tmp) / "out.csv"
            code, out = run_main([str(data), "-o", str(output), "--aliases", str(alias_file)])
            self.assertEqual(code, 1)
            self.assertIn("无法安全合并", out)
            self.assertFalse(output.exists())

    def test_extra_columns_beyond_header_get_unnamed_names(self):
        specs = self.specs([["a", "b"], ["1", "2", "3"]])
        self.assertEqual([s.display for s in specs], ["a", "b", "未命名列3"])

    def test_alias_applied(self):
        aliases = merge.parse_aliases_text("姓名 = 名字\n", "测试")
        specs = self.specs([["名字"], ["张三"]], aliases)
        self.assertEqual(specs[0].display, "姓名")
        self.assertEqual(specs[0].key, "姓名")

    def test_no_alias_config_means_no_merging_of_similar_names(self):
        # 不配别名时，「姓名」和「名字」就是两个不同的列 —— 工具不猜语义
        specs = self.specs([["姓名", "名字"], ["a", "b"]])
        self.assertEqual([s.display for s in specs], ["姓名", "名字"])

    def test_empty_table(self):
        self.assertEqual(self.specs([]), [])


class TestUnionSpecs(unittest.TestCase):
    def test_first_appearance_order_and_dedup(self):
        a = merge.build_header_specs([["姓名", "部门"], ["x", "y"]], {}, "a")
        b = merge.build_header_specs([["部门", "工号"], ["x", "y"]], {}, "b")
        union = merge.build_union_specs([a, b], with_source_column=False)
        self.assertEqual([s.display for s in union], ["姓名", "部门", "工号"])

    def test_source_column_comes_first(self):
        a = merge.build_header_specs([["姓名"], ["x"]], {}, "a")
        union = merge.build_union_specs([a], with_source_column=True)
        self.assertEqual(union[0].display, merge.SOURCE_COLUMN_NAME)
        self.assertEqual(union[0].key, merge.SOURCE_KEY)
        self.assertEqual([s.display for s in union], [merge.SOURCE_COLUMN_NAME, "姓名"])

    def test_source_column_can_be_off(self):
        a = merge.build_header_specs([["姓名"], ["x"]], {}, "a")
        union = merge.build_union_specs([a], with_source_column=False)
        self.assertEqual([s.display for s in union], ["姓名"])

    def test_same_target_across_files_is_allowed(self):
        # 跨文件归到同一列是别名表的目的：不同行，不覆盖也不丢数据
        aliases = merge.parse_aliases_text("姓名 = 名字\n", "测试")
        a = merge.build_header_specs([["姓名"], ["张三"]], aliases, "a")
        b = merge.build_header_specs([["名字"], ["李四"]], aliases, "b")
        union = merge.build_union_specs([a, b], with_source_column=False)
        self.assertEqual([s.display for s in union], ["姓名"])


# --------------------------------------------------------------------------
# 端到端：CSV
# --------------------------------------------------------------------------

class TestMergeCsv(TempCase):
    def test_basic_merge_by_column_name(self):
        a = self.write("a.csv", "姓名 , 部门 ,工号\n 张三 , 技术部 ,1001\n")
        b = self.write("b.csv", "工号,姓名,在职年限\n1002,李四,５\n")
        before_a = a.read_bytes()
        before_b = b.read_bytes()

        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0, out)

        output = self.tmp / "data_merged.csv"
        self.assertTrue(output.exists())
        rows = read_csv_rows(output)
        self.assertEqual(
            rows[0], ["来源文件", "姓名", "部门", "工号", "在职年限"]
        )
        self.assertEqual(rows[1], ["a.csv", "张三", "技术部", "1001", ""])
        self.assertEqual(rows[2], ["b.csv", "李四", "", "1002", "5"])

        # 源文件一个字节都不能变
        self.assertEqual(a.read_bytes(), before_a)
        self.assertEqual(b.read_bytes(), before_b)

    def test_output_has_utf8_bom(self):
        self.write("a.csv", "姓名\n张三\n")
        run_main([str(self.data), "--no-aliases"])
        raw = (self.tmp / "data_merged.csv").read_bytes()
        self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))

    def test_default_output_is_sibling_of_input_dir(self):
        self.write("a.csv", "姓名\n张三\n")
        run_main([str(self.data), "--no-aliases"])
        self.assertTrue((self.tmp / "data_merged.csv").exists())
        self.assertEqual(list(self.data.iterdir()), [self.data / "a.csv"])

    def test_explicit_output_and_nested_dir(self):
        self.write("a.csv", "姓名\n张三\n")
        target = self.tmp / "out" / "deep" / "结果.csv"
        code, _ = run_main([str(self.data), "-o", str(target), "--no-aliases"])
        self.assertEqual(code, 0)
        self.assertTrue(target.exists())

    def test_refuses_to_overwrite_existing_output(self):
        self.write("a.csv", "姓名\n张三\n")
        output = self.tmp / "data_merged.csv"
        output.write_text("旧内容", encoding="utf-8")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("--force", out)
        self.assertEqual(output.read_text(encoding="utf-8"), "旧内容")

    def test_force_allows_overwrite(self):
        self.write("a.csv", "姓名\n张三\n")
        output = self.tmp / "data_merged.csv"
        output.write_text("旧内容", encoding="utf-8")
        code, _ = run_main([str(self.data), "--no-aliases", "--force"])
        self.assertEqual(code, 0)
        self.assertNotEqual(output.read_text(encoding="utf-8"), "旧内容")

    def test_files_are_processed_in_name_order(self):
        self.write("b.csv", "姓名\n后写的\n")
        self.write("a.csv", "姓名\n先写的\n")
        run_main([str(self.data), "--no-aliases"])
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[1][0], "a.csv")
        self.assertEqual(rows[2][0], "b.csv")

    def test_same_result_on_second_run(self):
        self.write("a.csv", "姓名,部门\n张三,技术部\n")
        self.write("b.xlsm-not", "x")  # 无关文件，确认不影响结果
        out1 = self.tmp / "one.csv"
        out2 = self.tmp / "two.csv"
        run_main([str(self.data), "-o", str(out1), "--no-aliases"])
        run_main([str(self.data), "-o", str(out2), "--no-aliases"])
        self.assertEqual(out1.read_bytes(), out2.read_bytes())

    def test_blank_rows_skipped(self):
        self.write("a.csv", "a,b\n1,2\n,\n\n,\n")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(len(rows), 2)
        self.assertIn("跳过空白行 3 行", out)

    def test_short_row_padded_and_long_row_keeps_data(self):
        self.write("a.csv", "a,b\n9\n1,2,3\n")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[0], ["来源文件", "a", "b", "未命名列3"])
        self.assertEqual(rows[1], ["a.csv", "9", "", ""])
        self.assertEqual(rows[2], ["a.csv", "1", "2", "3"])

    def test_gb18030_input(self):
        self.write("gb.csv", "姓名,部门\n王五,销售部\n", encoding="gb18030")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[1], ["gb.csv", "王五", "销售部"])

    def test_no_source_column(self):
        self.write("a.csv", "姓名\n张三\n")
        run_main([str(self.data), "--no-aliases", "--no-source-col"])
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[0], ["姓名"])

    def test_alias_end_to_end(self):
        self.write("a.csv", "姓名\n张三\n")
        self.write("b.csv", "名字\n李四\n")
        alias_file = self.tmp / "my_aliases.txt"
        alias_file.write_text("姓名 = 名字\n", encoding="utf-8")
        code, out = run_main(
            [str(self.data), "--aliases", str(alias_file)]
        )
        self.assertEqual(code, 0)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[0], ["来源文件", "姓名"])
        self.assertEqual(rows[1], ["a.csv", "张三"])
        self.assertEqual(rows[2], ["b.csv", "李四"])

    def test_without_alias_same_column_stays_separate(self):
        self.write("a.csv", "姓名\n张三\n")
        self.write("b.csv", "名字\n李四\n")
        run_main([str(self.data), "--no-aliases"])
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[0], ["来源文件", "姓名", "名字"])

    def test_comment_only_aliases_file_maps_nothing(self):
        # 默认的 aliases.txt 就是这种空模板：一条有效行都没有 → 不做任何字段归并
        self.write("a.csv", "姓名\n张三\n")
        self.write("b.csv", "名字\n李四\n")
        alias_file = self.tmp / "空模板.txt"
        alias_file.write_text("# 全是注释\n# 姓名 = 名字\n", encoding="utf-8")
        code, out = run_main([str(self.data), "--aliases", str(alias_file)])
        self.assertEqual(code, 0)
        self.assertIn("没有任何有效行", out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[0], ["来源文件", "姓名", "名字"])

    def test_txt_counts_as_table(self):
        self.write("a.txt", "姓名\n张三\n")
        code, _ = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[1], ["a.txt", "张三"])

    def test_non_table_files_ignored(self):
        self.write("a.csv", "姓名\n张三\n")
        self.write("说明.md", "这不是表格")
        self.write("备注.txt.bak", "也不是")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        self.assertIn("非表格文件已忽略", out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(len(rows), 2)

    def test_xlsm_skipped_with_warning_and_run_continues(self):
        self.write("a.csv", "姓名\n张三\n")
        self.write("老表.xlsm", "随便什么内容，反正不会被打开")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        self.assertIn("跳过 老表.xlsm", out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(len(rows), 2)

    def test_previous_merged_output_not_eaten(self):
        self.write("a.csv", "姓名\n张三\n")
        self.write("old_merged.csv", "姓名\n上一次的结果\n")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        self.assertIn("上一次合并的产物", out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(len(rows), 2)

    def test_report_inside_input_dir_is_refused_on_second_run(self):
        # 第一次：报告文件还不存在，允许写进输入目录
        # 第二次：它已经存在、且后缀是表格（.txt）→ 硬闸门拒绝
        # 旧版本是把它从扫描里排除掉继续跑，那属于「用排除法掩盖冲突」，
        # 会把真正的危险藏起来（比如 -o 正好指向源表），已按要求删除该机制。
        self.write("a.csv", "姓名\n张三\n")
        report = self.data / "报告.txt"
        code, _ = run_main(
            [str(self.data), "-o", str(self.tmp / "one.csv"), "--no-aliases",
             "--report", str(report)]
        )
        self.assertEqual(code, 0)
        self.assertTrue(report.exists())
        before = report.read_bytes()

        code, out = run_main(
            [str(self.data), "-o", str(self.tmp / "two.csv"), "--no-aliases",
             "--report", str(report)]
        )
        self.assertEqual(code, 1)
        self.assertIn("原始表格文件", out)
        self.assertFalse((self.tmp / "two.csv").exists())
        self.assertEqual(report.read_bytes(), before)  # 报告那份文件一个字节没变

    def test_report_written_outside_dir(self):
        self.write("a.csv", "姓名\n张三\n")
        report = self.tmp / "统计报告.txt"
        code, out = run_main(
            [str(self.data), "--no-aliases", "--report", str(report)]
        )
        self.assertEqual(code, 0)
        text = report.read_text(encoding="utf-8")
        self.assertIn("合并完成", text)
        self.assertIn("结果文件", text)


# --------------------------------------------------------------------------
# 硬闸门：写入目标绝不能是输入目录里的原始表格
# --------------------------------------------------------------------------

class TestWriteTargetGate(TempCase):
    """
    README 承诺「永远不覆盖任何输入文件」，这一组用例就是守这句话的。

    历史 bug（已修）：run() 曾经在扫目录之前就把 output_path / report_path
    塞进 excluded_paths，于是「要写的文件恰好是源表」这种情况会先从输入集合里
    消失，后续所有安全检查都看不见它，最后被 write_csv() / write_text() 覆盖。
    修法是：扫描不再做排除，另设 check_write_targets() 硬闸门，
    且闸门与 --force 无关。
    """

    def setUp(self):
        super().setUp()
        self.a = self.write("a.csv", "姓名,部门\n张三,技术部\n")
        self.b = self.write("b.csv", "姓名,部门\n李四,销售部\n")
        self.c = self.write("c.txt", "姓名\n王五\n")
        self.original = {path: path.read_bytes() for path in (self.a, self.b, self.c)}

    def assert_sources_untouched(self):
        for path, raw in self.original.items():
            self.assertEqual(path.read_bytes(), raw, "源文件 {} 被改写了！".format(path.name))

    def test_A_output_on_existing_input_csv_refused_even_with_force(self):
        code, out = run_main(
            [str(self.data), "-o", str(self.a), "--force", "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertIn("永不覆盖", out)
        self.assertIn("原始表格文件", out)
        self.assert_sources_untouched()

    def test_B_report_on_existing_input_csv_refused(self):
        target = self.tmp / "result.csv"
        code, out = run_main(
            [str(self.data), "-o", str(target), "--report", str(self.a), "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertIn("原始表格文件", out)
        self.assertFalse(target.exists())
        self.assert_sources_untouched()

    def test_B2_report_on_existing_input_txt_refused(self):
        # .txt 也是本工具认的表格后缀，同样不许被报告覆盖
        target = self.tmp / "result.csv"
        code, out = run_main(
            [str(self.data), "-o", str(target), "--report", str(self.c), "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertIn("原始表格文件", out)
        self.assertFalse(target.exists())
        self.assert_sources_untouched()

    def test_B3_report_refused_even_with_force(self):
        target = self.tmp / "result.csv"
        code, out = run_main(
            [str(self.data), "-o", str(target), "--report", str(self.a),
             "--force", "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertFalse(target.exists())
        self.assert_sources_untouched()

    def test_C_default_output_still_works(self):
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0, out)
        self.assertTrue((self.tmp / "data_merged.csv").exists())
        self.assert_sources_untouched()

    def test_D_force_still_overwrites_a_real_existing_output(self):
        # --force 的本职工作没被误伤：覆盖自己以前的输出，仍然有效
        target = self.tmp / "结果.csv"
        target.write_text("这是上一次的输出，不是源表", encoding="utf-8")

        code, _ = run_main([str(self.data), "-o", str(target), "--no-aliases"])
        self.assertEqual(code, 1)  # 没有 --force：拒绝覆盖
        self.assertEqual(target.read_text(encoding="utf-8"), "这是上一次的输出，不是源表")

        code, out = run_main([str(self.data), "-o", str(target), "--no-aliases", "--force"])
        self.assertEqual(code, 0, out)  # 有 --force：允许覆盖自己的旧输出
        self.assertEqual(len(read_csv_rows(target)), 4)  # 表头 + a/b/c 各一行
        self.assert_sources_untouched()

    def test_E_symlink_to_input_file_refused(self):
        # 符号链接不能成为绕道：链接解析后的真实文件是源表，就必须拒绝
        link = self.tmp / "看起来无害.csv"
        try:
            link.symlink_to(self.a)
        except (OSError, NotImplementedError):
            self.skipTest("当前文件系统不支持创建符号链接")

        code, out = run_main(
            [str(self.data), "-o", str(link), "--force", "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertIn("符号链接", out)
        self.assert_sources_untouched()

    def test_E2_symlinked_report_to_input_file_refused(self):
        link = self.tmp / "报告别名.txt"
        try:
            link.symlink_to(self.c)
        except (OSError, NotImplementedError):
            self.skipTest("当前文件系统不支持创建符号链接")

        code, out = run_main(
            [str(self.data), "-o", str(self.tmp / "out.csv"), "--report", str(link),
             "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assert_sources_untouched()

    def test_merged_suffix_file_is_also_protected(self):
        # 以 _merged 结尾的文件扫目录时会被跳过，但它仍然是目录里既有的文件。
        # 如果它被当成写入目标，「跳过」就变成了覆盖它的通道 —— 同样必须拒绝。
        old = self.write("old_merged.csv", "姓名\n这是原始数据\n")
        before = old.read_bytes()
        code, out = run_main(
            [str(self.data), "-o", str(old), "--force", "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertEqual(old.read_bytes(), before)

    def test_refusal_happens_before_any_mkdir(self):
        # 拒绝必须发生在建目录 / 写文件之前：连 -o 的父目录都不许被创建出来
        target = self.data / "新建目录" / "结果.csv"
        code, out = run_main(
            [str(self.data), "-o", str(target), "--report", str(self.a), "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertFalse((self.data / "新建目录").exists())
        self.assert_sources_untouched()


# --------------------------------------------------------------------------
# 端到端：拒绝各种不该做的事情
# --------------------------------------------------------------------------

class TestRefusals(TempCase):
    def test_missing_directory(self):
        code, out = run_main([str(self.tmp / "不存在"), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("找不到输入目录", out)

    def test_input_must_be_directory(self):
        path = self.write("a.csv", "姓名\n张三\n")
        code, out = run_main([str(path), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("只接受目录", out)

    def test_empty_directory(self):
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("没有可合并的表格", out)

    def test_directory_with_only_non_tables(self):
        self.write("说明.md", "x")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("没有可合并的表格", out)

    def test_output_must_be_csv(self):
        self.write("a.csv", "姓名\n张三\n")
        code, out = run_main([str(self.data), "-o", str(self.tmp / "x.xlsx"), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("只输出 .csv", out)

    def test_report_path_conflicts_with_output(self):
        self.write("a.csv", "姓名\n张三\n")
        out_path = self.tmp / "同一份.csv"
        code, out = run_main(
            [str(self.data), "-o", str(out_path), "--report", str(out_path), "--no-aliases"]
        )
        self.assertEqual(code, 1)
        self.assertIn("重了", out)

    def test_header_only_file_is_ok_but_warns(self):
        self.write("a.csv", "姓名,部门\n")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0)
        self.assertIn("一张数据行都没读到", out)

    def test_source_column_clash_is_refused(self):
        # 输入自带「来源文件」列 → 和工具要加的那一列重名，输出会看不清，必须停机
        self.write("a.csv", "来源文件,姓名\nA,张三\n")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn(merge.SOURCE_COLUMN_NAME, out)
        self.assertIn("--no-source-col", out)
        self.assertFalse((self.tmp / "data_merged.csv").exists())

    def test_source_column_clash_ok_when_source_column_disabled(self):
        self.write("a.csv", "来源文件,姓名\nA,张三\n")
        code, out = run_main([str(self.data), "--no-aliases", "--no-source-col"])
        self.assertEqual(code, 0, out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[0], ["来源文件", "姓名"])
        self.assertEqual(rows[1], ["A", "张三"])


# --------------------------------------------------------------------------
# 端到端：Excel
# --------------------------------------------------------------------------

@unittest.skipUnless(HAS_OPENPYXL, "没装 openpyxl，跳过 Excel 相关用例")
class TestMergeExcel(TempCase):
    def test_single_sheet_label_is_file_name(self):
        make_xlsx(self.data / "c.xlsx", [("Sheet1", [["姓名", "部门"], ["张三", "技术部"]])])
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0, out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[1], ["c.xlsx", "张三", "技术部"])

    def test_multi_sheet_label_has_sheet_name(self):
        make_xlsx(
            self.data / "c.xlsx",
            [("一月", [["姓名"], ["张三"]]), ("二月", [["姓名"], ["李四"]])],
        )
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0, out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[1], ["c.xlsx#一月", "张三"])
        self.assertEqual(rows[2], ["c.xlsx#二月", "李四"])

    def test_csv_and_excel_merge_together(self):
        self.write("a.csv", "工号,姓名\n1001,张三\n")
        make_xlsx(self.data / "c.xlsx", [("S", [["姓名", "工号", "备注"], ["李四", 1002, "新入职"]])])
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0, out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[0], ["来源文件", "工号", "姓名", "备注"])
        self.assertEqual(rows[1], ["a.csv", "1001", "张三", ""])
        self.assertEqual(rows[2], ["c.xlsx", "1002", "李四", "新入职"])

    def test_excel_date_cell_written_as_date(self):
        make_xlsx(
            self.data / "d.xlsx",
            [("S", [["日期"], [datetime.datetime(2026, 9, 28, 0, 0)]])],
        )
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 0, out)
        rows = read_csv_rows(self.tmp / "data_merged.csv")
        self.assertEqual(rows[1][1], "2026-09-28")

    def test_formula_is_refused_and_no_output_written(self):
        make_xlsx(self.data / "f.xlsx", [("S", [["合计"], ["=1+1"]])])
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("安全检查没通过", out)
        self.assertIn("公式", out)
        self.assertFalse((self.tmp / "data_merged.csv").exists())

    def test_equal_text_cell_is_refused_and_no_output_written(self):
        path = self.data / "g.xlsx"
        make_xlsx(path, [("S", [["备注"], ["x"]])])

        # 读的时候由这里持有文件句柄。别改成 openpyxl.load_workbook(path)：
        # openpyxl 的 Workbook.close() 在非 read_only 模式下是空操作
        # （它自己的注释写着 "Only affects read-only and write-only modes"），
        # 加上读工作表时它自己不关成员文件对象，直接传路径就会漏一个
        # 打开的 .xlsx 文件描述符（ResourceWarning: unclosed file）。
        # 把已经打开的二进制句柄交进去，关闭责任就完整地落在这个 with 上。
        with path.open("rb") as handle:
            workbook = openpyxl.load_workbook(handle)
        try:
            cell = workbook.active["A2"]
            cell.value = "=不是公式只是文本"
            cell.data_type = "s"
            workbook.save(path)
        finally:
            # 顺手调一下（read_only 模式下它才真的有用）
            workbook.close()

        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("安全检查没通过", out)
        self.assertFalse((self.tmp / "data_merged.csv").exists())

    def test_good_file_does_not_rescue_bad_file(self):
        self.write("good.csv", "姓名\n张三\n")
        make_xlsx(self.data / "bad.xlsx", [("S", [["x"], ["=SUM(1,2)"]])])
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("bad.xlsx", out)
        self.assertFalse((self.tmp / "data_merged.csv").exists())

    def test_unreadable_xlsx_fails_loudly(self):
        (self.data / "broken.xlsx").write_text("这根本不是 Excel", encoding="utf-8")
        code, out = run_main([str(self.data), "--no-aliases"])
        self.assertEqual(code, 1)
        self.assertIn("打不开这个 Excel", out)
        self.assertIn("broken.xlsx", out)


if __name__ == "__main__":
    unittest.main(verbosity=2)
