#!/usr/bin/env python3
"""
Excel / CSV 批量合并与字段标准化工具 V1

把一个目录里所有表格（.csv / .txt / .xlsx）合并成一张 CSV。

V1 做四件事：
  1. 批量读入     —— 扫一个目录，按文件名排序逐个读。每个 .csv/.txt 算一张表，
                     .xlsx 的每张工作表各算一张表。
  2. 列名标准化   —— 去首尾空白 + 全角字母数字转半角 + 忽略大小写；
                     再按别名表把同义列名归并成同一个名字。
  3. 按列名横向对齐 —— 输出列 = 所有输入文件列名的并集，按「第一次出现」的顺序排。
                     某个文件没有的列，那一格留空（不会填 0 或 NULL）。
                     跨文件归到同一列是正常的（不同行、不同格子，不覆盖不丢数据）。
  4. 值标准化     —— 每个单元格去首尾空白 + 全角字母数字转半角。
                     整行都是空白的行会被跳过。

两条关于「不许猜」的硬规则：

  规则一：字段别名只认显式配置。
    工具不会去猜「名字」和「姓名」是不是同一个字段 —— 只有你在别名表里写下的
    对应关系才会生效。默认的 aliases.txt 是一条都没启用的空模板。

  规则二：同一张表里，两列归到同一个列名 = 直接报错停机。
    比如一个文件里同时有「姓名」和「名字」两列，而别名表把「名字」指向了「姓名」，
    那同一行的两个值就要抢同一个输出格子，必然丢数据。
    这种情况下工具不会自动改名、不会静默合并、也不会丢掉其中一列，
    而是打印清楚哪两列撞了、怎么改，然后整任务中止，不产出任何文件。
    （跨文件归到同一列不算撞车：那是不同行，不会互相覆盖。）

全角转半角只处理字母、数字和全角空格，不处理全角标点：
  「１２３ＡＢＣ　」→「123ABC 」
  「北京，上海」   →「北京，上海」（逗号是中文正文的一部分，动了就是改数据，不是标准化）

安全约定（写死在代码里，不靠自觉）：
  - 永远不覆盖任何输入文件。
  - 输出文件已经存在时，必须显式加 --force 才会覆盖。
  - 所有检查（读文件、查公式、查输出冲突）都发生在写盘之前，
    被拒绝时磁盘上不会多出任何文件。
  - 输入目录里的表格会被全部读一遍；读不了的会明确报错，不会静默跳过。

不支持 .xlsm —— 本工具要重建表，宏（VBA）没办法安全保留，硬做会静默丢宏。
不支持 .xls / .xlsb —— V1 不解析这两种老格式；目录里见到会跳过并警告。

Excel 额外有一道闸门：只接受「无公式的纯数据表」。
  发现真正的公式单元格 → 整个任务中止，不产出任何文件
  发现内容是 "=" 开头的文本单元格 → 同样中止
  （理由：本工具重建表时会丢公式和原始类型，硬做会算错）
  CSV / TXT 不受这道闸门影响。

用法示例：
  python3 merge.py 数据目录
  python3 merge.py 数据目录 -o 汇总.csv
  python3 merge.py 数据目录 --no-source-col --report 报告.txt
  python3 merge.py 数据目录 --aliases 我的别名表.txt

看详细帮助：python3 merge.py -h
"""

import argparse
import csv
import datetime
import sys
from collections import namedtuple
from pathlib import Path

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

CSV_SUFFIXES = {".csv", ".txt"}
EXCEL_SUFFIXES = {".xlsx"}

# 目录里见到这些后缀就跳过并警告（不是静默忽略）
UNSUPPORTED_TABLE_SUFFIXES = {".xlsm", ".xls", ".xlsb"}

# 读 CSV 时按这个顺序试编码：先 UTF-8（含 BOM），再 GB18030（GBK 的超集，兼容老中文表格）
CSV_INPUT_ENCODINGS = ("utf-8-sig", "gb18030")
# 写 CSV 用带 BOM 的 UTF-8，这样 Windows 版 Excel 双击打开不会中文乱码
CSV_OUTPUT_ENCODING = "utf-8-sig"
CSV_OUTPUT_SUFFIX = ".csv"

# 来源列：每一行标记它来自哪个文件（.xlsx 多表时是「文件名#表名」）
SOURCE_COLUMN_NAME = "来源文件"
SOURCE_SHEET_SEP = "#"
# 内部用的保留键，以 \x00 开头，真实列名永远撞不上它
SOURCE_KEY = "\x00source"

UNNAMED_COLUMN_TEMPLATE = "未命名列{}"
MERGED_NAME_SUFFIX = "_merged"
DEFAULT_ALIASES_FILENAME = "aliases.txt"

XLSM_HINT = (
    "不支持 .xlsm：本工具要重建表，宏（VBA）没办法安全保留，硬做会静默丢宏。\n"
    "请先在 Excel 里「另存为 .xlsx」，再处理另存出来的文件。"
)
XLSM_HINT_SHORT = "本工具要重建表，宏（VBA）没办法安全保留，硬做会静默丢宏；请先另存为 .xlsx"

# 一份列规格：display 是输出表头里显示的名字，key 是内部用来对齐的唯一键
ColumnSpec = namedtuple("ColumnSpec", "display key")


class MergeError(Exception):
    """本工具能对用户说清楚的问题。捕获后只打印提示，不打印 Python 堆栈。"""


# --------------------------------------------------------------------------
# 文本标准化
# --------------------------------------------------------------------------

def to_halfwidth(text):
    """
    全角字母、数字转半角；全角空格转半角空格。其它字符原样保留。

    故意不处理全角标点（，。：；！？（）【】等）：它们是中文正文的一部分，
    转换等于改了数据内容。只转字母数字，才能既让「１００１」和「1001」对得上，
    又不会把「北京，上海」变成「北京,上海」。
    """
    chars = []
    for ch in text:
        code = ord(ch)
        if (
            0xFF10 <= code <= 0xFF19  # ０-９
            or 0xFF21 <= code <= 0xFF3A  # Ａ-Ｚ
            or 0xFF41 <= code <= 0xFF5A  # ａ-ｚ
        ):
            chars.append(chr(code - 0xFEE0))
        elif code == 0x3000:  # 全角空格
            chars.append(" ")
        else:
            chars.append(ch)
    return "".join(chars)


def clean_text_value(value):
    """单元格值的标准化：全角字母数字转半角，然后去掉首尾空白。非字符串原样返回。"""
    if isinstance(value, str):
        return to_halfwidth(value).strip()
    return value


def normalize_key(name):
    """
    把列名压成「对齐用的键」：标准化之后统一转小写。

    只影响内部比对的键，不影响输出表头里显示的名字 ——
    所以第一个出现的写法（比如 “Name”）会被保留下来显示，不会被强制改成小写。
    """
    return clean_text_value(str(name)).casefold()


def cell_to_text(value):
    """把单元格的值变成要写进 CSV 的字符串。"""
    if value is None:
        return ""
    if isinstance(value, str):
        return clean_text_value(value)
    if isinstance(value, bool):
        # 注意：bool 要在 int 之前判断，因为 True 也是 int 的一种
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime.datetime):
        # Excel 里的日期常常读出来是「夜里 0 点整」的 datetime，
        # 直接 str() 会得到 “2026-09-28 00:00:00”，很吵；这种就只写日期。
        if value.time() == datetime.time(0, 0):
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, (datetime.date, datetime.time)):
        return value.isoformat()
    return str(value)


def is_blank_row(values):
    """一整行（已经标准化过）全是空字符串，就算空白行。"""
    return all(value == "" for value in values)


# --------------------------------------------------------------------------
# 别名表
# --------------------------------------------------------------------------

def parse_aliases_text(text, source_label):
    """
    解析别名表。每行一条，格式：

        标准名 = 别名1, 别名2

    - `#` 开头或空行忽略
    - 分隔符支持英文逗号、中文逗号、顿号，随便混用
    - 标准名自己也算自己的别名，所以不用重复写一遍
    - 同一个别名被指到两个不同的标准名 → 报错（宁可不做，也不猜你想用哪个）

    返回 {对齐键: 标准名} 的字典。
    """
    mapping = {}
    for lineno, raw_line in enumerate(text.splitlines(), 1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            raise MergeError(
                "{} 第 {} 行看不懂：{}\n"
                "每条别名要写成「标准名 = 别名1, 别名2」。".format(source_label, lineno, raw_line)
            )
        canonical_text, _, alias_text = line.partition("=")
        canonical = clean_text_value(canonical_text)
        if not canonical:
            raise MergeError("{} 第 {} 行等号左边是空的，缺标准名。".format(source_label, lineno))

        names = [canonical]
        for chunk in alias_text.replace("，", ",").replace("、", ",").split(","):
            name = clean_text_value(chunk)
            if name:
                names.append(name)

        for name in names:
            key = normalize_key(name)
            previous = mapping.get(key)
            if previous is not None and previous != canonical:
                raise MergeError(
                    "{} 里「{}」同时被指到「{}」和「{}」两个标准名，会对不上。\n"
                    "请把这一行删掉一个，或者改用不同的别名。".format(
                        source_label, name, previous, canonical
                    )
                )
            mapping[key] = canonical
    return mapping


def load_aliases(explicit_path, disabled):
    """
    决定到底用哪份别名表，返回 (映射字典, 说明文字)。

    explicit_path 由 --aliases 指定；没指定时，默认找脚本旁边的 aliases.txt。
    找不到就只有正向的标准化（去空白/全角转半角/忽略大小写），不做同义归并 ——
    这一点会在结果摘要里写明，不会让你以为别名生效了。
    """
    if disabled:
        return {}, "已关闭（--no-aliases）"

    if explicit_path is not None:
        path = Path(explicit_path).expanduser()
        if not path.exists():
            raise MergeError("--aliases 指的文件不存在：{}".format(path))
        if not path.is_file():
            raise MergeError("--aliases 指的不是一个普通文件：{}".format(path))
        label = str(path)
    else:
        path = Path(__file__).resolve().parent / DEFAULT_ALIASES_FILENAME
        if not path.exists():
            return {}, "未使用（没有找到 {}）".format(path)
        label = str(path)

    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        raise MergeError(
            "{} 不是 UTF-8 编码，读不了。\n"
            "别名表请用 UTF-8 保存（Linux 下用系统自带编辑器另存即可）。".format(label)
        )
    mapping = parse_aliases_text(text, label)
    if not mapping:
        return {}, "{} 里没有任何有效行".format(label)
    # 报「标准名」的个数，而不是键的个数 —— 一条 "姓名 = 名字" 展开成两个键，
    # 但对用户来说那就是定义了 1 个字段。
    return mapping, "{}（{} 个标准名）".format(label, len(set(mapping.values())))


# --------------------------------------------------------------------------
# 列规格：把一张表的表头算成「显示名 + 对齐键」
# --------------------------------------------------------------------------

def describe_collision(source_label, group, target_key):
    """
    拼出「同一张表里两列撞到同一个列名」的报错文本。

    必须讲清楚三件事：哪两列撞了、为什么危险、怎么改。
    用户看到这段话应该能自己动手修，不需要来问。
    """
    lines = [
        "{} 里有 {} 列会归到同一个列名「{}」，无法安全合并。".format(
            source_label, len(group), target_key
        ),
        "任务已中止，没有生成任何输出文件。",
        "",
        "撞车的列：",
    ]
    from_alias = False
    for index, raw_text, display in group:
        note = ""
        if raw_text == "":
            note = "（表头是空的，按位置取的名字）"
        elif normalize_key(raw_text) != normalize_key(display):
            note = "（表头原文是「{}」，被别名表归并到「{}」）".format(raw_text, display)
            from_alias = True
        else:
            note = "（表头原文就是「{}」）".format(raw_text)
        lines.append("    第 {} 列：{} {}".format(index + 1, display, note))

    lines.append("")
    lines.append(
        "为什么不能继续：同一行里这几个值要抢同一个输出格子，"
        "硬合并的结果一定是丢数据。本工具不会自动改名，也不会替你猜哪一列该留。"
    )
    lines.append("")
    lines.append("怎么改（改一个就行）：")
    lines.append("    1. 改源文件：把这几列的表头改成互不相同的名字，再重新合并。")
    if from_alias:
        lines.append(
            "    2. 改别名表：把造成归并的那条别名删掉，让这几列保持各自独立。"
        )
    return "\n".join(lines)


def build_header_specs(raw_rows, aliases, source_label):
    """
    算出这张表的列规格列表。

    raw_rows 的第 0 行当表头。宽度取「表头宽度」和「最宽的一行」的较大值：
      - 表头以外的多余列 → 叫「未命名列N」（不丢数据）
      - 表头里空的/缺的列 → 同样叫「未命名列N」

    如果同一张表里有两列标准化后归到同一个列名，直接抛 MergeError 中止整个任务：
    那种情况下同一行的两个值会抢同一个格子，必然丢数据，工具不做自动改名或静默合并。
    """
    if not raw_rows:
        return []

    header = raw_rows[0]
    width = max([len(row) for row in raw_rows] + [len(header)])

    # 先把每一列的「原文 + 归一后的显示名」算出来，再统一查重
    planned = []  # [(列号从0起, 表头原文, 最终显示名)]
    for index in range(width):
        raw_name = header[index] if index < len(header) else None
        text = cell_to_text(raw_name)
        if text == "":
            display = UNNAMED_COLUMN_TEMPLATE.format(index + 1)
        else:
            # 只认别名表里写过的对应关系，不做任何猜测
            display = aliases.get(normalize_key(text), text)
        planned.append((index, text, display))

    grouped = {}
    for index, text, display in planned:
        grouped.setdefault(normalize_key(display), []).append((index, text, display))

    for key, group in sorted(grouped.items()):
        if len(group) > 1:
            raise MergeError(describe_collision(source_label, group, key))

    return [ColumnSpec(display, normalize_key(display)) for _, _, display in planned]


def build_union_specs(all_specs, with_source_column):
    """
    把所有表的列规格合并成输出表的列顺序：取并集，按第一次出现的先后排。

    同一个键在不同文件里显示名不同时（比如一个叫 Name、一个叫 name），
    用第一次出现的那个写法显示。
    """
    union = []
    seen = set()
    if with_source_column:
        union.append(ColumnSpec(SOURCE_COLUMN_NAME, SOURCE_KEY))
        seen.add(SOURCE_KEY)

    for specs in all_specs:
        for spec in specs:
            if spec.key in seen:
                continue
            seen.add(spec.key)
            union.append(spec)
    return union


# --------------------------------------------------------------------------
# 读文件
# --------------------------------------------------------------------------

def read_csv_rows(path):
    """读 CSV / TXT，返回二维列表。会自动试 UTF-8 和 GB18030 两种编码。"""
    last_error = None
    for encoding in CSV_INPUT_ENCODINGS:
        try:
            with path.open("r", encoding=encoding, newline="") as f:
                return [row for row in csv.reader(f)]
        except UnicodeDecodeError as exc:
            last_error = exc
    raise MergeError(
        "读不了这个文件：{}。按 {} 依次尝试都解码失败。\n原始报错：{}".format(
            path, " / ".join(CSV_INPUT_ENCODINGS), last_error
        )
    )


def load_openpyxl():
    """按需导入 openpyxl，没装的话给一条能直接照抄的命令。"""
    try:
        import openpyxl
    except ImportError:
        raise MergeError(
            "目录里有 Excel 文件，读写需要 openpyxl，但当前 Python 里没有这个库。\n"
            "  Ubuntu 安装（先跑这一条）：sudo apt install python3-openpyxl\n"
            "  装完重新运行本工具即可。只想先合 CSV 的话可以把 .xlsx 挪出目录。"
        )
    return openpyxl


def read_excel_workbook(path):
    """
    读一遍 Excel，同时干两件事：
      1. 把每张工作表转成二维列表
      2. 记下两类「危险单元格」：真正的公式单元格、内容是 "=" 开头的文本单元格

    返回 ([(表名, 二维列表), ...], 公式单元格列表, 危险文本单元格列表)

    这里故意不用 values_only，而是逐个 cell 对象读：只有 cell.data_type 才能区分
    「真正的公式」和「内容是 = 开头的纯文本」。两者的 cell.value 长得一模一样，
    但前者重写后还是公式，后者会被 openpyxl 误当成公式。

    关于文件句柄：这个函数自己开文件、自己关（with 语句），不把开文件的事交给 openpyxl。
    原因见下面注释，别改回 load_workbook(path)。
    """
    openpyxl = load_openpyxl()

    # 为什么必须由我们自己持有文件句柄：
    #   1) openpyxl 的 Workbook.close() 在「非 read_only」模式下是空操作。
    #      openpyxl 自己的注释写着 "Only affects read-only and write-only modes"，
    #      所以读完再调 workbook.close() 是关不掉任何东西的（看着像关了，其实没有）。
    #   2) 如果直接 load_workbook(路径字符串)，底层 ZipFile 由 openpyxl 负责关闭。
    #      openpyxl 读工作表时会留下没有关闭的成员文件对象，使 zipfile 的内部引用计数
    #      (_fileRefCnt) 不归零，于是 ZipFile.close() 也不会真正关掉那个文件描述符，
    #      表现就是 Python 报 ResourceWarning: unclosed file ... .xlsx。
    #   3) 把「已经打开的二进制文件对象」交给 load_workbook 时，zipfile 认为句柄是
    #      外部传进来的（_filePassed），不会去关它 —— 关闭的责任完整地落回我们手里，
    #      用 with 就能保证：正常返回、抛异常、被安全检查拦截，任何路径下都会关闭。
    try:
        handle = path.open("rb")
    except OSError as exc:
        raise MergeError("打不开这个 Excel：{}\n原始报错：{}".format(path, exc))

    with handle:
        try:
            workbook = openpyxl.load_workbook(handle, data_only=False)
        except Exception as exc:
            # openpyxl 的异常类型有好几种（文件损坏、其实是 .xls、被占用等），统一转成一句人话
            raise MergeError("打不开这个 Excel：{}\n原始报错：{}".format(path, exc))
        try:
            tables = []
            formulas = []
            equal_texts = []
            for sheet in workbook.worksheets:
                rows = []
                for row in sheet.iter_rows():
                    values = []
                    for cell in row:
                        value = cell.value
                        values.append(value)
                        if isinstance(value, str) and value.startswith("="):
                            if cell.data_type == "f":
                                formulas.append((sheet.title, cell.coordinate, value))
                            else:
                                equal_texts.append((sheet.title, cell.coordinate, value))
                    rows.append(values)
                tables.append((sheet.title, rows))
            return tables, formulas, equal_texts
        finally:
            # 顺手调一下（read_only 模式下它才真的有用）；真正关文件的是上面的 with
            workbook.close()


def format_hits(hits, limit=5):
    """把「危险单元格」列表排版成几行，最多列 5 个，避免刷屏。"""
    lines = [
        "    [{0}] {1}: {2}".format(sheet, coord, value) for sheet, coord, value in hits[:limit]
    ]
    if len(hits) > limit:
        lines.append("    …… 另外还有 {} 处".format(len(hits) - limit))
    return lines


def check_excel_is_plain_data(problems):
    """
    Excel 安全闸门：只放行「无公式的纯数据表」。

    problems 是 [(文件名, 公式列表, 危险文本列表), ...]。
    只要有任何一个文件不合格，整个任务中止，不产出任何输出文件。
    这是有意的：批量合并里少了一个文件，结果看上去却「成功」了，比直接报错更危险。
    """
    bad = [item for item in problems if item[1] or item[2]]
    if not bad:
        return

    parts = ["Excel 安全检查没通过，整个任务已取消，没有生成任何输出文件。"]
    for name, formulas, equal_texts in bad:
        parts.append("文件：{}".format(name))
        if formulas:
            parts.append(
                "  发现真正的公式单元格，共 {} 处：\n{}".format(
                    len(formulas), "\n".join(format_hits(formulas))
                )
            )
        if equal_texts:
            parts.append(
                "  发现原本是文本、但内容以 = 开头的单元格，共 {} 处：\n{}".format(
                    len(equal_texts), "\n".join(format_hits(equal_texts))
                )
            )
    parts.append(
        "怎么改：在 Excel 里把公式选择性粘贴成数值（或先把那些 = 开头的文本改掉），"
        "另存为新的 .xlsx，再重新合并。"
    )
    raise MergeError("\n".join(parts))


# --------------------------------------------------------------------------
# 扫目录
# --------------------------------------------------------------------------

def scan_directory(directory):
    """
    扫目录，返回 (要处理的文件列表, 警告列表, 被忽略的非表格文件列表)。

    按文件名排序，保证同样的目录每次合并出来的结果完全一样（可复现）。
    只扫这一层，不递归子目录。

    注意：这里不再接受「输出/报告文件要排除掉」这种参数。
    排除法会掩盖真正的危险情况 —— 要写的路径如果本身就是源表，
    它一被排除，后面的安全检查就再也看不见它了，结果就是覆盖原始数据。
    写入目标的安全性由 check_write_targets() 单独硬闸门把关。
    """
    entries = sorted(
        (p for p in directory.iterdir() if p.is_file()),
        key=lambda p: (p.name.casefold(), p.name),
    )

    targets = []
    warnings = []
    ignored = []

    for path in entries:
        suffix = path.suffix.lower()
        if suffix in CSV_SUFFIXES or suffix in EXCEL_SUFFIXES:
            if path.stem.endswith(MERGED_NAME_SUFFIX):
                warnings.append(
                    "跳过 {}：文件名以「{}」结尾，像是上一次合并的产物。"
                    "如果它其实是你的原始数据，请改个名字再来。".format(path.name, MERGED_NAME_SUFFIX)
                )
                continue
            targets.append(path)
        elif suffix in UNSUPPORTED_TABLE_SUFFIXES:
            if suffix == ".xlsm":
                warnings.append("跳过 {}：{}".format(path.name, XLSM_HINT_SHORT))
            else:
                warnings.append("跳过 {}：V1 不解析 {} 这种老格式".format(path.name, suffix))
        else:
            ignored.append(path.name)

    return targets, warnings, ignored


def list_existing_table_files(directory):
    """
    列出输入目录里现有的、后缀属于表格的全部文件的真实路径（resolve 之后）。

    安全检查专用，故意不看 scan_directory() 的筛选结果：
    那个结果是给「要合并什么」用的，会被「*_merged.csv 跳过」之类的规则筛过，
    拿筛过的名单去查「谁不能被覆盖」会漏掉正好被筛掉的那个文件。
    这里收的是目录里所有表格后缀的文件，一个不漏。
    """
    existing = set()
    for entry in directory.iterdir():
        if entry.is_file() and entry.suffix.lower() in (CSV_SUFFIXES | EXCEL_SUFFIXES):
            existing.add(entry.resolve())
    return existing


def check_write_targets(directory, output_path, report_path):
    """
    硬闸门：要写入的路径绝不能是输入目录里现有的表格文件。

    这条闸门和 --force 无关：--force 只放宽「覆盖自己以前产出的结果」，
    永远不放宽「覆盖原始数据」。

    用 resolve() 之后的真实路径比对，所以符号链接也绕不过去：
    只要最终落到某个输入表上，就会被抓住。
    这个函数在 check_paths() 里被调用，也就是在所有 mkdir / 写文件之前。
    """
    existing = list_existing_table_files(directory)
    if not existing:
        return

    targets = [("输出文件（-o）", output_path)]
    if report_path is not None:
        targets.append(("报告文件（--report）", report_path))

    for label, path in targets:
        if path.resolve() in existing:
            raise MergeError(
                "{}{} 指向的是输入目录里的原始表格文件：\n"
                "    要写的路径：{}\n"
                "    解析后的真实文件：{}\n"
                "本工具承诺永不覆盖任何输入文件，所以任务已中止，没有改写任何文件。\n"
                "注意：这个拒绝和 --force 无关 —— --force 只允许覆盖你自己以前的输出，\n"
                "      绝不会允许覆盖原始数据。\n"
                "怎么改（二选一）：\n"
                "    1. 换一个写入路径，别落在输入目录里（默认输出就在输入目录旁边）；\n"
                "    2. 如果那个文件确实是你不要了的结果，先手动删掉或移走它。".format(
                    label, "" if not path.is_symlink() else "（是符号链接）",
                    path, path.resolve(),
                )
            )


# --------------------------------------------------------------------------
# 合并
# --------------------------------------------------------------------------

def convert_table(raw_rows, specs, source_label):
    """
    把一张原始二维表转成「按对齐键存值的字典」列表。

    跳过整行空白的行（Excel 里表尾常有一堆空行）。
    行比表头短 → 缺的格子留空；行比表头长 → 多出来的列已经在 specs 里了。
    """
    rows = []
    blank_skipped = 0
    for raw in raw_rows[1:]:
        values = [cell_to_text(cell) for cell in raw]
        if is_blank_row(values):
            blank_skipped += 1
            continue
        row = {SOURCE_KEY: source_label}
        for index, spec in enumerate(specs):
            row[spec.key] = values[index] if index < len(values) else ""
        rows.append(row)
    return rows, blank_skipped


def format_stats(read_rows, contributed, blank_skipped):
    return "读入 {} 行 → 贡献 {} 行；跳过空白行 {} 行".format(read_rows, contributed, blank_skipped)


def build_summary(
    directory,
    output_path,
    alias_note,
    with_source_column,
    union_specs,
    tables,
    total,
    detail_lines,
    warnings,
    ignored_names,
):
    """拼出给人看的摘要文本（终端打印和 --report 用的是同一份）。"""
    lines = [
        "合并完成",
        "  输入目录：{}".format(directory),
        "  结果文件：{}".format(output_path),
        "  别名映射：{}".format(alias_note),
        "  来源列：{}".format("已添加「{}」".format(SOURCE_COLUMN_NAME) if with_source_column else "已关闭"),
        "  汇总：读入 {} 张表、{} 行 → 输出 {} 行；"
        "输出 {} 列；跳过空白行 {} 行".format(
            len(tables), total["read_rows"], total["contributed"], len(union_specs), total["blank_skipped"]
        ),
        "  列顺序：{}".format(" → ".join(spec.display for spec in union_specs)),
        "  分表统计：",
    ]
    lines.extend(detail_lines)

    if total["contributed"] == 0:
        lines.append("  注意：一张数据行都没读到，检查一下这些文件是不是只有表头。")

    if ignored_names:
        shown = "、".join(ignored_names[:5])
        more = "" if len(ignored_names) <= 5 else " 等 {} 个".format(len(ignored_names))
        lines.append("  提示：目录里另有 {} 个非表格文件已忽略（{}{}）".format(
            len(ignored_names), shown, more
        ))

    if warnings:
        lines.append("  警告：")
        lines.extend("    - {}".format(text) for text in warnings)

    return "\n".join(lines)


def run(directory, output_path, report_path, aliases, alias_note, with_source_column):
    warnings = []

    targets, scan_warnings, ignored_names = scan_directory(directory)
    warnings.extend(scan_warnings)
    if not targets:
        raise MergeError(
            "目录里没有可合并的表格：{}\n支持的后缀：{}".format(
                directory, "、".join(sorted(CSV_SUFFIXES | EXCEL_SUFFIXES))
            )
        )

    # 先把所有文件读进来、把 Excel 安全检查做完，任何拒绝都发生在写盘之前
    tables = []
    excel_problems = []
    for path in targets:
        if path.suffix.lower() in CSV_SUFFIXES:
            tables.append((path.name, read_csv_rows(path)))
            continue
        sheets, formulas, equal_texts = read_excel_workbook(path)
        excel_problems.append((path.name, formulas, equal_texts))
        if len(sheets) == 1:
            # 只有一张表就直接用文件名，别再多加一个 #表名 显得啰嗦
            tables.append((path.name, sheets[0][1]))
        else:
            for sheet_title, rows in sheets:
                tables.append(("{}{}{}".format(path.name, SOURCE_SHEET_SEP, sheet_title), rows))
    check_excel_is_plain_data(excel_problems)

    all_specs = []
    for label, rows in tables:
        all_specs.append(build_header_specs(rows, aliases, label))

    # 输入里自带「来源文件」列的话，和工具要多加的来源列会重名。
    # 输出表头出现两个一模一样的名字，人是分不清哪个是哪个的 —— 同样按撞车处理。
    if with_source_column:
        source_key = normalize_key(SOURCE_COLUMN_NAME)
        clashed = [
            label
            for (label, _), specs in zip(tables, all_specs)
            for spec in specs
            if spec.key == source_key
        ]
        if clashed:
            raise MergeError(
                "这些输入表里本来就有叫「{}」的列：{}\n"
                "工具默认还要在最左边加一个同名的来源列，输出表头会出现两个「{}」，"
                "看的人分不清哪个是哪个，所以任务已中止，没有生成任何输出文件。\n"
                "怎么改（改一个就行）：\n"
                "    1. 加 --no-source-col，不要工具那个来源列；\n"
                "    2. 把源文件里这一列改成别的名字。".format(
                    SOURCE_COLUMN_NAME, "、".join(clashed), SOURCE_COLUMN_NAME
                )
            )

    union_specs = build_union_specs(all_specs, with_source_column)

    output_rows = []
    total = {"read_rows": 0, "contributed": 0, "blank_skipped": 0}
    detail_lines = []
    for (label, rows), specs in zip(tables, all_specs):
        converted, blank_skipped = convert_table(rows, specs, label)
        output_rows.extend(converted)
        total["read_rows"] += max(len(rows) - 1, 0)
        total["contributed"] += len(converted)
        total["blank_skipped"] += blank_skipped
        detail_lines.append(
            "    [{}] {}".format(
                label, format_stats(max(len(rows) - 1, 0), len(converted), blank_skipped)
            )
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    write_csv(output_path, union_specs, output_rows)

    summary = build_summary(
        directory,
        output_path,
        alias_note,
        with_source_column,
        union_specs,
        tables,
        total,
        detail_lines,
        warnings,
        ignored_names,
    )
    print(summary)

    if report_path is not None:
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(summary + "\n", encoding="utf-8")
        print("\n报告已写入：{}".format(report_path))


def write_csv(path, union_specs, rows):
    """
    写出结果 CSV。

    第一行是表头，后面每行按 union_specs 的顺序取值，缺的格子写空字符串。
    用带 BOM 的 UTF-8：Windows 版 Excel 双击打开不会中文乱码。
    """
    with path.open("w", encoding=CSV_OUTPUT_ENCODING, newline="") as f:
        writer = csv.writer(f)
        writer.writerow([spec.display for spec in union_specs])
        keys = [spec.key for spec in union_specs]
        for row in rows:
            writer.writerow([row.get(key, "") for key in keys])


# --------------------------------------------------------------------------
# 参数、路径检查、入口
# --------------------------------------------------------------------------

def default_output_path(directory):
    """
    默认把结果写到输入目录的「旁边」，而不是目录里面。

    这样第二次运行时不会把上一次的结果当成输入再合一遍。
    """
    return directory.parent / (directory.name + MERGED_NAME_SUFFIX + CSV_OUTPUT_SUFFIX)


def build_parser():
    parser = argparse.ArgumentParser(
        description=(
            "Excel / CSV 批量合并与字段标准化工具 V1：扫一个目录，"
            "按列名对齐合并成一张 CSV，永不覆盖原始文件。"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python3 merge.py 数据目录\n"
            "  python3 merge.py 数据目录 -o 汇总.csv\n"
            "  python3 merge.py 数据目录 --no-source-col --report 报告.txt\n"
            "  python3 merge.py 数据目录 --aliases 我的别名表.txt\n"
        ),
    )
    parser.add_argument("directory", help="要合并的目录（只扫这一层，不递归子目录）")
    parser.add_argument(
        "-o",
        "--output",
        help="结果写到哪；默认：在输入目录旁边生成「目录名_merged.csv」",
    )
    parser.add_argument("--report", help="把本次统计结果额外写一份到这个文本文件")
    parser.add_argument(
        "--aliases",
        help="别名表文件路径；默认自动找脚本旁边的 {}".format(DEFAULT_ALIASES_FILENAME),
    )
    parser.add_argument("--no-aliases", action="store_true", help="完全不用别名表")
    parser.add_argument(
        "--no-source-col",
        action="store_true",
        help="不要在结果最左边加「{}」列".format(SOURCE_COLUMN_NAME),
    )
    parser.add_argument("--force", action="store_true", help="输出文件已存在时，允许覆盖它")
    return parser


def check_paths(directory, output_path, report_path):
    """所有会写到磁盘之前的路径检查都放这里，任何一个不满足就直接报错退出。"""
    if not directory.exists():
        raise MergeError("找不到输入目录：{}".format(directory))
    if not directory.is_dir():
        raise MergeError(
            "这不是一个目录：{}\nV1 只接受目录（批量合并的入口是一个文件夹）。".format(directory)
        )
    if output_path.suffix.lower() != CSV_OUTPUT_SUFFIX:
        raise MergeError(
            "V1 只输出 {}，现在给的是「{}」。本工具不做格式转换。".format(
                CSV_OUTPUT_SUFFIX, output_path.suffix or "（没有后缀）"
            )
        )
    if output_path.resolve() == directory.resolve():
        raise MergeError("输出路径是个目录：{}，请用 -o 指定一个文件名。".format(output_path))
    if report_path is not None and report_path.resolve() == output_path.resolve():
        raise MergeError("--report 的路径和输出文件重了，换一个路径。")

    # 最关键的一条闸门，且必须在任何写盘动作之前：
    # 要写的路径不能是输入目录里现有的表格文件（原始数据）。与 --force 无关。
    check_write_targets(directory, output_path, report_path)


def main(argv=None):
    args = build_parser().parse_args(argv)

    directory = Path(args.directory).expanduser()
    if directory.exists():
        directory = directory.resolve()
    output_path = (
        Path(args.output).expanduser() if args.output else default_output_path(directory)
    )
    report_path = Path(args.report).expanduser() if args.report else None

    try:
        check_paths(directory, output_path, report_path)
        if output_path.exists() and not args.force:
            raise MergeError(
                "输出文件已存在：{}\n"
                "本工具默认不覆盖任何已有文件。确实要覆盖就加 --force，或者换个 -o 路径。".format(
                    output_path
                )
            )
        aliases, alias_note = load_aliases(args.aliases, args.no_aliases)
        run(
            directory,
            output_path,
            report_path,
            aliases,
            alias_note,
            with_source_column=not args.no_source_col,
        )
    except MergeError as exc:
        print("错误：{}".format(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
