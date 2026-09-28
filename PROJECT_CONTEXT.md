# PROJECT CONTEXT

> 给 AI 助手读的项目上下文。新会话先看这份文件，不要重新问一遍已经定下来的事。

## 是什么

Excel / CSV 批量合并与字段标准化工具 V1。
命令行工具：扫一个目录，把所有表格按列名对齐合并成一张 CSV。原文件永不改动。

## 版本

V1，2026-09-28 完成。79 项自检用例全过。

## 位置与环境（实测）

- 目录：`/home/mei/ai-automation/table-merger`
- Ubuntu 24.04 / Python 3.12.3（`/usr/bin/python3`）
- **本机没有 pip**（`No module named pip`），没有 pandas、没有 uv/pipx
- openpyxl 3.1.2 已装（apt 的 `python3-openpyxl`）
- **不引入任何第三方依赖**：只输出 CSV，用标准库 `csv/argparse/pathlib`；只有读 `.xlsx` 时才用 openpyxl
- 同一台机器上还有姊妹项目 `/home/mei/ai-automation/excel-cleaner`（Excel 清洗，已验收）。
  两个项目刻意保持同一套代码风格与安全约定：单文件 CLI + `CleanError`/`MergeError` 收口人话错误 +
  `PROJECT_CONTEXT.md` + `requirements.txt` + `.gitignore` + `examples/`。
- table-merger 目前**没有** `git init`（.gitignore 已备好，要建仓库直接 `git init` 即可）。

## 两条硬规则（小环明确要求，不许放宽、不许"优化"掉）

1. **字段别名映射必须基于显式配置，程序不许自行猜测字段含义。**
   → 只有别名表里写下的对应关系才生效。默认的 `aliases.txt` 是**全注释的空模板**，
   所以默认状态下零归并，不会出现"工具替我假设了字段语义"。
   → 任何形式的模糊匹配、相似度、同义词内置词典，都属于违反这条，不要加。
2. **同一张表里两列归到同一个列名 → 直接报错停机。**
   → 不许自动改名（早期版本曾改名成 `姓名(2)`，已按用户要求删除该行为）、
   不许静默合并、不许丢掉其中一列；必须打印出是哪两列撞了、怎么修，整任务中止，不产出任何文件。
   → **跨文件**归到同一列是允许的（不同行、不同格子，不覆盖不丢数据），这正是别名表的用途。
   若把跨文件也判为撞车，别名功能就自相矛盾了（一用就报错）—— 这是刻意的边界。
   → 同理，输入自带「来源文件」列时，与工具生成的来源列重名 → 也按撞车报错停机，提示用 `--no-source-col`。

3. **「永远不覆盖输入文件」是硬闸门（2026-09-28 由小环验收发现并修复的真实数据破坏 bug）。**
   → `-o` / `--report` 要写入的路径，都不允许指向输入目录里**现有的表格文件**（后缀属于表格的文件，
   包括 `.csv/.txt/.xlsx`，也包括名字以 `_merged` 结尾、扫描时会被跳过的那些）。
   → 按 `resolve()` 后的真实路径判断，符号链接绕不过去。
   → **与 `--force` 无关**：`--force` 只放宽「覆盖自己以前的输出」，永不放宽「覆盖原始数据」。
   → 必须在任何 `mkdir` / `write_csv` / `write_text` 之前判定（实际放在 `check_paths()` 里，
   早于 `main()` 里的 `--force` 检查，也早于 `run()` 里任何写盘动作）；拒绝时连父目录都不许创建。
   → **绝对不许恢复「把 output/report 加进 excluded_paths，让 scan_directory 排除掉」这种做法。**
   那正是这个 bug 的根因：要写的文件恰好是源表时，它先从输入集合里消失，
   后续所有安全检查都看不见它，最后被覆盖。安全判定必须看「目录里真实有什么」，
   不能看「扫描后剩下什么」。对应实现在 `list_existing_table_files()` + `check_write_targets()`。

## 历史 bug 记录（别改回去）

| bug | 根因 | 修法 |
|---|---|---|
| `-o data/a.csv --force` 覆盖源表；`--report data/a.csv` 覆盖源表 | `run()` 在 `scan_directory()` 之前把 output/report 塞进 `excluded_paths`，写入目标因此从输入集合里消失，安全检查看不见它 | 删除排除机制；新增 `list_existing_table_files()` + `check_write_targets()` 硬闸门，在 `check_paths()` 中调用 |
| ResourceWarning: unclosed file … .xlsx | `Workbook.close()` 在非 read_only 模式下是空操作（`hasattr(self,'_archive')` 为假），且 openpyxl 读工作表时留下未关闭的成员文件对象，使 `zipfile._fileRefCnt` 不归零，`ZipFile.close()` 关不掉底层 fd | `read_excel_workbook()` 自己 `with path.open("rb")` 持有句柄再交给 `load_workbook`；selftest 里读 xlsx 的用例同样处理 |

## 已定的架构决策（含理由，别推翻重问）

1. **输入 = 一个目录**（`python3 merge.py 目录`）。只扫这一层，不递归子目录。文件按文件名排序处理 → 结果可复现。
2. **按列名对齐，取并集，缺失留空**。输出列顺序 = 各列名**第一次出现**的顺序。显示名也取第一次出现的写法。
3. **标准化分两层**：
   - 列名：去首尾空白 + 全角字母数字转半角 + 比对时 casefold + 别名表归并（别名表默认空）
   - 值：去首尾空白 + 全角字母数字转半角
4. **全角转半角只处理字母、数字、全角空格（U+3000），不处理全角标点。**
   理由：`北京，上海` 的逗号是中文正文的一部分，转了就是改数据而不是标准化。写在 `to_halfwidth()` 里，有专门的用例守着。
5. **别名表 `aliases.txt` 默认自动加载**（位置：脚本旁边），`--aliases` 换文件、`--no-aliases` 关掉。
   格式 `标准名 = 别名1, 别名2`；同一别名指向两个标准名直接报错（不猜）。
   演示用的别名单独放 `examples/aliases_demo.txt`，必须显式 `--aliases` 才生效 —— 免得演示映射污染用户真实数据。
6. **默认输出到输入目录的旁边**（`目录名_merged.csv`），不是目录里面。
   理由：否则第二次运行会把上一次的输出当输入再合一遍。
   扫描时跳过名字以 `*_merged.csv` 结尾的文件（并打印警告），但这**不是**安全保障 ——
   安全保障是硬规则 3 的写入目标闸门（那个文件仍然受保护，不许被当成写入目标覆盖）。
   `--report` 的文件**不再**从扫描里排除：排除法掩盖冲突，已随 bug 修复一并删除。
7. **输出只有 CSV（UTF-8 带 BOM）**。`-o` 给非 `.csv` 后缀直接报错。BOM 是为了 Windows 版 Excel 双击不乱码。
8. **来源列**：默认在结果最左边加「来源文件」列（xlsx 多表时是 `文件名#表名`），`--no-source-col` 可关。
9. **空行**：整行标准化后全空的行跳过（Excel 表尾常见）。
10. **对齐的数据结构**：`ColumnSpec(display, key)`。
    - `key` = 内部对齐用的唯一键（`normalize_key`）
    - `display` = 输出表头显示的名字
    - 同一张表内 key 重复 → `MergeError`（见硬规则 2），不自动改名
    - 表头空/缺的列、以及比表头宽的行的额外列 → 叫「未命名列N」（N 是列号，不会自相撞车）
11. **批量语义**：不支持的格式（`.xlsm`/`.xls`/`.xlsb`）跳过 + 警告，继续跑；纯非表格文件忽略并计数说明。
    但「打不开」「含公式」「列名撞车」「来源列重名」是硬失败（整个任务中止，不产出任何文件）。
    理由：批量任务里少了一个文件却显示「成功」，比直接报错更危险。
12. **所有拒绝都发生在写盘之前**。被拒绝时磁盘上不多任何文件（连父目录都不会被创建）。有专门用例断言这一点。
    判定顺序：`check_paths()`（含写入目标闸门）→ `--force` 检查 → 读文件 / Excel 安全检查 / 列名撞车 → 最后才 mkdir + 写。

## 文件清单

| 文件 | 作用 |
|---|---|
| `merge.py` | 主程序（唯一逻辑文件） |
| `aliases.txt` | 别名表，默认全注释空模板（用户改这里来归并字段） |
| `selftest.py` | 自检，79 用例，纯标准库 unittest |
| `examples/input/` | 演示数据：2 个 CSV + 1 个 XLSX（列名和列顺序刻意都不一样） |
| `examples/aliases_demo.txt` | 演示用别名表（`姓名 = 名字, 员工姓名`），需显式指定 |
| `examples/sample_merged.csv` / `sample_report.txt` | 样例数据的预期输出 |
| `README.md` | 给人看的说明 |
| `requirements.txt` | 只有 openpyxl，且注明了用 apt 装 |
| `.gitignore` | 排除 `*_merged.csv`，放行 `examples/sample_merged.csv` |

## 验证方式

```bash
cd /home/mei/ai-automation/table-merger
python3 -W error::ResourceWarning selftest.py   # 79 用例，看到 OK 即全过，且不应有任何 ResourceWarning
python3 merge.py examples/input --aliases examples/aliases_demo.txt   # 端到端跑一遍样例
```

写入目标闸门的回归用例在 `TestWriteTargetGate`（10 个），覆盖：`-o` 指向输入 csv（含 `--force`）、
`--report` 指向输入 csv/txt（含 `--force`）、默认输出仍可用、`--force` 覆盖真旧输出仍有效、
符号链接绕道（`-o` 与 `--report` 各一）、`_merged` 结尾文件同样受保护、拒绝发生在 mkdir 之前。

## 已知限制（V1 有意为之）

- `.xlsm` / `.xls` / `.xlsb` 不支持；`.xlsx` 必须无公式（公式或 `=` 开头文本 → 整任务中止）
- Excel 只搬值：颜色、字体、列宽、合并单元格、公式全不保留
- 不做去重、不做类型识别（日期/数字只当文本清洗）
- 只扫一层目录；全量读进内存，无流式
- 输出只有 CSV，无格式转换

## 下一阶段候选（未决定）

- 跨文件去重（按主键，`--key 工号` 之类）
- 日期/数字类型标准化（需要先解决「哪一列是日期」的判断问题 —— 注意别踩硬规则 1）
- 大文件流式合并
- 递归子目录 / 输出 xlsx 多表
- 输出前预览（`--dry-run`）
