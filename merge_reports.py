#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OMNE 自动化测试报告汇总脚本
=============================
功能：
  - 遍历 ID / JP / PH 三个国家文件夹下的所有 ExecutionReport_*.html
  - 提取每份报告中的 jobs JSON 数据
  - 按 TestCaseId 去重，保留最新一次执行结果
  - 可选用 Sprint1.xlsx 覆盖最终状态（如 Sprint1.xlsx 中记录了最终判定）
  - 为每个国家生成：
      {country}/output/ExecutionReport_success.html
      {country}/output/ExecutionReport_failure.html

依赖安装：
  pip install beautifulsoup4 pandas openpyxl

运行：
  cd /path/to/OMNE_Report
  python merge_reports.py
"""

import os
import re
import json
import glob
from datetime import datetime

# ── 可选依赖 ────────────────────────────────────────────────────────────────
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    print("[警告] 未安装 pandas，将跳过 Sprint1.xlsx 校对逻辑。")
    print("       如需启用，请执行: pip install pandas openpyxl\n")

# ── 配置 ─────────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
COUNTRIES   = ["ID", "JP", "PH"]
SPRINT_FILE = os.path.join(BASE_DIR, "Sprint1.xlsx")

# Sprint1.xlsx 状态列名（根据实际列名调整）
XLSX_ID_COL     = "TestCaseId"   # 用例 ID 列
XLSX_STATUS_COL = "Status"       # 最终状态列（PASS / FAIL / SKIP）


# ═══════════════════════════════════════════════════════════════════════════════
# 1. 读取 Sprint1.xlsx，返回 {country: {testcase_id: status}} 映射
# ═══════════════════════════════════════════════════════════════════════════════
def load_sprint_xlsx():
    """
    读取 Sprint1.xlsx，每个 Sheet 对应一个国家。
    返回：{ "ID": {"TC_001": "PASS", ...}, "JP": {...}, "PH": {...} }
    """
    result = {}
    if not HAS_PANDAS:
        return result
    if not os.path.exists(SPRINT_FILE):
        print(f"[警告] Sprint1.xlsx 未找到：{SPRINT_FILE}，跳过校对。")
        return result

    try:
        xls = pd.ExcelFile(SPRINT_FILE)
        for sheet in xls.sheet_names:
            country = sheet.strip().upper()
            if country not in COUNTRIES:
                continue
            df = pd.read_excel(SPRINT_FILE, sheet_name=sheet)
            # 列名容错：忽略大小写、首尾空格
            df.columns = [c.strip() for c in df.columns]
            col_map = {c.lower(): c for c in df.columns}

            id_col  = col_map.get(XLSX_ID_COL.lower())
            sta_col = col_map.get(XLSX_STATUS_COL.lower())
            if id_col is None or sta_col is None:
                print(f"[警告] Sheet '{sheet}' 缺少列 '{XLSX_ID_COL}' 或 '{XLSX_STATUS_COL}'，跳过。")
                continue

            result[country] = {}
            for _, row in df.iterrows():
                tc_id  = str(row[id_col]).strip()
                status = str(row[sta_col]).strip().upper()
                if tc_id and status:
                    result[country][tc_id] = status
        print(f"[信息] Sprint1.xlsx 加载完成：{list(result.keys())}")
    except Exception as e:
        print(f"[错误] 读取 Sprint1.xlsx 失败：{e}")
    return result


# ═══════════════════════════════════════════════════════════════════════════════
# 2. 解析单个 HTML 报告，提取 jobs 列表和原始 HTML 模板
# ═══════════════════════════════════════════════════════════════════════════════
def _sanitize_jobs(jobs: list) -> list:
    """
    递归清理 jobs 中所有字符串值：
    将字面换行、制表符等控制字符替换为空格，
    防止它们破坏生成的 HTML 内嵌 JS。
    """
    def clean(obj):
        if isinstance(obj, str):
            # 将控制字符（U+0000–U+001F，除 \t \n \r 已由 json.dumps 转义外）
            # 替换为空格，保证嵌入 JS 字符串时不产生语法错误
            return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', ' ', obj)
        if isinstance(obj, dict):
            return {k: clean(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [clean(i) for i in obj]
        return obj
    return [clean(j) for j in jobs]


def parse_report_html(filepath):
    """
    从 ExecutionReport_*.html 中提取：
      - jobs     : list[dict]  — 测试用例数据列表
      - template : str         — 完整原始 HTML（用于后续重建报告）

    返回：(jobs_list, html_template_str) 或 ([], None) 表示解析失败
    """
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            html = f.read()
    except Exception as e:
        print(f"  [错误] 无法读取文件 {filepath}: {e}")
        return [], None

    jobs_match = re.search(r'const\s+jobs\s*=\s*(\[.*?\]);', html, re.DOTALL)
    if not jobs_match:
        print(f"  [警告] 未找到 jobs 数据：{os.path.basename(filepath)}")
        return [], html

    raw = jobs_match.group(1)

    # 原始报告有时在 JSON 字符串值里嵌入了字面换行/制表符（不合法的 JSON）。
    # 用状态机替换：只在双引号字符串内部（非转义）将控制字符转为 \n / \t / 空格。
    def _fix_raw_json(s: str) -> str:
        result, in_str, i = [], False, 0
        while i < len(s):
            c = s[i]
            if c == '\\' and in_str:          # 已转义字符，原样保留两个字符
                result.append(c)
                i += 1
                if i < len(s):
                    result.append(s[i])
            elif c == '"':
                in_str = not in_str
                result.append(c)
            elif in_str and c == '\n':
                result.append('\\n')
            elif in_str and c == '\r':
                result.append('\\r')
            elif in_str and c == '\t':
                result.append('\\t')
            elif in_str and ord(c) < 0x20:    # 其他控制字符替换为空格
                result.append(' ')
            else:
                result.append(c)
            i += 1
        return ''.join(result)

    try:
        jobs = json.loads(raw)
    except json.JSONDecodeError:
        try:
            jobs = json.loads(_fix_raw_json(raw))
        except json.JSONDecodeError as e:
            print(f"  [错误] jobs JSON 解析失败 {os.path.basename(filepath)}: {e}")
            return [], html

    return _sanitize_jobs(jobs), html


# ═══════════════════════════════════════════════════════════════════════════════
# 3. 从文件名解析时间戳（用于排序，确定"最新"执行）
# ═══════════════════════════════════════════════════════════════════════════════
def extract_timestamp_from_filename(filepath):
    """
    文件名格式：ExecutionReport_2026-06-01_16-28-27.html
    返回 datetime 对象，解析失败则返回 datetime.min
    """
    name = os.path.basename(filepath)
    # 匹配 yyyy-mm-dd_hh-mm-ss
    m = re.search(r'(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})', name)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y-%m-%d_%H-%M-%S")
        except ValueError:
            pass
    return datetime.min


# ═══════════════════════════════════════════════════════════════════════════════
# 4. 汇总一个国家的所有报告，返回去重后的 jobs 列表及模板 HTML
# ═══════════════════════════════════════════════════════════════════════════════
def collect_country_jobs(country_dir, sprint_status_map):
    """
    遍历 {country_dir}/output/ExecutionReport_*.html：
      - 按文件名时间戳从旧到新排序（新的覆盖旧的，实现"保留最新"）
      - 按 TestCaseId 去重
      - 若 sprint_status_map 不为空，用 Sprint1.xlsx 中的状态覆盖 ExecutionStatus

    返回：(deduped_jobs: list[dict], template_html: str | None)
    """
    output_dir = os.path.join(country_dir, "output")
    if not os.path.isdir(output_dir):
        print(f"  [警告] 目录不存在：{output_dir}")
        return [], None

    pattern = os.path.join(output_dir, "ExecutionReport_*.html")
    # 排除由本脚本生成的汇总文件，只处理带时间戳的原始报告
    EXCLUDE = {"ExecutionReport_success.html", "ExecutionReport_failure.html"}
    html_files = sorted(
        [f for f in glob.glob(pattern) if os.path.basename(f) not in EXCLUDE],
        key=extract_timestamp_from_filename
    )

    if not html_files:
        print(f"  [警告] 未找到任何 ExecutionReport_*.html：{output_dir}")
        return [], None

    print(f"  找到 {len(html_files)} 份报告，时间范围："
          f"{os.path.basename(html_files[0])} → {os.path.basename(html_files[-1])}")

    # 用 dict 去重：key=TestCaseId，value=job dict
    # 从旧到新遍历，新的自动覆盖旧的 → 最终保留最新执行结果
    deduped = {}
    latest_template = None  # 保存最新一份报告的完整 HTML 作为模板

    for fpath in html_files:
        jobs, html = parse_report_html(fpath)
        if html:
            latest_template = html  # 每次更新，最终保留最新文件的模板
        for job in jobs:
            tc_id = job.get("TestCaseId", "").strip()
            if tc_id:
                deduped[tc_id] = job

    jobs_list = list(deduped.values())

    # ── 用 Sprint1.xlsx 覆盖状态（可选）────────────────────────────────────
    # 逻辑：若 Sprint1.xlsx 中该 TestCaseId 有最终判定，以 xlsx 为准
    if sprint_status_map:
        overridden = 0
        for job in jobs_list:
            tc_id = job.get("TestCaseId", "").strip()
            if tc_id in sprint_status_map:
                xlsx_status = sprint_status_map[tc_id]
                # 仅当 xlsx 中状态与报告不一致时才覆盖
                if job.get("ExecutionStatus", "") != xlsx_status:
                    job["ExecutionStatus"] = xlsx_status
                    overridden += 1
        if overridden:
            print(f"  [Sprint1.xlsx] 覆盖了 {overridden} 条用例的最终状态")

    print(f"  去重后共 {len(jobs_list)} 条用例")
    return jobs_list, latest_template


# ═══════════════════════════════════════════════════════════════════════════════
# 5. 用原始 HTML 模板 + 新 jobs 列表生成汇总报告
# ═══════════════════════════════════════════════════════════════════════════════
def build_report_html(template_html, filtered_jobs, report_title, generated_time):
    """
    将 template_html 中的：
      - const jobs = [...];    替换为过滤后的新数据
      - const summary = {...}; 替换为新的统计数据
      - 页头标题 <h1>          替换为 report_title
      - Generated Time         替换为当前时间

    保持所有 CSS / JS 逻辑不变，仅更换数据部分。
    """
    if not template_html:
        return None

    # 计算新 summary
    total   = len(filtered_jobs)
    passed  = sum(1 for j in filtered_jobs if j.get("ExecutionStatus") == "PASS")
    failed  = sum(1 for j in filtered_jobs if j.get("ExecutionStatus") == "FAIL")
    skipped = sum(1 for j in filtered_jobs if j.get("ExecutionStatus") == "SKIPPED")
    summary = {"total": total, "passed": passed, "failed": failed, "skipped": skipped}

    # ensure_ascii=False 保留中文；同时让 json.dumps 自动把 \n \r \t 转义为 \\n 等，
    # 避免字面换行符嵌入 JS 字符串导致语法错误
    new_jobs_json    = json.dumps(filtered_jobs,    ensure_ascii=False, separators=(',', ':'), indent=None)
    new_summary_json = json.dumps(summary, ensure_ascii=False, separators=(',', ':'), indent=None)

    html = template_html

    # ── 替换 jobs 数组 ─────────────────────────────────────────────────────
    # 注意：必须用 lambda 而非字符串，否则 re.sub 会把 \n 等反斜杠序列
    # 再次展开为真实控制字符，破坏 json.dumps 已正确转义的 JSON。
    jobs_replacement = f'const jobs = {new_jobs_json};'
    html = re.sub(
        r'const\s+jobs\s*=\s*\[.*?\];',
        lambda _: jobs_replacement,
        html,
        count=1,
        flags=re.DOTALL
    )

    # ── 替换 summary 对象 ──────────────────────────────────────────────────
    summary_replacement = f'const summary = {new_summary_json};'
    html = re.sub(
        r'const\s+summary\s*=\s*\{.*?\};',
        lambda _: summary_replacement,
        html,
        count=1,
        flags=re.DOTALL
    )

    # ── 替换 Generated Time ────────────────────────────────────────────────
    html = re.sub(
        r'Generated Time:.*?</div>',
        f'Generated Time: {generated_time}</div>',
        html,
        count=1
    )

    # ── 替换页头标题（<h1> 中的文字）─────────────────────────────────────
    html = re.sub(
        r'(<h1>).*?(</h1>)',
        rf'\g<1>{report_title}\g<2>',
        html,
        count=1
    )

    return html


# ═══════════════════════════════════════════════════════════════════════════════
# 6. 主流程
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("OMNE 测试报告汇总脚本")
    print("=" * 60)

    # 6-1. 加载 Sprint1.xlsx（可选校对数据）
    sprint_data = load_sprint_xlsx()

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 6-2. 逐国家处理
    for country in COUNTRIES:
        print(f"\n[{country}] 开始处理...")
        country_dir = os.path.join(BASE_DIR, country)

        if not os.path.isdir(country_dir):
            print(f"  [跳过] 目录不存在：{country_dir}")
            continue

        # 该国在 Sprint1.xlsx 中的状态映射（没有则为空 dict）
        sprint_map = sprint_data.get(country, {})

        # 收集并去重所有 jobs
        all_jobs, template_html = collect_country_jobs(country_dir, sprint_map)

        if not all_jobs:
            print(f"  [跳过] {country} 无有效测试用例数据")
            continue

        if not template_html:
            print(f"  [跳过] {country} 无法获取 HTML 模板")
            continue

        output_dir = os.path.join(country_dir, "output")

        # 6-3. 按状态拆分
        success_jobs = [j for j in all_jobs if j.get("ExecutionStatus") == "PASS"]
        failure_jobs = [j for j in all_jobs if j.get("ExecutionStatus") == "FAIL"]
        skipped_jobs = [j for j in all_jobs if j.get("ExecutionStatus") == "SKIPPED"]

        print(f"  统计：PASS={len(success_jobs)}  FAIL={len(failure_jobs)}  SKIPPED={len(skipped_jobs)}")

        # 6-4. 生成 success 报告
        success_html = build_report_html(
            template_html,
            success_jobs,
            report_title=f"OMNE Report [{country}] — SUCCESS",
            generated_time=now_str
        )
        success_path = os.path.join(output_dir, "ExecutionReport_success.html")
        try:
            with open(success_path, "w", encoding="utf-8") as f:
                f.write(success_html)
            print(f"  ✅ 已生成：{success_path}")
        except Exception as e:
            print(f"  [错误] 写入 success 报告失败：{e}")

        # 6-5. 生成 failure 报告
        failure_html = build_report_html(
            template_html,
            failure_jobs,
            report_title=f"OMNE Report [{country}] — FAILURE",
            generated_time=now_str
        )
        failure_path = os.path.join(output_dir, "ExecutionReport_failure.html")
        try:
            with open(failure_path, "w", encoding="utf-8") as f:
                f.write(failure_html)
            print(f"  ❌ 已生成：{failure_path}")
        except Exception as e:
            print(f"  [错误] 写入 failure 报告失败：{e}")

    print("\n" + "=" * 60)
    print("全部完成！")
    print("=" * 60)


if __name__ == "__main__":
    main()
