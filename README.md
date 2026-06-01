# OMNE Test Execution Reports

Sprint 1 自动化测试执行报告汇总，覆盖三个国家：印尼（ID）、日本（JP）、菲律宾（PH）。

## 目录结构

```
├── ID/output/          # 印尼执行报告
├── JP/output/          # 日本执行报告
├── PH/output/          # 菲律宾执行报告
│   ├── ExecutionReport_yyyy-mm-dd_hh-mm-ss.html   # 原始报告（每次执行生成一份）
│   ├── ExecutionReport_success.html                # 汇总：所有 PASS 用例
│   └── ExecutionReport_failure.html                # 汇总：所有 FAIL 用例
└── merge_reports.py    # 汇总脚本
```

## 汇总结果（Sprint 1）

| 国家 | PASS | FAIL |
|------|------|------|
| ID   | 7    | 22   |
| JP   | 31   | 37   |
| PH   | 0    | 12   |

## 更新报告

有新报告时，将 `ExecutionReport_yyyy-mm-dd_hh-mm-ss.html` 放入对应国家的 `output/` 目录，然后运行：

```bash
python3 merge_reports.py
```

脚本会自动扫描所有原始报告、按 TestCaseId 去重（保留最新执行结果），重新生成各国的 success / failure 汇总报告。

依赖安装：
```bash
pip install pandas openpyxl
```
