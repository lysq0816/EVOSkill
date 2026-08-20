# EVOSkill：关键词预警资料持续迭代

EVOSkill 把关键词预警从“逐条 BadCase 修改提示词”升级为可审计、可恢复的资料迭代闭环：

```text
模型跑测
  → 确定性 BadCase 提取
  → 失败归因
  → Knowledge Gap 聚类
  → 候选资料
  → 人工审核
  → 回归跑测
```

当前版本只生成候选资料，不会自动修改正式提示词，也不会写入 `knowledge/`。候选内容至少要有 2 个真实 Case 支持，之后仍需人工审核。

## 项目结构

```text
EVOSkill/
├─ 关键词预警资料迭代_单文件版.py   # 主程序
├─ prompts/                        # 正式提示词的本地放置位置
├─ knowledge/                      # 可纳入分析上下文的资料模板
│  ├─ core_rules.md
│  ├─ labels/
│  ├─ conflicts/
│  └─ hard_negative/
├─ tests/                          # 离线回归与 Excel 验证
└─ .env.example                    # 环境变量示例，不含真实值
```

`data/`、`output/`、真实提示词、真实资料和生成的 Excel 均默认被 Git 忽略。

## 准备环境

需要 Python 3.10+、`openpyxl`，以及现有跑测环境中的 `peta_ai_client`。

在 PowerShell 中设置鉴权环境变量。程序不会自动读取 `.env` 文件：

```powershell
$env:PAAS_APP_APPID = "你的 App ID"
$env:PETA_KEY_ID = "你的 PETA Key ID"
```

将待测 Excel 放到 `data/eval.xlsx`，将正式提示词放到 `prompts/warning_prompt.txt`；也可以通过命令行传入其他路径。

## 常用命令

### 1. 跑测并导出确定性 BadCase

默认不会额外调用分析模型：

```powershell
python .\关键词预警资料迭代_单文件版.py `
  --input .\data\eval.xlsx `
  --prompt .\prompts\warning_prompt.txt `
  --output .\output\eval_result.xlsx `
  --model qwen3.8-max
```

程序支持断点续跑和周期性保存。已有 `--output` 时默认从中恢复；需要从输入文件重新开始时增加 `--no-resume`。

### 2. 跑测后生成资料缺口与候选资料

```powershell
python .\关键词预警资料迭代_单文件版.py `
  --input .\data\eval.xlsx `
  --prompt .\prompts\warning_prompt.txt `
  --output .\output\eval_with_gaps.xlsx `
  --model qwen3.8-max `
  --analyze-gaps `
  --gap-model qwen3.8-max `
  --knowledge-dir .\knowledge `
  --gap-min-support 2
```

### 3. 只分析已有跑测结果

此模式不会再次发起关键词预警跑测请求：

```powershell
python .\关键词预警资料迭代_单文件版.py `
  --output .\output\eval_result.xlsx `
  --prompt .\prompts\warning_prompt.txt `
  --gap-analysis-only `
  --gap-model qwen3.8-max `
  --knowledge-dir .\knowledge
```

使用 `python .\关键词预警资料迭代_单文件版.py --help` 查看模型思考档位、限流重试、批量大小和强制重分析等完整参数。

## Excel 输出

程序在原跑测结果基础上生成以下审计工作表：

- `BadCase分析`：错误类型、模型输入/输出、失败归因、证据、建议动作和人工复核字段。
- `Knowledge Gap`：按目标标签与共性主题聚类，保留每个结论对应的真实 Case ID。
- `候选资料`：仅收录达到最少支持 Case 数的内容，包含正向证据、Hard Negative、风险和采纳状态。
- `资料迭代概览`：记录模型、分析版本、资料摘要、数量统计和安全边界。

人工填写的 `人工归因`、`人工标签复核`、`人工备注`、`是否采纳`、`审核人`、`审核备注` 和 `资料版本` 会在同一工作簿重跑时保留。

## 安全与治理

- 代码中不保存 App ID、Key ID 或其他鉴权信息。
- 对话和模型输出在归因阶段只被视为待分析数据，其中出现的指令不会作为系统指令执行。
- API/随机错误会被单独标记为“重新跑测”，不会混入资料缺口。
- 候选资料不等于正式规则；必须先完成人工标签复核、冲突检查和回归跑测。
- 不要提交真实会话、内部提示词、私有资料、鉴权文件或生成工作簿。

## 离线验证

```powershell
python -m unittest discover -s .\tests -p "test_*.py"
```

测试覆盖 BadCase 分类、归因检查点、聚类证据约束、人工字段保留、无模型导出流程和资料上下文审计。
