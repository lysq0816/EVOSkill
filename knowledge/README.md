# 关键词预警资料库

这里存放经过人工审核的结构化资料。脚本会递归读取 `.md` 和 `.txt` 文件，并与主提示词一起用于 BadCase 归因和 Knowledge Gap 聚类。

建议按下面的目录组织：

```text
knowledge/
├── core_rules.md
├── labels/
│   └── _template.md
├── conflicts/
│   └── _template.md
├── hard_negative/
│   └── _template.md
└── private/              # 本地敏感资料，Git 默认忽略
```

只有在“候选资料”工作表中完成人工审核后，才把内容写入正式资料文件。不要直接把模型生成结果自动合并进来。
