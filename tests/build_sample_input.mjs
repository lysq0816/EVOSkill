import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";


const action = process.argv[2];
if (!action || !["build", "verify"].includes(action)) {
  throw new Error(
    "Usage: node build_sample_input.mjs build <output.xlsx> | verify <input.xlsx> <preview-dir>",
  );
}

if (action === "verify") {
  const inputPath = process.argv[3];
  const previewDir = process.argv[4];
  if (!inputPath || !previewDir) {
    throw new Error("verify requires <input.xlsx> <preview-dir>");
  }
  await fs.mkdir(previewDir, { recursive: true });
  const input = await FileBlob.load(inputPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheetInspection = await workbook.inspect({
    kind: "sheet",
    include: "id,name",
    maxChars: 5000,
  });
  console.log(sheetInspection.ndjson);
  const badcaseInspection = await workbook.inspect({
    kind: "table",
    range: "BadCase分析!A1:AF8",
    include: "values,formulas",
    tableMaxRows: 8,
    tableMaxCols: 32,
    maxChars: 12000,
  });
  console.log(badcaseInspection.ndjson);
  const errorScan = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 300 },
    summary: "final formula error scan",
    maxChars: 5000,
  });
  console.log(errorScan.ndjson);

  const expectedSheets = [
    "数据",
    "指标",
    "单标签统计",
    "BadCase分析",
    "Knowledge Gap",
    "候选资料",
    "资料迭代概览",
  ];
  for (const sheetName of expectedSheets) {
    const preview = await workbook.render({
      sheetName,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    const safeName = sheetName.replace(/[\\/:*?"<>| ]/g, "_");
    await fs.writeFile(
      path.join(previewDir, `${safeName}.png`),
      new Uint8Array(await preview.arrayBuffer()),
    );
  }
  process.exit(0);
}

const outputPath = process.argv[3];
if (!outputPath) {
  throw new Error("build requires <output.xlsx>");
}

await fs.mkdir(path.dirname(outputPath), { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("数据");
sheet.getRange("A1:J6").values = [
  [
    "keyid",
    "业务线",
    "给模型的输入对话",
    "原始标签",
    "是否预警",
    "判断的标签",
    "单条数据的从读取到判别需要时间",
    "判断的理由",
    "模型原始输出",
    "请求错误",
  ],
  [
    "k1",
    "酒店",
    "用户：我要把这件事曝光出去",
    "B1",
    "否",
    "",
    0.1,
    "未命中",
    "{\"是否需要预警\":\"否\",\"判断的标签\":\"\",\"判断的理由\":\"未命中\"}",
    "",
  ],
  [
    "k2",
    "酒店",
    "用户：请帮我查询订单状态",
    "",
    "是",
    "B1",
    0.1,
    "误判为舆情",
    "{\"是否需要预警\":\"是\",\"判断的标签\":\"B1\",\"判断的理由\":\"误判为舆情\"}",
    "",
  ],
  [
    "k3",
    "机票",
    "用户：客服服务很差，我准备公开投诉",
    "A4、B1",
    "是",
    "A4、B0",
    0.1,
    "标签混淆",
    "{\"是否需要预警\":\"是\",\"判断的标签\":\"A4、B0\",\"判断的理由\":\"标签混淆\"}",
    "",
  ],
  [
    "k4",
    "火车",
    "用户：这次体验让我有点失望",
    "【3-4】",
    "否",
    "【3-4】",
    0.1,
    "轻度负面",
    "{\"是否需要预警\":\"否\",\"判断的标签\":\"【3-4】\",\"判断的理由\":\"轻度负面\"}",
    "",
  ],
  [
    "k5",
    "酒店",
    "用户：问题仍然没有解决",
    "B0",
    "",
    "",
    0.1,
    "请求失败",
    "",
    "timeout",
  ],
];

sheet.getRange("A1:J1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
};
sheet.getRange("A2:J6").format = { verticalAlignment: "top", wrapText: true };
sheet.getRange("C1:C6").format.columnWidth = 46;
sheet.getRange("H1:J6").format.columnWidth = 32;
sheet.freezePanes.freezeRows(1);
sheet.showGridLines = false;

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
