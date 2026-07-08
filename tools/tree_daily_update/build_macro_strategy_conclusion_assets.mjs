import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = path.resolve("D:/AI/宏观框架分析/20260610_TREE宏观日报_趋势研判与资产含义.xlsx");

const trendTitle = "1. 趋势研判｜市场：国内信用偏弱但利润修复，美国通胀再抬头，资产配置宜重视结构而非全面进攻";

const trendRows = [
  ["模块", "研判", "结论"],
  [
    "境内｜流动性面",
    "中性偏弱：货币端稳定，信用端偏弱",
    "当前的中国流动性面边际变化主要体现在实体信用走弱，社会融资规模同比多增回落45995，新增人民币贷款同比多增回落35528，政府债券融资同比多增回落2617；价格型资金成本整体稳定，7天逆回购利率维持1.40%，DR007仅小幅上行0.0123个百分点。综合看，央行层面的流动性锚仍然稳定，但信用扩张动能不足，当前中国流动性更像是“货币端不紧、信用端偏弱”的组合。",
  ],
  [
    "境内｜政策面",
    "中性：政策有支撑，但资金释放仍不强",
    "当前的中国政策面边际变化主要体现在财政存款增加14788，财政资金阶段性沉淀，对实体和流动性的直接释放力度有限；财政赤字脉冲小幅改善0.0012，但历史样本不足，信号强度不高。综合看，政策面并非收缩，但当前更偏“政策预期存在、实际传导偏慢”。",
  ],
  [
    "境内｜增长面",
    "中性偏弱：外需支撑，内需和地产仍拖累",
    "当前的中国增长面边际变化主要体现在内需走弱，固定资产投资回落3.3个百分点，社会消费品零售总额回落1.5个百分点，PMI新订单回落0.7个百分点，房地产开发投资回落2.5个百分点；出口金额同比上行5.3个百分点，是增长端少数改善项。综合看，当前中国增长不是全面下行，而是“外需有支撑、内需偏弱、地产仍拖累”的结构。",
  ],
  [
    "境内｜价格利润面",
    "中性偏多：利润端改善强于需求端",
    "当前的中国价格利润面边际变化主要体现在PPI上行2.3个百分点，工业企业利润累计同比上行2.7个百分点，工业企业营业收入利润率上行0.32个百分点；原材料购价回落3.2，显示成本压力边际缓和。综合看，价格利润面比增长面更积极，当前更像是“利润修复好于需求修复”，但持续性仍取决于后续内需能否跟上。",
  ],
  [
    "境外｜美国流动性面",
    "中性：利率锚稳定，信用端略有分化",
    "当前的美国流动性面边际变化主要体现在美联储资产负债表增加7112，信用利差上行0.0006，银行贷款增速上行2.8个百分点；但联邦基金利率、SOFR、SOFR-IORB利差和IORB均基本持平。综合看，美国流动性并不是价格端明显放松，而是“利率端稳定、数量端小幅波动、信用端略有分化”。",
  ],
  [
    "境外｜美国经济基本面",
    "中性偏强但通胀约束上升",
    "当前的美国经济基本面边际变化主要体现在通胀重新抬头，CPI同比上行0.5个百分点，核心CPI上行0.1个百分点，PPI同比上行1.7个百分点；增长端仍有韧性，GDP同比和制造业ISM改善，但零售销售同比回落1.138个百分点，个人可支配收入同比回落0.3个百分点。综合看，美国经济呈现“生产和总量仍有韧性、消费边际降温、通胀粘性增强”的组合，对降息预期并不友好。",
  ],
];

const assetTitle = "2. 资产含义｜基于宏观面";

const assetRows = [
  ["资产", "配置倾向", "理由"],
  [
    "A股",
    "标配，结构优先",
    "国内流动性不紧，但信用扩张偏弱，决定了市场很难单靠总量信用形成全面上行；价格利润面改善对企业盈利有支撑，适合关注盈利修复、政策支持和外需链条，但不宜过度依赖强复苏交易。",
  ],
  [
    "港股",
    "标配偏谨慎",
    "港股受中国基本面和海外流动性共同影响。中国增长端仍偏弱，美国通胀约束降息预期，估值修复空间存在，但弹性需要等待信用或内需数据进一步改善。",
  ],
  [
    "美股",
    "标配偏谨慎",
    "美国增长仍有韧性，ISM和GDP改善对盈利有支撑，但CPI、核心CPI和PPI同步上行，会压制降息预期和估值扩张。美股更适合结构性配置，不适合单纯基于流动性宽松逻辑大幅加仓。",
  ],
  [
    "美债",
    "中性偏谨慎",
    "美国通胀边际抬头，利率端又没有明显宽松，长端利率下行逻辑不够顺畅。美债配置价值仍在，但久期仓位不宜过度激进。",
  ],
  [
    "黄金",
    "标配",
    "美国通胀粘性和全球不确定性仍对黄金有支撑，但如果降息预期被通胀压制，黄金上行动能也会受到阶段性扰动。当前更适合作为组合中的稳定器，而不是单边追涨资产。",
  ],
  [
    "中国债券",
    "标配偏多",
    "国内信用扩张偏弱、内需修复不足，对债券仍有支撑；但财政存款变化和政策预期会带来阶段性扰动。整体看，债券仍具备配置价值，但需要关注政策发力对利率的短期冲击。",
  ],
];

const finalJudgement =
  "当前宏观环境不是典型的全面风险偏好上行，而是“国内货币不紧但信用偏弱、盈利修复好于需求修复；海外增长仍有韧性但通胀约束增强”。因此配置上更适合保持均衡，权益看结构，债券看防御，黄金作为组合稳定器。";

function applyTableStyle(sheet, range, headerRange, widths) {
  sheet.getRange(range).format = {
    font: { size: 11, name: "Microsoft YaHei" },
    wrapText: true,
    verticalAlignment: "top",
  };
  sheet.getRange(headerRange).format = {
    fill: "#1F4E78",
    font: { color: "#FFFFFF", bold: true, size: 11, name: "Microsoft YaHei" },
    horizontalAlignment: "center",
    verticalAlignment: "middle",
  };
  sheet.getRange(range).format.borders = {
    preset: "all",
    style: "thin",
    color: "#D9E2F3",
  };
  for (const [col, width] of Object.entries(widths)) {
    sheet.getRange(`${col}:${col}`).format.columnWidthPx = width;
  }
  sheet.freezePanes.freezeRows(2);
}

const workbook = Workbook.create();

const trendSheet = workbook.worksheets.add("趋势研判");
trendSheet.getRange("A1:C1").values = [[trendTitle, "", ""]];
trendSheet.getRange(`A2:C${trendRows.length + 1}`).values = trendRows;
trendSheet.getRange("A1:C1").format = {
  fill: "#D9EAF7",
  font: { bold: true, size: 14, name: "Microsoft YaHei", color: "#002060" },
  verticalAlignment: "middle",
};
applyTableStyle(trendSheet, `A2:C${trendRows.length + 1}`, "A2:C2", {
  A: 170,
  B: 250,
  C: 900,
});

const assetSheet = workbook.worksheets.add("资产含义");
assetSheet.getRange("A1:C1").values = [[assetTitle, "", ""]];
assetSheet.getRange(`A2:C${assetRows.length + 1}`).values = assetRows;
assetSheet.getRange(`A${assetRows.length + 3}:C${assetRows.length + 3}`).values = [["一句话总判断", "", finalJudgement]];
assetSheet.getRange("A1:C1").format = {
  fill: "#D9EAF7",
  font: { bold: true, size: 14, name: "Microsoft YaHei", color: "#002060" },
  verticalAlignment: "middle",
};
applyTableStyle(assetSheet, `A2:C${assetRows.length + 1}`, "A2:C2", {
  A: 150,
  B: 180,
  C: 900,
});
assetSheet.getRange(`A${assetRows.length + 3}:C${assetRows.length + 3}`).format = {
  fill: "#FFF2CC",
  font: { bold: true, size: 11, name: "Microsoft YaHei" },
  wrapText: true,
  verticalAlignment: "top",
};
assetSheet.getRange(`A${assetRows.length + 3}:C${assetRows.length + 3}`).format.borders = {
  preset: "all",
  style: "thin",
  color: "#D6B656",
};

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
