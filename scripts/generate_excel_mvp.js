const path = require("path");
const ExcelJS = require("exceljs");

const outPath = path.join(
  process.cwd(),
  "artifacts",
  "单身期家庭资产诊断MVP.xlsx"
);

const dimensions = [
  { name: "储蓄能力", weight: 0.22, color: "21A67A", role: "收入转化为资产的核心发动机" },
  { name: "流动性", weight: 0.20, color: "3D8FE3", role: "防止突发事件击穿财务安全" },
  { name: "偿债能力", weight: 0.20, color: "E05A34", role: "控制负债规模和还款压力" },
  { name: "保障能力", weight: 0.15, color: "7D70E6", role: "转移重大风险，避免一次事故归零" },
  { name: "成长能力", weight: 0.15, color: "5A9E28", role: "判断资产是否具备长期增值能力" },
  { name: "财富自由度", weight: 0.08, color: "BF7B11", role: "结果性指标，单身期低权重保留" },
];

const dimensionMap = Object.fromEntries(dimensions.map((d) => [d.name, d]));

const indicators = [
  ["偿债能力", "资产负债率", 0.30, "总负债 / 总资产", "越低越好", 0.40, -10, "s", "20%优，40%中点，60%警示"],
  ["偿债能力", "债务偿付率", 0.45, "年债务本息 / 年收入", "越低越好", 0.25, -12, "s", "10%优，25%中点，40%警示"],
  ["偿债能力", "高息负债收入比", 0.25, "高息消费贷余额 / 年收入", "越低越好", 0.10, -25, "s", "0%优，10%警示，20%红灯"],
  ["流动性", "紧急备用金倍数", 0.50, "流动资产 / 月必要支出", "越高越好", 3, 1.1, "s", "1个月红灯，3个月中点，5-6个月优"],
  ["流动性", "净流动覆盖倍数", 0.25, "(流动资产 - 短期负债) / 月必要支出", "越高越好", 2, 1.2, "s", "0红灯，2中点，4优"],
  ["流动性", "流动资产占比", 0.25, "流动资产 / 总资产", "越高越好", 0.15, 12, "s", "5%红灯，15%中点，30%优"],
  ["储蓄能力", "总结余率", 0.45, "(收入 - 支出) / 收入", "越高越好", 0.20, 12, "s", "5%红灯，20%中点，35%优"],
  ["储蓄能力", "自由储蓄率", 0.35, "(收入 - 固定支出 - 债务本息) / 收入", "越高越好", 0.15, 12, "s", "0%红灯，15%中点，30%优"],
  ["储蓄能力", "支出刚性比率", 0.20, "固定支出 / 总支出", "越低越好", 0.60, -8, "s", "40%优，60%中点，80%警示"],
  ["成长能力", "生息资产占比", 0.35, "生息资产 / 总资产", "越高越好", 0.35, 8, "s", "15%红灯，35%中点，55%优"],
  ["成长能力", "长期投资资产占比", 0.35, "长期投资资产 / 可投资资产", "越高越好", 0.50, 8, "s", "20%红灯，50%中点，70%优"],
  ["成长能力", "投资报酬率", 0.30, "年投资收益 / 期初投资资产", "越高越好", 0.04, 12, "s", "0%低，4%中点，8%优"],
  ["保障能力", "基础保障完整度", 0.35, "已配置基础保障项 / 应配置基础保障项", "越高越好", 0.60, 8, "s", "0-30%红灯，60%中点，100%优"],
  ["保障能力", "保费收入比", 0.30, "年保费 / 年收入", "区间最优", 0.05, 0, "piecewise", "低于3%可能不足，5%-12%合理，高于20%过重"],
  ["保障能力", "重大风险覆盖倍数", 0.35, "有效保障额度 / 年必要支出", "越高越好", 3, 0.8, "s", "1倍红灯，3倍中点，6倍优"],
  ["财富自由度", "被动收入覆盖率", 0.50, "年被动收入 / 年支出", "越高越好", 0.10, 8, "s", "0%低，10%中点，30%优"],
  ["财富自由度", "净资产支出倍数", 0.50, "净资产 / 年支出", "越高越好", 1, 1.1, "s", "0低，1倍中点，3倍优"],
];

const sample = {
  totalAssets: 1050000,
  totalLiabilities: 280000,
  liquidAssets: 70000,
  shortTermLiabilities: 20000,
  interestBearingAssets: 260000,
  investableAssets: 330000,
  longTermInvestmentAssets: 230000,
  passiveIncome: 12000,
  investmentReturn: 12000,
  beginningInvestmentAssets: 230000,
  annualIncome: 360000,
  annualExpense: 300000,
  monthlyNecessaryExpense: 25000,
  fixedExpense: 210000,
  debtService: 110000,
  highInterestDebt: 60000,
  annualPremium: 16000,
  configuredProtectionItems: 2,
  requiredProtectionItems: 3,
  majorRiskCoverage: 900000,
};
sample.netAssets = sample.totalAssets - sample.totalLiabilities;

function logistic(x, m, k) {
  return 100 / (1 + Math.exp(-k * (x - m)));
}

function premiumScore(x) {
  if (x < 0.05) return logistic(x, 0.05, 50);
  if (x <= 0.12) return 95;
  if (x <= 0.20) return 95 - ((x - 0.12) / 0.08) * 35;
  return Math.max(20, 60 - ((x - 0.20) / 0.10) * 40);
}

function rawValue(name) {
  switch (name) {
    case "资产负债率":
      return sample.totalLiabilities / sample.totalAssets;
    case "债务偿付率":
      return sample.debtService / sample.annualIncome;
    case "高息负债收入比":
      return sample.highInterestDebt / sample.annualIncome;
    case "紧急备用金倍数":
      return sample.liquidAssets / sample.monthlyNecessaryExpense;
    case "净流动覆盖倍数":
      return (sample.liquidAssets - sample.shortTermLiabilities) / sample.monthlyNecessaryExpense;
    case "流动资产占比":
      return sample.liquidAssets / sample.totalAssets;
    case "总结余率":
      return (sample.annualIncome - sample.annualExpense) / sample.annualIncome;
    case "自由储蓄率":
      return (sample.annualIncome - sample.fixedExpense - sample.debtService) / sample.annualIncome;
    case "支出刚性比率":
      return sample.fixedExpense / sample.annualExpense;
    case "生息资产占比":
      return sample.interestBearingAssets / sample.totalAssets;
    case "长期投资资产占比":
      return sample.longTermInvestmentAssets / sample.investableAssets;
    case "投资报酬率":
      return sample.investmentReturn / sample.beginningInvestmentAssets;
    case "基础保障完整度":
      return sample.configuredProtectionItems / sample.requiredProtectionItems;
    case "保费收入比":
      return sample.annualPremium / sample.annualIncome;
    case "重大风险覆盖倍数":
      return sample.majorRiskCoverage / (sample.monthlyNecessaryExpense * 12);
    case "被动收入覆盖率":
      return sample.passiveIncome / sample.annualExpense;
    case "净资产支出倍数":
      return sample.netAssets / sample.annualExpense;
    default:
      return 0;
  }
}

function rawFormula(name) {
  const s = "'示例输入'!";
  switch (name) {
    case "资产负债率":
      return `${s}B6/${s}B5`;
    case "债务偿付率":
      return `${s}B22/${s}B18`;
    case "高息负债收入比":
      return `${s}B23/${s}B18`;
    case "紧急备用金倍数":
      return `${s}B8/${s}B20`;
    case "净流动覆盖倍数":
      return `(${s}B8-${s}B9)/${s}B20`;
    case "流动资产占比":
      return `${s}B8/${s}B5`;
    case "总结余率":
      return `(${s}B18-${s}B19)/${s}B18`;
    case "自由储蓄率":
      return `(${s}B18-${s}B21-${s}B22)/${s}B18`;
    case "支出刚性比率":
      return `${s}B21/${s}B19`;
    case "生息资产占比":
      return `${s}B10/${s}B5`;
    case "长期投资资产占比":
      return `${s}B12/${s}B11`;
    case "投资报酬率":
      return `${s}B14/${s}B15`;
    case "基础保障完整度":
      return `${s}B25/${s}B26`;
    case "保费收入比":
      return `${s}B24/${s}B18`;
    case "重大风险覆盖倍数":
      return `${s}B27/(${s}B20*12)`;
    case "被动收入覆盖率":
      return `${s}B13/${s}B19`;
    case "净资产支出倍数":
      return `${s}B7/${s}B19`;
    default:
      return "0";
  }
}

function scoreFormula(rowIndex, indicator) {
  if (indicator[7] === "piecewise") {
    return `IF(C${rowIndex}<0.05,100/(1+EXP(-50*(C${rowIndex}-0.05))),IF(C${rowIndex}<=0.12,95,IF(C${rowIndex}<=0.2,95-(C${rowIndex}-0.12)/0.08*35,MAX(20,60-(C${rowIndex}-0.2)/0.1*40))))`;
  }
  return `100/(1+EXP(${(-indicator[6])}*(C${rowIndex}-${indicator[5]})))`;
}

function scoreIndicator(row) {
  const x = rawValue(row[1]);
  if (row[7] === "piecewise") return premiumScore(x);
  return logistic(x, row[5], row[6]);
}

function status(score) {
  if (score >= 70) return "绿灯";
  if (score >= 55) return "黄灯";
  return "红灯";
}

function pct(v) {
  return `${(v * 100).toFixed(1)}%`;
}

function multiple(v) {
  return `${v.toFixed(1)}倍`;
}

function money(v) {
  return `${(v / 10000).toFixed(1)}万`;
}

function setTitle(ws, title, subtitle) {
  ws.mergeCells("A1:H1");
  ws.getCell("A1").value = title;
  ws.getCell("A1").font = { name: "Microsoft YaHei", bold: true, size: 18, color: { argb: "1F2937" } };
  ws.getCell("A1").alignment = { vertical: "middle" };
  ws.getRow(1).height = 30;
  if (subtitle) {
    ws.mergeCells("A2:H2");
    ws.getCell("A2").value = subtitle;
    ws.getCell("A2").font = { name: "Microsoft YaHei", size: 10, color: { argb: "6B7280" } };
  }
}

function styleHeader(row) {
  row.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "F3F0E8" } };
    cell.font = { name: "Microsoft YaHei", bold: true, color: { argb: "374151" } };
    cell.border = thinBorder();
    cell.alignment = { vertical: "middle" };
  });
}

function thinBorder() {
  return {
    top: { style: "thin", color: { argb: "D9D6CC" } },
    left: { style: "thin", color: { argb: "D9D6CC" } },
    bottom: { style: "thin", color: { argb: "D9D6CC" } },
    right: { style: "thin", color: { argb: "D9D6CC" } },
  };
}

function styleSheet(ws) {
  ws.views = [{ showGridLines: false }];
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      cell.font = cell.font || { name: "Microsoft YaHei", size: 10 };
      cell.alignment = cell.alignment || { vertical: "middle", wrapText: true };
    });
  });
}

async function main() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Codex";
  workbook.created = new Date();

  const readme = workbook.addWorksheet("README");
  readme.columns = [
    { width: 18 }, { width: 80 }, { width: 18 }, { width: 18 },
    { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 },
  ];
  setTitle(readme, "单身期家庭资产诊断 MVP", "使用两张家庭财务报表进行六维评分、风险识别、连锁诊断和改善建议。");
  readme.addRows([
    [],
    ["使用顺序", "1. 在“示例输入”中替换家庭数据；2. 查看“评分计算”和“六维总览”；3. 查看“连锁诊断”和“建议库”；4. 用样本家庭回测阈值。"],
    ["评分原则", "底层计算保持严格，展示层可以友好表达；红灯硬性风险不得被综合分稀释。"],
    ["MVP 边界", "本工具只输出方向性资产结构建议，不输出具体金融产品、保险产品或投资标的建议。"],
    ["同步文档", "对应 MD：docs/superpowers/specs/2026-05-20-single-stage-family-asset-diagnosis-design.md"],
  ]);

  const input = workbook.addWorksheet("示例输入");
  input.columns = [{ width: 26 }, { width: 18 }, { width: 18 }, { width: 42 }];
  setTitle(input, "示例输入", "当前填入一个中等收入单身样本，可直接替换 B 列数值进行测试。");
  input.addRow([]);
  input.addRow(["资产负债表", "金额", "单位", "说明"]);
  styleHeader(input.getRow(4));
  const assetRows = [
    ["总资产", sample.totalAssets, "元", "全部资产合计"],
    ["总负债", sample.totalLiabilities, "元", "房贷、车贷、消费贷等合计"],
    ["净资产", { formula: "B5-B6", result: sample.netAssets }, "元", "总资产 - 总负债"],
    ["流动资产", sample.liquidAssets, "元", "现金、活期、货币基金等"],
    ["短期负债", sample.shortTermLiabilities, "元", "一年内需偿还的短期债务"],
    ["生息资产", sample.interestBearingAssets, "元", "可产生利息、分红、收益的资产"],
    ["可投资资产", sample.investableAssets, "元", "流动资产 + 可配置投资资产"],
    ["长期投资资产", sample.longTermInvestmentAssets, "元", "基金、股票、养老金账户等长期资产"],
    ["年被动收入", sample.passiveIncome, "元", "理财、股息、租金等非工作收入"],
    ["年投资收益", sample.investmentReturn, "元", "当年投资收益"],
    ["期初投资资产", sample.beginningInvestmentAssets, "元", "用于计算投资报酬率"],
  ];
  input.addRows(assetRows);
  input.addRow([]);
  input.addRow(["收支储蓄表", "金额", "单位", "说明"]);
  styleHeader(input.getRow(16));
  const incomeRows = [
    ["年收入", sample.annualIncome, "元", "税后总收入"],
    ["年支出", sample.annualExpense, "元", "全年总支出"],
    ["月必要支出", sample.monthlyNecessaryExpense, "元", "必要生活成本、房租房贷、基本开销"],
    ["固定支出", sample.fixedExpense, "元", "房租/房贷、通勤、基本生活等刚性支出"],
    ["年债务本息支出", sample.debtService, "元", "全年还款本金和利息"],
    ["高息消费贷余额", sample.highInterestDebt, "元", "信用卡滚债、消费贷等"],
    ["年保费", sample.annualPremium, "元", "商业保险年保费"],
    ["已配置基础保障项", sample.configuredProtectionItems, "项", "医保、意外、重疾等"],
    ["应配置基础保障项", sample.requiredProtectionItems, "项", "单身期 MVP 默认 3 项"],
    ["有效重大风险保障额度", sample.majorRiskCoverage, "元", "可覆盖重大疾病、意外、收入中断等风险的额度"],
  ];
  input.addRows(incomeRows);
  input.getColumn(2).numFmt = "#,##0";

  const params = workbook.addWorksheet("指标参数");
  params.columns = [
    { width: 14 }, { width: 11 }, { width: 20 }, { width: 12 }, { width: 36 },
    { width: 12 }, { width: 10 }, { width: 10 }, { width: 12 }, { width: 36 },
  ];
  setTitle(params, "指标参数", "参数表是后续校准的核心：权重、m、k 和锚点都应可配置。");
  params.addRow([]);
  params.addRow(["维度", "维度权重", "指标", "维度内权重", "公式", "方向", "m", "k", "评分类型", "经验锚点"]);
  styleHeader(params.getRow(4));
  indicators.forEach((r) => {
    const d = dimensionMap[r[0]];
    params.addRow([r[0], d.weight, r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8]]);
  });
  params.getColumn(2).numFmt = "0%";
  params.getColumn(4).numFmt = "0%";
  params.getColumn(7).numFmt = "0.00";
  params.getColumn(8).numFmt = "0.00";

  const calc = workbook.addWorksheet("评分计算");
  calc.columns = [
    { width: 14 }, { width: 22 }, { width: 14 }, { width: 14 }, { width: 12 },
    { width: 10 }, { width: 12 }, { width: 14 }, { width: 50 },
  ];
  setTitle(calc, "评分计算", "根据示例输入计算原始指标、S 曲线得分、状态和加权分。");
  calc.addRow([]);
  calc.addRow(["维度", "指标", "原始值", "展示值", "指标得分", "状态", "维度内权重", "加权分", "公式口径"]);
  styleHeader(calc.getRow(4));
  indicators.forEach((r, index) => {
    const rowIndex = 5 + index;
    const x = rawValue(r[1]);
    const score = scoreIndicator(r);
    const display = r[1].includes("倍数") || r[1].includes("覆盖倍数") || r[1].includes("支出倍数")
      ? multiple(x)
      : pct(x);
    calc.addRow([
      r[0],
      r[1],
      { formula: rawFormula(r[1]), result: x },
      display,
      { formula: scoreFormula(rowIndex, r), result: score },
      { formula: `IF(E${rowIndex}>=70,"绿灯",IF(E${rowIndex}>=55,"黄灯","红灯"))`, result: status(score) },
      r[2],
      { formula: `E${rowIndex}*G${rowIndex}`, result: score * r[2] },
      r[3],
    ]);
  });
  calc.getColumn(3).numFmt = "0.000";
  calc.getColumn(5).numFmt = "0.0";
  calc.getColumn(7).numFmt = "0%";
  calc.getColumn(8).numFmt = "0.0";

  const overview = workbook.addWorksheet("六维总览");
  overview.columns = [
    { width: 16 }, { width: 12 }, { width: 13 }, { width: 10 }, { width: 46 }, { width: 18 },
  ];
  setTitle(overview, "六维总览", "综合分、六维分数、红黄绿状态和一句话诊断。");
  overview.addRow([]);
  overview.addRow(["综合健康分", "", "", "", "", ""]);
  overview.mergeCells("A4:B5");
  const scoresByDimension = {};
  indicators.forEach((r) => {
    const score = scoreIndicator(r);
    scoresByDimension[r[0]] = (scoresByDimension[r[0]] || 0) + score * r[2];
  });
  const overall = dimensions.reduce((sum, d) => sum + scoresByDimension[d.name] * d.weight, 0);
  overview.getCell("A4").value = { formula: "SUMPRODUCT(B8:B13,C8:C13)", result: Number(overall.toFixed(1)) };
  overview.getCell("A4").font = { name: "Microsoft YaHei", bold: true, size: 28, color: { argb: "111827" } };
  overview.getCell("C4").value = { formula: 'IF(A4>=70,"绿灯",IF(A4>=55,"黄灯","红灯"))', result: status(overall) };
  overview.getCell("C4").font = { name: "Microsoft YaHei", bold: true, color: { argb: overall >= 70 ? "15803D" : overall >= 55 ? "B45309" : "B91C1C" } };
  overview.mergeCells("D4:F5");
  overview.getCell("D4").value = "整体处于基础待改善状态：偿债压力和储蓄能力会影响后续成长能力，流动性安全垫需要优先观察。";
  overview.getCell("D4").alignment = { wrapText: true, vertical: "middle" };
  overview.addRow([]);
  overview.addRow(["维度", "权重", "维度得分", "状态", "定位", "颜色"]);
  styleHeader(overview.getRow(7));
  dimensions.forEach((d, index) => {
    const s = scoresByDimension[d.name];
    const rowIndex = 8 + index;
    overview.addRow([
      d.name,
      d.weight,
      { formula: `SUMIFS('评分计算'!H:H,'评分计算'!A:A,A${rowIndex})`, result: s },
      { formula: `IF(C${rowIndex}>=70,"绿灯",IF(C${rowIndex}>=55,"黄灯","红灯"))`, result: status(s) },
      d.role,
      "",
    ]);
    const row = overview.lastRow;
    row.getCell(6).fill = { type: "pattern", pattern: "solid", fgColor: { argb: d.color } };
  });
  overview.getColumn(2).numFmt = "0%";
  overview.getColumn(3).numFmt = "0.0";

  const chains = workbook.addWorksheet("连锁诊断");
  chains.columns = [
    { width: 12 }, { width: 24 }, { width: 36 }, { width: 36 }, { width: 52 }, { width: 52 }, { width: 10 },
  ];
  setTitle(chains, "连锁诊断", "解释维度之间的传导关系：先找起点问题，再给前置修复动作。");
  chains.addRow([]);
  chains.addRow(["强度", "传导链条", "触发条件", "当前证据", "影响解释", "优先行动", "是否触发"]);
  styleHeader(chains.getRow(4));
  const raw = Object.fromEntries(indicators.map((r) => [r[1], rawValue(r[1])]));
  const dim = scoresByDimension;
  const chainRows = [
    ["强", "储蓄能力 -> 成长能力", "储蓄能力低于60分，或自由储蓄率低于10%", `储蓄能力 ${dim["储蓄能力"].toFixed(1)} 分，自由储蓄率 ${pct(raw["自由储蓄率"])}`, "新增可投资资金不足，长期资产配置改善速度受限。", "先提高月度结余和自由储蓄率，再逐步增加长期投资资产。", dim["储蓄能力"] < 60 || raw["自由储蓄率"] < 0.10],
    ["中", "储蓄能力 -> 流动性", "总结余率低于10%，且紧急备用金低于3个月", `总结余率 ${pct(raw["总结余率"])}，备用金 ${multiple(raw["紧急备用金倍数"])}`, "很难靠未来现金流快速补足安全垫。", "压缩非必要支出，把新增结余优先补入流动账户。", raw["总结余率"] < 0.10 && raw["紧急备用金倍数"] < 3],
    ["强", "偿债能力 -> 流动性", "债务偿付率高于30%，或高息负债收入比高于10%", `债务偿付率 ${pct(raw["债务偿付率"])}，高息负债收入比 ${pct(raw["高息负债收入比"])}`, "还款占用现金流，收入波动时更容易出现短期资金缺口。", "优先处理高息负债，并避免继续扩大固定还款义务。", raw["债务偿付率"] > 0.30 || raw["高息负债收入比"] > 0.10],
    ["中", "偿债能力 -> 成长能力", "债务偿付率高于30%，且长期投资资产占比低于50%", `债务偿付率 ${pct(raw["债务偿付率"])}，长期投资资产占比 ${pct(raw["长期投资资产占比"])}`, "债务支出挤压长期投资，资产增长速度受限。", "降低还款压力后，再提高长期投资资产占比。", raw["债务偿付率"] > 0.30 && raw["长期投资资产占比"] < 0.50],
    ["强", "保障能力 -> 流动性", "基础保障完整度低于60%，且紧急备用金低于3个月", `保障完整度 ${pct(raw["基础保障完整度"])}，备用金 ${multiple(raw["紧急备用金倍数"])}`, "医疗、意外或收入中断事件可能快速消耗备用金。", "先补齐基础保障，再同步建立紧急备用金。", raw["基础保障完整度"] < 0.60 && raw["紧急备用金倍数"] < 3],
    ["中", "流动性 -> 成长能力", "紧急备用金低于3个月，且长期投资资产占比较高", `备用金 ${multiple(raw["紧急备用金倍数"])}，长期投资资产占比 ${pct(raw["长期投资资产占比"])}`, "安全垫不足时继续投资，突发支出时可能被迫卖出资产。", "暂停增加高波动资产，先把备用金补到3个月以上。", raw["紧急备用金倍数"] < 3 && raw["长期投资资产占比"] > 0.50],
    ["弱", "成长能力 -> 财富自由度", "成长能力低于60分，且被动收入覆盖率低于10%", `成长能力 ${dim["成长能力"].toFixed(1)} 分，被动收入覆盖率 ${pct(raw["被动收入覆盖率"])}`, "资产缺乏增值和现金流贡献，财富自由进度偏慢。", "在安全底座稳定后，逐步提高生息资产和长期投资资产占比。", dim["成长能力"] < 60 && raw["被动收入覆盖率"] < 0.10],
  ];
  const triggerFormulas = [
    'IF(OR(\'六维总览\'!C8<60,\'评分计算\'!C12<0.1),"是","否")',
    'IF(AND(\'评分计算\'!C11<0.1,\'评分计算\'!C8<3),"是","否")',
    'IF(OR(\'评分计算\'!C6>0.3,\'评分计算\'!C7>0.1),"是","否")',
    'IF(AND(\'评分计算\'!C6>0.3,\'评分计算\'!C15<0.5),"是","否")',
    'IF(AND(\'评分计算\'!C17<0.6,\'评分计算\'!C8<3),"是","否")',
    'IF(AND(\'评分计算\'!C8<3,\'评分计算\'!C15>0.5),"是","否")',
    'IF(AND(\'六维总览\'!C12<60,\'评分计算\'!C20<0.1),"是","否")',
  ];
  chains.addRows(chainRows.map((r, index) => [...r.slice(0, 6), { formula: triggerFormulas[index], result: r[6] ? "是" : "否" }]));

  const rec = workbook.addWorksheet("建议库");
  rec.columns = [{ width: 10 }, { width: 18 }, { width: 28 }, { width: 55 }, { width: 42 }];
  setTitle(rec, "建议库", "MVP 只输出方向性改善动作，不输出产品级推荐。");
  rec.addRow([]);
  rec.addRow(["优先级", "类别", "触发条件", "建议话术", "资产结构方向"]);
  styleHeader(rec.getRow(4));
  rec.addRows([
    [1, "急性风险", "月结余 < 0", "先让现金流转正，暂停非必要大额支出和高波动投资。", "提高现金类资产，减少消费性支出。"],
    [1, "流动性", "紧急备用金 < 3个月", "优先补足至少3个月必要支出的紧急备用金。", "新增结余优先进入现金/货币类账户。"],
    [1, "偿债", "高息负债收入比 > 10%", "优先偿还高息消费贷，避免利息侵蚀储蓄能力。", "降低负债，释放现金流。"],
    [2, "保障", "基础保障完整度 < 60%", "补齐医保、意外、重疾等基础保障，避免风险事件击穿资产。", "用合理保费换取重大风险转移。"],
    [3, "储蓄", "总结余率 < 20%", "优化支出结构，提高收入转化为资产的效率。", "提高每月可投资现金流。"],
    [4, "成长", "生息资产占比 < 35%", "在流动性安全后，逐步提高生息资产和长期资产占比。", "从闲置资金转向长期增值资产。"],
    [5, "自由度", "被动收入覆盖率 < 10%", "财富自由度是结果指标，先修复储蓄、负债、流动性和成长能力。", "长期提高资产现金流贡献。"],
  ]);

  [readme, input, params, calc, overview, chains, rec].forEach((ws) => {
    styleSheet(ws);
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = cell.border || thinBorder();
      });
    });
  });

  const fs = require("fs");
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  await workbook.xlsx.writeFile(outPath);
  console.log(outPath);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
