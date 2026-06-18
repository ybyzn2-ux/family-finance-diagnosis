const ExcelJS = require('exceljs');
const path = require('path');

const wb = new ExcelJS.Workbook();
wb.creator = '家庭资产健康度诊断工具';

// ===== Color palette =====
const COLORS = {
  primary: '1B4F72',
  primaryLight: 'D4E6F1',
  headerBg: '2C3E50',
  headerFont: 'FFFFFF',
  inputBg: 'FFF9C4',
  inputFont: '0000FF',
  lockBg: 'F5F5F5',
  green: '27AE60',
  yellow: 'F39C12',
  red: 'E74C3C',
  greenBg: 'E8F8F5',
  yellowBg: 'FEF9E7',
  redBg: 'FDEDEC',
  borderColor: 'BDC3C7',
  white: 'FFFFFF',
  lightGray: 'ECF0F1',
  dimText: '7F8C8D',
};

const FONT_MAIN = { name: 'Microsoft YaHei', size: 11 };
const FONT_TITLE = { name: 'Microsoft YaHei', size: 18, bold: true, color: { argb: COLORS.primary } };
const FONT_SUBTITLE = { name: 'Microsoft YaHei', size: 11, color: { argb: COLORS.dimText } };
const FONT_HEADER = { name: 'Microsoft YaHei', size: 11, bold: true, color: { argb: COLORS.headerFont } };
const FONT_INPUT = { name: 'Microsoft YaHei', size: 11, color: { argb: COLORS.inputFont } };
const FONT_SECTION = { name: 'Microsoft YaHei', size: 13, bold: true, color: { argb: COLORS.primary } };

const BORDER_THIN = {
  top: { style: 'thin', color: { argb: COLORS.borderColor } },
  bottom: { style: 'thin', color: { argb: COLORS.borderColor } },
  left: { style: 'thin', color: { argb: COLORS.borderColor } },
  right: { style: 'thin', color: { argb: COLORS.borderColor } },
};

function headerFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.headerBg } };
}
function inputFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.inputBg } };
}
function lockFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lockBg } };
}
function whiteFill() {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.white } };
}

function applyHeaderRow(ws, rowNum, colCount) {
  const row = ws.getRow(rowNum);
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.font = FONT_HEADER;
    cell.fill = headerFill();
    cell.border = BORDER_THIN;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
  row.height = 28;
}

function applyDataRow(ws, rowNum, colCount, opts = {}) {
  const row = ws.getRow(rowNum);
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    cell.font = { ...FONT_MAIN };
    cell.border = BORDER_THIN;
    cell.alignment = { vertical: 'middle', wrapText: true };
    if (opts.inputCols && opts.inputCols.includes(c)) {
      cell.fill = inputFill();
      cell.font = FONT_INPUT;
    } else {
      cell.fill = whiteFill();
    }
  }
  row.height = opts.height || 22;
}

// ===== Sheet 1: 使用说明 =====
const wsReadme = wb.addWorksheet('使用说明', { properties: { tabColor: { argb: COLORS.primary } } });
wsReadme.columns = [{ width: 4 }, { width: 18 }, { width: 70 }];

wsReadme.mergeCells('B2:C2');
wsReadme.getCell('B2').value = '家庭资产健康度诊断工具';
wsReadme.getCell('B2').font = FONT_TITLE;
wsReadme.getCell('B2').alignment = { vertical: 'middle' };
wsReadme.getRow(2).height = 40;

wsReadme.mergeCells('B3:C3');
wsReadme.getCell('B3').value = '基于 RFP（注册财务策划师）体系设计，六维度全面评估家庭资产结构健康状况。';
wsReadme.getCell('B3').font = FONT_SUBTITLE;

const steps = [
  ['第一步', '打开"资产填写"表，在黄色单元格中填入您的实际数据。所有金额单位为人民币元。'],
  ['第二步', '自动计算完成后，查看"诊断报告"表，了解六维度综合得分和各项指标状态。'],
  ['第三步', '查看"改善建议"表，获取基于您实际情况的针对性改善方向。'],
];

let row = 5;
wsReadme.mergeCells(`B${row}:C${row}`);
wsReadme.getCell(`B${row}`).value = '使用步骤';
wsReadme.getCell(`B${row}`).font = FONT_SECTION;
row++;

steps.forEach(([step, desc]) => {
  wsReadme.getCell(`B${row}`).value = step;
  wsReadme.getCell(`B${row}`).font = { ...FONT_MAIN, bold: true, color: { argb: COLORS.primary } };
  wsReadme.getCell(`C${row}`).value = desc;
  wsReadme.getCell(`C${row}`).font = FONT_MAIN;
  wsReadme.getCell(`C${row}`).alignment = { wrapText: true };
  wsReadme.getRow(row).height = 28;
  row++;
});

row++;
wsReadme.mergeCells(`B${row}:C${row}`);
wsReadme.getCell(`B${row}`).value = '重要说明';
wsReadme.getCell(`B${row}`).font = FONT_SECTION;
row++;

const notes = [
  '本工具只输出方向性资产结构建议，不构成具体金融产品或投资标的推荐。',
  '评分采用 S 曲线模型，基于 RFP 体系经验阈值，仅供参考。',
  '请勿修改"诊断报告"和"改善建议"中的公式，只需在"资产填写"表填写数据即可。',
  '绿灯(>=70分)=健康；黄灯(55-70分)=需关注；红灯(<55分)=需改善。',
];
notes.forEach(n => {
  wsReadme.getCell(`B${row}`).value = '·';
  wsReadme.getCell(`B${row}`).font = FONT_MAIN;
  wsReadme.getCell(`C${row}`).value = n;
  wsReadme.getCell(`C${row}`).font = FONT_MAIN;
  wsReadme.getCell(`C${row}`).alignment = { wrapText: true };
  wsReadme.getRow(row).height = 26;
  row++;
});

// ===== Sheet 2: 资产填写 =====
const wsInput = wb.addWorksheet('资产填写', { properties: { tabColor: { argb: 'F39C12' } } });
wsInput.columns = [
  { width: 4 },
  { width: 22 },
  { width: 16 },
  { width: 8 },
  { width: 50 },
];

wsInput.mergeCells('B1:E1');
wsInput.getCell('B1').value = '家庭资产数据填写';
wsInput.getCell('B1').font = FONT_TITLE;
wsInput.getRow(1).height = 36;

wsInput.mergeCells('B2:E2');
wsInput.getCell('B2').value = '请在下方黄色单元格中填入您的实际数据（所有金额单位：人民币元）';
wsInput.getCell('B2').font = FONT_SUBTITLE;

// Section 1: 资产负债
let r = 4;
wsInput.mergeCells(`B${r}:E${r}`);
wsInput.getCell(`B${r}`).value = '一、资产负债情况';
wsInput.getCell(`B${r}`).font = FONT_SECTION;
wsInput.getRow(r).height = 28;

r++;
['项目', '金额（元）', '单位', '填写说明'].forEach((h, i) => {
  wsInput.getCell(r, i + 2).value = h;
});
applyHeaderRow(wsInput, r, 5);

const assetItems = [
  ['总资产', 1050000, '元', '所有资产合计：现金+存款+理财+基金+股票+保险现金价值+房产+车辆等'],
  ['总负债', 120000, '元', '所有负债合计：房贷+车贷+消费贷+信用卡待还等'],
  ['流动资产', 205000, '元', '可随时变现的资产：现金、活期存款、货币基金、短期银行理财等'],
  ['短期负债', 36000, '元', '一年内需偿还的负债：信用卡待还、短期消费贷等'],
  ['生息资产', 260000, '元', '能产生利息/分红/收益的资产：定期存款、债券、基金、股票、理财产品等'],
  ['基金股票等长期投资', 680000, '元', '持有1年以上的投资资产：指数基金、股票、养老金账户等（不含流动资产）'],
  ['高息消费贷余额', 0, '元', '年化利率>10%的负债：信用卡循环、网贷、消费分期等'],
];

assetItems.forEach(([name, val, unit, desc]) => {
  r++;
  wsInput.getCell(r, 2).value = name;
  wsInput.getCell(r, 3).value = val;
  wsInput.getCell(r, 4).value = unit;
  wsInput.getCell(r, 5).value = desc;
  applyDataRow(wsInput, r, 5, { inputCols: [3] });
});
const assetStartRow = 6; // first data row
const assetEndRow = r;

// auto-calc: 净资产
r++;
wsInput.getCell(r, 2).value = '净资产（自动计算）';
wsInput.getCell(r, 2).font = { ...FONT_MAIN, italic: true, color: { argb: COLORS.dimText } };
wsInput.getCell(r, 3).value = { formula: `C${assetStartRow}-C${assetStartRow + 1}` };
wsInput.getCell(r, 3).numFmt = '#,##0';
wsInput.getCell(r, 4).value = '元';
wsInput.getCell(r, 5).value = '= 总资产 - 总负债';
applyDataRow(wsInput, r, 5);
wsInput.getCell(r, 3).fill = lockFill();
wsInput.getCell(r, 3).font = { ...FONT_MAIN, italic: true };
const netAssetRow = r;

// Section 2: 收支储蓄
r += 2;
wsInput.mergeCells(`B${r}:E${r}`);
wsInput.getCell(`B${r}`).value = '二、年度收支情况';
wsInput.getCell(`B${r}`).font = FONT_SECTION;
wsInput.getRow(r).height = 28;

r++;
['项目', '金额（元）', '单位', '填写说明'].forEach((h, i) => {
  wsInput.getCell(r, i + 2).value = h;
});
applyHeaderRow(wsInput, r, 5);

const incomeItems = [
  ['年税后总收入', 360000, '元', '全年到手总收入：工资+奖金+兼职+副业等（税后）'],
  ['年总支出', 300000, '元', '全年总支出：生活+房租/房贷+娱乐+购物等全部支出'],
  ['月必要支出', 25000, '元', '每月刚性支出：房租/房贷+餐饮+通勤+水电+必要生活费'],
  ['年固定支出', 150000, '元', '年度刚性支出：房租/房贷×12+通勤+基本生活费等无法压缩的支出'],
  ['年债务还款额', 110000, '元', '全年还款本金和利息总额（含房贷、车贷、消费贷等）'],
];

const incomeStartRow = r + 1;
incomeItems.forEach(([name, val, unit, desc]) => {
  r++;
  wsInput.getCell(r, 2).value = name;
  wsInput.getCell(r, 3).value = val;
  wsInput.getCell(r, 4).value = unit;
  wsInput.getCell(r, 5).value = desc;
  applyDataRow(wsInput, r, 5, { inputCols: [3] });
});
const incomeEndRow = r;

// Section 3: 投资与被动收入
r += 2;
wsInput.mergeCells(`B${r}:E${r}`);
wsInput.getCell(`B${r}`).value = '三、投资与被动收入';
wsInput.getCell(`B${r}`).font = FONT_SECTION;
wsInput.getRow(r).height = 28;

r++;
['项目', '金额（元）', '单位', '填写说明'].forEach((h, i) => {
  wsInput.getCell(r, i + 2).value = h;
});
applyHeaderRow(wsInput, r, 5);

const investItems = [
  ['年被动收入', 100000, '元', '非工作所得收入：理财利息+基金分红+股息+租金收入等'],
  ['年投资收益', 90000, '元', '今年投资账户的实际盈亏金额（可为负数）'],
  ['年初投资资产总值', 600000, '元', '今年1月1日投资账户的总市值（用于计算投资收益率）'],
];

const investStartRow = r + 1;
investItems.forEach(([name, val, unit, desc]) => {
  r++;
  wsInput.getCell(r, 2).value = name;
  wsInput.getCell(r, 3).value = val;
  wsInput.getCell(r, 4).value = unit;
  wsInput.getCell(r, 5).value = desc;
  applyDataRow(wsInput, r, 5, { inputCols: [3] });
});
const investEndRow = r;

// Section 4: 保障
r += 2;
wsInput.mergeCells(`B${r}:E${r}`);
wsInput.getCell(`B${r}`).value = '四、保险保障';
wsInput.getCell(`B${r}`).font = FONT_SECTION;
wsInput.getRow(r).height = 28;

r++;
['项目', '金额/数量', '单位', '填写说明'].forEach((h, i) => {
  wsInput.getCell(r, i + 2).value = h;
});
applyHeaderRow(wsInput, r, 5);

const insItems = [
  ['年保费总支出', 115000, '元', '全年商业保险保费支出（不含社保）'],
  ['已配置基础保障项数', 3, '项', '已配置的险种数（医保=1、意外险=1、重疾险=1，有则计数）'],
  ['应配置基础保障项数', 3, '项', '单身期建议至少3项：医保+意外险+重疾险'],
  ['有效重大风险保障额度', 1550000, '元', '重疾险+意外险+医疗险等能覆盖的最高赔付总额'],
];

const insStartRow = r + 1;
insItems.forEach(([name, val, unit, desc]) => {
  r++;
  wsInput.getCell(r, 2).value = name;
  wsInput.getCell(r, 3).value = val;
  wsInput.getCell(r, 4).value = unit;
  wsInput.getCell(r, 5).value = desc;
  applyDataRow(wsInput, r, 5, { inputCols: [3] });
});
const insEndRow = r;

// Add data validation for numeric inputs
const numericValidation = {
  type: 'whole',
  operator: 'greaterThanOrEqual',
  formulae: [0],
  showErrorMessage: true,
  errorTitle: '输入错误',
  error: '请输入大于等于0的数字',
};

// Apply to all input cells
for (let row = assetStartRow; row <= assetEndRow; row++) {
  wsInput.getCell(row, 3).dataValidation = numericValidation;
}
for (let row = incomeStartRow; row <= incomeEndRow; row++) {
  wsInput.getCell(row, 3).dataValidation = numericValidation;
}
for (let row = investStartRow; row <= investEndRow; row++) {
  if (row === investStartRow + 1) {
    // 年投资收益 can be negative
    wsInput.getCell(row, 3).dataValidation = {
      type: 'whole',
      operator: 'greaterThanOrEqual',
      formulae: [-999999999],
    };
  } else {
    wsInput.getCell(row, 3).dataValidation = numericValidation;
  }
}
for (let row = insStartRow; row <= insEndRow; row++) {
  wsInput.getCell(row, 3).dataValidation = numericValidation;
}

// Number format for money cells
for (const rows of [[assetStartRow, assetEndRow], [incomeStartRow, incomeEndRow], [investStartRow, investEndRow]]) {
  for (let row = rows[0]; row <= rows[1]; row++) {
    wsInput.getCell(row, 3).numFmt = '#,##0';
  }
}
wsInput.getCell(insStartRow, 3).numFmt = '#,##0';
wsInput.getCell(insStartRow + 3, 3).numFmt = '#,##0';

// ===== Named references for readability =====
// Build a cell reference map
// Asset section (assetStartRow = 6)
const REF = {
  totalAsset: `'资产填写'!C${assetStartRow}`,      // 总资产
  totalDebt: `'资产填写'!C${assetStartRow + 1}`,    // 总负债
  liquidAsset: `'资产填写'!C${assetStartRow + 2}`,  // 流动资产
  shortDebt: `'资产填写'!C${assetStartRow + 3}`,    // 短期负债
  incomeAsset: `'资产填写'!C${assetStartRow + 4}`,  // 生息资产
  longInvest: `'资产填写'!C${assetStartRow + 5}`,   // 长期投资
  highDebt: `'资产填写'!C${assetStartRow + 6}`,     // 高息消费贷
  netAsset: `'资产填写'!C${netAssetRow}`,            // 净资产

  annualIncome: `'资产填写'!C${incomeStartRow}`,     // 年收入
  annualExpense: `'资产填写'!C${incomeStartRow + 1}`, // 年支出
  monthExpense: `'资产填写'!C${incomeStartRow + 2}`,  // 月必要支出
  fixedExpense: `'资产填写'!C${incomeStartRow + 3}`,  // 年固定支出
  debtPayment: `'资产填写'!C${incomeStartRow + 4}`,   // 年债务还款

  passiveIncome: `'资产填写'!C${investStartRow}`,     // 年被动收入
  investReturn: `'资产填写'!C${investStartRow + 1}`,   // 年投资收益
  startInvest: `'资产填写'!C${investStartRow + 2}`,    // 年初投资资产

  annualPremium: `'资产填写'!C${insStartRow}`,         // 年保费
  coveredItems: `'资产填写'!C${insStartRow + 1}`,      // 已配置项数
  requiredItems: `'资产填写'!C${insStartRow + 2}`,     // 应配置项数
  riskCoverage: `'资产填写'!C${insStartRow + 3}`,      // 保障额度
};

// ===== Sheet 3: 诊断报告 =====
const wsReport = wb.addWorksheet('诊断报告', { properties: { tabColor: { argb: '27AE60' } } });
wsReport.columns = [
  { width: 4 },
  { width: 14 },
  { width: 20 },
  { width: 12 },
  { width: 14 },
  { width: 10 },
  { width: 12 },
  { width: 14 },
  { width: 36 },
];

wsReport.mergeCells('B1:I1');
wsReport.getCell('B1').value = '家庭资产健康度诊断报告';
wsReport.getCell('B1').font = FONT_TITLE;
wsReport.getRow(1).height = 36;

// ---- Overall score section ----
wsReport.mergeCells('B3:C3');
wsReport.getCell('B3').value = '综合健康分';
wsReport.getCell('B3').font = FONT_SECTION;

// The overall score will be computed after we define dimension scores
// We'll put dimensions first, then reference them

// ---- Six dimensions summary ----
wsReport.mergeCells('B5:I5');
wsReport.getCell('B5').value = '六维度总览';
wsReport.getCell('B5').font = FONT_SECTION;
wsReport.getRow(5).height = 28;

r = 6;
['维度', '权重', '维度得分', '状态', '核心定位'].forEach((h, i) => {
  wsReport.getCell(r, i + 2).value = h;
});
applyHeaderRow(wsReport, r, 6);

const dimensions = [
  ['偿债能力', 0.20, '控制负债规模和还款压力'],
  ['流动性', 0.20, '防止突发事件击穿财务安全'],
  ['储蓄能力', 0.22, '收入转化为资产的核心发动机'],
  ['成长能力', 0.15, '判断资产是否具备长期增值能力'],
  ['保障能力', 0.15, '转移重大风险，避免一次事故归零'],
  ['财富自由度', 0.08, '结果性指标，单身期低权重'],
];

const dimStartRow = 7;
dimensions.forEach(([name, weight, desc], idx) => {
  const row = dimStartRow + idx;
  wsReport.getCell(row, 2).value = name;
  wsReport.getCell(row, 3).value = weight;
  wsReport.getCell(row, 3).numFmt = '0%';
  // Score formula will be filled after indicators
  wsReport.getCell(row, 5).value = { formula: `IF(D${row}>=70,"绿灯",IF(D${row}>=55,"黄灯","红灯"))` };
  wsReport.getCell(row, 6).value = desc;
  applyDataRow(wsReport, row, 6, { height: 26 });
  wsReport.getCell(row, 2).font = { ...FONT_MAIN, bold: true };
});

// Overall score in B3
const dimEndRow = dimStartRow + dimensions.length - 1;
wsReport.getCell('D3').value = { formula: `SUMPRODUCT(C${dimStartRow}:C${dimEndRow},D${dimStartRow}:D${dimEndRow})` };
wsReport.getCell('D3').numFmt = '0.0';
wsReport.getCell('D3').font = { name: 'Microsoft YaHei', size: 24, bold: true };
wsReport.getCell('E3').value = { formula: `IF(D3>=70,"绿灯",IF(D3>=55,"黄灯","红灯"))` };
wsReport.getCell('E3').font = { name: 'Microsoft YaHei', size: 14, bold: true };

// ---- Detailed indicators section ----
let ir = dimEndRow + 2;
wsReport.mergeCells(`B${ir}:I${ir}`);
wsReport.getCell(`B${ir}`).value = '各项指标明细';
wsReport.getCell(`B${ir}`).font = FONT_SECTION;
wsReport.getRow(ir).height = 28;

ir++;
['维度', '指标', '原始值', '展示值', '得分', '状态', '维度内权重', '计算说明'].forEach((h, i) => {
  wsReport.getCell(ir, i + 2).value = h;
});
applyHeaderRow(wsReport, ir, 9);

// Now define all indicators with formulas
// Updated: removed 流动资产占比, redistributed weights
const indicators = [
  // 偿债能力
  {
    dim: '偿债能力', name: '资产负债率', dimWeight: 0.3,
    rawFormula: `${REF.totalDebt}/${REF.totalAsset}`,
    displayFormula: null, // will format as %
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(10*(${c}-0.4)))`,
    desc: '总负债 / 总资产',
  },
  {
    dim: '偿债能力', name: '债务偿付率', dimWeight: 0.45,
    rawFormula: `${REF.debtPayment}/${REF.annualIncome}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(12*(${c}-0.25)))`,
    desc: '年债务还款 / 年收入',
  },
  {
    dim: '偿债能力', name: '高息负债收入比', dimWeight: 0.25,
    rawFormula: `${REF.highDebt}/${REF.annualIncome}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(25*(${c}-0.1)))`,
    desc: '高息消费贷余额 / 年收入',
  },
  // 流动性 (REVISED: removed 流动资产占比, weight redistributed)
  {
    dim: '流动性', name: '紧急备用金倍数', dimWeight: 0.6,
    rawFormula: `${REF.liquidAsset}/${REF.monthExpense}`,
    displaySuffix: '个月',
    scoreFormula: (c) => `100/(1+EXP(-1.1*(${c}-3)))`,
    desc: '流动资产 / 月必要支出（3-6个月为健康区间）',
  },
  {
    dim: '流动性', name: '净流动覆盖倍数', dimWeight: 0.4,
    rawFormula: `(${REF.liquidAsset}-${REF.shortDebt})/${REF.monthExpense}`,
    displaySuffix: '个月',
    scoreFormula: (c) => `100/(1+EXP(-1.2*(${c}-2)))`,
    desc: '(流动资产-短期负债) / 月必要支出',
  },
  // 储蓄能力
  {
    dim: '储蓄能力', name: '总结余率', dimWeight: 0.45,
    rawFormula: `(${REF.annualIncome}-${REF.annualExpense})/${REF.annualIncome}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(-12*(${c}-0.2)))`,
    desc: '(年收入-年支出) / 年收入',
  },
  {
    dim: '储蓄能力', name: '自由储蓄率', dimWeight: 0.35,
    rawFormula: `(${REF.annualIncome}-${REF.fixedExpense}-${REF.debtPayment})/${REF.annualIncome}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(-12*(${c}-0.15)))`,
    desc: '(年收入-固定支出-债务还款) / 年收入',
  },
  {
    dim: '储蓄能力', name: '支出刚性比率', dimWeight: 0.2,
    rawFormula: `${REF.fixedExpense}/${REF.annualExpense}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(8*(${c}-0.6)))`,
    desc: '固定支出 / 总支出（越低越灵活）',
  },
  // 成长能力
  {
    dim: '成长能力', name: '生息资产占比', dimWeight: 0.35,
    rawFormula: `${REF.incomeAsset}/${REF.totalAsset}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(-8*(${c}-0.35)))`,
    desc: '生息资产 / 总资产',
  },
  {
    dim: '成长能力', name: '长期投资占总资产比', dimWeight: 0.35,
    rawFormula: `${REF.longInvest}/${REF.totalAsset}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(-8*(${c}-0.4)))`,
    desc: '长期投资资产 / 总资产',
  },
  {
    dim: '成长能力', name: '投资收益率', dimWeight: 0.3,
    rawFormula: `IF(${REF.startInvest}=0,0,${REF.investReturn}/${REF.startInvest})`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(-12*(${c}-0.04)))`,
    desc: '年投资收益 / 年初投资资产总值',
  },
  // 保障能力
  {
    dim: '保障能力', name: '基础保障完整度', dimWeight: 0.35,
    rawFormula: `${REF.coveredItems}/${REF.requiredItems}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(-8*(${c}-0.6)))`,
    desc: '已配置保障项 / 应配置保障项',
  },
  {
    dim: '保障能力', name: '保费收入比', dimWeight: 0.3,
    rawFormula: `${REF.annualPremium}/${REF.annualIncome}`,
    displaySuffix: '%',
    scoreFormula: (c) => `IF(${c}<0.05,100/(1+EXP(-50*(${c}-0.05))),IF(${c}<=0.12,95,IF(${c}<=0.2,95-(${c}-0.12)/0.08*35,MAX(20,60-(${c}-0.2)/0.1*40))))`,
    desc: '年保费 / 年收入（5%-12%为合理区间）',
  },
  {
    dim: '保障能力', name: '重大风险覆盖倍数', dimWeight: 0.35,
    rawFormula: `${REF.riskCoverage}/(${REF.monthExpense}*12)`,
    displaySuffix: '倍',
    scoreFormula: (c) => `100/(1+EXP(-0.8*(${c}-3)))`,
    desc: '有效保障额度 / 年必要支出',
  },
  // 财富自由度
  {
    dim: '财富自由度', name: '被动收入覆盖率', dimWeight: 0.5,
    rawFormula: `${REF.passiveIncome}/${REF.annualExpense}`,
    displaySuffix: '%',
    scoreFormula: (c) => `100/(1+EXP(-8*(${c}-0.1)))`,
    desc: '年被动收入 / 年支出',
  },
  {
    dim: '财富自由度', name: '净资产支出倍数', dimWeight: 0.5,
    rawFormula: `${REF.netAsset}/${REF.annualExpense}`,
    displaySuffix: '倍',
    scoreFormula: (c) => `100/(1+EXP(-1.1*(${c}-1)))`,
    desc: '净资产 / 年支出',
  },
];

const indicatorStartRow = ir + 1;
indicators.forEach((ind, idx) => {
  const row = indicatorStartRow + idx;
  const rawCell = `D${row}`;
  const scoreCell = `F${row}`;

  wsReport.getCell(row, 2).value = ind.dim;
  wsReport.getCell(row, 3).value = ind.name;
  wsReport.getCell(row, 4).value = { formula: ind.rawFormula };
  wsReport.getCell(row, 4).numFmt = '0.000';

  // Display value
  if (ind.displaySuffix === '%') {
    wsReport.getCell(row, 5).value = { formula: `TEXT(${rawCell}*100,"0.0")&"%"` };
  } else {
    wsReport.getCell(row, 5).value = { formula: `TEXT(${rawCell},"0.0")&"${ind.displaySuffix}"` };
  }

  // Score
  wsReport.getCell(row, 6).value = { formula: ind.scoreFormula(rawCell) };
  wsReport.getCell(row, 6).numFmt = '0.0';

  // Status
  wsReport.getCell(row, 7).value = { formula: `IF(${scoreCell}>=70,"绿灯",IF(${scoreCell}>=55,"黄灯","红灯"))` };

  // Weight
  wsReport.getCell(row, 8).value = ind.dimWeight;
  wsReport.getCell(row, 8).numFmt = '0%';

  // Desc
  wsReport.getCell(row, 9).value = ind.desc;

  applyDataRow(wsReport, row, 9, { height: 24 });
  wsReport.getCell(row, 2).font = { ...FONT_MAIN, color: { argb: COLORS.dimText } };
});
const indicatorEndRow = indicatorStartRow + indicators.length - 1;

// Now fill dimension scores by summing weighted indicator scores
const dimNames = ['偿债能力', '流动性', '储蓄能力', '成长能力', '保障能力', '财富自由度'];
dimNames.forEach((dimName, idx) => {
  const dimRow = dimStartRow + idx;
  // Sum of (score * weight) for indicators in this dimension
  // Use SUMPRODUCT with IF to match dimension name
  wsReport.getCell(dimRow, 4).value = {
    formula: `SUMPRODUCT((B${indicatorStartRow}:B${indicatorEndRow}="${dimName}")*F${indicatorStartRow}:F${indicatorEndRow}*H${indicatorStartRow}:H${indicatorEndRow})`
  };
  wsReport.getCell(dimRow, 4).numFmt = '0.0';
});

// Conditional formatting for status columns
const addStatusCF = (ws, cellRef) => {
  ws.addConditionalFormatting({
    ref: cellRef,
    rules: [
      {
        type: 'containsText',
        operator: 'containsText',
        text: '绿灯',
        style: { font: { color: { argb: COLORS.green }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.greenBg } } },
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: '黄灯',
        style: { font: { color: { argb: COLORS.yellow }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.yellowBg } } },
      },
      {
        type: 'containsText',
        operator: 'containsText',
        text: '红灯',
        style: { font: { color: { argb: COLORS.red }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.redBg } } },
      },
    ],
  });
};

addStatusCF(wsReport, `E${dimStartRow}:E${dimEndRow}`);
addStatusCF(wsReport, `G${indicatorStartRow}:G${indicatorEndRow}`);
addStatusCF(wsReport, `E3`);

// ===== Sheet 4: 改善建议 =====
const wsAdvice = wb.addWorksheet('改善建议', { properties: { tabColor: { argb: 'E74C3C' } } });
wsAdvice.columns = [
  { width: 4 },
  { width: 10 },
  { width: 14 },
  { width: 14 },
  { width: 40 },
  { width: 36 },
];

wsAdvice.mergeCells('B1:F1');
wsAdvice.getCell('B1').value = '个性化改善建议';
wsAdvice.getCell('B1').font = FONT_TITLE;
wsAdvice.getRow(1).height = 36;

wsAdvice.mergeCells('B2:F2');
wsAdvice.getCell('B2').value = '以下建议基于您的实际数据自动生成，按优先级排列。标记"需改善"的项目建议优先处理。';
wsAdvice.getCell('B2').font = FONT_SUBTITLE;

// Dynamic advice section
r = 4;
wsAdvice.mergeCells(`B${r}:F${r}`);
wsAdvice.getCell(`B${r}`).value = '改善建议（根据您的数据自动判断）';
wsAdvice.getCell(`B${r}`).font = FONT_SECTION;

r++;
['优先级', '类别', '是否触发', '建议', '资产结构方向'].forEach((h, i) => {
  wsAdvice.getCell(r, i + 2).value = h;
});
applyHeaderRow(wsAdvice, r, 6);

// Dynamic advice rules - each with a trigger formula
const adviceRules = [
  {
    priority: 1, category: '急性风险',
    trigger: `IF(${REF.annualIncome}-${REF.annualExpense}<0,"需改善","正常")`,
    advice: '月结余为负，先让现金流转正，暂停非必要大额支出和高波动投资。',
    direction: '提高现金类资产，减少消费性支出。',
  },
  {
    priority: 1, category: '流动性',
    trigger: `IF(${REF.liquidAsset}/${REF.monthExpense}<3,"需改善","正常")`,
    advice: '紧急备用金不足3个月，优先补足至少3个月必要支出的备用金。',
    direction: '新增结余优先进入现金/货币类账户。',
  },
  {
    priority: 1, category: '偿债',
    trigger: `IF(${REF.highDebt}/${REF.annualIncome}>0.1,"需改善","正常")`,
    advice: '高息负债占比过高，优先偿还高息消费贷，避免利息侵蚀储蓄能力。',
    direction: '降低负债，释放现金流。',
  },
  {
    priority: 2, category: '保障',
    trigger: `IF(${REF.coveredItems}/${REF.requiredItems}<0.6,"需改善","正常")`,
    advice: '基础保障不完整，补齐医保、意外、重疾等基础保障，避免风险事件击穿资产。',
    direction: '用合理保费换取重大风险转移。',
  },
  {
    priority: 2, category: '偿债',
    trigger: `IF(${REF.debtPayment}/${REF.annualIncome}>0.3,"需改善","正常")`,
    advice: '债务偿付率过高（>30%），考虑优化还款结构或降低负债规模。',
    direction: '降低每月固定还款压力，释放可投资现金流。',
  },
  {
    priority: 3, category: '储蓄',
    trigger: `IF((${REF.annualIncome}-${REF.annualExpense})/${REF.annualIncome}<0.2,"需改善","正常")`,
    advice: '总结余率低于20%，优化支出结构，提高收入转化为资产的效率。',
    direction: '提高每月可投资现金流。',
  },
  {
    priority: 4, category: '成长',
    trigger: `IF(${REF.incomeAsset}/${REF.totalAsset}<0.35,"需改善","正常")`,
    advice: '生息资产占比偏低，在流动性安全后，逐步提高生息资产比例。',
    direction: '从闲置资金转向长期增值资产。',
  },
  {
    priority: 5, category: '自由度',
    trigger: `IF(${REF.passiveIncome}/${REF.annualExpense}<0.1,"需改善","正常")`,
    advice: '被动收入覆盖率低，这是结果指标——先修复储蓄、负债、流动性和成长能力。',
    direction: '长期提高资产现金流贡献。',
  },
];

const adviceStartRow = r + 1;
adviceRules.forEach((rule, idx) => {
  const row = adviceStartRow + idx;
  wsAdvice.getCell(row, 2).value = rule.priority;
  wsAdvice.getCell(row, 3).value = rule.category;
  wsAdvice.getCell(row, 4).value = { formula: rule.trigger };
  wsAdvice.getCell(row, 5).value = rule.advice;
  wsAdvice.getCell(row, 6).value = rule.direction;
  applyDataRow(wsAdvice, row, 6, { height: 32 });
  wsAdvice.getCell(row, 5).alignment = { wrapText: true, vertical: 'middle' };
  wsAdvice.getCell(row, 6).alignment = { wrapText: true, vertical: 'middle' };
});
const adviceEndRow = adviceStartRow + adviceRules.length - 1;

// Conditional formatting for trigger column
wsAdvice.addConditionalFormatting({
  ref: `D${adviceStartRow}:D${adviceEndRow}`,
  rules: [
    {
      type: 'containsText',
      operator: 'containsText',
      text: '需改善',
      style: { font: { color: { argb: COLORS.red }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.redBg } } },
    },
    {
      type: 'containsText',
      operator: 'containsText',
      text: '正常',
      style: { font: { color: { argb: COLORS.green } }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.greenBg } } },
    },
  ],
});

// ---- Chain diagnosis section ----
r = adviceEndRow + 2;
wsAdvice.mergeCells(`B${r}:F${r}`);
wsAdvice.getCell(`B${r}`).value = '连锁风险诊断';
wsAdvice.getCell(`B${r}`).font = FONT_SECTION;

r++;
wsAdvice.mergeCells(`B${r}:F${r}`);
wsAdvice.getCell(`B${r}`).value = '维度之间存在传导关系：一个维度的问题可能影响其他维度。以下为自动检测结果。';
wsAdvice.getCell(`B${r}`).font = FONT_SUBTITLE;

r++;
['强度', '传导链', '是否触发', '影响', '优先行动'].forEach((h, i) => {
  wsAdvice.getCell(r, i + 2).value = h;
});
applyHeaderRow(wsAdvice, r, 6);

// Chain diagnosis formulas - reference the report sheet indicator rows
// We need to figure out which indicator rows correspond to which metrics
// Indicators are in order starting from indicatorStartRow
const indRowMap = {};
indicators.forEach((ind, idx) => {
  indRowMap[ind.name] = indicatorStartRow + idx;
});

const chainRules = [
  {
    strength: '强',
    chain: '储蓄能力 -> 成长能力',
    trigger: `IF(OR('诊断报告'!D${dimStartRow+2}<60,'诊断报告'!D${indRowMap['自由储蓄率']}<0.1),"已触发","未触发")`,
    impact: '储蓄能力不足时，可投资资金受限，长期资产配置改善速度慢。',
    action: '先提高月度结余和自由储蓄率，再逐步增加长期投资。',
  },
  {
    strength: '强',
    chain: '偿债能力 -> 流动性',
    trigger: `IF(OR('诊断报告'!D${indRowMap['债务偿付率']}>0.3,'诊断报告'!D${indRowMap['高息负债收入比']}>0.1),"已触发","未触发")`,
    impact: '还款占用现金流，收入波动时容易出现短期资金缺口。',
    action: '优先处理高息负债，避免继续扩大固定还款义务。',
  },
  {
    strength: '强',
    chain: '保障能力 -> 流动性',
    trigger: `IF(AND('诊断报告'!D${indRowMap['基础保障完整度']}<0.6,'诊断报告'!D${indRowMap['紧急备用金倍数']}<3),"已触发","未触发")`,
    impact: '保障缺失+备用金不足，意外事件可能快速消耗全部流动资产。',
    action: '先补齐基础保障，再同步建立紧急备用金。',
  },
  {
    strength: '中',
    chain: '储蓄能力 -> 流动性',
    trigger: `IF(AND('诊断报告'!D${indRowMap['总结余率']}<0.1,'诊断报告'!D${indRowMap['紧急备用金倍数']}<3),"已触发","未触发")`,
    impact: '结余率低+备用金不足，无法靠未来现金流快速补足安全垫。',
    action: '压缩非必要支出，新增结余优先补入流动账户。',
  },
  {
    strength: '中',
    chain: '流动性 -> 成长能力',
    trigger: `IF(AND('诊断报告'!D${indRowMap['紧急备用金倍数']}<3,'诊断报告'!D${indRowMap['长期投资占总资产比']}>0.4),"已触发","未触发")`,
    impact: '备用金不足时继续投资，突发支出可能被迫卖出资产。',
    action: '暂停增加高波动资产，先把备用金补到3个月以上。',
  },
  {
    strength: '弱',
    chain: '成长能力 -> 财富自由度',
    trigger: `IF(AND('诊断报告'!D${dimStartRow+3}<60,'诊断报告'!D${indRowMap['被动收入覆盖率']}<0.1),"已触发","未触发")`,
    impact: '资产缺乏增值和现金流贡献，财富自由进度偏慢。',
    action: '在安全底座稳定后，逐步提高生息资产占比。',
  },
];

const chainStartRow = r + 1;
chainRules.forEach((rule, idx) => {
  const row = chainStartRow + idx;
  wsAdvice.getCell(row, 2).value = rule.strength;
  wsAdvice.getCell(row, 3).value = rule.chain;
  wsAdvice.getCell(row, 4).value = { formula: rule.trigger };
  wsAdvice.getCell(row, 5).value = rule.impact;
  wsAdvice.getCell(row, 6).value = rule.action;
  applyDataRow(wsAdvice, row, 6, { height: 32 });
  wsAdvice.getCell(row, 5).alignment = { wrapText: true, vertical: 'middle' };
  wsAdvice.getCell(row, 6).alignment = { wrapText: true, vertical: 'middle' };
});
const chainEndRow = chainStartRow + chainRules.length - 1;

wsAdvice.addConditionalFormatting({
  ref: `D${chainStartRow}:D${chainEndRow}`,
  rules: [
    {
      type: 'containsText',
      operator: 'containsText',
      text: '已触发',
      style: { font: { color: { argb: COLORS.red }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.redBg } } },
    },
    {
      type: 'containsText',
      operator: 'containsText',
      text: '未触发',
      style: { font: { color: { argb: COLORS.green } }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COLORS.greenBg } } },
    },
  ],
});

// ===== Sheet 5: 指标参数 (hidden reference sheet) =====
const wsParam = wb.addWorksheet('指标参数', { properties: { tabColor: { argb: COLORS.dimText } } });
wsParam.state = 'hidden';
wsParam.columns = [{ width: 14 }, { width: 20 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 10 }, { width: 30 }];

wsParam.getCell('A1').value = '维度';
wsParam.getCell('B1').value = '指标';
wsParam.getCell('C1').value = '维度内权重';
wsParam.getCell('D1').value = 'm';
wsParam.getCell('E1').value = 'k';
wsParam.getCell('F1').value = '方向';
wsParam.getCell('G1').value = '经验锚点';

const paramData = [
  ['偿债能力', '资产负债率', 0.3, 0.4, -10, '越低越好', '20%优，40%中点，60%警示'],
  ['偿债能力', '债务偿付率', 0.45, 0.25, -12, '越低越好', '10%优，25%中点，40%警示'],
  ['偿债能力', '高息负债收入比', 0.25, 0.1, -25, '越低越好', '0%优，10%警示，20%红灯'],
  ['流动性', '紧急备用金倍数', 0.6, 3, -1.1, '越高越好', '1月红灯，3月中点，6月优'],
  ['流动性', '净流动覆盖倍数', 0.4, 2, -1.2, '越高越好', '0红灯，2中点，4优'],
  ['储蓄能力', '总结余率', 0.45, 0.2, 12, '越高越好', '5%红灯，20%中点，35%优'],
  ['储蓄能力', '自由储蓄率', 0.35, 0.15, 12, '越高越好', '0%红灯，15%中点，30%优'],
  ['储蓄能力', '支出刚性比率', 0.2, 0.6, -8, '越低越好', '40%优，60%中点，80%警示'],
  ['成长能力', '生息资产占比', 0.35, 0.35, 8, '越高越好', '15%红灯，35%中点，55%优'],
  ['成长能力', '长期投资占总资产比', 0.35, 0.4, 8, '越高越好', '15%红灯，40%中点，60%优'],
  ['成长能力', '投资收益率', 0.3, 0.04, 12, '越高越好', '0%低，4%中点，8%优'],
  ['保障能力', '基础保障完整度', 0.35, 0.6, 8, '越高越好', '0-30%红灯，60%中点，100%优'],
  ['保障能力', '保费收入比', 0.3, 0.05, 0, '区间最优', '5%-12%合理，>20%过重'],
  ['保障能力', '重大风险覆盖倍数', 0.35, 3, 0.8, '越高越好', '1倍红灯，3倍中点，6倍优'],
  ['财富自由度', '被动收入覆盖率', 0.5, 0.1, 8, '越高越好', '0%低，10%中点，30%优'],
  ['财富自由度', '净资产支出倍数', 0.5, 1, 1.1, '越高越好', '0低，1倍中点，3倍优'],
];

paramData.forEach((row, idx) => {
  row.forEach((val, col) => {
    wsParam.getCell(idx + 2, col + 1).value = val;
  });
});

// ===== Save =====
const outPath = path.join('C:/Users/Owner/Documents/family-asset-model/artifacts', '家庭资产健康度诊断工具-v2.xlsx');
wb.xlsx.writeFile(outPath).then(() => {
  console.log('Saved to:', outPath);
}).catch(err => {
  console.error('Error:', err);
});
