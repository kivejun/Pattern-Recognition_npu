import pptxgen from "pptxgenjs";
import fs from "node:fs/promises";
import fsSync from "node:fs";
import os from "node:os";
import path from "node:path";

const __dirname = path.dirname(new URL(import.meta.url).pathname).replace(/^\/([A-Za-z]:)/, "$1");
const outPath = path.join(__dirname, "pattern-recognition-polished-editable.pptx");
const workRoot = path.join(os.tmpdir(), "codex-presentations", "manual-pattern-recognition-polished");
const tmpDir = path.join(workRoot, "tmp");
const qaDir = path.join(tmpDir, "qa");
const fig = (...parts) => path.join(__dirname, "figures", ...parts);

await fs.mkdir(qaDir, { recursive: true });

await fs.writeFile(path.join(tmpDir, "source-notes.txt"), [
  "模式识别课程实验汇报 source notes",
  "",
  "Sources:",
  "- reports/main.tex: existing LaTeX experiment report in repository.",
  "- results/experiment1_bayes/metrics.csv and decision_rules.txt: Bayes classifier metrics and parameters.",
  "- results/experiment2_linear/metrics.csv and fisher_params.txt: Parzen/Fisher/LOO metrics.",
  "- results/experiment3_kl/metrics.csv and kl_params.txt: K-L/PCA metrics and directions.",
  "- results/experiment4_kl_image/metrics.csv and pca_params.txt: image PCA metrics and dataset information.",
  "- results/experiment5_clustering/metrics.csv and cluster_notes.txt: clustering metrics.",
  "- reports/figures/**: experiment figures converted to PNG in slides/pattern-recognition-report/figures.",
  "",
  "No external logos or unsourced online assets were used.",
].join("\n"), "utf8");

await fs.writeFile(path.join(tmpDir, "slide-plan.txt"), [
  "Create mode plan for polished editable PowerPoint",
  "",
  "Design:",
  "- Palette: deep navy #0B1220 as dominant, teal #0F766E as method accent, amber #F59E0B as highlight, slate neutrals for text and panels.",
  "- Fonts: Microsoft YaHei for Chinese body/headings; Aptos fallback is left to Office if unavailable.",
  "- Style: clean academic technology deck with dark section dividers, editable KPI cards, native tables, shape-based process diagrams, and independent image objects.",
  "",
  "Structure:",
  "1 cover",
  "2 narrative and method map",
  "3 data/environment",
  "4-7 experiment 1",
  "8-11 experiment 2",
  "12-15 experiment 3",
  "16-19 experiment 4",
  "20-23 experiment 5",
  "24 cross-experiment comparison",
  "25 conclusion",
  "26 Q&A",
].join("\n"), "utf8");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Pattern Recognition Coursework";
pptx.company = "NPU";
pptx.subject = "模式识别课程实验汇报";
pptx.title = "模式识别课程实验汇报";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  lang: "zh-CN",
};
pptx.defineLayout({ name: "LAYOUT_WIDE", width: 13.333, height: 7.5 });

const C = {
  ink: "0B1220",
  navy: "111827",
  blue: "1D4ED8",
  teal: "0F766E",
  teal2: "14B8A6",
  amber: "F59E0B",
  orange: "EA580C",
  green: "16A34A",
  rose: "E11D48",
  slate: "475569",
  slate2: "64748B",
  line: "CBD5E1",
  pale: "F8FAFC",
  paleTeal: "ECFDF5",
  paleAmber: "FFFBEB",
  white: "FFFFFF",
};

const W = 13.333;
const H = 7.5;
const M = 0.56;
let pageNum = 1;

function slideBase(slide, dark = false) {
  slide.background = { color: dark ? C.ink : C.white };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: W,
    h: H,
    fill: { color: dark ? C.ink : C.white },
    line: { color: dark ? C.ink : C.white, transparency: 100 },
  });
  if (!dark) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: 0,
      w: 0.12,
      h: H,
      fill: { color: C.teal },
      line: { color: C.teal, transparency: 100 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.12,
      y: 0,
      w: 0.035,
      h: H,
      fill: { color: C.amber },
      line: { color: C.amber, transparency: 100 },
    });
  }
}

function addTitle(slide, title, kicker = "") {
  slide.addText(kicker || "PATTERN RECOGNITION COURSEWORK", {
    x: M,
    y: 0.28,
    w: 5.5,
    h: 0.22,
    fontFace: "Microsoft YaHei",
    fontSize: 7.8,
    bold: true,
    color: C.teal,
    charSpace: 0.6,
    margin: 0,
  });
  slide.addText(title, {
    x: M,
    y: 0.55,
    w: 9.2,
    h: 0.52,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: C.ink,
    fit: "shrink",
    margin: 0,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: M,
    y: 1.16,
    w: 1.12,
    h: 0,
    line: { color: C.amber, width: 2.4 },
  });
}

function addFooter(slide, label = "模式识别课程实验汇报") {
  slide.addText(label, {
    x: M,
    y: 7.13,
    w: 4.6,
    h: 0.18,
    fontSize: 7,
    color: "94A3B8",
    margin: 0,
  });
  slide.addText(String(pageNum++).padStart(2, "0"), {
    x: 12.45,
    y: 7.07,
    w: 0.32,
    h: 0.2,
    fontSize: 8,
    bold: true,
    color: C.teal,
    align: "right",
    margin: 0,
  });
}

function newSlide(title, kicker = "") {
  const slide = pptx.addSlide();
  slideBase(slide);
  addTitle(slide, title, kicker);
  addFooter(slide);
  return slide;
}

function textBox(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    fontFace: "Microsoft YaHei",
    fontSize: opts.size ?? 11,
    bold: opts.bold ?? false,
    color: opts.color ?? C.slate,
    fit: opts.fit ?? "shrink",
    valign: opts.valign ?? "top",
    align: opts.align ?? "left",
    breakLine: false,
    margin: opts.margin ?? 0.04,
  });
}

function bullet(slide, items, x, y, w, h, opts = {}) {
  slide.addText(items.map((t) => ({ text: t, options: { bullet: { type: "ul" } } })), {
    x,
    y,
    w,
    h,
    fontFace: "Microsoft YaHei",
    fontSize: opts.size ?? 10.5,
    color: opts.color ?? C.slate,
    fit: "shrink",
    breakLine: false,
    paraSpaceAfterPt: opts.after ?? 6,
    margin: 0.06,
  });
}

function panel(slide, x, y, w, h, opts = {}) {
  const shape = {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: opts.fill ?? C.pale },
    line: { color: opts.line ?? "E2E8F0", width: opts.lineWidth ?? 0.75 },
  };
  if (opts.shadow) {
    shape.shadow = { type: "outer", color: "CBD5E1", opacity: 0.18, blur: 2, angle: 45, distance: 1 };
  }
  slide.addShape(pptx.ShapeType.roundRect, shape);
}

function card(slide, x, y, w, h, head, body, accent = C.teal, opts = {}) {
  panel(slide, x, y, w, h, { fill: opts.fill ?? C.white, line: opts.line ?? "E2E8F0", shadow: true });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w: 0.08,
    h,
    fill: { color: accent },
    line: { color: accent, transparency: 100 },
  });
  textBox(slide, head, x + 0.22, y + 0.16, w - 0.34, 0.25, { size: opts.headSize ?? 11.4, bold: true, color: C.ink });
  textBox(slide, body, x + 0.22, y + 0.5, w - 0.34, h - 0.58, { size: opts.bodySize ?? 9.2, color: opts.bodyColor ?? C.slate });
}

function metric(slide, x, y, w, h, value, label, accent = C.teal) {
  panel(slide, x, y, w, h, { fill: C.pale, line: "E2E8F0", shadow: true });
  textBox(slide, value, x + 0.18, y + 0.17, w - 0.35, 0.43, { size: 22, bold: true, color: accent, fit: "shrink" });
  textBox(slide, label, x + 0.2, y + 0.68, w - 0.35, 0.3, { size: 7.8, color: C.slate2 });
}

function img(slide, relPath, x, y, w, h, opts = {}) {
  if (!fsSync.existsSync(relPath)) throw new Error(`Missing image: ${relPath}`);
  panel(slide, x, y, w, h, { fill: C.white, line: "E2E8F0" });
  slide.addImage({
    path: relPath,
    x: x + (opts.pad ?? 0.04),
    y: y + (opts.pad ?? 0.04),
    w: w - 2 * (opts.pad ?? 0.04),
    h: h - 2 * (opts.pad ?? 0.04),
    sizing: { type: "contain", x: x + (opts.pad ?? 0.04), y: y + (opts.pad ?? 0.04), w: w - 2 * (opts.pad ?? 0.04), h: h - 2 * (opts.pad ?? 0.04) },
  });
}

function nativeTable(slide, rows, x, y, w, h, colW, opts = {}) {
  const tableRows = rows.map((row, r) => row.map((cell) => ({
    text: cell,
    options: {
      bold: r === 0,
      color: r === 0 ? C.white : C.ink,
      fill: r === 0 ? { color: opts.headerFill ?? C.teal } : { color: r % 2 ? "FFFFFF" : "F8FAFC" },
      valign: "mid",
      margin: 0.04,
    },
  })));
  slide.addTable(tableRows, {
    x,
    y,
    w,
    h,
    colW,
    border: { type: "solid", color: "CBD5E1", pt: 0.45 },
    fontFace: "Microsoft YaHei",
    fontSize: opts.size ?? 7.8,
    fit: "shrink",
    autoFit: false,
    rowH: Array(rows.length).fill(h / rows.length),
  });
}

function arrow(slide, x1, y1, x2, y2, color = C.teal) {
  slide.addShape(pptx.ShapeType.line, {
    x: x1,
    y: y1,
    w: x2 - x1,
    h: y2 - y1,
    line: { color, width: 2, beginArrowType: "none", endArrowType: "triangle" },
  });
}

function step(slide, x, y, w, title, desc, idx, color = C.teal) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x,
    y,
    w: 0.42,
    h: 0.42,
    fill: { color },
    line: { color, transparency: 100 },
  });
  textBox(slide, String(idx), x, y + 0.08, 0.42, 0.18, { size: 8.5, bold: true, color: C.white, align: "center" });
  textBox(slide, title, x + 0.55, y - 0.02, w - 0.55, 0.22, { size: 10.5, bold: true, color: C.ink });
  textBox(slide, desc, x + 0.55, y + 0.28, w - 0.55, 0.42, { size: 8.2, color: C.slate });
}

function sectionSlide(num, title, subtitle, accents = [C.teal, C.amber]) {
  const s = pptx.addSlide();
  slideBase(s, true);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: H, fill: { color: accents[0] }, line: { color: accents[0] } });
  s.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 0.06, h: H, fill: { color: accents[1] }, line: { color: accents[1] } });
  textBox(s, num, 0.82, 1.0, 1.8, 0.65, { size: 23, bold: true, color: accents[1] });
  textBox(s, title, 0.82, 2.05, 10.8, 0.8, { size: 34, bold: true, color: C.white });
  textBox(s, subtitle, 0.86, 3.08, 9.5, 0.45, { size: 15, color: "CBD5E1" });
  s.addShape(pptx.ShapeType.line, { x: 0.86, y: 4.0, w: 2.8, h: 0, line: { color: accents[1], width: 4 } });
  s.addText("模式识别课程实验汇报", { x: 0.86, y: 6.95, w: 4, h: 0.2, fontSize: 7.5, color: "94A3B8", margin: 0 });
}

function miniBar(slide, x, y, label, value, color = C.teal, max = 100) {
  textBox(slide, label, x, y, 2.0, 0.2, { size: 8.2, color: C.slate });
  slide.addShape(pptx.ShapeType.rect, { x: x + 2.1, y: y + 0.04, w: 2.4, h: 0.12, fill: { color: "E2E8F0" }, line: { color: "E2E8F0" } });
  slide.addShape(pptx.ShapeType.rect, { x: x + 2.1, y: y + 0.04, w: 2.4 * value / max, h: 0.12, fill: { color }, line: { color } });
  textBox(slide, `${value.toFixed(2)}%`, x + 4.62, y - 0.01, 0.65, 0.22, { size: 8, color: C.ink, bold: true, align: "right" });
}

// Cover
{
  const s = pptx.addSlide();
  slideBase(s, true);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: C.ink }, line: { color: C.ink } });
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: H, fill: { color: C.teal }, line: { color: C.teal } });
  s.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 0.06, h: H, fill: { color: C.amber }, line: { color: C.amber } });
  s.addText("PATTERN RECOGNITION", { x: 0.85, y: 1.55, w: 4.2, h: 0.24, fontSize: 8.5, bold: true, color: C.teal2, charSpace: 1.5, margin: 0 });
  s.addText("模式识别课程实验汇报", { x: 0.82, y: 2.15, w: 8.4, h: 0.78, fontSize: 38, bold: true, color: C.white, fit: "shrink", margin: 0 });
  s.addText("Bayes 分类器 · Fisher 判别 · K-L 变换 · 图像识别 · 聚类分析", { x: 0.88, y: 3.13, w: 8.2, h: 0.34, fontSize: 15, color: "CBD5E1", margin: 0 });
  card(s, 0.9, 4.55, 2.8, 0.76, "姓名", "请填写", C.teal2, { fill: "111827", line: "334155", bodyColor: "CBD5E1" });
  card(s, 4.1, 4.55, 2.8, 0.76, "学号", "请填写", C.amber, { fill: "111827", line: "334155", bodyColor: "CBD5E1" });
  card(s, 7.3, 4.55, 2.8, 0.76, "班级", "请填写", C.blue, { fill: "111827", line: "334155", bodyColor: "CBD5E1" });
  s.addShape(pptx.ShapeType.arc, { x: 9.5, y: 0.8, w: 4.8, h: 4.8, adjustPoint: 0.15, line: { color: "334155", width: 2 }, fill: { color: C.ink, transparency: 100 } });
}

{
  const s = newSlide("汇报结构：从分类到聚类的完整方法链", "TALK ROADMAP");
  const nodes = [
    ["01", "Bayes 分类", "概率密度估计与风险决策", C.teal],
    ["02", "Fisher / Parzen", "非参数估计与直接判别", C.blue],
    ["03", "K-L 特征提取", "协方差特征分解与降维", C.amber],
    ["04", "K-L 图像应用", "高维图像压缩与识别", C.orange],
    ["05", "聚类分析", "无监督样本结构发现", C.green],
  ];
  nodes.forEach((d, i) => {
    const x = 0.8 + i * 2.48;
    card(s, x, 1.65, 2.02, 1.38, `${d[0]} ${d[1]}`, d[2], d[3], { bodySize: 8.4 });
    if (i < nodes.length - 1) arrow(s, x + 2.05, 2.32, x + 2.42, 2.32, "94A3B8");
  });
  card(s, 1.0, 4.05, 2.25, 1.05, "理论知识", "先讲课程概念和公式", C.teal);
  card(s, 3.8, 4.05, 2.25, 1.05, "实验设置", "说明数据、特征与参数", C.blue);
  card(s, 6.6, 4.05, 2.25, 1.05, "实验结果", "展示指标与图像结果", C.amber);
  card(s, 9.4, 4.05, 2.25, 1.05, "分析总结", "解释现象与课程联系", C.orange);
}

{
  const s = newSlide("数据与实验环境", "SOURCE & SETUP");
  nativeTable(s, [
    ["数据集", "实验用途", "规模/说明"],
    ["FEMALE / MALE", "实验一、二、三、五", "50 女 + 50 男训练样本"],
    ["test1 / test2", "分类器泛化测试", "35 / 300 个测试样本"],
    ["飞机图像数据集", "实验四", "16 类，2530 训练图，636 测试图"],
  ], 0.75, 1.35, 6.25, 2.45, [1.8, 2.45, 2.0], { size: 8.2 });
  nativeTable(s, [
    ["工具", "作用"],
    ["Python 3.11", "实验脚本运行"],
    ["NumPy / SciPy", "数值计算与矩阵运算"],
    ["scikit-learn", "PCA、评估指标、分级聚类辅助"],
    ["Matplotlib", "分类边界、投影、聚类图输出"],
  ], 7.45, 1.35, 4.85, 2.65, [1.55, 3.3], { size: 8.2, headerFill: C.blue });
  metric(s, 0.95, 4.72, 2.4, 0.95, "5", "已完成实验任务", C.teal);
  metric(s, 3.65, 4.72, 2.4, 0.95, "18", "汇报中使用图表", C.blue);
  metric(s, 6.35, 4.72, 2.4, 0.95, "2D/1024D", "低维与图像高维特征", C.amber);
  metric(s, 9.05, 4.72, 2.4, 0.95, "分类+聚类", "监督与无监督覆盖", C.orange);
}

sectionSlide("01", "实验一：Bayes 性别分类器", "用身高/体重数据进行最小错误率与最小风险决策");

{
  const s = newSlide("实验一理论：概率决策框架", "BAYES CLASSIFIER");
  card(s, 0.75, 1.35, 3.65, 1.25, "后验概率决策", "选择 P(ωi|x) 最大的类别，本质是把观测特征映射到类别后验。", C.teal);
  card(s, 4.85, 1.35, 3.65, 1.25, "高斯参数估计", "在正态假设下估计均值 μ 与协方差 Σ，得到二次判别函数。", C.blue);
  card(s, 8.95, 1.35, 3.1, 1.25, "风险敏感决策", "用损失表刻画不同错误代价，选择条件风险最小的动作。", C.amber);
  panel(s, 1.0, 3.45, 10.95, 1.35, { fill: C.paleTeal, line: "A7F3D0" });
  textBox(s, "g_i(x) = -1/2 ln|Σ_i| - 1/2(x-μ_i)^TΣ_i^{-1}(x-μ_i) + ln P(ω_i)", 1.2, 3.84, 10.55, 0.42, { size: 20, bold: true, color: C.tealDark, align: "center" });
  bullet(s, ["课程相关：参数估计、先验概率、贝叶斯判别函数、最小风险决策", "实验目的：观察特征、先验、协方差假设和风险表对分类性能的影响"], 1.1, 5.35, 10.5, 0.92, { size: 10.2 });
}

{
  const s = newSlide("实验一设置：数据、特征与实现流程", "EXPERIMENT DESIGN");
  step(s, 0.9, 1.45, 4.4, "读取数据", "FEMALE/MALE 作为训练集；test1/test2 作为测试集", 1, C.teal);
  step(s, 0.9, 2.55, 4.4, "选择特征", "身高、体重、身高+体重三组对比", 2, C.blue);
  step(s, 0.9, 3.65, 4.4, "估计分布", "最大似然估计 μ、σ 或 Σ", 3, C.amber);
  step(s, 0.9, 4.75, 4.4, "分类评估", "准确率、召回率、混淆矩阵、错误分析", 4, C.orange);
  nativeTable(s, [
    ["设置项", "取值"],
    ["先验概率", "0.5/0.5，0.75/0.25，0.9/0.1"],
    ["协方差", "完整协方差 / 对角协方差"],
    ["风险表", "λ(M|F)=2，λ(F|M)=1"],
    ["非参数补充", "Parzen 窗、kNN"],
  ], 6.25, 1.55, 5.55, 2.7, [1.6, 3.95], { size: 8.5 });
  card(s, 6.55, 4.78, 4.9, 0.85, "实现要点", "所有分类器统一输出 train、test1、test2 上的性能，便于横向比较。", C.teal);
}

{
  const s = newSlide("实验一结果：数据分布与 Bayes 决策边界", "RESULT VISUALS");
  img(s, fig("experiment1", "data_scatter.png"), 0.78, 1.35, 5.65, 4.2);
  img(s, fig("experiment1", "bayes_2d_boundary.png"), 6.8, 1.35, 5.65, 4.2);
  metric(s, 0.95, 5.86, 2.2, 0.72, "88.00%", "二维 Bayes 训练准确率", C.teal);
  metric(s, 3.4, 5.86, 2.2, 0.72, "97.14%", "test1 准确率", C.blue);
  metric(s, 5.85, 5.86, 2.2, 0.72, "89.33%", "test2 准确率", C.amber);
  metric(s, 8.3, 5.86, 2.2, 0.72, "90.33%", "对角协方差 test2", C.orange);
}

{
  const s = newSlide("实验一分析：先验与风险改变决策偏向", "ANALYSIS");
  nativeTable(s, [
    ["方法", "训练集", "test1", "test2"],
    ["二维 Bayes 完整协方差", "88.00%", "97.14%", "89.33%"],
    ["二维 Bayes 对角协方差", "88.00%", "97.14%", "90.33%"],
    ["最小风险 Bayes", "84.00%", "94.29%", "84.33%"],
    ["Parzen Bayes", "87.00%", "97.14%", "90.33%"],
    ["kNN，k=5", "89.00%", "100.00%", "90.33%"],
  ], 0.8, 1.35, 6.65, 2.72, [2.65, 1.2, 1.2, 1.2], { size: 7.7 });
  img(s, fig("experiment1", "bayes_test2_confusion.png"), 8.1, 1.32, 3.55, 2.7);
  bullet(s, [
    "身高和体重组合能利用互补信息，整体比单特征更稳定",
    "test2 中男生比例更高，若人为设置女生先验过大，会显著降低整体准确率",
    "最小风险决策体现“错误代价不同”时的分类器设计思想",
  ], 0.95, 4.75, 11.1, 1.25, { size: 10.2 });
}

sectionSlide("02", "实验二：非参数估计与 Fisher 判别", "比较概率密度估计方法与直接设计线性分类器的方法");

{
  const s = newSlide("实验二理论：Parzen、Fisher 与留一法", "NONPARAMETRIC & LINEAR");
  card(s, 0.75, 1.35, 3.6, 1.35, "Parzen 窗估计", "不预设概率密度形式，用高斯核叠加估计类别条件密度。", C.teal);
  card(s, 4.85, 1.35, 3.6, 1.35, "Fisher 线性判别", "最大化类间离散度与类内离散度之比，得到线性投影方向。", C.blue);
  card(s, 8.95, 1.35, 3.1, 1.35, "留一法估计", "每次留出一个样本验证，适合小样本泛化误差估计。", C.amber);
  panel(s, 1.1, 3.55, 10.7, 1.2, { fill: "EFF6FF", line: "BFDBFE" });
  textBox(s, "J(w)=wᵀS_b w / wᵀS_w w,    w=S_w⁻¹(μ₁-μ₂)", 1.28, 3.95, 10.35, 0.38, { size: 22, bold: true, color: C.blue, align: "center" });
  bullet(s, ["课程重点：参数估计与非参数估计的差别；概率模型与判别模型的差别", "实验问题：非参数 Bayes 和 Fisher 是否能达到高斯 Bayes 的性能"], 1.0, 5.35, 10.8, 0.9, { size: 10.2 });
}

{
  const s = newSlide("实验二结果：线性边界与投影分布", "VISUAL COMPARISON");
  img(s, fig("experiment2", "fisher_bayes_boundaries.png"), 0.8, 1.35, 5.75, 4.35);
  img(s, fig("experiment2", "fisher_projection.png"), 6.85, 1.35, 5.55, 4.35);
  card(s, 1.1, 5.95, 5.1, 0.55, "观察 1", "Fisher 给出简单线性边界，Bayes 边界受密度估计影响。", C.blue);
  card(s, 6.95, 5.95, 5.1, 0.55, "观察 2", "投影后男女样本主体分离，但边界区域仍有重叠。", C.teal);
}

{
  const s = newSlide("实验二结果：错误率对比", "METRICS");
  nativeTable(s, [
    ["方法", "训练错误率", "test1 错误率", "test2 错误率", "留一法错误率"],
    ["Gaussian Bayes", "12.00%", "2.86%", "10.67%", "12.00%"],
    ["Parzen Bayes", "13.00%", "2.86%", "9.67%", "14.00%"],
    ["Fisher", "12.00%", "2.86%", "9.67%", "13.00%"],
  ], 0.9, 1.45, 11.4, 2.1, [2.5, 2.05, 2.15, 2.15, 2.15], { size: 8 });
  miniBar(s, 1.2, 4.35, "Gaussian Bayes test2 acc.", 89.33, C.teal);
  miniBar(s, 1.2, 4.82, "Parzen Bayes test2 acc.", 90.33, C.blue);
  miniBar(s, 1.2, 5.29, "Fisher test2 acc.", 90.33, C.amber);
  card(s, 7.1, 4.25, 4.65, 1.4, "结论", "三种方法表现接近。Fisher 用更简单的线性边界取得与非参数 Bayes 相当的 test2 准确率。", C.teal);
}

{
  const s = newSlide("实验二分析：模型复杂度与泛化", "TAKEAWAYS");
  card(s, 0.9, 1.45, 3.35, 1.5, "参数 Bayes", "形式清晰、样本效率高，但依赖正态分布假设。", C.teal);
  card(s, 4.9, 1.45, 3.35, 1.5, "Parzen Bayes", "密度形式灵活，但窗口宽度 h 和样本规模影响明显。", C.blue);
  card(s, 8.9, 1.45, 3.1, 1.5, "Fisher", "不估计密度，直接寻找判别方向，解释性强。", C.amber);
  panel(s, 1.0, 4.1, 10.9, 1.3, { fill: C.paleAmber, line: "FDE68A" });
  textBox(s, "留一法结果与 test2 错误率接近，说明它能在训练样本有限时给出较稳健的泛化误差估计。", 1.35, 4.5, 10.2, 0.45, { size: 15, bold: true, color: "92400E", align: "center" });
}

sectionSlide("03", "实验三：K-L 变换进行特征提取", "从无监督主成分方向理解身高体重样本的主要变化");

{
  const s = newSlide("实验三理论：K-L/PCA 与 Fisher 的差别", "FEATURE EXTRACTION");
  card(s, 0.85, 1.35, 4.9, 1.3, "K-L / PCA", "对协方差矩阵做特征分解，保留总体方差最大的方向。", C.teal);
  card(s, 6.15, 1.35, 4.9, 1.3, "Fisher", "利用类别标签，寻找最有利于分类的投影方向。", C.blue);
  panel(s, 1.0, 3.35, 10.85, 1.15, { fill: C.paleTeal, line: "A7F3D0" });
  textBox(s, "Σuᵢ = λᵢuᵢ,      zᵢ = uᵢᵀ(x-μ)", 1.25, 3.75, 10.3, 0.34, { size: 24, bold: true, color: C.tealDark, align: "center" });
  bullet(s, ["PCA 的目标是表示数据，不是直接优化分类", "当最大方差方向与类别差异方向一致时，PCA 也能得到较好分类效果"], 1.0, 5.3, 10.8, 0.85, { size: 10.5 });
}

{
  const s = newSlide("实验三设置：二维特征的主方向分析", "EXPERIMENT DESIGN");
  nativeTable(s, [
    ["项目", "数值/说明"],
    ["数据", "FEMALE + MALE，共 100 个训练样本"],
    ["特征", "身高、体重"],
    ["第一主成分解释率", "87.52%"],
    ["PC1", "(0.6269, 0.7791)"],
    ["类均值方向", "(0.6514, 0.7587)"],
  ], 0.85, 1.35, 4.65, 2.85, [1.8, 2.85], { size: 8.2 });
  img(s, fig("experiment3", "kl_directions.png"), 6.05, 1.2, 5.75, 4.45);
  card(s, 1.0, 5.08, 4.35, 0.75, "设置说明", "分别使用 PCA 第一主成分、类均值差方向、Fisher 和 Bayes 进行分类比较。", C.teal);
}

{
  const s = newSlide("实验三结果：投影分布与分类性能", "RESULTS");
  img(s, fig("experiment3", "pca_projection.png"), 0.82, 1.3, 5.6, 3.75);
  nativeTable(s, [
    ["方法", "训练集", "test1", "test2"],
    ["K-L PCA 第一主成分", "86.00%", "100.00%", "89.67%"],
    ["类均值方向", "86.00%", "100.00%", "89.67%"],
    ["Fisher", "88.00%", "97.14%", "90.33%"],
    ["Gaussian Bayes", "88.00%", "97.14%", "89.33%"],
  ], 6.75, 1.45, 5.25, 2.45, [2.25, 1.0, 1.0, 1.0], { size: 7.8 });
  bullet(s, ["PCA 第一主成分已经能较好区分男女样本", "Fisher 在 test2 上略优，体现监督判别方向的优势"], 6.88, 4.45, 5.0, 1.0, { size: 10.2 });
}

{
  const s = newSlide("实验三分析：最大方差不等于最佳分类", "TAKEAWAYS");
  card(s, 0.9, 1.45, 3.45, 1.35, "为什么 PCA 有效", "本数据中身高和体重的主要变化方向与性别差异方向接近。", C.teal);
  card(s, 4.9, 1.45, 3.45, 1.35, "为什么 Fisher 略优", "Fisher 直接利用类别标签，目标就是提高类别分离度。", C.blue);
  card(s, 8.9, 1.45, 3.15, 1.35, "一般规律", "无监督降维需要和监督分类器共同评价。", C.amber);
  img(s, fig("experiment3", "mean_projection.png"), 2.1, 3.4, 8.8, 2.7);
}

sectionSlide("04", "实验四：K-L 变换的图像识别应用", "把 K-L/PCA 从二维样本扩展到高维飞机图像");

{
  const s = newSlide("实验四理论：高维图像的低维表示", "IMAGE RECOGNITION");
  step(s, 0.9, 1.45, 4.2, "图像预处理", "灰度化、缩放到 32×32", 1, C.teal);
  step(s, 0.9, 2.55, 4.2, "向量化", "每张图像展开为 1024 维向量", 2, C.blue);
  step(s, 0.9, 3.65, 4.2, "K-L 降维", "提取前 k 个主成分作为图像特征", 3, C.amber);
  step(s, 0.9, 4.75, 4.2, "低维分类", "NearestCentroid 与 1NN 对比", 4, C.orange);
  panel(s, 6.35, 1.55, 4.95, 1.0, { fill: "EFF6FF", line: "BFDBFE" });
  textBox(s, "z = U_kᵀ(x-μ)", 6.55, 1.87, 4.5, 0.34, { size: 24, bold: true, color: C.blue, align: "center" });
  nativeTable(s, [
    ["项目", "设置"],
    ["类别数", "16 类"],
    ["训练集", "2530 张"],
    ["测试集", "636 张"],
    ["主成分", "8 / 16 / 32 / 64 / 128"],
  ], 6.35, 3.05, 4.95, 2.1, [1.65, 3.3], { size: 8.4, headerFill: C.blue });
}

{
  const s = newSlide("实验四结果：样本图像与 PCA 主成分", "VISUAL RESULTS");
  img(s, fig("experiment4", "sample_images.png"), 0.8, 1.35, 5.7, 4.35);
  img(s, fig("experiment4", "pca_components.png"), 6.85, 1.35, 5.55, 4.35);
  metric(s, 1.1, 5.92, 2.3, 0.7, "1024→64", "维度压缩", C.blue);
  metric(s, 3.85, 5.92, 2.3, 0.7, "89.88%", "64 维解释方差", C.teal);
  metric(s, 6.6, 5.92, 2.3, 0.7, "16", "飞机类别数", C.amber);
  metric(s, 9.35, 5.92, 2.3, 0.7, "636", "测试图像数", C.orange);
}

{
  const s = newSlide("实验四结果：主成分数与识别准确率", "METRICS");
  img(s, fig("experiment4", "accuracy_vs_components.png"), 0.85, 1.3, 5.8, 4.45);
  nativeTable(s, [
    ["方法", "主成分数", "训练集", "测试集"],
    ["PCA+NearestCentroid", "64", "20.47%", "19.97%"],
    ["PCA+1NN", "32", "100.00%", "54.09%"],
    ["PCA+1NN", "64", "100.00%", "53.93%"],
    ["PCA+1NN", "128", "100.00%", "52.99%"],
  ], 7.0, 1.55, 4.95, 2.35, [2.15, 1.0, 0.9, 0.9], { size: 7.7, headerFill: C.orange });
  bullet(s, ["1NN 明显优于最近质心，说明飞机类别内部变化复杂", "主成分数增加到 32 后收益变小，后续成分可能包含背景和噪声"], 7.1, 4.45, 4.9, 1.05, { size: 9.8 });
}

{
  const s = newSlide("实验四分析：PCA 能压缩，但不是完整识别方案", "ANALYSIS");
  img(s, fig("experiment4", "confusion_matrix.png"), 0.9, 1.35, 5.8, 4.45);
  card(s, 7.15, 1.55, 4.25, 1.05, "混淆原因", "低分辨率灰度图中，部分飞机外形相近，姿态和背景也会干扰。", C.orange);
  card(s, 7.15, 3.05, 4.25, 1.05, "方法局限", "PCA 是无监督降维，优先保留总体方差，不保证保留最强判别信息。", C.blue);
  card(s, 7.15, 4.55, 4.25, 1.05, "改进方向", "可结合 HOG/SIFT/ORB、数据增强或 CNN 提升识别性能。", C.teal);
}

sectionSlide("05", "实验五：C 均值与分级聚类分析", "不使用类别标签，观察身高体重样本的自然结构");

{
  const s = newSlide("实验五理论：从监督分类到无监督聚类", "CLUSTERING");
  card(s, 0.85, 1.35, 4.6, 1.3, "C 均值聚类", "独立编程实现，通过样本分配和中心更新最小化类内平方误差。", C.teal);
  card(s, 6.0, 1.35, 4.6, 1.3, "分级聚类", "使用 Ward 方法形成层次结构，用树状图观察样本合并过程。", C.blue);
  panel(s, 1.1, 3.45, 10.6, 1.1, { fill: C.paleTeal, line: "A7F3D0" });
  textBox(s, "J = Σⱼ Σₓ∈ωⱼ ||x - μⱼ||²", 1.3, 3.83, 10.2, 0.36, { size: 25, bold: true, color: C.tealDark, align: "center" });
  bullet(s, ["聚类不使用性别标签，评价时再与真实标签比较", "特征标准化很重要，否则体重/身高尺度会影响距离计算"], 1.0, 5.3, 10.8, 0.85, { size: 10.5 });
}

{
  const s = newSlide("实验五设置：两组数据与多种聚类数", "EXPERIMENT DESIGN");
  nativeTable(s, [
    ["设置项", "说明"],
    ["数据一", "FEMALE + MALE，共 100 个训练样本"],
    ["数据二", "训练集 + test2，共 400 个样本"],
    ["特征", "身高、体重，先标准化"],
    ["C 均值类别数", "2、3、4、5"],
    ["初始值", "多个随机种子比较稳定性"],
  ], 0.85, 1.35, 5.55, 3.0, [1.75, 3.8], { size: 8.2 });
  step(s, 7.0, 1.55, 4.5, "初始化中心", "随机选择 C 个中心", 1, C.teal);
  step(s, 7.0, 2.55, 4.5, "样本分配", "分配到最近中心", 2, C.blue);
  step(s, 7.0, 3.55, 4.5, "中心更新", "计算各簇均值", 3, C.amber);
  step(s, 7.0, 4.55, 4.5, "指标评估", "SSE、轮廓系数、ARI、匹配率", 4, C.orange);
}

{
  const s = newSlide("实验五结果：C=2 聚类与初始值影响", "C-MEANS RESULTS");
  img(s, fig("experiment5", "cmeans_train_2clusters.png"), 0.8, 1.35, 5.65, 4.25);
  img(s, fig("experiment5", "init_variation.png"), 6.85, 1.35, 5.55, 4.25);
  metric(s, 1.1, 5.85, 2.25, 0.72, "85.00%", "训练集 C=2 匹配率", C.teal);
  metric(s, 3.75, 5.85, 2.25, 0.72, "85-87%", "不同初始值波动", C.blue);
  metric(s, 6.4, 5.85, 2.25, 0.72, "0.5108", "轮廓系数", C.amber);
  metric(s, 9.05, 5.85, 2.25, 0.72, "0.4850", "ARI", C.orange);
}

{
  const s = newSlide("实验五结果：类别数选择与分级聚类", "MODEL SELECTION");
  img(s, fig("experiment5", "cmeans_k_metrics_train.png"), 0.8, 1.35, 5.65, 4.2);
  img(s, fig("experiment5", "hierarchical_dendrogram_train.png"), 6.85, 1.35, 5.55, 4.2);
  card(s, 1.0, 5.85, 5.0, 0.62, "类别数判断", "SSE 随 C 增大下降，但轮廓系数在 C=2 最高，因此两类更合理。", C.teal);
  card(s, 6.95, 5.85, 5.0, 0.62, "分级聚类", "Ward 方法匹配率 81.00%，略低于 C 均值，但能展示层次结构。", C.blue);
}

{
  const s = newSlide("实验五分析：加入 test2 后结构保持但边界更复杂", "COMBINED DATA");
  img(s, fig("experiment5", "cmeans_combined_2clusters.png"), 0.8, 1.3, 5.55, 3.85);
  img(s, fig("experiment5", "cmeans_k_metrics_combined.png"), 6.75, 1.3, 5.55, 3.85);
  nativeTable(s, [
    ["数据", "方法", "轮廓系数", "ARI", "匹配率"],
    ["训练集", "C 均值", "0.5108", "0.4850", "85.00%"],
    ["训练集", "Ward", "0.5049", "0.3789", "81.00%"],
    ["训练集+test2", "C 均值", "0.4850", "0.4693", "84.50%"],
    ["训练集+test2", "Ward", "0.4739", "0.4172", "82.50%"],
  ], 0.85, 5.55, 11.35, 0.95, [1.8, 1.8, 1.9, 1.5, 1.5], { size: 6.9 });
}

{
  const s = newSlide("五个实验的横向总结", "SYNTHESIS");
  nativeTable(s, [
    ["实验", "课程知识点", "监督信息", "主要目标"],
    ["Bayes 分类", "后验概率、参数估计、风险决策", "使用标签", "最小错误率或最小风险"],
    ["Fisher / Parzen", "非参数估计、线性判别、留一法", "使用标签", "比较概率模型与直接判别"],
    ["K-L 特征提取", "PCA、协方差特征分解", "可不使用标签", "降维与主方向分析"],
    ["K-L 图像应用", "高维图像特征压缩", "使用标签", "验证 PCA 在图像识别中的作用"],
    ["聚类分析", "C 均值、分级聚类", "不使用标签", "发现样本自然结构"],
  ], 0.75, 1.35, 11.75, 4.05, [1.65, 4.1, 1.9, 4.1], { size: 7.6 });
  panel(s, 1.0, 5.85, 11.1, 0.62, { fill: C.paleTeal, line: "A7F3D0" });
  textBox(s, "主线：从概率决策到线性判别，从监督分类到无监督聚类，从低维特征到高维图像识别。", 1.22, 6.02, 10.6, 0.24, { size: 12, bold: true, color: C.tealDark, align: "center" });
}

{
  const s = newSlide("总体结论与汇报收束", "CONCLUSION");
  card(s, 0.9, 1.35, 3.5, 1.25, "特征选择", "身高+体重比单特征更稳定；图像任务中 PCA 可压缩高维特征。", C.teal);
  card(s, 4.9, 1.35, 3.5, 1.25, "模型假设", "正态假设、窗口宽度、线性边界都会影响最终性能。", C.blue);
  card(s, 8.9, 1.35, 3.1, 1.25, "评价指标", "准确率、错误率、召回率、混淆矩阵和聚类指标需要综合分析。", C.amber);
  panel(s, 1.05, 3.75, 10.95, 1.25, { fill: C.paleAmber, line: "FDE68A" });
  textBox(s, "核心体会：模式识别方法没有绝对最优，算法选择必须结合数据分布、任务目标和评价指标。", 1.42, 4.18, 10.2, 0.42, { size: 17, bold: true, color: "92400E", align: "center" });
  bullet(s, ["Bayes 强调概率决策，Fisher 强调判别方向，K-L 强调表示与降维，聚类强调自然结构。"], 1.2, 5.65, 10.6, 0.55, { size: 10.2 });
}

{
  const s = pptx.addSlide();
  slideBase(s, true);
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.22, h: H, fill: { color: C.teal }, line: { color: C.teal } });
  s.addShape(pptx.ShapeType.rect, { x: 0.22, y: 0, w: 0.06, h: H, fill: { color: C.amber }, line: { color: C.amber } });
  textBox(s, "谢谢", 0.9, 2.45, 4.5, 0.75, { size: 44, bold: true, color: C.white });
  textBox(s, "欢迎老师和同学批评指正", 0.94, 3.45, 6.0, 0.35, { size: 18, color: "CBD5E1" });
  s.addShape(pptx.ShapeType.line, { x: 0.94, y: 4.18, w: 2.65, h: 0, line: { color: C.amber, width: 4 } });
}

await pptx.writeFile({ fileName: outPath });

await fs.writeFile(path.join(qaDir, "visual-qa.txt"), [
  "Visual QA ledger for polished editable PPTX",
  "",
  `Final PPTX: ${outPath}`,
  "Checks performed in generation script:",
  "- Uses native PowerPoint text boxes for headings, bullets, cards, formulas, and conclusions.",
  "- Uses native PowerPoint tables for metric tables.",
  "- Uses native editable shapes for panels, step diagrams, progress bars, and KPI cards.",
  "- Embeds experiment figures as separate image objects, not as full-slide bitmaps.",
  "- No external image assets or unverified logos were used.",
  "",
  "Manual follow-up checks should inspect PowerPoint rendering on the target machine.",
].join("\n"), "utf8");

console.log(outPath);
console.log(`WORKSPACE=${workRoot}`);
