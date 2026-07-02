import pptxgen from "pptxgenjs";
import fs from "node:fs";
import path from "node:path";

const __dirname = path.dirname(new URL(import.meta.url).pathname).replace(/^\/([A-Za-z]:)/, "$1");
const outPath = path.join(__dirname, "pattern-recognition-editable.pptx");
const fig = (...parts) => path.join(__dirname, "figures", ...parts);

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Pattern Recognition Coursework";
pptx.subject = "模式识别课程实验汇报";
pptx.title = "模式识别课程实验汇报";
pptx.company = "NPU";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  lang: "zh-CN",
};
pptx.defineLayout({ name: "LAYOUT_WIDE", width: 13.333, height: 7.5 });

const C = {
  navy: "0F172A",
  teal: "0F766E",
  tealDark: "115E59",
  cyan: "155E75",
  slate: "475569",
  light: "F8FAFC",
  line: "D8E2E8",
  white: "FFFFFF",
  orange: "EA580C",
  green: "16A34A",
};

const W = 13.333;
const H = 7.5;
const margin = 0.55;

function addBg(slide) {
  slide.background = { color: C.white };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: W,
    h: H,
    fill: { color: C.white },
    line: { color: C.white, transparency: 100 },
  });
}

function title(slide, text, subtitle = "") {
  slide.addText(text, {
    x: margin,
    y: 0.32,
    w: 8.2,
    h: 0.45,
    fontFace: "Microsoft YaHei",
    fontSize: 24,
    bold: true,
    color: C.navy,
    margin: 0,
    breakLine: false,
  });
  slide.addShape(pptx.ShapeType.line, {
    x: margin,
    y: 0.93,
    w: 1.15,
    h: 0,
    line: { color: C.teal, width: 3 },
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 10.2,
      y: 0.38,
      w: 2.4,
      h: 0.28,
      fontSize: 8.5,
      color: C.slate,
      align: "right",
      margin: 0,
    });
  }
}

function footer(slide, idx) {
  slide.addText(`模式识别课程实验汇报 · ${idx}`, {
    x: margin,
    y: 7.12,
    w: 4,
    h: 0.18,
    fontSize: 7.5,
    color: "94A3B8",
    margin: 0,
  });
}

function bullet(slide, lines, x, y, w, h, opts = {}) {
  slide.addText(
    lines.map((t) => ({ text: t, options: { bullet: { type: "ul" } } })),
    {
      x,
      y,
      w,
      h,
      fontFace: "Microsoft YaHei",
      fontSize: opts.fontSize ?? 14,
      color: opts.color ?? C.navy,
      fit: "shrink",
      breakLine: false,
      paraSpaceAfterPt: 7,
      margin: 0.08,
    },
  );
}

function card(slide, x, y, w, h, head, body, accent = C.teal) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: C.light },
    line: { color: C.line, width: 1 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w: 0.08,
    h,
    fill: { color: accent },
    line: { color: accent, transparency: 100 },
  });
  slide.addText(head, {
    x: x + 0.18,
    y: y + 0.18,
    w: w - 0.32,
    h: 0.28,
    fontSize: 13,
    bold: true,
    color: C.navy,
    margin: 0,
  });
  slide.addText(body, {
    x: x + 0.18,
    y: y + 0.58,
    w: w - 0.32,
    h: h - 0.7,
    fontSize: 10.2,
    color: C.slate,
    fit: "shrink",
    breakLine: false,
    margin: 0,
  });
}

function image(slide, file, x, y, w, h) {
  if (!fs.existsSync(file)) throw new Error(`Missing image: ${file}`);
  slide.addImage({ path: file, x, y, w, h, sizingCrop: true });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w,
    h,
    fill: { transparency: 100 },
    line: { color: "CBD5E1", width: 0.8 },
  });
}

function table(slide, rows, x, y, w, h, colW, fontSize = 9.5) {
  slide.addTable(rows, {
    x,
    y,
    w,
    h,
    colW,
    border: { type: "solid", color: "CBD5E1", pt: 0.6 },
    margin: 0.05,
    color: C.navy,
    fontFace: "Microsoft YaHei",
    fontSize,
    valign: "mid",
    fit: "shrink",
    autoFit: false,
    fill: "FFFFFF",
    rowH: Array(rows.length).fill(h / rows.length),
  });
}

function section(name, sub, num) {
  const s = pptx.addSlide();
  addBg(s);
  s.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: W,
    h: H,
    fill: { color: C.light },
    line: { color: C.light },
  });
  s.addText(num, {
    x: 0.72,
    y: 1.1,
    w: 1.4,
    h: 0.6,
    fontSize: 21,
    bold: true,
    color: C.teal,
    margin: 0,
  });
  s.addText(name, {
    x: 0.72,
    y: 2.0,
    w: 10.5,
    h: 0.8,
    fontSize: 34,
    bold: true,
    color: C.navy,
    margin: 0,
  });
  s.addText(sub, {
    x: 0.76,
    y: 3.03,
    w: 9.5,
    h: 0.45,
    fontSize: 16,
    color: C.slate,
    margin: 0,
  });
  s.addShape(pptx.ShapeType.line, {
    x: 0.76,
    y: 4.0,
    w: 3.0,
    h: 0,
    line: { color: C.teal, width: 4 },
  });
}

let n = 1;
function addSlide(titleText, subtitle = "") {
  const s = pptx.addSlide();
  addBg(s);
  title(s, titleText, subtitle);
  footer(s, n++);
  return s;
}

{
  const s = pptx.addSlide();
  addBg(s);
  s.addText("模式识别课程实验汇报", {
    x: 0.72,
    y: 2.45,
    w: 8.4,
    h: 0.75,
    fontSize: 38,
    bold: true,
    color: C.navy,
    margin: 0,
  });
  s.addText("Bayes 分类器、线性判别、K-L 变换、图像识别与聚类分析", {
    x: 0.75,
    y: 3.35,
    w: 7.8,
    h: 0.32,
    fontSize: 16,
    color: C.slate,
    margin: 0,
  });
  card(s, 0.75, 4.55, 3.4, 0.72, "姓名", "请填写");
  card(s, 4.95, 4.55, 3.4, 0.72, "学号", "请填写");
  card(s, 9.15, 4.55, 3.4, 0.72, "班级", "请填写");
}

{
  const s = addSlide("汇报大纲", "五个实验的统一讲述结构");
  const tasks = [
    ["实验一", "Bayes 性别分类器"],
    ["实验二", "非参数估计与 Fisher"],
    ["实验三", "K-L 特征提取"],
    ["实验四", "K-L 图像识别"],
    ["实验五", "聚类分析"],
  ];
  tasks.forEach((t, i) => card(s, 0.65 + i * 2.48, 1.55, 2.1, 1.15, t[0], t[1], i % 2 ? C.cyan : C.teal));
  card(s, 1.2, 3.5, 2.2, 1.1, "1 理论知识", "结合课程中的概率决策、线性判别、降维和聚类理论");
  card(s, 3.9, 3.5, 2.2, 1.1, "2 实验设置", "说明数据、特征、模型参数和评价指标");
  card(s, 6.6, 3.5, 2.2, 1.1, "3 实验结果", "展示准确率、错误率、混淆矩阵和可视化结果");
  card(s, 9.3, 3.5, 2.2, 1.1, "4 分析总结", "解释现象并联系模式识别课程知识点");
}

{
  const s = addSlide("数据与实验环境");
  table(s, [
    ["数据集", "用途", "状态"],
    ["FEMALE / MALE", "性别分类、K-L 特征提取、聚类", "已使用"],
    ["test1 / test2", "分类器泛化测试", "已使用"],
    ["飞机图像数据集", "K-L 图像识别应用", "已使用"],
  ], 0.8, 1.4, 5.8, 2.25, [1.6, 3.0, 1.2], 9);
  table(s, [
    ["工具", "作用"],
    ["Python 3.11", "实验主环境"],
    ["NumPy / SciPy", "数值计算"],
    ["scikit-learn", "评估、PCA、分级聚类辅助"],
    ["Matplotlib", "实验图表输出"],
  ], 7.1, 1.4, 5.2, 2.65, [1.7, 3.5], 9);
  card(s, 1.0, 5.0, 11.1, 0.95, "说明", "核心实验均已脚本化执行，输出包括指标 CSV、实验图、LaTeX 报告以及本次汇报 PPT。", C.teal);
}

section("实验一", "用身高/体重数据建立 Bayes 性别分类器", "01");

{
  const s = addSlide("实验一：理论知识");
  card(s, 0.75, 1.25, 5.7, 1.1, "最小错误率 Bayes 决策", "选择后验概率最大的类别：argmax p(x|ωi)P(ωi)。", C.teal);
  card(s, 0.75, 2.65, 5.7, 1.25, "正态分布参数估计", "使用训练样本最大似然估计类别均值和协方差，建立高斯判别函数。", C.cyan);
  card(s, 0.75, 4.22, 5.7, 1.25, "最小风险 Bayes 决策", "引入损失函数 λ(αi|ωj)，不再只统计错误个数，而是最小化总体风险。", C.orange);
  s.addText("判别函数", { x: 7.0, y: 1.3, w: 4.5, h: 0.3, fontSize: 16, bold: true, color: C.navy, margin: 0 });
  s.addText("g_i(x) = -1/2 ln|Σ_i| - 1/2(x-μ_i)^TΣ_i^{-1}(x-μ_i) + ln P(ω_i)", {
    x: 7.0,
    y: 1.9,
    w: 4.95,
    h: 1.3,
    fontSize: 18,
    color: C.tealDark,
    fit: "shrink",
    breakLine: false,
    margin: 0.05,
  });
  bullet(s, ["参数估计对应课程中的概率密度估计", "先验概率体现类别分布假设", "风险决策体现不同错误类型的代价差异"], 7.05, 3.75, 4.8, 1.8);
}

{
  const s = addSlide("实验一：实验设置与实现");
  bullet(s, [
    "训练集：50 个女生、50 个男生",
    "测试集：test1 共 35 个样本，test2 共 300 个样本",
    "特征：身高、体重、身高+体重",
    "先验概率：0.5/0.5、0.75/0.25、0.9/0.1",
    "协方差假设：完整协方差与对角协方差",
  ], 0.85, 1.25, 5.0, 4.6);
  card(s, 6.6, 1.25, 5.2, 0.85, "步骤 1", "读取训练集和测试集");
  card(s, 6.6, 2.3, 5.2, 0.85, "步骤 2", "估计均值、方差或协方差矩阵");
  card(s, 6.6, 3.35, 5.2, 0.85, "步骤 3", "计算 Bayes 判别函数和风险函数");
  card(s, 6.6, 4.4, 5.2, 0.85, "步骤 4", "输出准确率、召回率和混淆矩阵");
}

{
  const s = addSlide("实验一：数据分布与决策边界");
  image(s, fig("experiment1", "data_scatter.png"), 0.75, 1.25, 5.9, 4.7);
  image(s, fig("experiment1", "bayes_2d_boundary.png"), 6.85, 1.25, 5.75, 4.7);
  s.addText("男生样本整体身高、体重更高，但边界区域存在重叠，因此分类器无法完全无误。", {
    x: 0.85,
    y: 6.15,
    w: 10.5,
    h: 0.34,
    fontSize: 13,
    color: C.slate,
    margin: 0,
  });
}

{
  const s = addSlide("实验一：结果与分析");
  table(s, [
    ["方法", "训练集", "test1", "test2"],
    ["二维 Bayes，完整协方差", "88.00%", "97.14%", "89.33%"],
    ["二维 Bayes，对角协方差", "88.00%", "97.14%", "90.33%"],
    ["最小风险 Bayes", "84.00%", "94.29%", "84.33%"],
    ["Parzen Bayes", "87.00%", "97.14%", "90.33%"],
    ["kNN，k=5", "89.00%", "100.00%", "90.33%"],
  ], 0.75, 1.25, 6.8, 2.45, [3.2, 1.2, 1.2, 1.2], 8.4);
  image(s, fig("experiment1", "bayes_test2_confusion.png"), 8.2, 1.4, 3.8, 3.15);
  bullet(s, [
    "两个特征整体优于单个特征",
    "先验概率偏离测试集分布时性能下降明显",
    "非参数方法在 test2 上与参数 Bayes 表现接近",
  ], 0.9, 4.45, 6.2, 1.5);
  card(s, 7.5, 5.05, 4.7, 0.88, "课程联系", "Bayes 决策把概率密度估计、先验知识和损失函数统一到同一判别框架中。", C.teal);
}

section("实验二", "非参数估计、Fisher 线性判别与留一法", "02");

{
  const s = addSlide("实验二：理论知识");
  card(s, 0.75, 1.25, 3.7, 1.45, "Parzen 窗估计", "不假定概率密度形式，用训练样本附近的高斯核叠加估计 p(x|ω)。");
  card(s, 4.85, 1.25, 3.7, 1.45, "Fisher 线性判别", "寻找投影方向，使类间离散度相对类内离散度最大。", C.cyan);
  card(s, 8.95, 1.25, 3.2, 1.45, "留一法", "每次留出一个训练样本验证，用于估计泛化误差。", C.orange);
  s.addText("Fisher 目标函数", { x: 1.0, y: 3.35, w: 3.0, h: 0.3, fontSize: 16, bold: true, color: C.navy, margin: 0 });
  s.addText("J(w)=w^T S_b w / w^T S_w w,    w=S_w^{-1}(μ_1-μ_2)", {
    x: 1.0,
    y: 3.95,
    w: 10.8,
    h: 0.55,
    fontSize: 22,
    color: C.tealDark,
    fit: "shrink",
    margin: 0,
  });
  bullet(s, ["非参数估计强调样本局部结构", "Fisher 强调直接设计判别面", "留一法连接训练误差与泛化性能"], 1.0, 5.0, 10.5, 1.1);
}

{
  const s = addSlide("实验二：决策边界与投影");
  image(s, fig("experiment2", "fisher_bayes_boundaries.png"), 0.75, 1.25, 5.9, 4.65);
  image(s, fig("experiment2", "fisher_projection.png"), 6.85, 1.25, 5.75, 4.65);
  s.addText("Bayes 边界可为非线性；Fisher 直接给出线性投影和阈值，模型更简单、解释性更强。", {
    x: 0.85,
    y: 6.12,
    w: 10.8,
    h: 0.32,
    fontSize: 12.5,
    color: C.slate,
    margin: 0,
  });
}

{
  const s = addSlide("实验二：结果与分析");
  table(s, [
    ["方法", "训练错误率", "test1 错误率", "test2 错误率", "留一法错误率"],
    ["Gaussian Bayes", "12.00%", "2.86%", "10.67%", "12.00%"],
    ["Parzen Bayes", "13.00%", "2.86%", "9.67%", "14.00%"],
    ["Fisher", "12.00%", "2.86%", "9.67%", "13.00%"],
  ], 0.85, 1.35, 11.4, 1.85, [2.4, 2.0, 2.2, 2.2, 2.2], 9);
  card(s, 0.95, 3.75, 3.4, 1.3, "Parzen", "不依赖正态分布假设，但对窗口宽度和样本规模敏感。");
  card(s, 4.95, 3.75, 3.4, 1.3, "Fisher", "不估计概率密度，直接优化线性可分性。", C.cyan);
  card(s, 8.95, 3.75, 3.4, 1.3, "留一法", "结果与 test2 更接近，能较稳健估计泛化误差。", C.orange);
}

section("实验三", "K-L 变换进行特征提取", "03");

{
  const s = addSlide("实验三：理论知识");
  card(s, 0.8, 1.25, 4.8, 1.35, "K-L / PCA", "对协方差矩阵特征分解，寻找最大方差方向，是无监督特征提取方法。");
  card(s, 6.0, 1.25, 4.8, 1.35, "Fisher 对比", "Fisher 使用类别标签，目标是增强类别可分性。", C.cyan);
  s.addText("Σu_i = λ_i u_i,     z_i = u_i^T(x-μ)", {
    x: 1.05,
    y: 3.35,
    w: 10.8,
    h: 0.55,
    fontSize: 25,
    color: C.tealDark,
    margin: 0,
  });
  bullet(s, ["PCA 保留总体方差最大的信息", "最大方差方向不一定等于最佳分类方向", "本数据中二者比较接近，因此 PCA 分类效果较好"], 1.0, 4.55, 10.5, 1.25);
}

{
  const s = addSlide("实验三：实验设置与投影");
  table(s, [
    ["项目", "数值/说明"],
    ["数据", "FEMALE + MALE，共 100 个训练样本"],
    ["特征", "身高、体重"],
    ["第一主成分解释率", "87.52%"],
    ["PC1", "(0.6269, 0.7791)"],
    ["类均值方向", "(0.6514, 0.7587)"],
  ], 0.75, 1.25, 4.8, 2.75, [1.8, 3.0], 8.7);
  image(s, fig("experiment3", "kl_directions.png"), 6.0, 1.15, 5.95, 4.3);
  s.addText("主成分方向与类均值差方向接近，说明身高和体重的主要变化也与性别差异相关。", {
    x: 6.1,
    y: 5.7,
    w: 5.5,
    h: 0.42,
    fontSize: 12,
    color: C.slate,
    margin: 0,
  });
}

{
  const s = addSlide("实验三：结果与分析");
  image(s, fig("experiment3", "pca_projection.png"), 0.75, 1.25, 5.6, 3.65);
  table(s, [
    ["方法", "训练集", "test1", "test2"],
    ["K-L PCA 第一主成分", "86.00%", "100.00%", "89.67%"],
    ["类均值方向", "86.00%", "100.00%", "89.67%"],
    ["Fisher", "88.00%", "97.14%", "90.33%"],
    ["Gaussian Bayes", "88.00%", "97.14%", "89.33%"],
  ], 6.7, 1.35, 5.45, 2.45, [2.4, 1.0, 1.0, 1.0], 8.5);
  bullet(s, ["K-L 适合降维和特征提取", "Fisher 使用标签，在 test2 上略优", "当主方差方向与类别差异一致时，PCA 也有较好分类效果"], 6.85, 4.35, 5.2, 1.45);
}

section("实验四", "K-L 变换在飞机图像识别中的应用", "04");

{
  const s = addSlide("实验四：理论知识与设置");
  card(s, 0.75, 1.25, 5.2, 1.15, "图像向量化", "灰度图缩放为 32×32，再展开为 1024 维向量。");
  card(s, 0.75, 2.75, 5.2, 1.15, "K-L 图像特征", "学习训练图像的主要变化方向，类似“特征脸”思想。", C.cyan);
  card(s, 0.75, 4.25, 5.2, 1.15, "低维识别", "将图像投影到 PCA 空间后，用最近质心和 1NN 分类。", C.orange);
  table(s, [
    ["项目", "设置"],
    ["类别数", "16 类"],
    ["训练集", "2530 张"],
    ["测试集", "636 张"],
    ["主成分数", "8、16、32、64、128"],
    ["64 维解释率", "89.88%"],
  ], 6.75, 1.4, 4.9, 2.95, [1.7, 3.2], 9);
}

{
  const s = addSlide("实验四：样本与主成分");
  image(s, fig("experiment4", "sample_images.png"), 0.75, 1.25, 5.9, 4.65);
  image(s, fig("experiment4", "pca_components.png"), 6.85, 1.25, 5.75, 4.65);
  s.addText("PCA 主成分提取到外形、亮度和纹理等主要变化，但其中也会包含背景和姿态信息。", {
    x: 0.85,
    y: 6.12,
    w: 10.7,
    h: 0.32,
    fontSize: 12,
    color: C.slate,
    margin: 0,
  });
}

{
  const s = addSlide("实验四：结果与分析");
  table(s, [
    ["方法", "主成分数", "训练集", "测试集"],
    ["PCA+NearestCentroid", "64", "20.47%", "19.97%"],
    ["PCA+1NN", "32", "100.00%", "54.09%"],
    ["PCA+1NN", "64", "100.00%", "53.93%"],
    ["PCA+1NN", "128", "100.00%", "52.99%"],
  ], 0.8, 1.25, 5.65, 2.25, [2.45, 1.1, 1.0, 1.1], 8.7);
  image(s, fig("experiment4", "accuracy_vs_components.png"), 6.85, 1.25, 5.55, 3.55);
  bullet(s, ["1NN 明显优于最近质心，说明类内变化复杂", "主成分过多会保留背景和噪声", "PCA 可有效压缩特征，但单独使用不足以高精度图像分类"], 0.9, 4.15, 5.3, 1.55);
  image(s, fig("experiment4", "confusion_matrix.png"), 7.55, 4.85, 3.8, 1.45);
}

section("实验五", "C 均值与分级聚类分析", "05");

{
  const s = addSlide("实验五：理论知识与设置");
  card(s, 0.75, 1.25, 5.3, 1.35, "C 均值聚类", "即 k-means，通过样本分配和中心更新迭代最小化类内平方误差。");
  card(s, 0.75, 2.9, 5.3, 1.35, "分级聚类", "逐步合并或分裂样本簇，形成层次结构，可用树状图观察。", C.cyan);
  card(s, 0.75, 4.55, 5.3, 1.35, "课程联系", "从监督分类过渡到无监督学习，关注样本自然结构。", C.orange);
  table(s, [
    ["项目", "设置"],
    ["数据一", "FEMALE + MALE，100 个样本"],
    ["数据二", "训练集 + test2，400 个样本"],
    ["特征", "身高、体重"],
    ["预处理", "标准化"],
    ["类别数", "2、3、4、5"],
  ], 6.8, 1.4, 4.9, 2.95, [1.6, 3.3], 9);
}

{
  const s = addSlide("实验五：C 均值聚类结果");
  image(s, fig("experiment5", "cmeans_train_2clusters.png"), 0.75, 1.25, 5.75, 4.4);
  image(s, fig("experiment5", "init_variation.png"), 6.85, 1.25, 5.55, 4.4);
  s.addText("训练集取 C=2 时，与真实性别标签的最佳匹配率约为 85.00%；不同初始值下约在 85.00% 到 87.00% 之间波动。", {
    x: 0.85,
    y: 5.95,
    w: 11.0,
    h: 0.4,
    fontSize: 12,
    color: C.slate,
    fit: "shrink",
    margin: 0,
  });
}

{
  const s = addSlide("实验五：类别数选择与分级聚类");
  image(s, fig("experiment5", "cmeans_k_metrics_train.png"), 0.75, 1.25, 5.75, 4.35);
  image(s, fig("experiment5", "hierarchical_dendrogram_train.png"), 6.85, 1.25, 5.55, 4.35);
  card(s, 0.95, 5.85, 5.2, 0.55, "C=2 更合理", "轮廓系数最高，继续增加类别数虽降低 SSE，但解释性下降。");
  card(s, 6.95, 5.85, 5.2, 0.55, "Ward 聚类", "不依赖随机初始中心，可观察层次关系。", C.cyan);
}

{
  const s = addSlide("实验五：加入 test2 后的变化");
  image(s, fig("experiment5", "cmeans_combined_2clusters.png"), 0.75, 1.25, 5.75, 4.35);
  image(s, fig("experiment5", "cmeans_k_metrics_combined.png"), 6.85, 1.25, 5.55, 4.35);
  table(s, [
    ["数据", "方法", "轮廓系数", "ARI", "性别匹配率"],
    ["训练集", "C 均值，2 类", "0.5108", "0.4850", "85.00%"],
    ["训练集", "Ward，2 类", "0.5049", "0.3789", "81.00%"],
    ["训练集+test2", "C 均值，2 类", "0.4850", "0.4693", "84.50%"],
    ["训练集+test2", "Ward，2 类", "0.4739", "0.4172", "82.50%"],
  ], 0.85, 5.85, 11.35, 0.85, [2.0, 2.8, 1.7, 1.5, 1.8], 7.2);
}

{
  const s = addSlide("五个实验的横向关系");
  table(s, [
    ["实验", "课程知识点", "监督信息", "主要目标"],
    ["Bayes 分类", "后验概率、参数估计、风险决策", "使用标签", "最小错误率或最小风险"],
    ["Fisher / Parzen", "非参数估计、线性判别、留一法", "使用标签", "比较概率模型与直接判别"],
    ["K-L 特征提取", "PCA、协方差特征分解", "可不使用标签", "降维与主方向分析"],
    ["K-L 图像应用", "高维图像特征压缩", "使用标签", "验证 PCA 图像识别作用"],
    ["聚类分析", "C 均值、分级聚类", "不使用标签", "发现样本自然结构"],
  ], 0.75, 1.25, 11.75, 4.15, [1.7, 4.0, 2.0, 4.05], 8.2);
  card(s, 1.25, 5.9, 10.6, 0.6, "主线", "从概率决策到线性判别，从监督分类到无监督聚类，从低维特征到高维图像识别。", C.teal);
}

{
  const s = addSlide("总体结论");
  bullet(s, [
    "特征选择会直接影响分类器性能",
    "先验概率和损失函数体现任务背景",
    "非参数方法灵活，但依赖样本规模和超参数",
    "K-L 适合降维，不一定最适合分类",
    "聚类能揭示结构，但解释需要结合真实标签",
  ], 0.95, 1.35, 5.4, 3.5, { fontSize: 13.5 });
  card(s, 6.85, 1.45, 4.9, 1.1, "核心体会", "模式识别方法没有绝对最优，算法选择必须结合数据分布、任务目标和评价指标。", C.teal);
  card(s, 6.85, 3.05, 4.9, 1.1, "汇报落点", "五个实验分别覆盖分类、估计、降维、图像识别和聚类，形成完整的方法链条。", C.cyan);
}

{
  const s = pptx.addSlide();
  addBg(s);
  s.addText("谢谢", {
    x: 0.8,
    y: 2.65,
    w: 5.0,
    h: 0.8,
    fontSize: 44,
    bold: true,
    color: C.navy,
    margin: 0,
  });
  s.addText("欢迎老师和同学批评指正", {
    x: 0.85,
    y: 3.65,
    w: 6.0,
    h: 0.38,
    fontSize: 18,
    color: C.slate,
    margin: 0,
  });
  s.addShape(pptx.ShapeType.line, {
    x: 0.85,
    y: 4.35,
    w: 2.8,
    h: 0,
    line: { color: C.teal, width: 4 },
  });
}

await pptx.writeFile({ fileName: outPath });
console.log(outPath);
