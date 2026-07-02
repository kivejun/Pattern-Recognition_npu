---
theme: default
title: 模式识别课程实验汇报
info: |
  基于男女身高体重数据、飞机图像数据与聚类分析的模式识别课程实验汇报。
class: text-slate-800
highlighter: shiki
lineNumbers: false
drawings:
  persist: false
transition: slide-left
mdc: true
canvasWidth: 1280
aspectRatio: 16/9
---

# 模式识别课程实验汇报

Bayes 分类器、线性判别、K-L 变换、图像识别与聚类分析

<div class="mt-12 grid grid-cols-3 gap-4 text-sm">
  <div class="info-box">姓名：请填写</div>
  <div class="info-box">学号：请填写</div>
  <div class="info-box">班级：请填写</div>
</div>

---

# 汇报思路

<div class="grid grid-cols-5 gap-4 mt-8">
  <div class="task-card"><b>实验一</b><br/>Bayes 分类器</div>
  <div class="task-card"><b>实验二</b><br/>非参数估计与 Fisher</div>
  <div class="task-card"><b>实验三</b><br/>K-L 特征提取</div>
  <div class="task-card"><b>实验四</b><br/>K-L 图像识别应用</div>
  <div class="task-card"><b>实验五</b><br/>聚类分析</div>
</div>

<div class="mt-10 text-xl">
每个实验按照同一条主线展开：
</div>

```mermaid
flowchart LR
  A[课程理论] --> B[实验设置]
  B --> C[实现方法]
  C --> D[实验结果]
  D --> E[分析总结]
```

---

# 数据与实验环境

<div class="grid grid-cols-2 gap-8 mt-4">
<div>

## 主要数据

| 数据集 | 用途 |
| --- | --- |
| FEMALE / MALE | 性别分类、特征提取、聚类 |
| test1 / test2 | 分类器泛化测试 |
| 飞机图像数据集 | K-L 图像识别应用 |

</div>
<div>

## 实验工具

| 工具 | 作用 |
| --- | --- |
| Python 3.11 | 实验主环境 |
| NumPy / SciPy | 数值计算 |
| scikit-learn | 评估、PCA、层次聚类辅助 |
| Matplotlib | 结果可视化 |

</div>
</div>

<div class="note mt-8">
核心实验代码均已脚本化执行，输出包括指标 CSV、图像结果和 LaTeX 报告。
</div>

---
layout: section
---

# 实验一

用身高/体重数据建立 Bayes 性别分类器

---

# 实验一：课程理论

Bayes 分类器根据后验概率进行决策：

$$
\hat{\omega}=\arg\max_{\omega_i} P(\omega_i|\mathbf{x})
=\arg\max_{\omega_i} p(\mathbf{x}|\omega_i)P(\omega_i)
$$

在正态分布假设下：

$$
g_i(\mathbf{x})=-\frac{1}{2}\ln|\Sigma_i|
-\frac{1}{2}(\mathbf{x}-\mu_i)^T\Sigma_i^{-1}(\mathbf{x}-\mu_i)
+\ln P(\omega_i)
$$

<div class="grid grid-cols-3 gap-4 mt-6">
  <div class="concept"><b>参数估计</b><br/>最大似然估计均值和协方差</div>
  <div class="concept"><b>先验概率</b><br/>影响决策边界偏移</div>
  <div class="concept"><b>最小风险</b><br/>用损失函数代替单纯错误率</div>
</div>

---

# 实验一：实验设置与实现

<div class="grid grid-cols-2 gap-8">
<div>

## 设置

- 训练集：50 个女生、50 个男生
- 测试集：`test1` 35 个样本，`test2` 300 个样本
- 特征：身高、体重、身高+体重
- 先验：0.5/0.5、0.75/0.25、0.9/0.1
- 协方差：完整协方差、对角协方差

</div>
<div>

## 实现流程

```mermaid
flowchart TD
  A[读取训练/测试数据] --> B[估计均值与协方差]
  B --> C[计算判别函数]
  C --> D[输出类别预测]
  D --> E[计算准确率与混淆矩阵]
```

</div>
</div>

---

# 实验一：数据分布与决策边界

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment1/data_scatter.png" class="figure" />
  <img src="./figures/experiment1/bayes_2d_boundary.png" class="figure" />
</div>

<div class="note mt-4">
男生样本整体身高、体重更高，但边界区域存在重叠，因此无法做到完全无误分类。
</div>

---

# 实验一：结果与分析

| 方法 | 训练集 | test1 | test2 |
| --- | ---: | ---: | ---: |
| 二维 Bayes，完整协方差，等先验 | 88.00% | 97.14% | 89.33% |
| 二维 Bayes，对角协方差，等先验 | 88.00% | 97.14% | 90.33% |
| 最小风险 Bayes | 84.00% | 94.29% | 84.33% |
| Parzen Bayes | 87.00% | 97.14% | 90.33% |
| kNN，k=5 | 89.00% | 100.00% | 90.33% |

<div class="grid grid-cols-2 gap-6 mt-4">
<div>

## 现象

- 两个特征整体优于单个特征
- 先验概率偏离测试集分布时，性能明显下降
- 非参数方法与参数 Bayes 表现接近

</div>
<img src="./figures/experiment1/bayes_test2_confusion.png" class="figure-small" />
</div>

---
layout: section
---

# 实验二

非参数估计、Fisher 线性判别与留一法

---

# 实验二：课程理论

Parzen 窗估计不预设密度形式：

$$
\hat{p}(\mathbf{x}|\omega_i)=
\frac{1}{N_i}\sum_{k=1}^{N_i}
\frac{1}{(2\pi)^{d/2}h^d}
\exp\left(-\frac{\|\mathbf{x}-\mathbf{x}_k\|^2}{2h^2}\right)
$$

Fisher 线性判别寻找投影方向，使类间离散度相对类内离散度最大：

$$
J(\mathbf{w})=\frac{\mathbf{w}^TS_b\mathbf{w}}{\mathbf{w}^TS_w\mathbf{w}},
\quad
\mathbf{w}=S_w^{-1}(\mu_1-\mu_2)
$$

留一法用于在样本较少时估计泛化误差。

---

# 实验二：实验设置与实现

<div class="grid grid-cols-3 gap-4 mt-6">
  <div class="task-card"><b>Parzen Bayes</b><br/>高斯核，窗口宽度 h=5.0</div>
  <div class="task-card"><b>Fisher 判别</b><br/>同时使用身高和体重</div>
  <div class="task-card"><b>留一法</b><br/>每次留出一个训练样本验证</div>
</div>

<div class="mt-8">

```mermaid
flowchart LR
  A[训练样本] --> B[Parzen 密度估计]
  A --> C[Fisher 投影方向]
  B --> D[Bayes 分类]
  C --> E[线性分类]
  D --> F[测试集评估]
  E --> F
  F --> G[与留一法误差比较]
```

</div>

---

# 实验二：决策边界与投影

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment2/fisher_bayes_boundaries.png" class="figure" />
  <img src="./figures/experiment2/fisher_projection.png" class="figure" />
</div>

<div class="note mt-4">
Bayes 边界可以是非线性的；Fisher 直接给出线性投影与阈值，解释性更强。
</div>

---

# 实验二：结果与分析

| 方法 | 训练错误率 | test1 错误率 | test2 错误率 | 留一法错误率 |
| --- | ---: | ---: | ---: | ---: |
| Gaussian Bayes | 12.00% | 2.86% | 10.67% | 12.00% |
| Parzen Bayes | 13.00% | 2.86% | 9.67% | 14.00% |
| Fisher | 12.00% | 2.86% | 9.67% | 13.00% |

## 分析

- Parzen 不依赖正态分布假设，但对窗口宽度敏感
- Fisher 不估计概率密度，直接优化线性可分性
- 留一法结果与 test2 更接近，说明其能较稳健地估计泛化误差

---
layout: section
---

# 实验三

K-L 变换进行特征提取

---

# 实验三：课程理论

K-L 变换即 PCA，从协方差矩阵的特征分解中寻找最大方差方向：

$$
\Sigma \mathbf{u}_i=\lambda_i \mathbf{u}_i
$$

投影形式为：

$$
z_i=\mathbf{u}_i^T(\mathbf{x}-\mu)
$$

<div class="grid grid-cols-2 gap-6 mt-6">
<div class="concept">
<b>K-L / PCA</b><br/>
无监督，目标是保留数据总体方差。
</div>
<div class="concept">
<b>Fisher</b><br/>
有监督，目标是增强类别可分性。
</div>
</div>

---

# 实验三：实验设置与实现

<div class="grid grid-cols-2 gap-8">
<div>

## 设置

- 数据：FEMALE + MALE 共 100 个训练样本
- 特征：身高、体重
- 方法一：不考虑类别信息，做 PCA
- 方法二：利用类均值差方向投影
- 对比：Fisher 与 Gaussian Bayes

</div>
<div>

## 关键参数

| 项目 | 数值 |
| --- | --- |
| 第一主成分解释率 | 87.52% |
| PC1 | (0.6269, 0.7791) |
| 类均值方向 | (0.6514, 0.7587) |

</div>
</div>

---

# 实验三：投影方向与样本分布

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment3/kl_directions.png" class="figure" />
  <img src="./figures/experiment3/pca_projection.png" class="figure" />
</div>

<div class="note mt-4">
本数据中最大方差方向与性别差异方向比较接近，因此 PCA 一维投影也能取得较好分类效果。
</div>

---

# 实验三：结果与分析

| 方法 | 训练集 | test1 | test2 |
| --- | ---: | ---: | ---: |
| K-L PCA 第一主成分 | 86.00% | 100.00% | 89.67% |
| 类均值方向 | 86.00% | 100.00% | 89.67% |
| Fisher | 88.00% | 97.14% | 90.33% |
| Gaussian Bayes | 88.00% | 97.14% | 89.33% |

## 分析

- K-L 适合降维和特征提取，但不是专门为分类优化
- Fisher 使用类别标签，因此在 test2 上略优
- 当主方差方向与类别差异方向一致时，PCA 也可以获得较好分类效果

---
layout: section
---

# 实验四

K-L 变换在飞机图像识别中的应用

---

# 实验四：课程理论

图像可以看作高维向量。K-L 变换在图像识别中的典型思想是：

```mermaid
flowchart LR
  A[图像矩阵] --> B[灰度化/缩放]
  B --> C[展开为高维向量]
  C --> D[PCA 提取低维特征]
  D --> E[分类器识别]
```

与“特征脸”类似，本实验学习的是飞机图像的主要变化方向。

$$
\mathbf{z}=U_k^T(\mathbf{x}-\mu)
$$

---

# 实验四：实验设置与实现

<div class="grid grid-cols-2 gap-8">
<div>

## 数据与预处理

- 数据集：飞机分类图像
- 类别数：16 类
- 训练集：2530 张
- 测试集：636 张
- 预处理：灰度化、缩放到 32×32、展开为 1024 维向量

</div>
<div>

## 模型设置

- 标准化后进行 PCA
- 主成分数：8、16、32、64、128
- 分类器：Nearest Centroid 与 1NN
- 主结果：64 维 PCA + 1NN

</div>
</div>

---

# 实验四：样本与主成分

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment4/sample_images.png" class="figure" />
  <img src="./figures/experiment4/pca_components.png" class="figure" />
</div>

<div class="note mt-4">
64 个主成分保留约 89.88% 的训练集方差信息，将 1024 维图像压缩到 64 维。
</div>

---

# 实验四：结果与分析

<div class="grid grid-cols-2 gap-6">
<div>

| 方法 | 主成分数 | 训练集 | 测试集 |
| --- | ---: | ---: | ---: |
| PCA+NearestCentroid | 64 | 20.47% | 19.97% |
| PCA+1NN | 32 | 100.00% | 54.09% |
| PCA+1NN | 64 | 100.00% | 53.93% |
| PCA+1NN | 128 | 100.00% | 52.99% |

## 分析

- 1NN 明显优于最近质心，说明类内变化较复杂
- 主成分过多后，背景和噪声也会被保留
- PCA 是有效特征压缩方法，但单独使用不足以完成高精度图像分类

</div>
<img src="./figures/experiment4/accuracy_vs_components.png" class="figure" />
</div>

---

# 实验四：混淆矩阵观察

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment4/confusion_matrix.png" class="figure" />
  <div>
    <h2>主要现象</h2>
    <ul>
      <li>部分飞机类别外形相近，在低分辨率灰度图中容易混淆</li>
      <li>姿态、背景、拍摄尺度会进入 PCA 主成分</li>
      <li>后续可结合 HOG/SIFT/ORB 或 CNN 提升识别能力</li>
    </ul>
  </div>
</div>

---
layout: section
---

# 实验五

C 均值与分级聚类分析

---

# 实验五：课程理论

C 均值聚类即 k-means，通过迭代最小化类内平方误差：

$$
J=\sum_{j=1}^{C}\sum_{\mathbf{x}_i\in \omega_j}
\|\mathbf{x}_i-\mu_j\|^2
$$

迭代步骤：

```mermaid
flowchart LR
  A[初始化中心] --> B[按距离分配样本]
  B --> C[更新聚类中心]
  C --> D{是否收敛}
  D -- 否 --> B
  D -- 是 --> E[输出聚类结果]
```

分级聚类则通过样本间距离形成层次结构，可用树状图观察聚类关系。

---

# 实验五：实验设置与实现

<div class="grid grid-cols-2 gap-8">
<div>

## 数据设置

- 数据一：FEMALE + MALE，共 100 个样本
- 数据二：训练集 + test2，共 400 个样本
- 特征：身高、体重
- 预处理：标准化

</div>
<div>

## 方法设置

- C 均值：独立编程实现
- 类别数：2、3、4、5
- 初始值：多随机种子比较
- 分级聚类：Ward 方法

</div>
</div>

---

# 实验五：C 均值聚类结果

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment5/cmeans_train_2clusters.png" class="figure" />
  <img src="./figures/experiment5/init_variation.png" class="figure" />
</div>

<div class="note mt-4">
训练集取 C=2 时，与真实性别标签的最佳匹配率约为 85.00%；不同初始值下约在 85.00% 到 87.00% 之间波动。
</div>

---

# 实验五：类别数选择与分级聚类

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment5/cmeans_k_metrics_train.png" class="figure" />
  <img src="./figures/experiment5/hierarchical_dendrogram_train.png" class="figure" />
</div>

<div class="grid grid-cols-2 gap-6 mt-4">
<div>
<b>C 均值，训练集 C=2</b><br/>
轮廓系数 0.5108，ARI 0.4850，性别匹配率 85.00%
</div>
<div>
<b>Ward 分级聚类，训练集 2 类</b><br/>
轮廓系数 0.5049，ARI 0.3789，性别匹配率 81.00%
</div>
</div>

---

# 实验五：加入 test2 后的变化

<div class="grid grid-cols-2 gap-6">
  <img src="./figures/experiment5/cmeans_combined_2clusters.png" class="figure" />
  <img src="./figures/experiment5/cmeans_k_metrics_combined.png" class="figure" />
</div>

<div class="note mt-4">
加入 test2 后样本变为 400 个，男生比例更高；C=2 的匹配率为 84.50%，整体结构仍然支持两类划分，但边界样本更多。
</div>

---

# 实验五：分析总结

| 数据 | 方法 | 轮廓系数 | ARI | 性别匹配率 |
| --- | --- | ---: | ---: | ---: |
| 训练集 | C 均值，2 类 | 0.5108 | 0.4850 | 85.00% |
| 训练集 | Ward，2 类 | 0.5049 | 0.3789 | 81.00% |
| 训练集+test2 | C 均值，2 类 | 0.4850 | 0.4693 | 84.50% |
| 训练集+test2 | Ward，2 类 | 0.4739 | 0.4172 | 82.50% |

## 体会

- 聚类不使用标签，但能发现与性别相关的自然分组
- C 均值需要指定类别数，对初始化和特征尺度敏感
- 分级聚类更适合观察层次关系，但大样本计算成本更高

---

# 五个实验的横向关系

| 实验 | 课程知识点 | 监督信息 | 主要目标 |
| --- | --- | --- | --- |
| Bayes 分类 | 后验概率、参数估计、风险决策 | 使用标签 | 最小错误率或最小风险 |
| Fisher / Parzen | 非参数估计、线性判别、留一法 | 使用标签 | 比较概率模型与直接判别 |
| K-L 特征提取 | PCA、协方差特征分解 | 可不使用标签 | 降维与主方向分析 |
| K-L 图像应用 | 高维图像特征压缩 | 使用标签做识别 | 验证 PCA 在图像识别中的作用 |
| 聚类分析 | C 均值、分级聚类 | 不使用标签 | 发现样本自然结构 |

---

# 总体结论

<div class="grid grid-cols-2 gap-8">
<div>

## 实验认识

- 特征选择会直接影响分类器性能
- 先验概率和损失函数体现任务背景
- 非参数方法灵活，但依赖样本和超参数
- K-L 适合降维，不一定最适合分类
- 聚类能揭示结构，但解释需要结合真实标签

</div>
<div>

## 课程联系

- 从概率决策到线性判别
- 从监督分类到无监督聚类
- 从低维特征到高维图像识别
- 从训练误差到泛化性能评估

</div>
</div>

<div class="closing mt-10">
核心体会：模式识别方法没有绝对最优，算法选择必须结合数据分布、任务目标和评价指标。
</div>

---
layout: end
---

# 谢谢

欢迎老师和同学批评指正

<style>
.slidev-layout {
  font-size: 28px;
  line-height: 1.55;
}
.slidev-layout h1 {
  font-size: 2.25rem;
  color: #0f172a;
}
.slidev-layout h2 {
  font-size: 1.25rem;
  color: #1f4f66;
  margin-top: 0.6rem;
}
.info-box, .task-card, .concept {
  border: 1px solid #d8e2e8;
  border-radius: 8px;
  background: #f8fafc;
  padding: 1rem;
}
.task-card {
  min-height: 120px;
}
.concept {
  min-height: 95px;
}
.note {
  border-left: 5px solid #0f766e;
  background: #f0fdfa;
  padding: 0.8rem 1rem;
  color: #134e4a;
}
.figure {
  width: 100%;
  max-height: 430px;
  object-fit: contain;
  border: 1px solid #e2e8f0;
  border-radius: 8px;
  background: white;
}
.figure-small {
  width: 100%;
  max-height: 260px;
  object-fit: contain;
  border: 1px solid #e2e8f0;
  border-radius: 8px;
  background: white;
}
.closing {
  font-size: 1.35rem;
  color: #0f766e;
  font-weight: 700;
}
table {
  font-size: 0.86em;
}
</style>
