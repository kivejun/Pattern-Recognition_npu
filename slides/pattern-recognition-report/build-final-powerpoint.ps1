param(
  [string]$Output = "F:\Pattern-Recognition_npu\slides\pattern-recognition-report\pattern-recognition-final-polished-editable.pptx"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$figRoot = Join-Path $root "figures"
$tmpRoot = Join-Path $env:TEMP "codex-presentations\manual-pattern-recognition-final"
$previewDir = Join-Path $tmpRoot "tmp\preview"
$qaDir = Join-Path $tmpRoot "tmp\qa"
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null
New-Item -ItemType Directory -Force -Path $qaDir | Out-Null

Set-Content -Path (Join-Path $tmpRoot "tmp\source-notes.txt") -Encoding UTF8 -Value @"
模式识别课程实验汇报 source notes

Sources:
- reports/main.tex: existing LaTeX experiment report in repository.
- results/experiment1_bayes/metrics.csv and decision_rules.txt: Bayes classifier metrics and parameters.
- results/experiment2_linear/metrics.csv and fisher_params.txt: Parzen/Fisher/LOO metrics.
- results/experiment3_kl/metrics.csv and kl_params.txt: K-L/PCA metrics and directions.
- results/experiment4_kl_image/metrics.csv and pca_params.txt: image PCA metrics and dataset information.
- results/experiment5_clustering/metrics.csv and cluster_notes.txt: clustering metrics.
- slides/pattern-recognition-report/figures/**: rendered experiment figures used as independent image objects.

No external logos, online images, or unverified identity assets were used.
"@

Set-Content -Path (Join-Path $tmpRoot "tmp\slide-plan.txt") -Encoding UTF8 -Value @"
Create-mode plan for final polished editable PowerPoint

Output:
- pattern-recognition-final-polished-editable.pptx

Design:
- Dark navy section dividers, light content slides, teal/blue/amber/orange accents.
- Microsoft YaHei for Chinese headings and body text.
- Native PowerPoint text boxes, shapes, tables, and independent image objects.
- Each experiment follows: theory, setup, results, analysis.

QA:
- Generate with PowerPoint COM.
- Reopen final PPTX through PowerPoint COM.
- Export all slides to PNG previews.
- Inspect package structure for slide XML and media.
"@

function Rgb([string]$hex) {
  $h = $hex.TrimStart("#")
  $r = [Convert]::ToInt32($h.Substring(0,2), 16)
  $g = [Convert]::ToInt32($h.Substring(2,2), 16)
  $b = [Convert]::ToInt32($h.Substring(4,2), 16)
  return $r + ($g * 256) + ($b * 65536)
}

$C = @{
  Ink = Rgb "0B1220"; Navy = Rgb "111827"; Teal = Rgb "0F766E"; Teal2 = Rgb "14B8A6"
  Blue = Rgb "1D4ED8"; Amber = Rgb "F59E0B"; Orange = Rgb "EA580C"; Green = Rgb "16A34A"
  Slate = Rgb "475569"; Slate2 = Rgb "64748B"; Line = Rgb "CBD5E1"; Pale = Rgb "F8FAFC"
  PaleTeal = Rgb "ECFDF5"; PaleAmber = Rgb "FFFBEB"; White = Rgb "FFFFFF"; Black = Rgb "000000"
}

$msoFalse = 0
$msoTrue = -1
$ppLayoutBlank = 12
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeOval = 9
$msoShapeLine = 9
$ppSaveAsOpenXMLPresentation = 24

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$pres = $app.Presentations.Add($msoTrue)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

$script:Page = 1

function Add-Slide($dark = $false) {
  $slide = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)
  $bg = $slide.Shapes.AddShape($msoShapeRectangle, 0, 0, 960, 540)
  $bg.Fill.ForeColor.RGB = $(if ($dark) { $C.Ink } else { $C.White })
  $bg.Line.Visible = $msoFalse
  if (-not $dark) {
    $bar = $slide.Shapes.AddShape($msoShapeRectangle, 0, 0, 9, 540)
    $bar.Fill.ForeColor.RGB = $C.Teal
    $bar.Line.Visible = $msoFalse
    $bar2 = $slide.Shapes.AddShape($msoShapeRectangle, 9, 0, 3, 540)
    $bar2.Fill.ForeColor.RGB = $C.Amber
    $bar2.Line.Visible = $msoFalse
  }
  return $slide
}

function Add-Text($slide, [string]$text, [double]$x, [double]$y, [double]$w, [double]$h, [int]$size = 12, $color = $null, [bool]$bold = $false, [string]$align = "left") {
  if ($null -eq $color) { $color = $C.Slate }
  $box = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $x, $y, $w, $h)
  $tr = $box.TextFrame.TextRange
  $tr.Text = $text
  $tr.Font.Name = "Microsoft YaHei"
  $tr.Font.Size = $size
  $tr.Font.Color.RGB = $color
  $tr.Font.Bold = $(if ($bold) { $msoTrue } else { $msoFalse })
  if ($align -eq "center") { $tr.ParagraphFormat.Alignment = 2 }
  elseif ($align -eq "right") { $tr.ParagraphFormat.Alignment = 3 }
  else { $tr.ParagraphFormat.Alignment = 1 }
  $box.TextFrame.MarginLeft = 0
  $box.TextFrame.MarginRight = 0
  $box.TextFrame.MarginTop = 0
  $box.TextFrame.MarginBottom = 0
  return $box
}

function Add-Title($slide, [string]$title, [string]$kicker = "PATTERN RECOGNITION COURSEWORK") {
  Add-Text $slide $kicker 42 20 390 18 7 $C.Teal $true | Out-Null
  Add-Text $slide $title 42 42 650 42 24 $C.Ink $true | Out-Null
  $line = $slide.Shapes.AddLine(42, 84, 122, 84)
  $line.Line.ForeColor.RGB = $C.Amber
  $line.Line.Weight = 2.5
}

function Add-Footer($slide) {
  Add-Text $slide "模式识别课程实验汇报" 42 514 260 12 7 (Rgb "94A3B8") $false | Out-Null
  Add-Text $slide ("{0:D2}" -f $script:Page) 895 510 28 14 8 $C.Teal $true "right" | Out-Null
  $script:Page += 1
}

function New-ContentSlide([string]$title, [string]$kicker = "PATTERN RECOGNITION COURSEWORK") {
  $s = Add-Slide $false
  Add-Title $s $title $kicker
  Add-Footer $s
  return $s
}

function Add-Panel($slide, [double]$x, [double]$y, [double]$w, [double]$h, $fill = $null, $line = $null) {
  if ($null -eq $fill) { $fill = $C.Pale }
  if ($null -eq $line) { $line = Rgb "E2E8F0" }
  $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $x, $y, $w, $h)
  $shape.Fill.ForeColor.RGB = $fill
  $shape.Line.ForeColor.RGB = $line
  $shape.Line.Weight = 0.7
  return $shape
}

function Add-Card($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$head, [string]$body, $accent = $null, $fill = $null) {
  if ($null -eq $accent) { $accent = $C.Teal }
  if ($null -eq $fill) { $fill = $C.White }
  Add-Panel $slide $x $y $w $h $fill | Out-Null
  $bar = $slide.Shapes.AddShape($msoShapeRectangle, $x, $y, 6, $h)
  $bar.Fill.ForeColor.RGB = $accent
  $bar.Line.Visible = $msoFalse
  Add-Text $slide $head ($x+15) ($y+10) ($w-24) 18 10 $C.Ink $true | Out-Null
  Add-Text $slide $body ($x+15) ($y+35) ($w-24) ($h-42) 8 $C.Slate $false | Out-Null
}

function Add-Metric($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$value, [string]$label, $accent = $null) {
  if ($null -eq $accent) { $accent = $C.Teal }
  Add-Panel $slide $x $y $w $h $C.Pale | Out-Null
  Add-Text $slide $value ($x+12) ($y+8) ($w-22) 24 20 $accent $true | Out-Null
  Add-Text $slide $label ($x+12) ($y+43) ($w-22) 16 7 $C.Slate2 $false | Out-Null
}

function Add-Bullets($slide, [string[]]$items, [double]$x, [double]$y, [double]$w, [double]$h, [int]$size = 10) {
  $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $x, $y, $w, $h)
  $text = [string]::Join("`r", $items)
  $tr = $shape.TextFrame.TextRange
  $tr.Text = $text
  $tr.Font.Name = "Microsoft YaHei"
  $tr.Font.Size = $size
  $tr.Font.Color.RGB = $C.Slate
  $tr.ParagraphFormat.Bullet.Visible = $msoTrue
  $tr.ParagraphFormat.SpaceAfter = 5
  $shape.TextFrame.MarginLeft = 8
  $shape.TextFrame.MarginRight = 2
  $shape.TextFrame.MarginTop = 2
  $shape.TextFrame.MarginBottom = 2
}

function Add-Table($slide, $rows, [double]$x, [double]$y, [double]$w, [double]$h, $headerColor = $null) {
  if ($null -eq $headerColor) { $headerColor = $C.Teal }
  $rCount = $rows.Count
  $cCount = $rows[0].Count
  $tblShape = $slide.Shapes.AddTable($rCount, $cCount, $x, $y, $w, $h)
  $tbl = $tblShape.Table
  for ($r=1; $r -le $rCount; $r++) {
    for ($c=1; $c -le $cCount; $c++) {
      $cell = $tbl.Cell($r,$c)
      $cell.Shape.TextFrame.TextRange.Text = [string]$rows[$r-1][$c-1]
      $cell.Shape.TextFrame2.TextRange.Font.Name = "Microsoft YaHei"
      $cell.Shape.TextFrame2.TextRange.Font.Size = $(if ($r -eq 1) { 8 } else { 7 })
      $cell.Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $(if ($r -eq 1) { $C.White } else { $C.Ink })
      $cell.Shape.TextFrame2.TextRange.Font.Bold = $(if ($r -eq 1) { $msoTrue } else { $msoFalse })
      try {
        $cell.Shape.Fill.ForeColor.RGB = $(if ($r -eq 1) { $headerColor } elseif ($r % 2 -eq 0) { $C.White } else { $C.Pale })
      } catch {
        # Keep default cell fill if the Office COM table cell facade is incomplete.
      }
      try {
        $cell.Borders(1).ForeColor.RGB = $C.Line
        $cell.Borders(2).ForeColor.RGB = $C.Line
        $cell.Borders(3).ForeColor.RGB = $C.Line
        $cell.Borders(4).ForeColor.RGB = $C.Line
      } catch {
        # Some Office builds expose table border formatting differently.
      }
      $cell.Shape.TextFrame.MarginLeft = 3
      $cell.Shape.TextFrame.MarginRight = 3
      $cell.Shape.TextFrame.MarginTop = 2
      $cell.Shape.TextFrame.MarginBottom = 2
    }
  }
  return $tblShape
}

function Add-Image($slide, [string]$file, [double]$x, [double]$y, [double]$w, [double]$h) {
  if (-not (Test-Path $file)) { throw "Missing image $file" }
  Add-Panel $slide $x $y $w $h $C.White | Out-Null
  $pic = $slide.Shapes.AddPicture($file, $msoFalse, $msoTrue, $x+4, $y+4, -1, -1)
  $pic.LockAspectRatio = $msoTrue
  $maxW = $w - 8
  $maxH = $h - 8
  $scaleW = $maxW / $pic.Width
  $scaleH = $maxH / $pic.Height
  $scale = [Math]::Min($scaleW, $scaleH)
  $pic.Width = $pic.Width * $scale
  $pic.Height = $pic.Height * $scale
  $pic.Left = $x + ($w - $pic.Width) / 2
  $pic.Top = $y + ($h - $pic.Height) / 2
  return $pic
}

function Add-Step($slide, [double]$x, [double]$y, [string]$num, [string]$head, [string]$body, $color = $null) {
  if ($null -eq $color) { $color = $C.Teal }
  $oval = $slide.Shapes.AddShape($msoShapeOval, $x, $y, 30, 30)
  $oval.Fill.ForeColor.RGB = $color
  $oval.Line.Visible = $msoFalse
  Add-Text $slide $num ($x+1) ($y+7) 28 12 8 $C.White $true "center" | Out-Null
  Add-Text $slide $head ($x+42) ($y-1) 260 16 10 $C.Ink $true | Out-Null
  Add-Text $slide $body ($x+42) ($y+20) 260 28 8 $C.Slate $false | Out-Null
}

function Add-Section($num, $title, $subtitle) {
  $s = Add-Slide $true
  $bar = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 16, 540)
  $bar.Fill.ForeColor.RGB = $C.Teal
  $bar.Line.Visible = $msoFalse
  $bar2 = $s.Shapes.AddShape($msoShapeRectangle, 16, 0, 5, 540)
  $bar2.Fill.ForeColor.RGB = $C.Amber
  $bar2.Line.Visible = $msoFalse
  Add-Text $s $num 58 82 130 44 22 $C.Amber $true | Out-Null
  Add-Text $s $title 58 150 760 64 32 $C.White $true | Out-Null
  Add-Text $s $subtitle 62 225 700 30 14 (Rgb "CBD5E1") $false | Out-Null
  $ln = $s.Shapes.AddLine(62, 292, 260, 292)
  $ln.Line.ForeColor.RGB = $C.Amber
  $ln.Line.Weight = 4
}

function Add-MiniBar($slide, [double]$x, [double]$y, [string]$label, [double]$value, $color = $null) {
  if ($null -eq $color) { $color = $C.Teal }
  Add-Text $slide $label $x $y 150 12 7 $C.Slate $false | Out-Null
  $bg = $slide.Shapes.AddShape($msoShapeRectangle, $x+160, $y+4, 170, 7)
  $bg.Fill.ForeColor.RGB = Rgb "E2E8F0"
  $bg.Line.Visible = $msoFalse
  $fg = $slide.Shapes.AddShape($msoShapeRectangle, $x+160, $y+4, 170 * $value / 100, 7)
  $fg.Fill.ForeColor.RGB = $color
  $fg.Line.Visible = $msoFalse
  Add-Text $slide ("{0:N2}%" -f $value) ($x+338) ($y-1) 50 14 7 $C.Ink $true "right" | Out-Null
}

# Cover
$s = Add-Slide $true
$bar = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 16, 540); $bar.Fill.ForeColor.RGB = $C.Teal; $bar.Line.Visible = $msoFalse
$bar2 = $s.Shapes.AddShape($msoShapeRectangle, 16, 0, 5, 540); $bar2.Fill.ForeColor.RGB = $C.Amber; $bar2.Line.Visible = $msoFalse
Add-Text $s "PATTERN RECOGNITION" 62 112 320 18 8 $C.Teal2 $true | Out-Null
Add-Text $s "模式识别课程实验汇报" 62 158 600 58 36 $C.White $true | Out-Null
Add-Text $s "Bayes 分类器 · Fisher 判别 · K-L 变换 · 图像识别 · 聚类分析" 66 230 650 26 14 (Rgb "CBD5E1") $false | Out-Null
Add-Panel $s 68 328 200 55 (Rgb "111827") (Rgb "CBD5E1") | Out-Null
$b=$s.Shapes.AddShape($msoShapeRectangle,68,328,6,55);$b.Fill.ForeColor.RGB=$C.Teal2;$b.Line.Visible=$msoFalse
Add-Text $s "姓名" 84 340 160 12 8 (Rgb "CBD5E1") $true | Out-Null
Add-Text $s "请填写" 84 362 160 12 8 (Rgb "E2E8F0") $false | Out-Null
Add-Panel $s 305 328 200 55 (Rgb "111827") (Rgb "CBD5E1") | Out-Null
$b=$s.Shapes.AddShape($msoShapeRectangle,305,328,6,55);$b.Fill.ForeColor.RGB=$C.Amber;$b.Line.Visible=$msoFalse
Add-Text $s "学号" 321 340 160 12 8 (Rgb "CBD5E1") $true | Out-Null
Add-Text $s "请填写" 321 362 160 12 8 (Rgb "E2E8F0") $false | Out-Null
Add-Panel $s 542 328 200 55 (Rgb "111827") (Rgb "CBD5E1") | Out-Null
$b=$s.Shapes.AddShape($msoShapeRectangle,542,328,6,55);$b.Fill.ForeColor.RGB=$C.Blue;$b.Line.Visible=$msoFalse
Add-Text $s "班级" 558 340 160 12 8 (Rgb "CBD5E1") $true | Out-Null
Add-Text $s "请填写" 558 362 160 12 8 (Rgb "E2E8F0") $false | Out-Null

$s = New-ContentSlide "汇报结构：从分类到聚类的完整方法链" "TALK ROADMAP"
$nodes = @(
  @("01 Bayes 分类","概率密度估计与风险决策",$C.Teal),
  @("02 Fisher/Parzen","非参数估计与直接判别",$C.Blue),
  @("03 K-L 特征提取","协方差特征分解与降维",$C.Amber),
  @("04 K-L 图像应用","高维图像压缩与识别",$C.Orange),
  @("05 聚类分析","无监督样本结构发现",$C.Green)
)
for($i=0;$i -lt $nodes.Count;$i++){ Add-Card $s (56+$i*178) 118 145 86 $nodes[$i][0] $nodes[$i][1] $nodes[$i][2] }
Add-Card $s 76 300 170 78 "理论知识" "先讲课程概念、公式和方法思想" $C.Teal
Add-Card $s 278 300 170 78 "实验设置" "说明数据、特征、模型参数和指标" $C.Blue
Add-Card $s 480 300 170 78 "实验结果" "展示准确率、错误率、混淆矩阵和图像" $C.Amber
Add-Card $s 682 300 170 78 "分析总结" "解释现象并联系课程理论" $C.Orange

$s = New-ContentSlide "数据与实验环境" "SOURCE & SETUP"
Add-Table $s @(
  @("数据集","实验用途","规模/说明"),
  @("FEMALE / MALE","实验一、二、三、五","50 女 + 50 男训练样本"),
  @("test1 / test2","分类器泛化测试","35 / 300 个测试样本"),
  @("飞机图像数据集","实验四","16 类，2530 训练图，636 测试图")
) 58 105 455 176 $C.Teal | Out-Null
Add-Table $s @(
  @("工具","作用"),
  @("Python 3.11","实验脚本运行"),
  @("NumPy / SciPy","数值计算与矩阵运算"),
  @("scikit-learn","PCA、评估指标、分级聚类辅助"),
  @("Matplotlib","分类边界、投影、聚类图输出")
) 545 105 350 192 $C.Blue | Out-Null
Add-Metric $s 78 354 160 56 "5" "已完成实验任务" $C.Teal
Add-Metric $s 278 354 160 56 "18" "汇报中使用图表" $C.Blue
Add-Metric $s 478 354 160 56 "2D/1024D" "低维与图像高维特征" $C.Amber
Add-Metric $s 678 354 160 56 "分类+聚类" "监督与无监督覆盖" $C.Orange

Add-Section "01" "实验一：Bayes 性别分类器" "用身高/体重数据进行最小错误率与最小风险决策"
$s = New-ContentSlide "实验一理论：概率决策框架" "BAYES CLASSIFIER"
Add-Card $s 58 108 260 86 "后验概率决策" "选择 P(ωi|x) 最大的类别，本质是把观测特征映射到类别后验。" $C.Teal
Add-Card $s 350 108 260 86 "高斯参数估计" "在正态假设下估计均值 μ 与协方差 Σ，得到二次判别函数。" $C.Blue
Add-Card $s 642 108 230 86 "风险敏感决策" "用损失表刻画不同错误代价，选择条件风险最小的动作。" $C.Amber
Add-Panel $s 78 260 804 70 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "g_i(x) = -1/2 ln|Σ_i| - 1/2(x-μ_i)^TΣ_i^{-1}(x-μ_i) + ln P(ω_i)" 100 286 760 20 16 $C.Teal $true "center" | Out-Null
Add-Bullets $s @("课程相关：参数估计、先验概率、贝叶斯判别函数、最小风险决策","实验目的：观察特征、先验、协方差假设和风险表对分类性能的影响") 80 382 780 60 9

$s = New-ContentSlide "实验一设置：数据、特征与实现流程" "EXPERIMENT DESIGN"
Add-Step $s 70 118 "1" "读取数据" "FEMALE/MALE 作为训练集；test1/test2 作为测试集" $C.Teal
Add-Step $s 70 198 "2" "选择特征" "身高、体重、身高+体重三组对比" $C.Blue
Add-Step $s 70 278 "3" "估计分布" "最大似然估计 μ、σ 或 Σ" $C.Amber
Add-Step $s 70 358 "4" "分类评估" "准确率、召回率、混淆矩阵、错误分析" $C.Orange
Add-Table $s @(
  @("设置项","取值"),
  @("先验概率","0.5/0.5，0.75/0.25，0.9/0.1"),
  @("协方差","完整协方差 / 对角协方差"),
  @("风险表","λ(M|F)=2，λ(F|M)=1"),
  @("非参数补充","Parzen 窗、kNN")
) 468 120 390 185 $C.Teal | Out-Null
Add-Card $s 488 356 350 56 "实现要点" "所有分类器统一输出 train、test1、test2 上的性能，便于横向比较。" $C.Teal

$s = New-ContentSlide "实验一结果：数据分布与 Bayes 决策边界" "RESULT VISUALS"
Add-Image $s (Join-Path $figRoot "experiment1\data_scatter.png") 58 110 410 300 | Out-Null
Add-Image $s (Join-Path $figRoot "experiment1\bayes_2d_boundary.png") 492 110 410 300 | Out-Null
Add-Metric $s 80 430 150 52 "88.00%" "二维 Bayes 训练准确率" $C.Teal
Add-Metric $s 260 430 150 52 "97.14%" "test1 准确率" $C.Blue
Add-Metric $s 440 430 150 52 "89.33%" "test2 准确率" $C.Amber
Add-Metric $s 620 430 150 52 "90.33%" "对角协方差 test2" $C.Orange

$s = New-ContentSlide "实验一分析：先验与风险改变决策偏向" "ANALYSIS"
Add-Table $s @(
  @("方法","训练集","test1","test2"),
  @("二维 Bayes 完整协方差","88.00%","97.14%","89.33%"),
  @("二维 Bayes 对角协方差","88.00%","97.14%","90.33%"),
  @("最小风险 Bayes","84.00%","94.29%","84.33%"),
  @("Parzen Bayes","87.00%","97.14%","90.33%"),
  @("kNN，k=5","89.00%","100.00%","90.33%")
) 60 105 500 190 $C.Teal | Out-Null
Add-Image $s (Join-Path $figRoot "experiment1\bayes_test2_confusion.png") 618 105 220 190 | Out-Null
Add-Bullets $s @("身高和体重组合能利用互补信息，整体比单特征更稳定","test2 中男生比例更高，若人为设置女生先验过大，会显著降低整体准确率","最小风险决策体现错误代价不同时的分类器设计思想") 72 350 780 85 9

Add-Section "02" "实验二：非参数估计与 Fisher 判别" "比较概率密度估计方法与直接设计线性分类器的方法"
$s = New-ContentSlide "实验二理论：Parzen、Fisher 与留一法" "NONPARAMETRIC & LINEAR"
Add-Card $s 58 108 260 90 "Parzen 窗估计" "不预设概率密度形式，用高斯核叠加估计类别条件密度。" $C.Teal
Add-Card $s 350 108 260 90 "Fisher 线性判别" "最大化类间离散度与类内离散度之比，得到线性投影方向。" $C.Blue
Add-Card $s 642 108 230 90 "留一法估计" "每次留出一个样本验证，适合小样本泛化误差估计。" $C.Amber
Add-Panel $s 90 260 780 62 (Rgb "EFF6FF") (Rgb "BFDBFE") | Out-Null
Add-Text $s "J(w)=wᵀS_b w / wᵀS_w w,    w=S_w⁻¹(μ₁-μ₂)" 110 282 740 18 16 $C.Blue $true "center" | Out-Null
Add-Bullets $s @("课程重点：参数估计与非参数估计的差别；概率模型与判别模型的差别","实验问题：非参数 Bayes 和 Fisher 是否能达到高斯 Bayes 的性能") 80 382 780 60 9

$s = New-ContentSlide "实验二结果：线性边界与投影分布" "VISUAL COMPARISON"
Add-Image $s (Join-Path $figRoot "experiment2\fisher_bayes_boundaries.png") 58 110 410 300 | Out-Null
Add-Image $s (Join-Path $figRoot "experiment2\fisher_projection.png") 492 110 410 300 | Out-Null
Add-Card $s 80 430 360 50 "观察 1" "Fisher 给出简单线性边界，Bayes 边界受密度估计影响。" $C.Blue
Add-Card $s 520 430 360 50 "观察 2" "投影后男女样本主体分离，但边界区域仍有重叠。" $C.Teal

$s = New-ContentSlide "实验二结果：错误率对比" "METRICS"
Add-Table $s @(
  @("方法","训练错误率","test1 错误率","test2 错误率","留一法错误率"),
  @("Gaussian Bayes","12.00%","2.86%","10.67%","12.00%"),
  @("Parzen Bayes","13.00%","2.86%","9.67%","14.00%"),
  @("Fisher","12.00%","2.86%","9.67%","13.00%")
) 70 122 820 150 $C.Blue | Out-Null
Add-MiniBar $s 90 330 "Gaussian Bayes test2 acc." 89.33 $C.Teal
Add-MiniBar $s 90 365 "Parzen Bayes test2 acc." 90.33 $C.Blue
Add-MiniBar $s 90 400 "Fisher test2 acc." 90.33 $C.Amber
Add-Card $s 555 320 300 90 "结论" "三种方法表现接近。Fisher 用更简单的线性边界取得与非参数 Bayes 相当的 test2 准确率。" $C.Teal

$s = New-ContentSlide "实验二分析：模型复杂度与泛化" "TAKEAWAYS"
Add-Card $s 70 125 245 100 "参数 Bayes" "形式清晰、样本效率高，但依赖正态分布假设。" $C.Teal
Add-Card $s 360 125 245 100 "Parzen Bayes" "密度形式灵活，但窗口宽度 h 和样本规模影响明显。" $C.Blue
Add-Card $s 650 125 225 100 "Fisher" "不估计密度，直接寻找判别方向，解释性强。" $C.Amber
Add-Panel $s 90 315 780 70 $C.PaleAmber (Rgb "FDE68A") | Out-Null
Add-Text $s "留一法结果与 test2 错误率接近，说明它能在训练样本有限时给出较稳健的泛化误差估计。" 118 340 720 18 13 (Rgb "92400E") $true "center" | Out-Null

Add-Section "03" "实验三：K-L 变换进行特征提取" "从无监督主成分方向理解身高体重样本的主要变化"
$s = New-ContentSlide "实验三理论：K-L/PCA 与 Fisher 的差别" "FEATURE EXTRACTION"
Add-Card $s 70 115 350 90 "K-L / PCA" "对协方差矩阵做特征分解，保留总体方差最大的方向。" $C.Teal
Add-Card $s 495 115 350 90 "Fisher" "利用类别标签，寻找最有利于分类的投影方向。" $C.Blue
Add-Panel $s 100 265 760 62 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "Σuᵢ = λᵢuᵢ,      zᵢ = uᵢᵀ(x-μ)" 120 286 720 18 18 $C.Teal $true "center" | Out-Null
Add-Bullets $s @("PCA 的目标是表示数据，不是直接优化分类","当最大方差方向与类别差异方向一致时，PCA 也能得到较好分类效果") 80 380 760 60 9

$s = New-ContentSlide "实验三设置：二维特征的主方向分析" "EXPERIMENT DESIGN"
Add-Table $s @(
  @("项目","数值/说明"),
  @("数据","FEMALE + MALE，共 100 个训练样本"),
  @("特征","身高、体重"),
  @("第一主成分解释率","87.52%"),
  @("PC1","(0.6269, 0.7791)"),
  @("类均值方向","(0.6514, 0.7587)")
) 65 110 350 200 $C.Teal | Out-Null
Add-Image $s (Join-Path $figRoot "experiment3\kl_directions.png") 455 105 415 310 | Out-Null
Add-Card $s 80 388 315 56 "设置说明" "分别使用 PCA 第一主成分、类均值差方向、Fisher 和 Bayes 进行分类比较。" $C.Teal

$s = New-ContentSlide "实验三结果：投影分布与分类性能" "RESULTS"
Add-Image $s (Join-Path $figRoot "experiment3\pca_projection.png") 65 110 410 260 | Out-Null
Add-Table $s @(
  @("方法","训练集","test1","test2"),
  @("K-L PCA 第一主成分","86.00%","100.00%","89.67%"),
  @("类均值方向","86.00%","100.00%","89.67%"),
  @("Fisher","88.00%","97.14%","90.33%"),
  @("Gaussian Bayes","88.00%","97.14%","89.33%")
) 515 122 380 180 $C.Teal | Out-Null
Add-Bullets $s @("PCA 第一主成分已经能较好区分男女样本","Fisher 在 test2 上略优，体现监督判别方向的优势") 525 340 360 62 9

$s = New-ContentSlide "实验三分析：最大方差不等于最佳分类" "TAKEAWAYS"
Add-Card $s 70 120 245 90 "为什么 PCA 有效" "本数据中身高和体重的主要变化方向与性别差异方向接近。" $C.Teal
Add-Card $s 360 120 245 90 "为什么 Fisher 略优" "Fisher 直接利用类别标签，目标就是提高类别分离度。" $C.Blue
Add-Card $s 650 120 225 90 "一般规律" "无监督降维需要和监督分类器共同评价。" $C.Amber
Add-Image $s (Join-Path $figRoot "experiment3\mean_projection.png") 160 270 640 180 | Out-Null

Add-Section "04" "实验四：K-L 变换的图像识别应用" "把 K-L/PCA 从二维样本扩展到高维飞机图像"
$s = New-ContentSlide "实验四理论：高维图像的低维表示" "IMAGE RECOGNITION"
Add-Step $s 80 120 "1" "图像预处理" "灰度化、缩放到 32×32" $C.Teal
Add-Step $s 80 200 "2" "向量化" "每张图像展开为 1024 维向量" $C.Blue
Add-Step $s 80 280 "3" "K-L 降维" "提取前 k 个主成分作为图像特征" $C.Amber
Add-Step $s 80 360 "4" "低维分类" "NearestCentroid 与 1NN 对比" $C.Orange
Add-Panel $s 515 128 315 68 (Rgb "EFF6FF") (Rgb "BFDBFE") | Out-Null
Add-Text $s "z = U_kᵀ(x-μ)" 540 150 265 20 18 $C.Blue $true "center" | Out-Null
Add-Table $s @(
  @("项目","设置"),
  @("类别数","16 类"),
  @("训练集","2530 张"),
  @("测试集","636 张"),
  @("主成分","8 / 16 / 32 / 64 / 128")
) 515 240 315 165 $C.Blue | Out-Null

$s = New-ContentSlide "实验四结果：样本图像与 PCA 主成分" "VISUAL RESULTS"
Add-Image $s (Join-Path $figRoot "experiment4\sample_images.png") 58 110 410 300 | Out-Null
Add-Image $s (Join-Path $figRoot "experiment4\pca_components.png") 492 110 410 300 | Out-Null
Add-Metric $s 80 430 150 52 "1024→64" "维度压缩" $C.Blue
Add-Metric $s 260 430 150 52 "89.88%" "64 维解释方差" $C.Teal
Add-Metric $s 440 430 150 52 "16" "飞机类别数" $C.Amber
Add-Metric $s 620 430 150 52 "636" "测试图像数" $C.Orange

$s = New-ContentSlide "实验四结果：主成分数与识别准确率" "METRICS"
Add-Image $s (Join-Path $figRoot "experiment4\accuracy_vs_components.png") 65 110 420 300 | Out-Null
Add-Table $s @(
  @("方法","主成分数","训练集","测试集"),
  @("PCA+NearestCentroid","64","20.47%","19.97%"),
  @("PCA+1NN","32","100.00%","54.09%"),
  @("PCA+1NN","64","100.00%","53.93%"),
  @("PCA+1NN","128","100.00%","52.99%")
) 525 125 350 168 $C.Orange | Out-Null
Add-Bullets $s @("1NN 明显优于最近质心，说明飞机类别内部变化复杂","主成分数增加到 32 后收益变小，后续成分可能包含背景和噪声") 535 340 340 65 9

$s = New-ContentSlide "实验四分析：PCA 能压缩，但不是完整识别方案" "ANALYSIS"
Add-Image $s (Join-Path $figRoot "experiment4\confusion_matrix.png") 75 110 420 310 | Out-Null
Add-Card $s 555 120 300 80 "混淆原因" "低分辨率灰度图中，部分飞机外形相近，姿态和背景也会干扰。" $C.Orange
Add-Card $s 555 235 300 80 "方法局限" "PCA 是无监督降维，优先保留总体方差，不保证保留最强判别信息。" $C.Blue
Add-Card $s 555 350 300 80 "改进方向" "可结合 HOG/SIFT/ORB、数据增强或 CNN 提升识别性能。" $C.Teal

Add-Section "05" "实验五：C 均值与分级聚类分析" "不使用类别标签，观察身高体重样本的自然结构"
$s = New-ContentSlide "实验五理论：从监督分类到无监督聚类" "CLUSTERING"
Add-Card $s 70 115 350 90 "C 均值聚类" "独立编程实现，通过样本分配和中心更新最小化类内平方误差。" $C.Teal
Add-Card $s 495 115 350 90 "分级聚类" "使用 Ward 方法形成层次结构，用树状图观察样本合并过程。" $C.Blue
Add-Panel $s 100 265 760 62 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "J = Σⱼ Σₓ∈ωⱼ ||x - μⱼ||²" 120 286 720 18 18 $C.Teal $true "center" | Out-Null
Add-Bullets $s @("聚类不使用性别标签，评价时再与真实标签比较","特征标准化很重要，否则体重/身高尺度会影响距离计算") 80 380 760 60 9

$s = New-ContentSlide "实验五设置：两组数据与多种聚类数" "EXPERIMENT DESIGN"
Add-Table $s @(
  @("设置项","说明"),
  @("数据一","FEMALE + MALE，共 100 个训练样本"),
  @("数据二","训练集 + test2，共 400 个样本"),
  @("特征","身高、体重，先标准化"),
  @("C 均值类别数","2、3、4、5"),
  @("初始值","多个随机种子比较稳定性")
) 70 110 400 210 $C.Teal | Out-Null
Add-Step $s 545 120 "1" "初始化中心" "随机选择 C 个中心" $C.Teal
Add-Step $s 545 200 "2" "样本分配" "分配到最近中心" $C.Blue
Add-Step $s 545 280 "3" "中心更新" "计算各簇均值" $C.Amber
Add-Step $s 545 360 "4" "指标评估" "SSE、轮廓系数、ARI、匹配率" $C.Orange

$s = New-ContentSlide "实验五结果：C=2 聚类与初始值影响" "C-MEANS RESULTS"
Add-Image $s (Join-Path $figRoot "experiment5\cmeans_train_2clusters.png") 58 110 410 295 | Out-Null
Add-Image $s (Join-Path $figRoot "experiment5\init_variation.png") 492 110 410 295 | Out-Null
Add-Metric $s 80 428 150 52 "85.00%" "训练集 C=2 匹配率" $C.Teal
Add-Metric $s 260 428 150 52 "85-87%" "不同初始值波动" $C.Blue
Add-Metric $s 440 428 150 52 "0.5108" "轮廓系数" $C.Amber
Add-Metric $s 620 428 150 52 "0.4850" "ARI" $C.Orange

$s = New-ContentSlide "实验五结果：类别数选择与分级聚类" "MODEL SELECTION"
Add-Image $s (Join-Path $figRoot "experiment5\cmeans_k_metrics_train.png") 58 110 410 295 | Out-Null
Add-Image $s (Join-Path $figRoot "experiment5\hierarchical_dendrogram_train.png") 492 110 410 295 | Out-Null
Add-Card $s 80 430 360 50 "类别数判断" "SSE 随 C 增大下降，但轮廓系数在 C=2 最高，因此两类更合理。" $C.Teal
Add-Card $s 520 430 360 50 "分级聚类" "Ward 方法匹配率 81.00%，略低于 C 均值，但能展示层次结构。" $C.Blue

$s = New-ContentSlide "实验五分析：加入 test2 后结构保持但边界更复杂" "COMBINED DATA"
Add-Image $s (Join-Path $figRoot "experiment5\cmeans_combined_2clusters.png") 58 108 405 275 | Out-Null
Add-Image $s (Join-Path $figRoot "experiment5\cmeans_k_metrics_combined.png") 498 108 405 275 | Out-Null
Add-Table $s @(
  @("数据","方法","轮廓系数","ARI","匹配率"),
  @("训练集","C 均值","0.5108","0.4850","85.00%"),
  @("训练集","Ward","0.5049","0.3789","81.00%"),
  @("训练集+test2","C 均值","0.4850","0.4693","84.50%"),
  @("训练集+test2","Ward","0.4739","0.4172","82.50%")
) 70 410 820 72 $C.Teal | Out-Null

$s = New-ContentSlide "五个实验的横向总结" "SYNTHESIS"
Add-Table $s @(
  @("实验","课程知识点","监督信息","主要目标"),
  @("Bayes 分类","后验概率、参数估计、风险决策","使用标签","最小错误率或最小风险"),
  @("Fisher / Parzen","非参数估计、线性判别、留一法","使用标签","比较概率模型与直接判别"),
  @("K-L 特征提取","PCA、协方差特征分解","可不使用标签","降维与主方向分析"),
  @("K-L 图像应用","高维图像特征压缩","使用标签","验证 PCA 图像识别作用"),
  @("聚类分析","C 均值、分级聚类","不使用标签","发现样本自然结构")
) 58 110 844 290 $C.Teal | Out-Null
Add-Panel $s 85 430 790 48 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "主线：从概率决策到线性判别，从监督分类到无监督聚类，从低维特征到高维图像识别。" 110 446 740 16 11 $C.Teal $true "center" | Out-Null

$s = New-ContentSlide "总体结论与汇报收束" "CONCLUSION"
Add-Card $s 70 120 245 95 "特征选择" "身高+体重比单特征更稳定；图像任务中 PCA 可压缩高维特征。" $C.Teal
Add-Card $s 360 120 245 95 "模型假设" "正态假设、窗口宽度、线性边界都会影响最终性能。" $C.Blue
Add-Card $s 650 120 225 95 "评价指标" "准确率、错误率、召回率、混淆矩阵和聚类指标需要综合分析。" $C.Amber
Add-Panel $s 100 300 760 76 $C.PaleAmber (Rgb "FDE68A") | Out-Null
Add-Text $s "核心体会：模式识别方法没有绝对最优，算法选择必须结合数据分布、任务目标和评价指标。" 125 326 710 20 14 (Rgb "92400E") $true "center" | Out-Null
Add-Bullets $s @("Bayes 强调概率决策，Fisher 强调判别方向，K-L 强调表示与降维，聚类强调自然结构。") 120 420 720 35 9

$s = Add-Slide $true
$bar = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 16, 540); $bar.Fill.ForeColor.RGB = $C.Teal; $bar.Line.Visible = $msoFalse
$bar2 = $s.Shapes.AddShape($msoShapeRectangle, 16, 0, 5, 540); $bar2.Fill.ForeColor.RGB = $C.Amber; $bar2.Line.Visible = $msoFalse
Add-Text $s "谢谢" 70 198 300 60 42 $C.White $true | Out-Null
Add-Text $s "欢迎老师和同学批评指正" 74 275 440 25 17 (Rgb "CBD5E1") $false | Out-Null
$ln = $s.Shapes.AddLine(74, 326, 260, 326); $ln.Line.ForeColor.RGB = $C.Amber; $ln.Line.Weight = 4

if (Test-Path $Output) { Remove-Item -LiteralPath $Output -Force }
$pres.SaveAs($Output, $ppSaveAsOpenXMLPresentation)
$slideCount = $pres.Slides.Count
$pres.Close()
$app.Quit()

# Reopen and render to PNG previews.
$app2 = New-Object -ComObject PowerPoint.Application
$app2.Visible = $msoTrue
$pres2 = $app2.Presentations.Open($Output, $msoTrue, $msoFalse, $msoFalse)
$verifiedSlides = $pres2.Slides.Count
$pres2.Export($previewDir, "PNG", 1280, 720)
$pres2.Close()
$app2.Quit()

Set-Content -Path (Join-Path $qaDir "visual-qa.txt") -Encoding UTF8 -Value @"
Visual QA ledger for final polished editable PPTX

Final PPTX: $Output
Created with PowerPoint COM native shapes, text boxes, tables, and images.
Generated slide count: $slideCount
Reopened slide count: $verifiedSlides
Preview export directory: $previewDir

Checks:
- Final PPTX can be opened by PowerPoint COM.
- All slides exported to PNG previews.
- Text, tables, shapes and images are independent PowerPoint objects.
- No external online assets or unverified logos were used.
"@

Write-Output "FINAL_PPTX=$Output"
Write-Output "SLIDES=$verifiedSlides"
Write-Output "PREVIEW_DIR=$previewDir"
