param(
  [string]$Output = "F:\Pattern-Recognition_npu\slides\pattern-recognition-report\experiment1-bayes-detailed.pptx"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$figRoot = Join-Path $root "figures\experiment1"
$tmpRoot = Join-Path $env:TEMP "codex-presentations\experiment1-bayes-detailed"
$previewDir = Join-Path $tmpRoot "preview"
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null

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
  PaleTeal = Rgb "ECFDF5"; PaleAmber = Rgb "FFFBEB"; White = Rgb "FFFFFF"
}

$msoFalse = 0
$msoTrue = -1
$ppLayoutBlank = 12
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeOval = 9
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

function Add-Title($slide, [string]$title, [string]$kicker = "EXPERIMENT 1") {
  Add-Text $slide $kicker 42 20 390 18 7 $C.Teal $true | Out-Null
  Add-Text $slide $title 42 42 720 42 23 $C.Ink $true | Out-Null
  $line = $slide.Shapes.AddLine(42, 84, 122, 84)
  $line.Line.ForeColor.RGB = $C.Amber
  $line.Line.Weight = 2.5
}

function Add-Footer($slide) {
  Add-Text $slide "实验一：Bayes 性别分类器" 42 514 260 12 7 (Rgb "94A3B8") $false | Out-Null
  Add-Text $slide ("{0:D2}" -f $script:Page) 895 510 28 14 8 $C.Teal $true "right" | Out-Null
  $script:Page += 1
}

function New-ContentSlide([string]$title, [string]$kicker = "EXPERIMENT 1") {
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
  Add-Text $slide $body ($x+15) ($y+34) ($w-24) ($h-40) 8 $C.Slate $false | Out-Null
}

function Add-Metric($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$value, [string]$label, $accent = $null) {
  if ($null -eq $accent) { $accent = $C.Teal }
  Add-Panel $slide $x $y $w $h $C.Pale | Out-Null
  Add-Text $slide $value ($x+12) ($y+8) ($w-22) 24 19 $accent $true | Out-Null
  Add-Text $slide $label ($x+12) ($y+41) ($w-22) 16 7 $C.Slate2 $false | Out-Null
}

function Add-Bullets($slide, [string[]]$items, [double]$x, [double]$y, [double]$w, [double]$h, [int]$size = 9) {
  $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $x, $y, $w, $h)
  $tr = $shape.TextFrame.TextRange
  $tr.Text = [string]::Join("`r", $items)
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
      try { $cell.Shape.Fill.ForeColor.RGB = $(if ($r -eq 1) { $headerColor } elseif ($r % 2 -eq 0) { $C.White } else { $C.Pale }) } catch {}
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
  $scale = [Math]::Min($maxW / $pic.Width, $maxH / $pic.Height)
  $pic.Width = $pic.Width * $scale
  $pic.Height = $pic.Height * $scale
  $pic.Left = $x + ($w - $pic.Width) / 2
  $pic.Top = $y + ($h - $pic.Height) / 2
}

function Add-Step($slide, [double]$x, [double]$y, [string]$num, [string]$head, [string]$body, $color = $null) {
  if ($null -eq $color) { $color = $C.Teal }
  $oval = $slide.Shapes.AddShape($msoShapeOval, $x, $y, 30, 30)
  $oval.Fill.ForeColor.RGB = $color
  $oval.Line.Visible = $msoFalse
  Add-Text $slide $num ($x+1) ($y+7) 28 12 8 $C.White $true "center" | Out-Null
  Add-Text $slide $head ($x+42) ($y-1) 275 16 10 $C.Ink $true | Out-Null
  Add-Text $slide $body ($x+42) ($y+20) 275 30 8 $C.Slate $false | Out-Null
}

function Add-Section($num, $title, $subtitle) {
  $s = Add-Slide $true
  $bar = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 16, 540); $bar.Fill.ForeColor.RGB = $C.Teal; $bar.Line.Visible = $msoFalse
  $bar2 = $s.Shapes.AddShape($msoShapeRectangle, 16, 0, 5, 540); $bar2.Fill.ForeColor.RGB = $C.Amber; $bar2.Line.Visible = $msoFalse
  Add-Text $s $num 58 82 130 44 22 $C.Amber $true | Out-Null
  Add-Text $s $title 58 150 760 64 31 $C.White $true | Out-Null
  Add-Text $s $subtitle 62 225 700 30 14 (Rgb "CBD5E1") $false | Out-Null
  $ln = $s.Shapes.AddLine(62, 292, 260, 292); $ln.Line.ForeColor.RGB = $C.Amber; $ln.Line.Weight = 4
}

function Add-MiniBar($slide, [double]$x, [double]$y, [string]$label, [double]$value, $color = $null) {
  if ($null -eq $color) { $color = $C.Teal }
  Add-Text $slide $label $x $y 158 12 7 $C.Slate $false | Out-Null
  $bg = $slide.Shapes.AddShape($msoShapeRectangle, $x+168, $y+4, 180, 7); $bg.Fill.ForeColor.RGB = Rgb "E2E8F0"; $bg.Line.Visible = $msoFalse
  $fg = $slide.Shapes.AddShape($msoShapeRectangle, $x+168, $y+4, 180 * $value / 100, 7); $fg.Fill.ForeColor.RGB = $color; $fg.Line.Visible = $msoFalse
  Add-Text $slide ("{0:N2}%" -f $value) ($x+356) ($y-1) 54 14 7 $C.Ink $true "right" | Out-Null
}

# Cover
$s = Add-Slide $true
$bar = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 16, 540); $bar.Fill.ForeColor.RGB = $C.Teal; $bar.Line.Visible = $msoFalse
$bar2 = $s.Shapes.AddShape($msoShapeRectangle, 16, 0, 5, 540); $bar2.Fill.ForeColor.RGB = $C.Amber; $bar2.Line.Visible = $msoFalse
Add-Text $s "PATTERN RECOGNITION · EXPERIMENT 1" 62 112 360 18 8 $C.Teal2 $true | Out-Null
Add-Text $s "Bayes 性别分类器实验" 62 158 600 58 36 $C.White $true | Out-Null
Add-Text $s "理论原理 · 实验设计 · 软件实现 · 实验结果分析" 66 230 650 26 14 (Rgb "CBD5E1") $false | Out-Null
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

$s = New-ContentSlide "本实验要解决什么问题？" "OVERVIEW"
Add-Card $s 60 115 250 95 "输入" "每个样本包含身高、体重两个特征，部分测试样本带有真实性别标签。" $C.Teal
Add-Card $s 355 115 250 95 "模型" "在正态分布假设下建立 Bayes 分类器，并比较不同特征、先验和风险设置。" $C.Blue
Add-Card $s 650 115 230 95 "输出" "预测性别，并用准确率、召回率、混淆矩阵等指标评价分类效果。" $C.Amber
Add-Panel $s 95 285 770 90 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "核心问题：如何根据身高/体重特征 x，判断样本属于女生 F 还是男生 M？" 120 318 720 22 16 $C.Teal $true "center" | Out-Null
Add-Bullets $s @("这是典型的两类模式分类问题。","课程知识点包括 Bayes 决策、参数估计、先验概率、风险函数和非参数估计。") 90 420 760 55 9

Add-Section "01" "理论原理" "从 Bayes 决策到高斯判别函数、风险决策与非参数估计"

$s = New-ContentSlide "Bayes 决策：从先验到后验" "THEORY"
Add-Card $s 60 112 250 90 "先验概率 P(ωi)" "观察样本前对类别出现概率的估计，例如男女先验相等。" $C.Teal
Add-Card $s 355 112 250 90 "类条件密度 p(x|ωi)" "给定类别后，身高/体重特征出现的概率密度。" $C.Blue
Add-Card $s 650 112 230 90 "后验概率 P(ωi|x)" "观察到特征 x 后，样本属于某类别的概率。" $C.Amber
Add-Panel $s 90 260 780 68 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "ω* = argmax P(ωi|x) = argmax p(x|ωi)P(ωi)" 120 284 720 18 17 $C.Teal $true "center" | Out-Null
Add-Bullets $s @("p(x) 对所有类别相同，因此分类时只比较 p(x|ωi)P(ωi)。","Bayes 分类器本质上是在概率意义下选择最可能的类别。") 90 388 760 58 9

$s = New-ContentSlide "正态分布假设与参数估计" "THEORY"
Add-Card $s 58 110 260 100 "一维特征" "只使用身高或体重时，假设每类服从一维正态分布 N(μi, σi²)。" $C.Teal
Add-Card $s 350 110 260 100 "二维特征" "同时使用身高和体重时，假设每类服从二维高斯分布 N(μi, Σi)。" $C.Blue
Add-Card $s 642 110 230 100 "最大似然估计" "用训练样本估计每个类别的均值、方差或协方差矩阵。" $C.Amber
Add-Panel $s 70 280 395 90 (Rgb "EFF6FF") (Rgb "BFDBFE") | Out-Null
Add-Text $s "μi = (1/Ni) Σ xk" 95 305 350 18 15 $C.Blue $true "center" | Out-Null
Add-Text $s "Σi = (1/Ni) Σ (xk-μi)(xk-μi)^T" 95 335 350 18 15 $C.Blue $true "center" | Out-Null
Add-Panel $s 500 280 350 90 $C.PaleAmber (Rgb "FDE68A") | Out-Null
Add-Text $s "本实验估计的是女生和男生各自的分布参数，然后构造分类边界。" 525 310 300 36 11 (Rgb "92400E") $true "center" | Out-Null

$s = New-ContentSlide "二维高斯 Bayes 判别函数" "THEORY"
Add-Panel $s 70 115 820 105 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "gi(x) = -1/2 ln|Σi| - 1/2(x-μi)^TΣi⁻¹(x-μi) + ln P(ωi)" 95 150 770 24 15 $C.Teal $true "center" | Out-Null
Add-Card $s 90 275 235 90 "均值项 μi" "决定类别中心位置。男生均值整体高于女生均值。" $C.Teal
Add-Card $s 365 275 235 90 "协方差 Σi" "决定类别分布形状和身高体重是否相关。" $C.Blue
Add-Card $s 640 275 215 90 "先验 P(ωi)" "决定分类边界向哪个类别偏移。" $C.Amber
Add-Text $s "决策规则：若 gF(x) ≥ gM(x)，判为女生；否则判为男生。" 120 420 720 18 13 $C.Ink $true "center" | Out-Null

$s = New-ContentSlide "完整协方差与对角协方差" "THEORY"
Add-Card $s 70 112 360 115 "完整协方差假设：相关" "保留协方差矩阵中的非对角元素，表示身高和体重之间可能有关联。判别边界通常更灵活。" $C.Teal
Add-Card $s 530 112 330 115 "对角协方差假设：不相关" "只保留方差项，忽略身高和体重之间的相关性。参数更少，估计更稳定。" $C.Blue
Add-Table $s @(
  @("比较项","完整协方差","对角协方差"),
  @("参数数量","更多","更少"),
  @("表达能力","更强，可表达相关性","较弱，假设特征独立"),
  @("小样本稳定性","可能受估计方差影响","通常更稳")
) 128 285 704 130 $C.Teal | Out-Null
Add-Text $s "实验中两种假设都要比较，不能只凭理论复杂度判断好坏。" 140 448 680 18 12 $C.Slate $true "center" | Out-Null

$s = New-ContentSlide "先验概率与决策边界偏移" "THEORY"
Add-Card $s 70 112 250 100 "等先验" "P(F)=0.5, P(M)=0.5，认为男女出现概率相同。" $C.Teal
Add-Card $s 355 112 250 100 "女生先验较大" "P(F)=0.75 或 0.9 时，分类器更容易把样本判为女生。" $C.Blue
Add-Card $s 640 112 250 100 "与测试分布有关" "如果测试集中男生占多数，过高女生先验会造成大量男生误判。" $C.Orange
Add-Panel $s 100 292 760 78 $C.PaleAmber (Rgb "FDE68A") | Out-Null
Add-Text $s "先验不是随意调参，它应来自真实类别比例、任务背景或已有知识。" 130 320 700 20 15 (Rgb "92400E") $true "center" | Out-Null
Add-Bullets $s @("test2 中男生样本远多于女生。","当人为提高女生先验时，test2 总体准确率明显下降。") 120 412 720 48 9

$s = New-ContentSlide "最小风险 Bayes 决策" "THEORY"
Add-Panel $s 80 112 780 68 (Rgb "EFF6FF") (Rgb "BFDBFE") | Out-Null
Add-Text $s "R(αi|x) = Σ λ(αi|ωj) P(ωj|x)" 120 138 700 18 17 $C.Blue $true "center" | Out-Null
Add-Table $s @(
  @("损失项","含义","取值"),
  @("λ(F|F)","女生判为女生","0"),
  @("λ(M|M)","男生判为男生","0"),
  @("λ(F|M)","男生误判为女生","1"),
  @("λ(M|F)","女生误判为男生","2")
) 110 230 740 150 $C.Blue | Out-Null
Add-Bullets $s @("最小错误率只统计错几个；最小风险还考虑错的代价。","本实验中女生误判为男生的代价更高，因此分类器会更谨慎地判为男生。") 110 420 740 55 9

$s = New-ContentSlide "非参数估计：Parzen 窗与 kNN" "THEORY"
Add-Card $s 70 112 360 115 "Parzen 窗 Bayes" "不假设正态分布，用训练样本周围的高斯核叠加估计 p(x|ωi)。本实验取 h=5.0。" $C.Teal
Add-Card $s 530 112 330 115 "k 近邻分类" "不显式估计概率密度，直接找距离最近的 k 个训练样本投票。本实验取 k=5。" $C.Blue
Add-Table $s @(
  @("方法","优点","局限"),
  @("参数 Bayes","形式清晰、样本效率高","依赖分布假设"),
  @("Parzen","密度形式灵活","依赖窗口宽度和样本量"),
  @("kNN","实现直观、利用局部结构","预测时计算量较大")
) 110 290 740 130 $C.Teal | Out-Null

Add-Section "02" "实验设计与软件实现" "数据集、实验分组、代码流程与评价指标"

$s = New-ContentSlide "数据集与样本划分" "EXPERIMENT DESIGN"
Add-Table $s @(
  @("文件","用途","样本说明"),
  @("FEMALE.TXT","训练集","50 个女生的身高、体重"),
  @("MALE.TXT","训练集","50 个男生的身高、体重"),
  @("test1.txt","测试集","35 个样本，15 女 20 男"),
  @("test2.txt","测试集","300 个样本，50 女 250 男")
) 92 112 776 180 $C.Teal | Out-Null
Add-Metric $s 110 335 160 58 "100" "训练样本总数" $C.Teal
Add-Metric $s 310 335 160 58 "35" "test1 样本数" $C.Blue
Add-Metric $s 510 335 160 58 "300" "test2 样本数" $C.Amber
Add-Metric $s 710 335 160 58 "2" "类别数量" $C.Orange

$s = New-ContentSlide "实验分组设计" "EXPERIMENT DESIGN"
Add-Card $s 70 110 250 88 "单特征实验" "分别只用身高或体重，建立一维高斯 Bayes 分类器。" $C.Teal
Add-Card $s 355 110 250 88 "双特征实验" "同时用身高和体重，比较完整协方差和对角协方差。" $C.Blue
Add-Card $s 640 110 230 88 "先验概率实验" "比较 0.5/0.5、0.75/0.25、0.9/0.1 三种先验。" $C.Amber
Add-Card $s 70 255 250 88 "最小风险实验" "给定损失表，按照条件风险最小原则分类。" $C.Orange
Add-Card $s 355 255 250 88 "参数估计实例" "使用最大似然估计计算均值和协方差。" $C.Green
Add-Card $s 640 255 230 88 "非参数估计实例" "实现 Parzen 窗 Bayes 和 kNN 进行对比。" $C.Blue
Add-Text $s "实验设计覆盖了任务一要求中的特征选择、先验影响、风险决策、参数估计和非参数估计。" 100 425 760 18 12 $C.Ink $true "center" | Out-Null

$s = New-ContentSlide "软件实现流程" "SOFTWARE IMPLEMENTATION"
Add-Step $s 85 118 "1" "读取数据" "解析训练文件和测试文件，统一标签为 F/M。" $C.Teal
Add-Step $s 85 198 "2" "选择特征" "按实验选择身高、体重或二维特征。" $C.Blue
Add-Step $s 85 278 "3" "训练模型" "估计均值、方差/协方差，或保存非参数样本。" $C.Amber
Add-Step $s 520 118 "4" "进行预测" "计算判别函数、风险函数或近邻投票。" $C.Orange
Add-Step $s 520 198 "5" "评价性能" "输出 accuracy、precision、recall、confusion matrix。" $C.Green
Add-Step $s 520 278 "6" "生成图表" "绘制散点图、决策边界和混淆矩阵。" $C.Teal
Add-Card $s 180 392 600 62 "对应脚本" "scripts/experiment1_bayes.py：包含 fit_gaussian、predict_gaussian、predict_min_risk、predict_parzen、predict_knn 等核心函数。" $C.Blue

$s = New-ContentSlide "评价指标说明" "EVALUATION"
Add-Card $s 70 112 250 90 "准确率 Accuracy" "预测正确样本数 / 总样本数，用于整体性能比较。" $C.Teal
Add-Card $s 355 112 250 90 "精确率 Precision" "预测为某类的样本中，真正属于该类的比例。" $C.Blue
Add-Card $s 640 112 230 90 "召回率 Recall" "真实属于某类的样本中，被正确找出的比例。" $C.Amber
Add-Panel $s 115 270 730 130 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "混淆矩阵" 135 292 680 18 14 $C.Teal $true "center" | Out-Null
Add-Text $s "用于观察女生判错为男生、男生判错为女生的具体数量；比单一准确率更能解释分类器错误类型。" 155 330 650 35 12 $C.Teal $true "center" | Out-Null

Add-Section "03" "实验结果" "参数、图像、表格与不同设置下的性能比较"

$s = New-ContentSlide "训练集估计参数" "PARAMETER ESTIMATION"
Add-Table $s @(
  @("类别","均值向量 μ","协方差矩阵 Σ"),
  @("女生 F","[162.84, 52.596]","[[43.934, 15.525], [15.525, 31.129]]"),
  @("男生 M","[173.92, 65.502]","[[20.754, 23.058], [23.058, 59.898]]")
) 70 118 820 145 $C.Teal | Out-Null
Add-Bullets $s @("男生平均身高和平均体重均高于女生，这是分类的主要依据。","协方差矩阵非对角元素不为 0，说明身高和体重之间存在一定相关性。","协方差估计是否可靠，会影响完整协方差 Bayes 的测试表现。") 92 315 760 90 9

$s = New-ContentSlide "样本分布图：身高与体重的可分性" "RESULT FIGURE"
Add-Image $s (Join-Path $figRoot "data_scatter.png") 78 110 520 340
Add-Card $s 640 130 230 80 "观察 1" "男生样本整体位于更高身高、更高体重区域。" $C.Teal
Add-Card $s 640 245 230 80 "观察 2" "两类样本在边界区域存在重叠，因此无法完全无误分类。" $C.Blue
Add-Card $s 640 360 230 80 "观察 3" "test1 样本基本落在训练集分布附近，测试准确率较高。" $C.Amber

$s = New-ContentSlide "二维 Bayes 决策边界" "RESULT FIGURE"
Add-Image $s (Join-Path $figRoot "bayes_2d_boundary.png") 75 108 535 350
Add-Card $s 650 120 230 90 "边界含义" "黑色曲线是 gF(x)=gM(x) 的位置，两侧分别判为女生或男生。" $C.Teal
Add-Card $s 650 245 230 90 "为何非线性" "两类协方差不同，因此高斯 Bayes 形成二次判别边界。" $C.Blue
Add-Card $s 650 370 230 70 "分类错误来源" "边界附近样本重叠，是主要错误来源。" $C.Orange

$s = New-ContentSlide "主要分类结果总表" "RESULT TABLE"
Add-Table $s @(
  @("方法","设置","训练","test1","test2"),
  @("单特征 Bayes","只用身高，等先验","86.00%","94.29%","91.00%"),
  @("单特征 Bayes","只用体重，等先验","82.00%","94.29%","85.33%"),
  @("二维相关 Bayes","完整协方差，等先验","88.00%","97.14%","89.33%"),
  @("二维不相关 Bayes","对角协方差，等先验","88.00%","97.14%","90.33%"),
  @("最小风险 Bayes","自定义损失表","84.00%","94.29%","84.33%"),
  @("Parzen Bayes","h=5.0","87.00%","97.14%","90.33%"),
  @("kNN","k=5","89.00%","100.00%","90.33%")
) 38 105 885 318 $C.Teal | Out-Null
Add-Text $s "表中结果说明：二维特征整体更稳定，非参数方法与参数 Bayes 表现接近。" 70 452 820 18 12 $C.Ink $true "center" | Out-Null

$s = New-ContentSlide "单特征对比：身高优于体重" "RESULT ANALYSIS"
Add-Metric $s 95 125 180 70 "91.00%" "只用身高 test2" $C.Teal
Add-Metric $s 320 125 180 70 "85.33%" "只用体重 test2" $C.Blue
Add-Metric $s 545 125 180 70 "+5.67%" "身高相对提升" $C.Amber
Add-Panel $s 100 260 750 95 $C.PaleTeal (Rgb "A7F3D0") | Out-Null
Add-Text $s "解释：在该数据集中，男女性别差异在身高上更明显；体重受个体差异影响更大，因此单独使用体重分类效果较弱。" 130 292 690 34 14 $C.Teal $true "center" | Out-Null
Add-Bullets $s @("身高一维 Bayes 在 test2 上达到 91.00%。","体重一维 Bayes 在 test2 上为 85.33%。","这说明特征选择会显著影响模式分类性能。") 100 400 760 60 9

$s = New-ContentSlide "二维特征与协方差假设对比" "RESULT ANALYSIS"
Add-Table $s @(
  @("模型","训练准确率","test1准确率","test2准确率"),
  @("完整协方差 Bayes","88.00%","97.14%","89.33%"),
  @("对角协方差 Bayes","88.00%","97.14%","90.33%")
) 110 120 740 115 $C.Blue | Out-Null
Add-Card $s 120 285 300 95 "完整协方差" "能表达身高和体重相关性，模型更灵活，但参数估计方差较大。" $C.Teal
Add-Card $s 540 285 300 95 "对角协方差" "忽略相关性，参数更少，在 test2 上反而略好，说明小样本下更稳定。" $C.Blue
Add-Text $s "结论：模型复杂度并非越高越好，需要结合样本规模和测试表现判断。" 130 430 700 18 12 $C.Ink $true "center" | Out-Null

$s = New-ContentSlide "先验概率影响实验" "RESULT ANALYSIS"
Add-Table $s @(
  @("先验设置","训练准确率","test1准确率","test2准确率"),
  @("P(F)=0.50, P(M)=0.50","88.00%","97.14%","89.33%"),
  @("P(F)=0.75, P(M)=0.25","86.00%","85.71%","80.00%"),
  @("P(F)=0.90, P(M)=0.10","79.00%","80.00%","67.33%")
) 98 118 760 140 $C.Orange | Out-Null
Add-MiniBar $s 135 310 "等先验 test2 acc." 89.33 $C.Teal
Add-MiniBar $s 135 350 "女生先验 0.75 test2 acc." 80.00 $C.Blue
Add-MiniBar $s 135 390 "女生先验 0.90 test2 acc." 67.33 $C.Orange
Add-Card $s 570 315 260 80 "原因" "test2 中男生占多数，女生先验过高会导致男生被大量误判为女生。" $C.Orange

$s = New-ContentSlide "最小风险 Bayes：准确率不是唯一目标" "RESULT ANALYSIS"
Add-Table $s @(
  @("方法","训练","test1","test2"),
  @("最小错误率 Bayes","88.00%","97.14%","89.33%"),
  @("最小风险 Bayes","84.00%","94.29%","84.33%")
) 115 115 730 105 $C.Blue | Out-Null
Add-Panel $s 112 275 736 80 $C.PaleAmber (Rgb "FDE68A") | Out-Null
Add-Text $s "最小风险 Bayes 的目标不是最高准确率，而是降低加权错误代价。" 142 306 680 20 15 (Rgb "92400E") $true "center" | Out-Null
Add-Bullets $s @("当 λ(M|F)=2 时，女生误判为男生的代价更高。","分类器会更保守地判为男生，因此总体准确率可能下降。","这体现了实际任务中根据错误代价设计分类器的思想。") 112 400 736 65 9

$s = New-ContentSlide "非参数方法结果：Parzen 与 kNN" "RESULT ANALYSIS"
Add-Table $s @(
  @("方法","训练准确率","test1准确率","test2准确率"),
  @("Parzen Bayes, h=5.0","87.00%","97.14%","90.33%"),
  @("kNN, k=5","89.00%","100.00%","90.33%"),
  @("二维对角 Bayes","88.00%","97.14%","90.33%")
) 105 118 750 130 $C.Teal | Out-Null
Add-Card $s 115 305 300 90 "Parzen" "不依赖正态分布假设，可利用样本局部密度，但受窗口宽度 h 影响。" $C.Teal
Add-Card $s 545 305 300 90 "kNN" "直接利用最近邻投票，直观有效，但预测时需要计算与训练样本的距离。" $C.Blue
Add-Text $s "结论：在该二维数据集上，非参数方法能达到与参数 Bayes 接近的性能。" 135 445 690 18 12 $C.Ink $true "center" | Out-Null

$s = New-ContentSlide "test2 混淆矩阵：错误类型观察" "RESULT FIGURE"
Add-Image $s (Join-Path $figRoot "bayes_test2_confusion.png") 92 110 350 300
Add-Card $s 520 125 300 85 "结果解读" "二维完整协方差 Bayes 在 test2 上正确识别 49 个女生和 219 个男生。" $C.Teal
Add-Card $s 520 245 300 85 "错误类型" "女生误判为男生 1 个，男生误判为女生 31 个。" $C.Blue
Add-Card $s 520 365 300 85 "原因分析" "test2 男生样本更多，边界附近男生更容易被判入女生区域。" $C.Orange

Add-Section "04" "实验总结" "从结果回到模式识别理论"

$s = New-ContentSlide "任务一核心结论" "SUMMARY"
Add-Card $s 70 115 250 90 "特征选择" "身高单特征优于体重；身高+体重整体更稳定。" $C.Teal
Add-Card $s 355 115 250 90 "模型假设" "完整协方差表达力强，但对角协方差在 test2 上略好，说明小样本估计稳定性重要。" $C.Blue
Add-Card $s 640 115 230 90 "先验概率" "先验应符合真实分布或任务背景，设置不当会明显降低性能。" $C.Amber
Add-Card $s 180 280 250 90 "风险决策" "最小风险 Bayes 体现不同错误代价下的分类器设计思想。" $C.Orange
Add-Card $s 525 280 250 90 "非参数估计" "Parzen 和 kNN 可作为参数模型的重要补充。" $C.Green
Add-Text $s "总体体会：Bayes 分类器把概率密度估计、先验知识和决策代价统一到一个清晰的理论框架中。" 95 440 770 20 13 $C.Ink $true "center" | Out-Null

$s = Add-Slide $true
$bar = $s.Shapes.AddShape($msoShapeRectangle, 0, 0, 16, 540); $bar.Fill.ForeColor.RGB = $C.Teal; $bar.Line.Visible = $msoFalse
$bar2 = $s.Shapes.AddShape($msoShapeRectangle, 16, 0, 5, 540); $bar2.Fill.ForeColor.RGB = $C.Amber; $bar2.Line.Visible = $msoFalse
Add-Text $s "实验一学习完成" 70 190 520 55 34 $C.White $true | Out-Null
Add-Text $s "下一步可以继续完善实验二：非参数估计、Fisher 线性判别与留一法。" 74 270 650 28 15 (Rgb "CBD5E1") $false | Out-Null
$ln = $s.Shapes.AddLine(74, 326, 300, 326); $ln.Line.ForeColor.RGB = $C.Amber; $ln.Line.Weight = 4

if (Test-Path $Output) { Remove-Item -LiteralPath $Output -Force }
$pres.SaveAs($Output, $ppSaveAsOpenXMLPresentation)
$slideCount = $pres.Slides.Count
$pres.Close()
$app.Quit()

$app2 = New-Object -ComObject PowerPoint.Application
$app2.Visible = $msoTrue
$pres2 = $app2.Presentations.Open($Output, $msoTrue, $msoFalse, $msoFalse)
$verifiedSlides = $pres2.Slides.Count
$pres2.Export($previewDir, "PNG", 1280, 720)
$pres2.Close()
$app2.Quit()

Write-Output "FINAL_PPTX=$Output"
Write-Output "SLIDES=$verifiedSlides"
Write-Output "PREVIEW_DIR=$previewDir"
