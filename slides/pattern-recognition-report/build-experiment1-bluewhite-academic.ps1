param(
  [string]$Output = "F:\Pattern-Recognition_npu\slides\pattern-recognition-report\experiment1-bayes-bluewhite-academic.pptx"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$figRoot = Join-Path $root "figures\experiment1"
$tmpRoot = Join-Path $env:TEMP "codex-presentations\experiment1-bluewhite-academic"
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
  Blue = Rgb "006BB7"; Deep = Rgb "004A82"; Sky = Rgb "3A9BD9"; Cloud = Rgb "E3F2FD"
  Gold = Rgb "D4A84B"; White = Rgb "FAFCFF"; Ink = Rgb "1A2E44"; Text = Rgb "333D4A"
  Muted = Rgb "6B7B8C"; Line = Rgb "C9DDEE"; Pale = Rgb "F4FAFF"; Green = Rgb "0F766E"
  Orange = Rgb "EA580C"
}

$msoFalse = 0
$msoTrue = -1
$ppLayoutBlank = 12
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeParallelogram = 2
$msoShapeOval = 9
$ppSaveAsOpenXMLPresentation = 24

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$pres = $app.Presentations.Add($msoTrue)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540
$script:Page = 1

function AddText($slide, [string]$text, [double]$x, [double]$y, [double]$w, [double]$h, [int]$size = 12, $color = $null, [bool]$bold = $false, [string]$align = "left") {
  if ($null -eq $color) { $color = $C.Text }
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

function AddBackground($slide, [bool]$dark = $false) {
  $bg = $slide.Shapes.AddShape($msoShapeRectangle, 0, 0, 960, 540)
  $bg.Fill.ForeColor.RGB = $(if ($dark) { $C.Deep } else { $C.White })
  $bg.Line.Visible = $msoFalse

  if (-not $dark) {
    $head = $slide.Shapes.AddShape($msoShapeParallelogram, -25, -22, 1040, 82)
    $head.Fill.ForeColor.RGB = $C.Blue
    $head.Line.Visible = $msoFalse
    $head.Adjustments.Item(1) = 0.16
    $accent = $slide.Shapes.AddShape($msoShapeParallelogram, 705, -8, 310, 74)
    $accent.Fill.ForeColor.RGB = $C.Sky
    $accent.Fill.Transparency = 0.2
    $accent.Line.Visible = $msoFalse
    $gold = $slide.Shapes.AddShape($msoShapeRectangle, 52, 74, 92, 4)
    $gold.Fill.ForeColor.RGB = $C.Gold
    $gold.Line.Visible = $msoFalse
    $wave = $slide.Shapes.AddShape($msoShapeParallelogram, -90, 500, 520, 46)
    $wave.Fill.ForeColor.RGB = $C.Cloud
    $wave.Line.Visible = $msoFalse
    $wave.Adjustments.Item(1) = 0.28
  }
}

function AddFooter($slide) {
  AddText $slide "模式识别课程实验一 · Bayes 性别分类器" 54 511 360 12 7 $C.Muted $false | Out-Null
  AddText $slide ("{0:D2}" -f $script:Page) 888 507 32 14 8 $C.Blue $true "right" | Out-Null
  $script:Page += 1
}

function AddSlide($title, $kicker = "EXPERIMENT 1") {
  $slide = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)
  AddBackground $slide $false
  AddText $slide $kicker 58 18 360 16 7 $C.White $true | Out-Null
  AddText $slide $title 58 38 700 28 20 $C.White $true | Out-Null
  AddFooter $slide
  return $slide
}

function AddPanel($slide, $x, $y, $w, $h, $fill = $null, $line = $null) {
  if ($null -eq $fill) { $fill = $C.White }
  if ($null -eq $line) { $line = $C.Line }
  $p = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $x, $y, $w, $h)
  $p.Fill.ForeColor.RGB = $fill
  $p.Line.ForeColor.RGB = $line
  $p.Line.Weight = 0.8
  return $p
}

function AddCard($slide, $x, $y, $w, $h, $head, $body, $accent = $null, $fill = $null) {
  if ($null -eq $accent) { $accent = $C.Blue }
  if ($null -eq $fill) { $fill = $C.White }
  AddPanel $slide $x $y $w $h $fill | Out-Null
  $bar = $slide.Shapes.AddShape($msoShapeRectangle, $x, $y, 5, $h)
  $bar.Fill.ForeColor.RGB = $accent
  $bar.Line.Visible = $msoFalse
  AddText $slide $head ($x+16) ($y+10) ($w-26) 16 10 $C.Ink $true | Out-Null
  AddText $slide $body ($x+16) ($y+34) ($w-26) ($h-38) 8 $C.Text $false | Out-Null
}

function AddMetric($slide, $x, $y, $w, $h, $value, $label, $accent = $null) {
  if ($null -eq $accent) { $accent = $C.Blue }
  AddPanel $slide $x $y $w $h $C.Pale | Out-Null
  AddText $slide $value ($x+12) ($y+8) ($w-20) 22 18 $accent $true | Out-Null
  AddText $slide $label ($x+12) ($y+39) ($w-20) 14 7 $C.Muted $false | Out-Null
}

function AddBullets($slide, [string[]]$items, $x, $y, $w, $h, $size = 9) {
  $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $x, $y, $w, $h)
  $tr = $shape.TextFrame.TextRange
  $tr.Text = [string]::Join("`r", $items)
  $tr.Font.Name = "Microsoft YaHei"
  $tr.Font.Size = $size
  $tr.Font.Color.RGB = $C.Text
  $tr.ParagraphFormat.Bullet.Visible = $msoTrue
  $tr.ParagraphFormat.SpaceAfter = 5
  $shape.TextFrame.MarginLeft = 8
  $shape.TextFrame.MarginRight = 2
  $shape.TextFrame.MarginTop = 2
  $shape.TextFrame.MarginBottom = 2
}

function AddTable($slide, $rows, $x, $y, $w, $h, $headerColor = $null) {
  if ($null -eq $headerColor) { $headerColor = $C.Blue }
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
}

function AddImage($slide, $file, $x, $y, $w, $h) {
  if (-not (Test-Path $file)) { throw "Missing image $file" }
  AddPanel $slide $x $y $w $h $C.White | Out-Null
  $pic = $slide.Shapes.AddPicture($file, $msoFalse, $msoTrue, $x+5, $y+5, -1, -1)
  $pic.LockAspectRatio = $msoTrue
  $scale = [Math]::Min(($w-10)/$pic.Width, ($h-10)/$pic.Height)
  $pic.Width = $pic.Width * $scale
  $pic.Height = $pic.Height * $scale
  $pic.Left = $x + ($w - $pic.Width) / 2
  $pic.Top = $y + ($h - $pic.Height) / 2
}

function AddStep($slide, $x, $y, $num, $head, $body, $color = $null) {
  if ($null -eq $color) { $color = $C.Blue }
  $oval = $slide.Shapes.AddShape($msoShapeOval, $x, $y, 30, 30)
  $oval.Fill.ForeColor.RGB = $color
  $oval.Line.Visible = $msoFalse
  AddText $slide $num ($x+1) ($y+7) 28 12 8 $C.White $true "center" | Out-Null
  AddText $slide $head ($x+42) ($y-1) 280 16 10 $C.Ink $true | Out-Null
  AddText $slide $body ($x+42) ($y+20) 280 30 8 $C.Text $false | Out-Null
}

function AddSection($num, $title, $subtitle) {
  $s = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)
  AddBackground $s $true
  $diag = $s.Shapes.AddShape($msoShapeParallelogram, 585, -20, 430, 160)
  $diag.Fill.ForeColor.RGB = $C.Sky
  $diag.Fill.Transparency = 0.15
  $diag.Line.Visible = $msoFalse
  $diag.Adjustments.Item(1) = 0.22
  AddText $s $num 70 95 130 42 22 $C.Gold $true | Out-Null
  AddText $s $title 70 165 680 58 30 $C.White $true | Out-Null
  AddText $s $subtitle 74 236 680 28 13 (Rgb "D9ECFF") $false | Out-Null
  $ln = $s.Shapes.AddLine(74, 300, 250, 300)
  $ln.Line.ForeColor.RGB = $C.Gold
  $ln.Line.Weight = 4
}

# Cover
$s = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)
AddBackground $s $true
$diag = $s.Shapes.AddShape($msoShapeParallelogram, -80, 345, 520, 230)
$diag.Fill.ForeColor.RGB = $C.Blue
$diag.Line.Visible = $msoFalse
$diag.Adjustments.Item(1) = 0.2
$diag2 = $s.Shapes.AddShape($msoShapeParallelogram, 635, -20, 420, 155)
$diag2.Fill.ForeColor.RGB = $C.Sky
$diag2.Fill.Transparency = 0.2
$diag2.Line.Visible = $msoFalse
$diag2.Adjustments.Item(1) = 0.18
AddText $s "PATTERN RECOGNITION · EXPERIMENT 1" 72 96 420 18 8 (Rgb "BFE8FF") $true | Out-Null
AddText $s "Bayes 性别分类器实验" 72 156 610 56 34 $C.White $true | Out-Null
AddText $s "理论原理 · 实验设计 · 软件实现 · 实验结果分析" 76 228 640 26 14 (Rgb "D9ECFF") $false | Out-Null
AddPanel $s 72 352 205 54 (Rgb "0E3556") (Rgb "75BDEB") | Out-Null
AddText $s "姓名" 90 363 160 12 8 (Rgb "D9ECFF") $true | Out-Null
AddText $s "请填写" 90 383 160 12 8 $C.White $false | Out-Null
AddPanel $s 322 352 205 54 (Rgb "0E3556") (Rgb "75BDEB") | Out-Null
AddText $s "学号" 340 363 160 12 8 (Rgb "D9ECFF") $true | Out-Null
AddText $s "请填写" 340 383 160 12 8 $C.White $false | Out-Null
AddPanel $s 572 352 205 54 (Rgb "0E3556") (Rgb "75BDEB") | Out-Null
AddText $s "班级" 590 363 160 12 8 (Rgb "D9ECFF") $true | Out-Null
AddText $s "请填写" 590 383 160 12 8 $C.White $false | Out-Null

$s = AddSlide "汇报逻辑：从理论到结果分析" "ROADMAP"
AddCard $s 70 130 190 85 "1 理论原理" "Bayes 决策、正态假设、参数估计、最小风险与非参数估计。" $C.Blue
AddCard $s 290 130 190 85 "2 实验设计" "数据集、特征组合、先验设置、风险表和评价指标。" $C.Sky
AddCard $s 510 130 190 85 "3 软件实现" "读取数据、训练模型、分类预测、绘图和结果保存。" $C.Gold
AddCard $s 730 130 150 85 "4 结果分析" "准确率、混淆矩阵、决策边界和误差来源。" $C.Green
AddPanel $s 118 300 724 92 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "核心问题：如何根据身高/体重特征 x，判断样本属于女生 F 还是男生 M？" 146 332 670 20 15 $C.Deep $true "center" | Out-Null
AddBullets $s @("这是一个典型的两类模式分类问题。","任务一要求同时考察特征、先验、协方差假设、风险决策、参数估计和非参数估计。") 112 422 730 52 9

AddSection "01" "理论原理" "Bayes 决策、参数估计、风险函数与非参数方法"

$s = AddSlide "Bayes 决策：从先验到后验" "THEORY"
AddCard $s 70 122 250 88 "先验概率 P(ωi)" "观察样本前对类别出现概率的估计，例如男女先验相等。" $C.Blue
AddCard $s 355 122 250 88 "类条件密度 p(x|ωi)" "给定类别后，身高/体重特征出现的概率密度。" $C.Sky
AddCard $s 640 122 230 88 "后验概率 P(ωi|x)" "观察到特征 x 后，样本属于某类别的概率。" $C.Gold
AddPanel $s 100 275 760 74 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "ω* = argmax P(ωi|x) = argmax p(x|ωi)P(ωi)" 130 302 700 18 17 $C.Deep $true "center" | Out-Null
AddBullets $s @("p(x) 对所有类别相同，因此分类时只比较 p(x|ωi)P(ωi)。","Bayes 分类器是在概率意义下选择最可能的类别。") 108 398 740 52 9

$s = AddSlide "正态分布假设与最大似然估计" "THEORY"
AddCard $s 62 120 250 92 "一维特征" "只使用身高或体重时，假设每类服从一维正态分布 N(μi, σi²)。" $C.Blue
AddCard $s 356 120 250 92 "二维特征" "同时使用身高和体重时，假设每类服从二维高斯分布 N(μi, Σi)。" $C.Sky
AddCard $s 650 120 230 92 "最大似然估计" "用训练样本估计每个类别的均值、方差或协方差矩阵。" $C.Gold
AddPanel $s 95 288 360 88 (Rgb "EEF7FF") (Rgb "B6DDF5") | Out-Null
AddText $s "μi = (1/Ni) Σ xk" 115 315 320 18 15 $C.Blue $true "center" | Out-Null
AddText $s "Σi = (1/Ni) Σ (xk-μi)(xk-μi)^T" 115 345 320 18 15 $C.Blue $true "center" | Out-Null
AddPanel $s 510 288 330 88 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "估计女生和男生各自的分布参数，再构造分类边界。" 540 318 270 32 12 $C.Deep $true "center" | Out-Null

$s = AddSlide "二维高斯 Bayes 判别函数" "THEORY"
AddPanel $s 72 120 815 92 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "gi(x) = -1/2 ln|Σi| - 1/2(x-μi)^TΣi⁻¹(x-μi) + ln P(ωi)" 100 154 760 20 15 $C.Deep $true "center" | Out-Null
AddCard $s 95 270 225 84 "均值 μi" "决定类别中心位置。男生均值整体高于女生均值。" $C.Blue
AddCard $s 368 270 225 84 "协方差 Σi" "决定类别分布形状和身高体重是否相关。" $C.Sky
AddCard $s 640 270 220 84 "先验 P(ωi)" "决定分类边界向哪个类别偏移。" $C.Gold
AddText $s "决策规则：若 gF(x) ≥ gM(x)，判为女生；否则判为男生。" 130 420 700 18 12 $C.Ink $true "center" | Out-Null

$s = AddSlide "协方差假设与先验概率" "THEORY"
AddCard $s 70 116 360 100 "完整协方差：考虑相关" "保留协方差矩阵非对角元素，表示身高和体重之间可能有关联，判别边界更灵活。" $C.Blue
AddCard $s 530 116 330 100 "对角协方差：忽略相关" "只保留方差项，假设特征相互独立；参数更少，小样本估计更稳定。" $C.Sky
AddTable $s @(
  @("先验设置","含义","预期影响"),
  @("0.5 / 0.5","男女先验相等","决策主要由似然决定"),
  @("0.75 / 0.25","女生先验较大","边界向男生侧偏移"),
  @("0.9 / 0.1","女生先验极大","更容易把样本判为女生")
) 110 280 740 128 $C.Blue
AddText $s "先验概率应来自真实类别比例或任务背景，不能随意设置。" 130 444 700 18 12 $C.Deep $true "center" | Out-Null

$s = AddSlide "最小风险 Bayes 与非参数估计" "THEORY"
AddPanel $s 80 116 780 60 (Rgb "EEF7FF") (Rgb "B6DDF5") | Out-Null
AddText $s "R(αi|x) = Σ λ(αi|ωj) P(ωj|x)" 120 138 700 18 16 $C.Blue $true "center" | Out-Null
AddTable $s @(
  @("内容","本实验设置","意义"),
  @("损失表","λ(M|F)=2, λ(F|M)=1","女生误判为男生代价更高"),
  @("Parzen 窗","Gaussian kernel, h=5.0","不预设正态分布，用核密度估计"),
  @("kNN","k=5","根据最近邻样本投票")
) 100 230 760 125 $C.Blue
AddBullets $s @("最小风险 Bayes 的目标是降低总代价，不一定追求最高准确率。","非参数方法不依赖固定分布假设，可作为参数 Bayes 的补充。") 105 405 740 55 9

AddSection "02" "实验设计与软件实现" "数据、特征、流程和评价指标"

$s = AddSlide "数据集与样本划分" "EXPERIMENT DESIGN"
AddTable $s @(
  @("文件","用途","样本说明"),
  @("FEMALE.TXT","训练集","50 个女生的身高、体重"),
  @("MALE.TXT","训练集","50 个男生的身高、体重"),
  @("test1.txt","测试集","35 个样本，15 女 20 男"),
  @("test2.txt","测试集","300 个样本，50 女 250 男")
) 92 118 776 180 $C.Blue
AddMetric $s 112 338 155 55 "100" "训练样本" $C.Blue
AddMetric $s 318 338 155 55 "35" "test1 样本" $C.Sky
AddMetric $s 524 338 155 55 "300" "test2 样本" $C.Gold
AddMetric $s 730 338 155 55 "2" "类别数量" $C.Green

$s = AddSlide "实验分组设计" "EXPERIMENT DESIGN"
AddCard $s 70 115 248 86 "单特征实验" "分别只用身高或体重，建立一维高斯 Bayes 分类器。" $C.Blue
AddCard $s 356 115 248 86 "双特征实验" "同时用身高和体重，比较完整协方差和对角协方差。" $C.Sky
AddCard $s 642 115 228 86 "先验概率实验" "比较 0.5/0.5、0.75/0.25、0.9/0.1。" $C.Gold
AddCard $s 70 260 248 86 "最小风险实验" "给定损失表，按照条件风险最小原则分类。" $C.Orange
AddCard $s 356 260 248 86 "参数估计实例" "使用最大似然估计计算均值和协方差。" $C.Green
AddCard $s 642 260 228 86 "非参数估计实例" "实现 Parzen 窗 Bayes 和 kNN 进行对比。" $C.Blue
AddText $s "实验设计覆盖任务一要求中的特征选择、先验影响、风险决策、参数估计和非参数估计。" 100 430 760 18 12 $C.Deep $true "center" | Out-Null

$s = AddSlide "软件实现流程" "SOFTWARE IMPLEMENTATION"
AddStep $s 85 115 "1" "读取数据" "解析训练文件和测试文件，统一标签为 F/M。" $C.Blue
AddStep $s 85 195 "2" "选择特征" "按实验选择身高、体重或二维特征。" $C.Sky
AddStep $s 85 275 "3" "训练模型" "估计均值、方差/协方差，或保存非参数样本。" $C.Gold
AddStep $s 520 115 "4" "进行预测" "计算判别函数、风险函数或近邻投票。" $C.Orange
AddStep $s 520 195 "5" "评价性能" "输出 accuracy、precision、recall、confusion matrix。" $C.Green
AddStep $s 520 275 "6" "生成图表" "绘制散点图、决策边界和混淆矩阵。" $C.Blue
AddCard $s 180 392 600 58 "对应脚本" "scripts/experiment1_bayes.py：包含 fit_gaussian、predict_gaussian、predict_min_risk、predict_parzen、predict_knn 等核心函数。" $C.Blue

$s = AddSlide "评价指标：不仅看准确率" "EVALUATION"
AddCard $s 70 120 250 84 "Accuracy" "预测正确样本数 / 总样本数，用于整体性能比较。" $C.Blue
AddCard $s 355 120 250 84 "Precision" "预测为某类的样本中，真正属于该类的比例。" $C.Sky
AddCard $s 640 120 230 84 "Recall" "真实属于某类的样本中，被正确找出的比例。" $C.Gold
AddPanel $s 112 282 736 100 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "混淆矩阵用于观察女生判错为男生、男生判错为女生的具体数量；比单一准确率更能解释错误类型。" 150 318 660 35 13 $C.Deep $true "center" | Out-Null

AddSection "03" "实验结果与分析" "参数估计、图表结果、分类性能和现象解释"

$s = AddSlide "训练集估计参数" "PARAMETER ESTIMATION"
AddTable $s @(
  @("类别","均值向量 μ","协方差矩阵 Σ"),
  @("女生 F","[162.84, 52.596]","[[43.934, 15.525], [15.525, 31.129]]"),
  @("男生 M","[173.92, 65.502]","[[20.754, 23.058], [23.058, 59.898]]")
) 70 124 820 140 $C.Blue
AddBullets $s @("男生平均身高和平均体重均高于女生，这是分类的主要依据。","协方差矩阵非对角元素不为 0，说明身高和体重存在一定相关性。","协方差估计是否可靠，会影响完整协方差 Bayes 的测试表现。") 96 325 760 85 9

$s = AddSlide "样本分布：身高与体重的可分性" "RESULT FIGURE"
AddImage $s (Join-Path $figRoot "data_scatter.png") 72 112 520 340
AddCard $s 640 132 230 78 "观察 1" "男生样本整体位于更高身高、更高体重区域。" $C.Blue
AddCard $s 640 240 230 78 "观察 2" "两类样本在边界区域存在重叠，因此无法完全无误分类。" $C.Sky
AddCard $s 640 348 230 78 "观察 3" "test1 样本基本落在训练集分布附近，测试准确率较高。" $C.Gold

$s = AddSlide "二维 Bayes 决策边界" "RESULT FIGURE"
AddImage $s (Join-Path $figRoot "bayes_2d_boundary.png") 72 108 535 350
AddCard $s 650 120 230 82 "边界含义" "黑色曲线是 gF(x)=gM(x) 的位置，两侧分别判为女生或男生。" $C.Blue
AddCard $s 650 235 230 82 "为何非线性" "两类协方差不同，因此高斯 Bayes 形成二次判别边界。" $C.Sky
AddCard $s 650 350 230 70 "错误来源" "边界附近样本重叠，是主要错误来源。" $C.Orange

$s = AddSlide "主要分类结果总表" "RESULT TABLE"
AddTable $s @(
  @("方法","设置","训练","test1","test2"),
  @("单特征 Bayes","只用身高，等先验","86.00%","94.29%","91.00%"),
  @("单特征 Bayes","只用体重，等先验","82.00%","94.29%","85.33%"),
  @("二维相关 Bayes","完整协方差，等先验","88.00%","97.14%","89.33%"),
  @("二维不相关 Bayes","对角协方差，等先验","88.00%","97.14%","90.33%"),
  @("最小风险 Bayes","自定义损失表","84.00%","94.29%","84.33%"),
  @("Parzen Bayes","h=5.0","87.00%","97.14%","90.33%"),
  @("kNN","k=5","89.00%","100.00%","90.33%")
) 42 110 875 308 $C.Blue
AddText $s "表中结果说明：二维特征整体更稳定，非参数方法与参数 Bayes 表现接近。" 70 452 820 18 12 $C.Deep $true "center" | Out-Null

$s = AddSlide "单特征对比：身高优于体重" "RESULT ANALYSIS"
AddMetric $s 105 126 180 68 "91.00%" "只用身高 test2" $C.Blue
AddMetric $s 340 126 180 68 "85.33%" "只用体重 test2" $C.Sky
AddMetric $s 575 126 180 68 "+5.67%" "身高相对提升" $C.Gold
AddPanel $s 110 265 740 92 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "解释：在该数据集中，男女性别差异在身高上更明显；体重受个体差异影响更大，因此单独使用体重分类效果较弱。" 140 297 680 34 13 $C.Deep $true "center" | Out-Null
AddBullets $s @("身高一维 Bayes 在 test2 上达到 91.00%。","体重一维 Bayes 在 test2 上为 85.33%。","这说明特征选择会显著影响分类性能。") 110 405 740 58 9

$s = AddSlide "二维特征与协方差假设对比" "RESULT ANALYSIS"
AddTable $s @(
  @("模型","训练准确率","test1准确率","test2准确率"),
  @("完整协方差 Bayes","88.00%","97.14%","89.33%"),
  @("对角协方差 Bayes","88.00%","97.14%","90.33%")
) 110 125 740 112 $C.Blue
AddCard $s 120 290 300 88 "完整协方差" "能表达身高和体重相关性，模型更灵活，但参数估计方差较大。" $C.Blue
AddCard $s 540 290 300 88 "对角协方差" "忽略相关性，参数更少，在 test2 上反而略好，说明小样本下更稳定。" $C.Sky
AddText $s "结论：模型复杂度并非越高越好，需要结合样本规模和测试表现判断。" 130 430 700 18 12 $C.Deep $true "center" | Out-Null

$s = AddSlide "先验概率影响实验" "RESULT ANALYSIS"
AddTable $s @(
  @("先验设置","训练准确率","test1准确率","test2准确率"),
  @("P(F)=0.50, P(M)=0.50","88.00%","97.14%","89.33%"),
  @("P(F)=0.75, P(M)=0.25","86.00%","85.71%","80.00%"),
  @("P(F)=0.90, P(M)=0.10","79.00%","80.00%","67.33%")
) 98 120 760 135 $C.Blue
AddMetric $s 122 315 150 58 "89.33%" "等先验 test2" $C.Blue
AddMetric $s 322 315 150 58 "80.00%" "女生先验 0.75" $C.Sky
AddMetric $s 522 315 150 58 "67.33%" "女生先验 0.90" $C.Orange
AddCard $s 700 306 170 84 "原因" "test2 中男生占多数，女生先验过高会导致男生被大量误判为女生。" $C.Orange

$s = AddSlide "最小风险 Bayes：准确率不是唯一目标" "RESULT ANALYSIS"
AddTable $s @(
  @("方法","训练","test1","test2"),
  @("最小错误率 Bayes","88.00%","97.14%","89.33%"),
  @("最小风险 Bayes","84.00%","94.29%","84.33%")
) 115 120 730 105 $C.Blue
AddPanel $s 112 275 736 78 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "最小风险 Bayes 的目标不是最高准确率，而是降低加权错误代价。" 142 306 680 20 14 $C.Deep $true "center" | Out-Null
AddBullets $s @("当 λ(M|F)=2 时，女生误判为男生的代价更高。","分类器会更保守地判为男生，因此总体准确率可能下降。","这体现了实际任务中根据错误代价设计分类器的思想。") 112 400 736 65 9

$s = AddSlide "非参数方法结果：Parzen 与 kNN" "RESULT ANALYSIS"
AddTable $s @(
  @("方法","训练准确率","test1准确率","test2准确率"),
  @("Parzen Bayes, h=5.0","87.00%","97.14%","90.33%"),
  @("kNN, k=5","89.00%","100.00%","90.33%"),
  @("二维对角 Bayes","88.00%","97.14%","90.33%")
) 105 120 750 128 $C.Blue
AddCard $s 115 305 300 88 "Parzen" "不依赖正态分布假设，可利用样本局部密度，但受窗口宽度 h 影响。" $C.Blue
AddCard $s 545 305 300 88 "kNN" "直接利用最近邻投票，直观有效，但预测时需要计算与训练样本的距离。" $C.Sky
AddText $s "结论：在该二维数据集上，非参数方法能达到与参数 Bayes 接近的性能。" 135 445 690 18 12 $C.Deep $true "center" | Out-Null

$s = AddSlide "test2 混淆矩阵：错误类型观察" "RESULT FIGURE"
AddImage $s (Join-Path $figRoot "bayes_test2_confusion.png") 95 110 350 300
AddCard $s 525 130 300 76 "结果解读" "二维完整协方差 Bayes 在 test2 上正确识别 49 个女生和 219 个男生。" $C.Blue
AddCard $s 525 240 300 76 "错误类型" "女生误判为男生 1 个，男生误判为女生 31 个。" $C.Sky
AddCard $s 525 350 300 76 "原因分析" "test2 男生样本更多，边界附近男生更容易被判入女生区域。" $C.Orange

AddSection "04" "实验总结" "从结果回到模式识别理论"

$s = AddSlide "任务一核心结论" "SUMMARY"
AddCard $s 70 120 250 84 "特征选择" "身高单特征优于体重；身高+体重整体更稳定。" $C.Blue
AddCard $s 355 120 250 84 "模型假设" "完整协方差表达力强，但对角协方差在 test2 上略好，说明小样本估计稳定性重要。" $C.Sky
AddCard $s 640 120 230 84 "先验概率" "先验应符合真实分布或任务背景，设置不当会明显降低性能。" $C.Gold
AddCard $s 180 280 250 84 "风险决策" "最小风险 Bayes 体现不同错误代价下的分类器设计思想。" $C.Orange
AddCard $s 525 280 250 84 "非参数估计" "Parzen 和 kNN 可作为参数模型的重要补充。" $C.Green
AddPanel $s 110 430 740 42 $C.Cloud (Rgb "B6DDF5") | Out-Null
AddText $s "总体体会：Bayes 分类器把概率密度估计、先验知识和决策代价统一到一个清晰的理论框架中。" 128 443 704 14 11 $C.Deep $true "center" | Out-Null

$s = $pres.Slides.Add($pres.Slides.Count + 1, $ppLayoutBlank)
AddBackground $s $true
$diag = $s.Shapes.AddShape($msoShapeParallelogram, -80, 345, 520, 230)
$diag.Fill.ForeColor.RGB = $C.Blue
$diag.Line.Visible = $msoFalse
$diag.Adjustments.Item(1) = 0.2
AddText $s "谢谢" 78 190 300 54 40 $C.White $true | Out-Null
AddText $s "实验一学习完成，下一步可继续完善实验二。" 82 270 600 28 15 (Rgb "D9ECFF") $false | Out-Null
$ln = $s.Shapes.AddLine(82, 326, 275, 326); $ln.Line.ForeColor.RGB = $C.Gold; $ln.Line.Weight = 4

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
