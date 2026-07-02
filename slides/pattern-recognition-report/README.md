# 模式识别课程实验汇报 Slidev deck

本目录用于课堂汇报幻灯片。

## 文件说明

- `pattern-recognition-final-polished-editable.pptx`：PowerPoint 原生生成的精美可编辑最终版，推荐用于汇报和提交。
- `pattern-recognition-polished-editable.pptx`：PptxGenJS 生成的实验版，已不推荐优先使用。
- `pattern-recognition-editable.pptx`：早期可编辑草稿版，不优先推荐。
- `slides-export.pptx`：Slidev 导出的 PPTX，适合展示但可编辑性较弱。
- `slides-export.pdf`：Slidev 导出的 PDF 预览版。
- `slides.md`：Slidev 源文件。

## 常用命令

```powershell
npm.cmd install
npm.cmd run dev
npm.cmd run export:pptx
npm.cmd run export:pdf
node build-editable-pptx.mjs
PowerShell -ExecutionPolicy Bypass -File .\build-final-powerpoint.ps1
```

开发预览默认地址为 <http://127.0.0.1:3030>。

导出 PPTX/PDF 需要 `playwright-chromium`，已在 `package.json` 中声明。
