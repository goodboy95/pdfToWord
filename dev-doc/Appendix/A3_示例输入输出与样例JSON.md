````markdown
# A3_示例输入输出与样例JSON

> 本附录提供“从用户输入 → 任务参数快照 → 中间产物（IR/表格结构/请求与响应 JSON）→ 输出 docx”的端到端示例。  
> 目的：  
> - 便于开发联调（各模块按同一 JSON 协议对接）  
> - 便于测试回归（用固定样例对比 IR 快照）  
> - 便于诊断包复现问题（定位是表格结构还是 Gemini OCR）  
>
> 注意：示例内容为虚构文本，不代表真实业务数据。

---

## A3.1 示例 1：典型任务（含页码范围、去页眉页脚、跟随 PDF）

### A3.1.1 用户输入（UI）
- 输入 PDF：`C:\Docs\report_scan.pdf`
- 页码范围：`1,3,5-10`
- 去页眉页脚：`RemoveBoth`
- 页眉裁切：`6%`
- 页脚裁切：`6%`
- 页面尺寸：`FollowPdf`
- DPI：`300`
- deskew：开启
- 页并发：`2`
- Gemini 并发：`2`
- 输出目录：`C:\Docs\output\`

### A3.1.2 任务参数快照（meta.json / DocumentMeta.Options）
```json
{
  "jobId": "20251227_001",
  "sourcePdfPath": "C:\\Docs\\report_scan.pdf",
  "generatedAtUtc": "2025-12-27T03:00:00Z",
  "options": {
    "dpi": 300,
    "pageRange": "1,3,5-10",
    "headerFooterMode": "RemoveBoth",
    "headerPercent": 0.06,
    "footerPercent": 0.06,
    "pageSizeMode": "FollowPdf",
    "pageConcurrency": 2,
    "geminiConcurrency": 2,
    "enableDeskew": true
  },
  "output": {
    "outputDirectory": "C:\\Docs\\output\\",
    "outputDocxPath": "C:\\Docs\\output\\report_scan.converted.docx"
  }
}
````

### A3.1.3 解析后的页码列表（内部）

```json
{
  "totalPages": 12,
  "selectedPages": [1,3,5,6,7,8,9,10],
  "warnings": []
}
```

---

## A3.2 示例 2：表格检测输出（table_detect.json）

> 每页可能有 0..N 个表格 bbox。此文件用于诊断包与预览叠加。

```json
{
  "pageNumber": 1,
  "pageSizePx": {"w": 2480, "h": 3088},
  "tables": [
    {
      "tableIndex": 0,
      "bboxInPage": {"x": 120, "y": 480, "w": 2240, "h": 900},
      "structureMeta": {"engine": "OpenCV", "lineScore": 0.82}
    }
  ]
}
```

---

## A3.3 示例 3：表格网格与单元格结构输出（p001_t00_grid.json）

> 该文件对应《07》的输出，用于 owner 校验与 Gemini 回填。
> 坐标 `bboxInTable` 基于表格裁剪图（左上为 0,0）。

```json
{
  "pageNumber": 1,
  "tableIndex": 0,
  "tableSizePx": {"w": 2240, "h": 900},
  "bboxInPage": {"x": 120, "y": 480, "w": 2240, "h": 900},
  "nCols": 4,
  "nRows": 3,
  "colLinesX": [0, 560, 1120, 1680, 2239],
  "rowLinesY": [0, 260, 560, 899],
  "cells": [
    {"id":"p1_t0_r0_c0","row":0,"col":0,"rowspan":1,"colspan":1,"bboxInTable":{"x":0,"y":0,"w":560,"h":260}},
    {"id":"p1_t0_r0_c1","row":0,"col":1,"rowspan":1,"colspan":1,"bboxInTable":{"x":560,"y":0,"w":560,"h":260}},
    {"id":"p1_t0_r0_c2","row":0,"col":2,"rowspan":1,"colspan":1,"bboxInTable":{"x":1120,"y":0,"w":560,"h":260}},
    {"id":"p1_t0_r0_c3","row":0,"col":3,"rowspan":1,"colspan":1,"bboxInTable":{"x":1680,"y":0,"w":559,"h":260}},

    {"id":"p1_t0_r1_c0","row":1,"col":0,"rowspan":2,"colspan":1,"bboxInTable":{"x":0,"y":260,"w":560,"h":639}},
    {"id":"p1_t0_r1_c1","row":1,"col":1,"rowspan":1,"colspan":3,"bboxInTable":{"x":560,"y":260,"w":1679,"h":300}},

    {"id":"p1_t0_r2_c1","row":2,"col":1,"rowspan":1,"colspan":1,"bboxInTable":{"x":560,"y":560,"w":560,"h":339}},
    {"id":"p1_t0_r2_c2","row":2,"col":2,"rowspan":1,"colspan":1,"bboxInTable":{"x":1120,"y":560,"w":560,"h":339}},
    {"id":"p1_t0_r2_c3","row":2,"col":3,"rowspan":1,"colspan":1,"bboxInTable":{"x":1680,"y":560,"w":559,"h":339}}
  ],
  "validation": {
    "ownerMatrixOk": true,
    "conflicts": []
  }
}
```

> 说明：上例展示了常见合并：
>
> * `p1_t0_r1_c0`：rowspan=2（竖向合并）
> * `p1_t0_r1_c1`：colspan=3（横向合并）
>   其余被合并覆盖的格子不再作为主 cell 出现在 cells 列表（实现层可只列主 cell）。

---

## A3.4 示例 4：Gemini 表格 OCR 请求与响应（T1）

### A3.4.1 请求 payload（概念示例）

> 实际 API 结构取决于你使用的 Gemini SDK/REST，本处只强调“图片 + cells 列表 + 严格 JSON 输出”。

**Prompt 中嵌入的 cells JSON：**

```json
{
  "cells": [
    {"id":"p1_t0_r0_c0","x":0,"y":0,"w":560,"h":260},
    {"id":"p1_t0_r0_c1","x":560,"y":0,"w":560,"h":260},
    {"id":"p1_t0_r0_c2","x":1120,"y":0,"w":560,"h":260},
    {"id":"p1_t0_r0_c3","x":1680,"y":0,"w":559,"h":260},

    {"id":"p1_t0_r1_c0","x":0,"y":260,"w":560,"h":639},
    {"id":"p1_t0_r1_c1","x":560,"y":260,"w":1679,"h":300},

    {"id":"p1_t0_r2_c1","x":560,"y":560,"w":560,"h":339},
    {"id":"p1_t0_r2_c2","x":1120,"y":560,"w":560,"h":339},
    {"id":"p1_t0_r2_c3","x":1680,"y":560,"w":559,"h":339}
  ]
}
```

### A3.4.2 响应 JSON（必须严格）

```json
{
  "cells": [
    {"id":"p1_t0_r0_c0","text":"项目"},
    {"id":"p1_t0_r0_c1","text":"负责人"},
    {"id":"p1_t0_r0_c2","text":"部门"},
    {"id":"p1_t0_r0_c3","text":"日期"},

    {"id":"p1_t0_r1_c0","text":"A-001"},
    {"id":"p1_t0_r1_c1","text":"张三 / 研发一部 / 2025-12-01"},

    {"id":"p1_t0_r2_c1","text":"李四"},
    {"id":"p1_t0_r2_c2","text":"研发二部"},
    {"id":"p1_t0_r2_c3","text":"2025-12-02"}
  ]
}
```

### A3.4.3 校验失败示例（缺 id）

```json
{
  "cells": [
    {"id":"p1_t0_r0_c0","text":"项目"}
  ]
}
```

对应处理：

* 触发 `E_GEMINI_OCR_MISSING_IDS` → 重试一次（更强 prompt）→ 仍失败则 D1（保留结构但空单元格）或 D2（纯文本表格）。

---

## A3.5 示例 5：Gemini 正文 OCR 请求与响应（T2）

### A3.5.1 输入说明

* 输入图片为“遮罩表格后的整页图”（表格区域填白/马赛克）
* prompt 强调忽略空白遮罩区域，只输出正文

### A3.5.2 响应 JSON（严格）

```json
{
  "paragraphs": [
    {"role":"title","text":"项目周报"},
    {"role":"body","text":"一、总体进展\n本周完成了需求评审与原型确认。"},
    {"role":"body","text":"二、风险与问题\n1. 供应商接口变更频繁\n2. 测试资源排期紧张"}
  ]
}
```

### A3.5.3 降级响应（单段文本）

当 `paragraphs` 为空或不稳定时，可切到降级协议：

```json
{
  "text": "项目周报\n一、总体进展\n本周完成了需求评审与原型确认。\n二、风险与问题\n1. 供应商接口变更频繁\n2. 测试资源排期紧张"
}
```

---

## A3.6 示例 6：页 IR（p001_ir.json）

> 页 IR 用于回归测试与 docx 写入前检查。块顺序按阅读顺序排列。

```json
{
  "pageNumber": 1,
  "originalWidthPx": 2480,
  "originalHeightPx": 3508,
  "widthPx": 2480,
  "heightPx": 3088,
  "crop": {"mode":"RemoveBoth","cropTopPx":210,"cropBottomPx":210},
  "blocks": [
    {
      "type":"paragraph",
      "role":"title",
      "text":"项目周报",
      "source":{"producer":"GeminiText","attempt":1}
    },
    {
      "type":"table",
      "tableBBox":{"x":120,"y":480,"w":2240,"h":900},
      "nCols":4,
      "rows":[
        {"cells":[
          {"text":"项目","rowspan":1,"colspan":1,"cellId":"p1_t0_r0_c0"},
          {"text":"负责人","rowspan":1,"colspan":1,"cellId":"p1_t0_r0_c1"},
          {"text":"部门","rowspan":1,"colspan":1,"cellId":"p1_t0_r0_c2"},
          {"text":"日期","rowspan":1,"colspan":1,"cellId":"p1_t0_r0_c3"}
        ]},
        {"cells":[
          {"text":"A-001","rowspan":2,"colspan":1,"cellId":"p1_t0_r1_c0"},
          {"text":"张三 / 研发一部 / 2025-12-01","rowspan":1,"colspan":3,"cellId":"p1_t0_r1_c1"}
        ]},
        {"cells":[
          {"text":"李四","rowspan":1,"colspan":1,"cellId":"p1_t0_r2_c1"},
          {"text":"研发二部","rowspan":1,"colspan":1,"cellId":"p1_t0_r2_c2"},
          {"text":"2025-12-02","rowspan":1,"colspan":1,"cellId":"p1_t0_r2_c3"}
        ]}
      ],
      "source":{"producer":"OpenCvTable+GeminiCells","attempt":1}
    },
    {
      "type":"paragraph",
      "role":"body",
      "text":"一、总体进展\n本周完成了需求评审与原型确认。",
      "source":{"producer":"GeminiText","attempt":1}
    }
  ]
}
```

---

## A3.7 示例 7：文档 IR（doc_ir.json）

> 最终用于 docx 写入的完整 IR。

```json
{
  "version": "1.0",
  "meta": {
    "sourcePdfPath": "C:\\Docs\\report_scan.pdf",
    "generatedAtUtc": "2025-12-27T03:00:00Z",
    "options": {
      "dpi": 300,
      "pageRange": "1,3,5-10",
      "headerFooterMode": "RemoveBoth",
      "headerPercent": 0.06,
      "footerPercent": 0.06,
      "pageSizeMode": "FollowPdf"
    }
  },
  "pages": [
    { "pageNumber": 1, "widthPx": 2480, "heightPx": 3088, "blocks": [/*...*/] },
    { "pageNumber": 3, "widthPx": 2480, "heightPx": 3088, "blocks": [/*...*/] }
  ]
}
```

---

## A3.8 示例 8：输出（docx）行为说明

### A3.8.1 A4 模式输出

* 全文 A4
* 每 PDF 页内容后插入 PageBreak
* 适合统一打印与企业模板

### A3.8.2 FollowPdf 模式输出

* 每页一个 Section（SectionBreak NextPage）
* 页面宽高由像素→twips换算得到（基于裁切后尺寸）
* 适合最大程度贴近原 PDF 版面

---

## A3.9 示例 9：失败与降级示例（表格结构冲突 → 纯文本表格）

### A3.9.1 触发条件

* `E_TABLE_GRID_CONFLICT`（owner 矩阵冲突）或 `E_TABLE_GRID_INSUFFICIENT_LINES`

### A3.9.2 降级后的 IR（表格变段落块）

```json
{
  "type":"paragraph",
  "role":"body",
  "text":"【表格（已降级为纯文本）】\n项目\t负责人\t部门\t日期\nA-001\t张三\t研发一部\t2025-12-01\n...",
  "source":{"producer":"GeminiTableTextFallback","attempt":1}
}
```

### A3.9.3 UI 提示示例

* 警告：部分表格已以纯文本方式输出（可读但不可编辑为表格）
* 建议：提高 DPI 或开启 deskew 重试

---

## A3.10 建议的样例文件清单（用于回归测试）

* `sample_01_basic_tables.pdf`：标准线条表格 + 少量合并
* `sample_02_many_merges.pdf`：合并单元格多
* `sample_03_thin_lines.pdf`：细线易断
* `sample_04_header_footer_noise.pdf`：页眉页脚干扰明显
* `sample_05_landscape_pages.pdf`：横向页
* `sample_06_small_text_dense.pdf`：小字密集

每个样例建议保存：

* 期望页码范围、建议 DPI
* 抽检点（关键 cell 文本、关键段落句子）
* 允许降级的范围（例如某些表格允许降级为文本）

---

```
```
