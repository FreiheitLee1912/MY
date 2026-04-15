# 上传文件分析:Jira CSV → PPTX 甘特图生成器

## 总览

上传的 5 个文件构成了一个**完整的 Web + CLI 双模式甘特图生成器**,从 Jira CSV 导出生成可直接用于 PPT 演示的 `.pptx` 文件。

| 文件 | 角色 | 行数 |
|------|------|------|
| `index.html` | Web 前端 UI(日语界面) | 196 |
| `styles.css` | 样式(玻璃拟态 + 亮/暗主题) | 1029 |
| `app.js` | 前端逻辑:CSV 解析 + Gantt 渲染 + 调用后端 API | 871 |
| `server.py` | HTTP 服务器 + PPTX 生成(亮色主题) | 631 |
| `csv_to_gantt.py` | 独立 CLI 版:直接 CSV → PPTX(深色主题) | 733 |

## 架构:两种运行模式

### 模式 A:浏览器 + 服务器(`server.py` + 前端)

```
用户
  │ 拖放 CSV
  ▼
index.html / app.js
  │ • 前端解析 CSV
  │ • 渲染交互式 Gantt(可切换 月/季/年 视图、分组、项目筛选)
  │ • 点 "PPTX 出力"
  │ POST /api/generate-pptx (JSON)
  ▼
server.py (port 8090)
  │ python-pptx 生成 .pptx
  ▼
./output/*.pptx
```

### 模式 B:命令行(`csv_to_gantt.py` 独立运行)

```
python csv_to_gantt.py [file.csv ...]
  │ 读取 Downloads/ 里的 Jira 导出 CSV
  │ 直接生成 PPTX
  ▼
./output/*.pptx
```

---

## CSV 格式(Jira 导出假设)

按**列位置**读取(不按列名),顺序固定:

| 列索引 | 字段 | 示例 |
|--------|------|------|
| 0 | Key | `DSSC-123` |
| 1 | Type | `Milestone` / `Task` / `Goal` / ... |
| 2 | Parent | `DSSC-100`(父任务 Key) |
| 3 | Summary | 任务名称 |
| 4 | Status | `完了` / `Done` / `In Progress` / ... |
| 5 | Assignee | 担当人 |
| 6 | Start date | `20/4/26`(DD/M/YY) |
| 7 | Est. Start | 预计开始 |
| 8 | Deadline | 截止 |
| 9 | Est. Deadline | 预计截止 |

**注意点**:
- 日期格式是 `DD/M/YY`(日-月-年2位),这是**日本 Jira 导出**的特殊格式,例如 `20/4/26` = 2026-04-20
- 时间后缀 `12:00 午前 / 午後` 会被正则剥离
- Project 前缀从 Key 自动提取(`DSSC-123` → `DSSC`)

---

## 任务类型与颜色体系

`server.py` / `csv_to_gantt.py` / `app.js` 共享同一套类型枚举:

| Type | 颜色 | 徽章 |
|------|------|------|
| Milestone | 橙 `#f2994a` | MS |
| Goal | 绿 `#4caf50` / `#6fcf97` | Goal |
| Review | 青 `#2996d6` | Rev |
| Event | 紫 `#9c27b0` | Evt |
| Task / タスク | 蓝 `#427ad6` / `#8ab4f8` | Task |
| Sub task | 淡蓝 | Sub |
| Validation | 红 `#e53e3e` | Val |
| SOP | 金 `#e6a817` | SOP |
| Certificate | 蓝 `#2f80ed` | Cert |
| PPAP | 红 `#e54b4b` | PPAP |
| Top-level initiative | 蓝 | Init |
| **完了 / Done** | 绿 `#27ae60`(覆盖类型色) | — |

---

## PPTX 生成技术要点(`python-pptx`)

### 幻灯片规格

```python
SLIDE_WIDTH  = Inches(13.33)   # 16:9 widescreen
SLIDE_HEIGHT = Inches(7.5)
TASKS_PER_SLIDE = 18           # 分页:每页 18 个任务
```

### 布局

```
┌─────────────────────────────────────────────────────────┐
│ Title (14-16pt bold)                                    │
│ YYYY/MM - YYYY/MM  (subtitle)                           │
├─────────────┬───────────────────────────────────────────┤
│  Labels     │  Month headers (proportional width)      │ ← HEADER_HEIGHT
│  (3.2")     │  ───────────┼───────────┼────────────    │   + grid lines
│             │                                           │
│  Key  Badge │       ● ━━━━━━━━━                         │ ← ROW_HEIGHT
│  Summary    │                                           │    × 18 rows
│             │            ▼ Today (red line)             │
│    ...      │                                           │
├─────────────┴───────────────────────────────────────────┤
│ ■ Milestone ■ Goal ■ Review ... | Today                  │ ← Legend
└─────────────────────────────────────────────────────────┘
```

### 绘图技巧

1. **无原生图表**,全部用 `add_shape()` 画矩形/圆角矩形/菱形,手动计算坐标。
2. **月份比例宽度**:`month_width = CHART_WIDTH × (days_in_month / total_days)`
3. **任务条定位**:
   ```python
   bar_x = CHART_LEFT + CHART_WIDTH * ((task_start - bounds_start).days / total_days)
   bar_w = CHART_WIDTH * (duration_days / total_days)
   ```
4. **Milestone 用菱形**(`MSO_SHAPE.DIAMOND`),其它用圆角矩形
5. **徽章背景色 = 类型色 + 白色混合**(0.2 / 0.15 alpha)实现柔和的标签
6. **Today 红线**:一个 2pt 宽的窄矩形
7. **垂直居中文字**:直接操作 XML — `tf._txBody.bodyPr.set('anchor', 'ctr')`
8. **日文字体** `'Meiryo UI'` 统一指定
9. **分页保留分组头**:每页 18 个任务,分组标题不计入任务数

---

## 主题差异(需要注意)

| | `server.py`(浏览器后端) | `csv_to_gantt.py`(CLI) |
|---|---|---|
| 背景 | 白色 `#ffffff`(**适合 PPT 打印/投影**) | 深色 `#0f1117`(**屏幕演示**) |
| 文本 | 深色 `#1a1a2e` | 浅色 `#e8eaed` |
| 行交替 | 淡灰 `#f8f9fb` | 深灰 `#1a1c28` |

⚠ **两套代码各画一遍 PPT** — 逻辑几乎一致,但 ~400 行重复。

---

## 交互功能(前端)

| 功能 | 实现 |
|------|------|
| CSV 拖放/点选 | `FileReader` + `parseCSVGenerator`(手写 CSV parser,支持引号/转义) |
| 日期解析 | `DD/M/YY` → Date,兼容 `YYYY-MM-DD` |
| 项目筛选 | `project` 下拉(从 Key 前缀自动提取) |
| 视图模式 | 月 / 四半期(季度)/ 年 → 调整 `dayWidth` (4/2/1 px) |
| 分组 | 类型别 / 親タスク別 / なし |
| 时间范围自定义 | `<input type="date">` start/end |
| PNG 出力 | `html2canvas` 截屏 |
| PPTX 出力 | POST 到后端,后端用 `python-pptx` 生成 |
| 主题切换 | `data-theme="light"/"dark"` |
| 悬停提示 | 自定义 tooltip DOM |

---

## 发现的潜在问题

### 1. 代码重复严重
`server.py` 与 `csv_to_gantt.py` 的 `generate_pptx_*` 函数几乎一样,只是颜色常量不同。**可以抽出 `gantt_renderer.py` 模块**,两边都导入。

### 2. 缺少依赖声明
没有 `requirements.txt`。运行 `server.py` 需要:
```
python-pptx>=0.6.21
```

### 3. CSV 列按位置硬编码
`row[0]..row[9]` 强依赖 Jira 导出列顺序。如果用户导出时列顺序变化或缺失列,会错位。`app.js` 里有 `getValue(names)` 按 header 名查找的逻辑但没用上。

### 4. 日期格式假设特殊
`DD/M/YY` 是日本 Jira 的本地化格式。**其它地区的 Jira 会是 `M/D/YYYY` 或 `YYYY-MM-DD`**,代码里用 `a > 31` / `c <= 99` 的启发式判断可能误判。

### 5. 图例可能溢出
18 行任务 × 0.3" + 0.6" 顶 + 0.3" header ≈ 6.3",slide 总高 7.5",图例在 `7.5 - 0.35 = 7.15"` — **边缘紧凑,长 summary 可能压到图例**。

### 6. `csv_to_gantt.py` 主题是深色
生成的 pptx 背景是深蓝黑,**打印/投影时对比度差**。如果只做演示 OK;若要打印,应改用 `server.py` 的亮色主题。

### 7. `server.py` 的 grid line 用 `Pt(0.5)` 作为宽度
Pt 是 `12700 EMU`,而 Inches 是 `914400 EMU`。Pt(0.5) 作为 shape 宽度 ≈ 0.007" ≈ 1 像素,视觉上可能太细/看不到,取决于 PPT 缩放。建议 `Pt(1)` 或 `Emu(9525)` (1px)。

### 8. 文件名清洗过度激进
```python
filename = re.sub(r'[^\w\-_\.]', '_', filename)
```
`/` 和空格会变 `_`,但日文字符 `\w` 在 Python re 里默认不匹配 — 日文文件名会被整段替换为 `_`。应加 `re.UNICODE` 或允许更多字符。

---

## 关于"pptx 输出技能"的核心知识点提炼

从这份代码可以抽出 **python-pptx 高级用法** 作为技能基础:

### A. 精确坐标 / 尺寸体系

- `Inches(x)`、`Pt(x)`、`Emu(x)` — 都返回 EMU 整数值
- 1 Inch = 914400 EMU, 1 Pt = 12700 EMU
- 加减乘除 EMU 整数来算比例布局
- `//` 整除避免浮点

### B. 复合布局的套路

1. 定义 `SLIDE_WIDTH / HEIGHT`
2. 定义 `LEFT_MARGIN / TOP_MARGIN / LABEL_WIDTH / CHART_WIDTH`
3. **迭代绘制** — month headers → today line → rows(交替背景 → text → badge → summary → bar / diamond)→ legend
4. 分页模式:`TASKS_PER_SLIDE = N`,遍历时计数

### C. 用形状绘制业务图形

- 条形 → `MSO_SHAPE.ROUNDED_RECTANGLE`
- 里程碑 → `MSO_SHAPE.DIAMOND`
- 分隔线 → 窄 `MSO_SHAPE.RECTANGLE`(Pt 宽)
- 徽章 → `MSO_SHAPE.ROUNDED_RECTANGLE` + 小号文字

### D. 文本精细控制

- 水平:`p.alignment = PP_ALIGN.CENTER / LEFT / RIGHT`
- 垂直:`tf._txBody.bodyPr.set('anchor', 'ctr')`(raw XML)
- 日文字体:`p.font.name = 'Meiryo UI'`
- `tf.word_wrap = False` 防止窄 bar 内文字换行
- `p.space_before = Pt(0); p.space_after = Pt(0)` 去除默认间距

### E. 颜色混合(柔和徽章背景)

```python
# 类型色 * 0.15 + 白色 * 0.85 → 淡色背景
r, g, b = type_rgb[0], type_rgb[1], type_rgb[2]
softer = RGBColor(
    int(r * 0.15 + 255 * 0.85),
    int(g * 0.15 + 255 * 0.85),
    int(b * 0.15 + 255 * 0.85),
)
```

### F. 背景色设置

```python
bg = slide.background
bg.fill.solid()
bg.fill.fore_color.rgb = RGBColor(0xff, 0xff, 0xff)
```

### G. HTTP API 返回 JSON + 文件落地

- `http.server` 继承写 `do_POST`
- 输出文件到 `./output/`,API 返回 `{ success, filename, path }`
- 前端 `fetch('/api/generate-pptx', {method:'POST', body:JSON.stringify(...)})`

---

## 下一步建议

根据用途选方向:

| 目标 | 行动 |
|------|------|
| **只是想看懂/学习** | 本文档已够,可直接读 `server.py:185-540`(最精华的 PPTX 生成逻辑) |
| **继续使用这套工具** | 加 `requirements.txt`,写 `README.md` 启动说明,修复#2 和 #8 |
| **重构为可复用模块** | 抽 `gantt_renderer.py`,两边 import;通过参数传主题/布局常量 |
| **扩展为通用 pptx 技能** | 把 A-G 的套路整理成模板脚本,放到 `scripts/` 和 `templates/` |
| **替换/对比方案** | 研究 `pptxgenjs`(纯浏览器生成 pptx)、`spire.presentation`、`Aspose.Slides` |

请告诉我要哪个方向,我继续推进。
