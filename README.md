# PDF-Comparison

一个基于 Python 的 PDF 差异对比工具，当前实现重点是文本语义级对比，而不是纯像素差分。

项目提供桌面 GUI（PySide6），可对两份 PDF 执行文本差异分析、坐标级高亮标注，并输出结构化对比报告 PDF。

## 当前实现状态

当前主实现文件是 py_PDF_compare_gui.py，对比链路已升级为：

1. 从 PDF 中提取文本 span（含坐标）
2. 使用 diff-match-patch 执行语义级文本 diff
3. 将新增/删除差异映射回页面坐标
4. 生成可视化标注页与结构化报告页

已不再使用旧版依赖 OpenCV 轮廓阈值作为核心差异识别策略。

## 核心功能

### 1. 文本语义对比

- 基于 diff-match-patch 进行文本差异计算
- 支持语义清理流程（Semantic / Efficiency cleanup）
- 支持最小差异 token 长度过滤，减少噪声变更

### 2. 精准坐标高亮

- 删除内容在旧文档页面标红
- 新增内容在新文档页面标绿
- 结果来自文本坐标映射，不是图像轮廓估计

### 3. 结构化报告页

输出 PDF 会附带文字版差异汇总，字段包括：

- 原文档描述
- 原文档页数
- 新文档描述
- 新文档页数

报告支持自动换行与跨页续写，避免长路径/长文本被截断。

### 4. 多种输出视图

可在设置中控制输出页类型：

- New Copy
- Old Copy
- Markup
- Difference
- Overlay

## 技术栈

- Python 3.9+
- PySide6（桌面 GUI）
- PyMuPDF / fitz（PDF 解析与渲染）
- Pillow（图像绘制与合成）
- diff-match-patch（语义 diff）

## 项目结构

- py_PDF_compare_gui.py: 主程序与核心比较引擎
- settings.json: 运行参数持久化配置
- 文档对比.html: 前端原型页面（用于交互和策略参考）
- README.md: 当前说明文档
- README.zh-CN.md: 中文辅助说明文档

## 安装与运行

### 1. 创建虚拟环境（Windows PowerShell）

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2. 安装依赖

```powershell
pip install PySide6 PyMuPDF Pillow diff-match-patch
```

### 3. 启动程序

```powershell
python py_PDF_compare_gui.py
```

## 使用说明

1. 启动 GUI
2. 拖拽或选择两份 PDF（主文件 + 次文件）
3. 点击 Settings 按需调整：
	- 输出类型
	- DPI / 页面尺寸
	- 文本对比参数（例如最小 token 长度、是否规范化文本）
4. 点击 Compare 执行比较
5. 在日志窗口查看进度，等待输出 PDF 生成

## 关键配置项（settings.json）

- INCLUDE_IMAGES: 控制输出页类型
- DPI / DPI_LEVEL / DPI_LEVELS: 渲染精度
- PAGE_SIZE / PAGE_SIZES: 页面尺寸
- MAIN_PAGE: 主文档选择（New Document / Old Document）
- TEXT_MIN_DIFF_LENGTH: 最小差异 token 长度
- NORMALIZE_TEXT: 文本规范化开关
- OUTPUT_PATH: 输出路径（为空时回落源文件目录）

## 已知限制

- 对扫描件或图片型 PDF（无可提取文本）效果有限
- 复杂排版（多栏、表格、浮动对象）可能引入额外噪声差异
- 大文档在高 DPI 下可能需要较多内存与时间

## 后续优化建议

- 增加 OCR 流程以支持图片型 PDF
- 增强替换差异聚合策略（长段落重排场景）
- 增加 JSON/CSV 差异导出

## 许可说明

当前仓库未包含单独 License 文件。
