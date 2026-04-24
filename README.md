# final-exam-ppt-review-handout

面向中国高校期末复习场景的 OpenClaw / ClawHub skill：把老师给的课程 `.pptx/.pptm` 课件整理成**考前复习大纲、Word 讲义和 PDF**。

它不是泛泛的 PPT 总结工具，而是专门服务于这类需求：

> “我的课程 PPT 在某个路径里，请你自己拉取这个 skill，分析 PPT，整理成适合期末复习的讲义，输出 Word 和 PDF。”

> 说明：旧版二进制 `.ppt` 不能被 `python-pptx` 直接解析，需要先用 PowerPoint、WPS 或 LibreOffice 转为 `.pptx`。

---

## 这个 skill 能做什么？

- 批量处理一门课的多章 `.pptx/.pptm`，也能处理单个总复习课件或一组独立专题课件。
- 支持 `--recursive` 递归扫描子目录。
- 不要求文件必须命名为 `第一章、第二章`；AI 应根据文件名、页内标题和实际内容自行判断顺序与分组。
- 提取标题、正文、表格、备注、图示数量和图示页提示。
- 对图片较多的课件给出报告警告，提醒调用方结合原 PPT 或视觉/OCR 模型复核。
- 让调用它的 AI 根据课件内容进行理解、合并和提炼，而不是机械复制文字。
- 输出每个章节/专题一份 Word 复习讲义，并可导出 PDF、打包 Word zip、生成报告。

默认推荐使用 **review-margin 批注式讲义版式**：正文区域在左侧，右侧预留约四分之一空白批注区，中间用竖线隔开，方便打印后手写补充、复核图示页和标注重点。也可以通过 `--layout standard` 切换为普通满版讲义。

---

## 最简单用法：直接告诉 OpenClaw / Agent

```text
去 GitHub 拉取这个 skill：
https://github.com/CZ3744/final-exam-ppt-review-handout

我的课程 PPTX/PPTM 放在：<你的PPT目录>
请按这个 skill 的 SKILL.md 流程执行：
1. 先提取 PPT 内容；
2. 阅读 compact.md / slides.json，判断资料是连续章节、单个总复习课件，还是独立专题课件；
3. 自行决定合理顺序和分组，不要只按文件名机械排序；
4. 整理成期末复习讲义 handout.json；
5. 再导出 Word 和 PDF；
6. 输出到：<你的输出目录>；
7. 默认使用 review-margin 批注式讲义版式；
8. 完成后检查 report.md，并把结果路径告诉我。

注意：不要机械逐页复制 PPT，要按考前复习逻辑整理成知识点大纲、名词解释、对比表、流程、易考点和速记总结。
```

---

## 手动安装与运行

```bash
pip install -e .
```

### 1. 提取 PPTX/PPTM

```bash
python -m ppt_review_handout.cli_v2 extract \
  --input ./course_ppts \
  --workspace ./workspace \
  --config examples/sample_config.json
```

如果课件在多层目录中：

```bash
python -m ppt_review_handout.cli_v2 extract \
  --input ./course_ppts \
  --workspace ./workspace \
  --recursive
```

输出：

```text
workspace/extracted/*.slides.json
workspace/extracted/*.compact.md
workspace/report.md
workspace/report.json
```

### 2. 让调用方 AI 写 handout.json

调用方 AI 应阅读：

```text
workspace/extracted/*.compact.md
workspace/extracted/*.slides.json
```

并写入：

```text
workspace/analysis/*.handout.json
```

`render` 阶段会故意只接受 `*.handout.json`，避免把原始 `slides.json` 误当成复习讲义渲染。

### 3. 渲染 Word/PDF

```bash
python -m ppt_review_handout.cli_v2 render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --export-pdf \
  --zip-word \
  --config examples/sample_config.json
```

普通满版讲义：

```bash
python -m ppt_review_handout.cli_v2 render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --export-pdf \
  --zip-word \
  --layout standard
```

PDF 导出需要本机安装 LibreOffice 或 `soffice`。没有 PDF 环境时，Word 仍然可以生成，报告会写明原因。

---

## Fallback 模式

```bash
python -m ppt_review_handout.cli_v2 build \
  --input ./course_ppts \
  --output ./outputs \
  --mode handout \
  --export-pdf \
  --zip-word \
  --keep-intermediate
```

Fallback 只适合 smoke test 或粗略草稿，不替代真正的语义分析。生成的文档和报告会标记为 fallback 粗略版。

---

## 常见资料组织方式

### 连续章节课件

```text
绪论.pptx
第一章 材料性能.pptx
第二章 晶体结构.pptx
第三章 钢的热处理.pptx
```

推荐输出：按课程逻辑输出每章讲义，通常为：绪论 → 第一章 → 第二章 → 第三章。

### 单个总复习课件

```text
机械工程材料总复习.pptx
```

推荐输出：生成一份完整总复习讲义，内部按知识主题重组。

### 独立专题课件

```text
金属材料复习.pptx
热处理专题.pptx
有色金属专题.pptx
实验复习.pptx
```

推荐输出：每个专题单独生成一份讲义，不要强行合并为“第一章、第二章”。

---

## 开发

```bash
pip install -e . pytest
pytest -q
```

---

## License

MIT-0.
