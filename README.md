# final-exam-ppt-review-handout

面向中国高校期末复习场景的 OpenClaw / ClawHub skill：把老师给的课程 `.pptx` / `.pptm` 课件整理成**考前复习大纲、Word 讲义和 PDF**。

> 旧版二进制 `.ppt` 文件会被检测并写入报告，但需要先转换成 `.pptx`，因为 `python-pptx` 不能直接解析 `.ppt`。

它不是泛泛的 PPT 总结工具，而是专门服务于这类需求：

> “我的课程 PPTX/PPTM 在某个路径里，请你自己拉取这个 skill，分析课件，整理成适合期末复习的讲义，输出 Word 和 PDF。”

---

## 这个 skill 能做什么？

- 批量处理一门课的多章 `.pptx` / `.pptm`，也能处理单个总复习课件或一组独立专题课件。
- 不要求文件必须命名为 `第一章、第二章`；AI 会根据文件名、页内标题和实际内容自行判断顺序与分组。
- 提取课件中的标题、正文、表格、备注和图示页信息。
- 对图片占比较高的课件在报告中提示调用方结合原 PPT 或视觉/OCR 步骤复核。
- 让调用它的 AI 根据课件内容进行理解、合并和提炼，而不是机械复制文字。
- 输出每个章节/专题一份 Word 复习讲义。
- 输出每个章节/专题一份 PDF。
- 可把所有 Word 打包成 zip。
- 生成运行报告，说明处理了哪些文件、输出了哪些产物、有没有失败或警告。

默认推荐使用 **review-margin 批注式讲义版式**：正文区域在左侧，右侧预留约四分之一空白批注区，中间用竖线隔开，方便打印后手写补充、复核图示页和标注重点。也可以通过 `--layout standard` 切换为普通满版讲义。

最终目标是得到这种资料：

```text
第一章 绪论_复习讲义版.docx
第一章 绪论_复习讲义版.pdf
材料性能专题_复习讲义版.docx
材料性能专题_复习讲义版.pdf
review_handouts_docx.zip
report.md
report.json
```

---

## 最简单用法：直接告诉 OpenClaw / Agent

普通用户不需要理解内部流程。你可以直接对 OpenClaw、Claude Code、Codex 或其他本地 agent 说：

```text
去 GitHub 拉取这个 skill：
https://github.com/CZ3744/final-exam-ppt-review-handout

我的课程 PPTX/PPTM 放在：<你的课件目录>
请按这个 skill 的 SKILL.md 流程执行：
1. 先提取课件内容；
2. 你自己阅读提取结果，判断这些资料是连续章节、单个总复习课件，还是一组独立专题课件；
3. 你自己决定合理顺序和分组，不要只按文件名机械排序；
4. 整理成期末复习讲义；
5. 再导出 Word 和 PDF；
6. 输出到：<你的输出目录>；
7. 默认使用推荐的 review-margin 批注式讲义版式；
8. 完成后检查 report.md，并把结果路径告诉我。

注意：不要机械逐页复制 PPT，要按考前复习逻辑整理成知识点大纲、名词解释、对比表、流程、易考点和速记总结。
```

如果你想输出到原课件同目录，可以把最后一行改成：

```text
输出到课件目录下新建的 final_review_outputs 文件夹。
```

---

## 常见资料组织方式

这个 skill 不要求固定命名。调用它的 AI 应自己判断资料结构。

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

### 混合或命名不清楚

```text
绪论.pptx
期末重点.pptx
补充资料.pptx
实验复习.pptx
```

推荐做法：AI 根据内容决定顺序，并在结果或报告中简要说明排序假设。

---

## 适合什么场景？

适合：

- 中国高校期末复习；
- 老师发了一整套 `.pptx` / `.pptm` 课程课件；
- 想把课件变成复习大纲、知识点讲义、Word/PDF；
- 不想一页一页手动复制文字；
- 想让 AI 帮你理解、提炼和排版；
- 想得到方便打印批注的复习讲义版式。

不适合：

- 只想原样把 PPT 转成 PDF；
- 课件基本全是扫描图片，且没有 OCR 或视觉模型辅助；
- 需要 100% 保留 PPT 原始视觉设计；
- 直接处理旧版 `.ppt` 二进制文件。

---

## 给 Agent 的核心要求

这个项目采用“调用方 AI 自己分析”的设计。

也就是说：

```text
skill 负责：提取 PPTX/PPTM、生成中间文件、渲染 Word/PDF、打包、生成报告。
调用它的 AI 负责：真正理解课件、判断资料顺序、合并知识点、提炼易考点、写复习讲义内容。
```

因此，Agent 不能只运行一条“自动 build”命令就结束。高质量结果应该走：

```text
extract → AI 阅读 compact.md / slides.json → AI 判断顺序与分组 → AI 写 handout.json → render
```

`render` 阶段只接受 `*.handout.json`，会拒绝把原始 `slides.json` 当作讲义渲染。

详细执行规则写在 `SKILL.md` 中。

---

## 手动安装与运行

如果你不是通过 OpenClaw 调用，而是想自己本地跑：

```bash
pip install -e .
```

提取课件：

```bash
ppt-review-handout extract \
  --input ./course_ppts \
  --workspace ./workspace
```

如果课件在子目录中，可以加递归扫描：

```bash
ppt-review-handout extract \
  --input ./course_ppts \
  --workspace ./workspace \
  --recursive
```

然后让 AI 阅读：

```text
workspace/extracted/*.compact.md
workspace/extracted/*.slides.json
```

并写入：

```text
workspace/analysis/*.handout.json
```

最后渲染，默认就是推荐的批注式讲义版式：

```bash
ppt-review-handout render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --export-pdf \
  --zip-word
```

如果需要普通满版讲义：

```bash
ppt-review-handout render \
  --analysis ./workspace/analysis \
  --output ./outputs \
  --export-pdf \
  --zip-word \
  --layout standard
```

PDF 导出需要本机安装 LibreOffice 或 soffice。没有 PDF 环境时，Word 仍然可以生成，报告会写明原因。

---

## 示例与配置

通用示例配置在：

```text
examples/sample_config.json
```

该配置不包含任何学校或课程专属清洗词。若你要为某一门课删除固定页眉/页脚，可另建自己的 config，并使用精确文本匹配；需要正则时使用 `re:` 前缀。

仓库中可放置无版权 demo 素材到：

```text
examples/demo/
docs/assets/
```

为了避免课件版权问题，公开仓库中通常不建议放原始老师 PPT；完整 Word/PDF 产物如果体积较大，也更适合放在 GitHub Release。

---

## Development

```bash
pip install -e . pytest
pytest -q
```

CI 会运行单元测试和一个最小端到端链路测试：自动生成 PPTX，执行 `extract -> handout.json -> render -> zip`。

---

## License

MIT-0.
