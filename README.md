# leak-check

两个 Claude Code skill：AI 写完论文、课程作业、实验报告之后，交上去之前查一遍有没有**露馅**或**一眼不对**的句子。

| skill | 调用 | 做什么 |
|---|---|---|
| `leak-check` | `/leak-check` | 只查不改：列出每处问题的位置、原句、原因、建议改法 |
| `leak-fix` | `/leak-fix` | 查完直接修：做最小改动，保留 Word 格式，先备份原文件，修完复扫，交付修改清单 |

两者共用 `shared/scripts` 和 `shared/references`。各 skill 目录里的 `scripts`、`references` 是指向 `shared/` 的符号链接。

## 查什么

**雷点**，就是一眼看出是 AI 代写或者没收尾的句子：

- 元话语：根据用户的要求…… / 用户说…… / 老师要求…… / 根据你提供的资料 / 本文共计 3200 字
- 草稿残留：之后要替换掉这里的内容 / 此处插入流程图 / 【此处填写结论】 / TODO / 待补充 / 学号：XXX / 示例数据
- 版本标记：改进版 / 最终版 / （修订版） / 修改前：…… 修改后：……
- 对话残留与 AI 自曝：好的，以下是…… / 希望对你有帮助 / 如果还需要调整请告诉我 / 作为一个 AI
- 格式残留：Word 或 PDF 里的 `##`、`**`、`$公式$`、`\frac`
- Word 残留：批注、未接受的修订、荧光高亮、隐藏文字

**硬伤**，就是一眼看出不对的地方：

- 图 1、图 3，中间没有图 2；正文写"如图 5"，但没有图 5；同一个表号出现两次
- 引用了 [12]，文献表只有 10 条；文献编号跳号；文献年份比今年还晚
- 错误！未找到引用源。
- 标题编号 1 → 3；空章节；两段话一字不差
- "共回收问卷 200 份"和"共回收问卷 180 份"
- 括号不配对；引号写成两个右引号 ”…”

脚本能查出能靠字符串和结构判断的部分；同义变体和语义矛盾（结论与正文对不上、可疑文献、残句、时间穿越等），由 AI 按 [semantic-checklist.md](shared/references/semantic-checklist.md) 逐句通读找出。完整清单和每一类的"什么时候其实没问题"写在 [patterns.md](shared/references/patterns.md) 里。

## 误报控制

- 需求分析、产品设计类文章里，"用户希望按价格筛选"是正常写法。脚本发现全文有 3 段以上把"用户"当领域词用时，会把"用户……"类命中降为 WARN。但如果同一句话还提到了报告本身，比如"用户说这一节要写详细，所以本报告分三点展开"，仍然报 ERROR。
- 以下写法不会误报：待定系数法、方向待定、改进版 A* 算法、限于篇幅省略推导（详见文献[3]）、参考文献里的（第 2 版）、C 代码里的 `arr[5]`、讲义里以 `#` 开头的运算符说明、"1）2）3）"式编号。
- 同一句话命中多条规则时，合并成一条；同类问题多处出现时，汇总成一条。

## 安装

```bash
git clone https://github.com/Jonty-Zhang/leak-check ~/.claude/skill-repos/leak-check
ln -s ~/.claude/skill-repos/leak-check/leak-check ~/.claude/skills/leak-check
ln -s ~/.claude/skill-repos/leak-check/leak-fix ~/.claude/skills/leak-fix
pip install -r ~/.claude/skill-repos/leak-check/requirements.txt
brew install poppler   # 可选：查 PDF 需要 pdftotext
```

不能用符号链接的环境（如 Windows），把 `shared/scripts` 和 `shared/references` 复制进两个 skill 目录即可。

## 使用

在 Claude Code 里输入 `/leak-check 论文.docx` 或 `/leak-fix 论文.docx`。也可以直接说"交之前帮我查一下有没有雷""把露馅的句子都改掉"。

也可以单独用脚本：

```bash
python3 shared/scripts/scan_leaks.py 论文.docx                 # 扫描，有 ERROR 时退出码为 1
python3 shared/scripts/scan_leaks.py 论文.docx --dump          # 带段号的全文
python3 shared/scripts/scan_leaks.py a.md b.tex --json         # 多个文件，输出 JSON
cat 草稿.txt | python3 shared/scripts/scan_leaks.py -           # 从标准输入读
python3 shared/scripts/fix_docx.py 论文.docx edits.json --dry-run
python3 shared/scripts/fix_docx.py 论文.docx edits.json --inplace --accept-revisions --remove-comments
python3 shared/scripts/selftest.py                             # 自测
```

`edits.json` 的格式写在 [fix_docx.py](shared/scripts/fix_docx.py) 的文件开头。按段号替换时，文字沿用原 run 的格式；跨 run 的命中也能改。只要有一条没命中，脚本就整份不写入。

## 路线图

- v1（当前）：雷点和硬伤
- v2：加入 AI 味检查，包括套话词汇、三段式排比、空洞的升华句、让步式自贬

## License

MIT
