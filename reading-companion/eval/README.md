# Reading Companion — 评测集 (eval set)

固定的一批标准样本,用来对 `prompt-template.txt` 做**可比、可回归**的质量评测。改 prompt 后重跑同一批,才能判断是进步还是退步。

## 文件
- `eval_set.json` — 26 条策划案例(本目录)。

## 每条案例的字段
| 字段 | 含义 |
|---|---|
| `id` | 稳定编号 |
| `domain` | 领域(cs / law / economics / science / medicine / philosophy / literature …) |
| `selection` | 被查的词或短语 |
| `context_sentence` | 它所在的句子(送模型的锚点) |
| `article_text` | 一段自足的上下文(生产环境送全文;评测用代表性段落保证可复现) |
| `native_language` | 解释应输出的语言 |
| `tags` | 分层与陷阱标签(见下) |
| `should_have_distinction` | **期望** `distinction` 是否应触发(测触发恰当性的金标准) |
| `distinction_hint` | 若应触发,期望的易混邻居(给评委参考,不要求逐字匹配) |
| `check_notes` | 这条要重点检查什么 |

## 标签词表 (tags)
- 领域类:`cs` `ml` `law` `economics` `finance` `science` `physics` `medicine` `philosophy` `social-science` `literature`
- 词形类:`single-word` `phrase` `idiom` `latin` `french` `japanese`
- 弱点/陷阱类:
  - `jargon-prone` — 易用行话解释行话(考 meaning 是否零背景可懂)
  - `confusable` — 有常见易混邻居(应触发 distinction)
  - `polysemous` — 一词多义(考是否取此处的义项)
  - `trap` — 专门制造的陷阱(如 `repeated-word`)
  - `repeated-word` — 该词在段落多次出现(考 sentence 锚定是否取对句)

## 覆盖情况(26 条)
- `should_have_distinction`: **true 17 / false 9**(既测"该触发时命中",也测"不该触发时正确留空")
- 领域:cs 6、literature 5、law 3、economics 3、science 3、medicine 2、philosophy 2、finance 1、social-science 1
- 陷阱:`confusable` 17、`polysemous` 3、`jargon-prone` 3、`idiom` 2、`repeated-word` 1

## 这批案例怎么对应四个模块的评分

| 模块 | 主要靠哪些案例施压 |
|---|---|
| **meaning_in_context**(含义简洁) | `jargon-prone`(002/004/026 等)测"不能行话套行话";`polysemous`(008/026)、`repeated-word`(025)测"取对此处义项/此处句子" |
| **substance**(机制详解) | 深度型:019 dialectic、014 entropy、011 moral hazard 等,测是否讲清底层机制而非复述含义 |
| **related_terms**(锚定词语) | 全部案例都产出锚点;重点看锚点是否比原词更熟悉、桥接是否真帮理解;`confusable` 案例额外测**不得与 distinction 重复**(009 是典型) |
| **distinction**(对比概念) | `should_have_distinction=true` 的 17 条测"命中且区分清楚";=false 的 9 条测"正确留空、不硬凑" |

## 陷阱案例速查(专抓已知弱点)
- **不该有 distinction** → 001 dropout、013 catalyst、019 dialectic、021 bittersweet、022 elephant、023 raison d'être、008 consideration、025 model、026 field
- **jargon 套 jargon** → 002 JSON Schema、004 idempotent、026 field
- **多义取错** → 008 consideration、026 field、025 model
- **同词多次出现(取错句)** → 025 model
- **纠正常见误解型 distinction** → 014 entropy(vs 通俗的"disorder")

## 用法(建议流程)
1. 对每条跑当前 prompt → 收集结构化输出。
2. **确定性检查**(便宜、先跑):schema 合法、`related_terms` 与 `distinction.vs` 无重复、输出语言 = `native_language`、`should_have_distinction` 与实际是否触发一致、声称来自文章的锚点确在 `article_text` 中。
3. **LLM-as-judge**:按 rubric 对四个模块逐维度打分(1–4,带理由),评委需看到 article+sentence+selection。
4. 聚合:按模块/维度/领域出均分 + 失败清单;每次改 prompt 对比上一版,防回归。

## 扩展建议
- 先用这 26 条跑通;之后按薄弱领域补充(如各领域再加 3–5 条)。
- 可为其中 5–8 条写"理想答案要点"用于评委校准(算与人工的一致性)。
