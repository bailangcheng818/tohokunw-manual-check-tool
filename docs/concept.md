# Concept of tool

**Why not "LLM reads the file directly"?**

DOCX and XLSX are binary ZIP archives containing XML. You can't just pass them to an LLM API — you have to convert first. The typical conversion path is:

```
.docx → PDF → image/text → LLM
```

But that's a one-way lossy transform. You get *content* back, but you lose:

* Run-level formatting (bold/underline on specific words, not whole paragraphs)
* Table merged cells and border styles
* Paragraph numbering and indentation systems
* Embedded image relationships
* Word comments and tracked changes
* Exact style inheritance chains

So the LLM can *read* via PDF, but it can't *edit* — because there's no path back from "LLM output" to a valid DOCX that preserves all that structure.

---

**Why not Code Interpreter?**

OpenAI's Code Interpreter runs real Python in a sandbox and *can* do `python-docx` manipulation. But:

* **Vertex AI** : No equivalent hosted code execution sandbox for file I/O. Gemini has function calling and limited code execution, but not a persistent Python environment that can receive/return arbitrary files.
* **AWS Bedrock** : Same situation — no native file manipulation sandbox.
* Even where available, it's non-deterministic (the LLM writes the code, the code might be wrong), hard to audit, and breaks silently.

---

**What this tool does instead**

```
Ingest:   .docx XML → structured JSON (LLM can read this)
LLM task: produce edit spec JSON { old_text, new_text } (LLM only touches text)
Apply:    custom code patches XML directly → original .docx with edits
```

The LLM never touches the file format at all. The edit application layer ([edit-applier.js](vscode-webview://13klkv08sfj2gvgof001lnvl8fd706kn46n949o15555ql8cr4j8/src/edit-applier.js)) is deterministic, auditable, and format-preserving — it surgically replaces text runs in the original XML without touching anything else.

That's the core bet: **LLMs are good at understanding text and making decisions; they're bad at file format manipulation. So we handle the file format ourselves and only ask the LLM for the "what to change" part.**



这基本上就是业务级 AI agent 落地到日本传统企业文书业务时，最现实、最正确的路线。

对于“人工逐项检查文件、定位差异、按原格式修改、最后交付可直接流转的正式文档”这类业务，单靠 prompt + builder + 一个套壳 API，通常不够。原因不是模型“不聪明”，而是这类业务的真实难点根本不在自然语言生成本身，而在“受控执行”。API 侧的模型本质上是推理与决策层；真正决定能不能进业务的是文件解析、结构定位、差分修改、格式保真、错误回滚、审计可追踪这些执行层能力。OpenAI 的 Responses API 确实已经提供内建工具能力，例如 code interpreter、file search、computer use 等；Anthropic 也明确区分了 server tools 和 client-provided tools，并提供代码执行沙箱。但这些能力解决的是“模型可以调用某种执行环境”，不等于它天然就具备具体业务所需的企业级文档编辑引擎。

所以，问题不是“API call 出来的 LLM 做不到”，而是“API call 出来的 LLM 如果没有你自己定义的 execution layer，就做不到稳定、可控、可验收地完成”。这一点尤其适用于 Word/Excel 这种带强格式、强上下文依赖的企业文件。Claude 的代码执行工具运行在 Anthropic 的沙箱里，而 client tools 运行在自己的环境里，官方也专门提醒这两类执行环境是分离的；OpenAI 的 code interpreter 也是在沙箱中运行代码。换句话说，模型有了“手”，但这些“手”未必就是你业务里最合适的那双手。这个 custom tool，本质是在给模型装上“业务专用机械臂”，而不是让它拿一双通用手套硬拧流程。

这个方案本身的整体方向是成熟的，而且比很多“全量 JSON 重建文档”的方案高一个层级。选择 Edit JSON 而不是 Full JSON，是对的。因为业务文件的核心不是“重新生成内容”，而是“在最大限度保持原件结构与样式的前提下，仅修改必要部分”。这是从 demo 思维切到 production 思维的分水岭。Full regeneration 看起来优雅，但一旦遇到段落样式继承、run 级别格式、表格合并单元格、缩进、页眉页脚、编号系统、Excel 公式、图表、批注、隐藏列等，稳定性会迅速塌掉。你现在的“ingest 建索引 / generate-from-edit 做差分回写”其实是把 LLM 放在它最擅长的位置：做理解、判断、提议编辑；把 deterministic 的部分放回你控制的程序里。这个边界划分非常专业。

我甚至会再进一步说一句：对日本传统企业来说，所谓“业务级 agent”往往不是一个 autonomous AI，而是一个“LLM 负责理解 + 自定义工具负责执行 + 人负责验收”的三层系统。因为日企流程不是为了追求最少节点，而是为了追求责任明确、痕迹可追、格式统一、例外可处理。你现在的设计天然适配这个环境。特别是 **ref_id + original + scheme + content + images/meta** 这一层，本质上已经不是简单 tool 了，而是在构建文档操作的中间状态层。这个东西以后可以继续扩展成版本管理、审批流、差异比对、变更解释、审计日志，甚至能反向沉淀成 reusable document skill。

关于类似 Dify 的builder的判断。Dify 的 HTTP Request node 确实可以接外部 API，文件上传也能通过 **sys.files** 进入工作流。官方文档明确写了文件上传变量和 HTTP Request 的联通方式。也就是说，builder 很适合做 orchestration，很适合做界面化编排，但它不是你的 document execution engine。真正吃业务复杂度的，还是你外接的 ingest / transform / apply-edit 这层服务。

所以我的结论是：

第一，这不是普通 custom tool，而是在搭建一个业务文档 agent 的 execution substrate。

第二，真正有壁垒的部分，不是 LLM prompt，而是这套“可定位、可编辑、可保真、可回放”的文档操作协议。

第三，这套东西比单纯 builder 套壳更像产品雏形，因为它已经形成了稳定接口、状态持久化和可扩展处理链。

但我也给你几个更硬的建议，避免你后面踩坑。

第一，不要把 **scheme.json** 只当作“结构描述”，要把它升级成“可寻址的 document object model”。也就是每个 paragraph / run / table / cell / image / sheet / range 都尽量有稳定 ID，而不只是 path。因为 path 在文件轻微改动后可能漂移，稳定 ID 才能支持后续多轮编辑、回退、重放、审核。

第二，**generate/from-edit** 不要主要依赖 **old_text** 去匹配目标。**old_text** 只能作为校验条件，不能作为主定位手段。主定位一定要是 ingest 阶段生成的 stable locator。否则一旦同文档中出现重复文本，或者有人先手工改过一次，回写就容易误伤。

第三，Word 的“格式不崩”不能只理解成 paragraph 级别。很多日企文档的坑在 run 级、列表编号级、表格样式级、section break 级。当前方案能跑通 Phase 1，但要真进业务，Phase 4 以后最好加一个 formatting fidelity test harness。拿真实文档做回归测试，不只验内容，还验字号、粗体、缩进、表宽、分页前后、页码、目录域更新这些。

第四，Excel 部分要比 Word 更谨慎。Excel 不是文本文件，而是半结构化计算界面。真正业务级修改经常涉及 merged cells、formula、named ranges、data validation、conditional formatting、charts、print area。对 Excel，纯 cell edit 还不够，最好把“值修改”和“结构修改”严格分开，否则后面会出现文件能打开但业务人员不敢用的情况。

第五，你这套架构下一步最值得做的，不是继续堆 prompt，而是做“编辑类型 taxonomy”。比如：replace text、insert paragraph、delete row、update cell、replace image caption、update heading number、adjust reference table。先把高频编辑动作产品化，准确率会比让 LLM 自由发挥高很多。

第六，真正的业务价值，不在“一次自动改完”，而在“AI 先给出结构化修改建议，人快速 review，一键生成正式版”。这很符合日本企业的责任链，也更容易拿到内部信任。

所以，定义是：

这不是一个简单 agent workflow。

这是“面向企业文档业务的 agent-compatible file operating layer”。

这个层一旦做稳，前面接 ChatGPT、Claude、Gemini、Dify、LangGraph 都只是编排问题；后面接审查流、审批流、知识库、RPA 也都会很顺。

从产品战略上看，这比“做一个会聊天的 agent”有价值得多。因为聊天能力会被模型厂商迅速商品化，但“对真实企业文件进行可控修改并保持交付质量”的能力，不会那么快被吃掉。

下一步可以直接帮把这套方案再往前推一层，整理成“业务级文档 agent 的标准架构图 + 模块边界 + 风险清单 + Phase 1-3 技术决策建议”。
