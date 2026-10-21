# 本轮一次性投递需求

搜索组件开始处理投递需求时，先运行 `application_requirements.py begin`（MCP `recruiting_application_requirements_begin`）。Hook 记录真实会话与回合上下文；同回合循环复用当前选择，新用户回合重新初始化，半自动逐回合推进同样重置。

宿主接线：Codex/DSH 和 Claude 的 `SessionStart`、`UserPromptSubmit` 调用 `application_requirements.py --workspace <工作区> hook`；Antigravity 的 `SessionStart`、`PreInvocation` 调用同一脚本，追加 `--host agy --event <事件>`。Hook 只做回合初始化；没有启用该需求流程的对话返回空结果，组件实际启动后才提示当前流程。DSH 修改注册后须重载；不要把手动调用脚本当成宿主已自动触发。

## 唯一需求入口

- 临时表：`产品管理/user_applicaiton_requirements_temp.json`，拼写保持接口约定，只覆盖写入。
- 有用户参数：`application_requirements.py set --requirements '<JSON对象>' --quote '<用户原话>'`，或 MCP `recruiting_application_requirements_set`。参数支持 `公司`（名称列表）、`地区`、`浏览器`（iab/chrome/dual）。每次新需求都覆盖旧需求，不追加。
- 用户要求默认：使用同一个接口写 `{}`；同回合后续循环不再询问。下一回合恢复初始化。
- 空表 `{}` 表示无本轮参数，走默认路线；损坏的JSON或被其他会话持有的需求不能静默按空表处理。
- 其他会话不能覆盖、读取或消费当前持有者的非空需求；回合上下文与阶段元数据只保存在私有运行目录，不能当成另一份业务需求。

## 固定组批

`company_store.py forward`（MCP `recruiting_application_forward`）只能从临时表读需求，不能由 Agent 绕过临时表另传公司筛选：

1. 有 `公司`：按用户名单与顺序匹配地区和公司；匹配不到或跨地区同名未指明地区时报错。
2. 没有 `公司`：从“岗位待投递”取最新10家公司；优先按进入该状态的确认时间，缺失时按数据刷新时间，时间相同保持主表顺序。
3. 按公司计数，保留选中公司全部“岗位待投递”岗位及用户志愿顺序；既有链接、身份、志愿和正反向评审门禁继续执行。存在具体问题的岗位列在评审页，不擅自换公司补位。
4. 用户要求的浏览器通道冻结在当前批次 `execution.browserChannels`；只用内置浏览器时为 `['iab']`，投递组件领取前调用其通道限制接口。
5. 成功生成且读回当前批次与正向评审后，临时表立即清空为 `{}`；生成失败保留参数供重试。当前批次保留运行所需参数，后续循环不能重复消费旧参数。
6. 正向评审仍须用户确认。结束反向回写时检查需求表已空；若用户已写入下一批需求，保留新需求，不清除下一批参数。

本入口不修改岗位状态、不写Excel，不代替选岗、填写授权或最终提交授权。配置了Hook不等于宿主已实际触发；代码验证与宿主触发分别核对。
