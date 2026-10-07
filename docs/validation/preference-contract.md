# 独立偏好与执行证据维护验证

基线 main：`3ab6cec960bef31e950d864cbbf97655a3cf3285`。仅新分支/Draft PR；不合并、不修改招聘网站或公共模板中的个人取值。保留上游布局、Excel 列/枚举、首轮/二轮范围、批量分阶段、自动化授权与提交审核。

## 可复现检查

```sh
python3 -m unittest discover -s '【校招岗位收集与投递】campus-job-application/scripts' -p 'test_*.py'
python3 docs/validation/check_package.py .
```

Python 合成测试：37 项通过，无跳过。包括不限城市不排除、硬城市范围、弱优先/并列/未知原样保留、跨轴不代选/不加权、AI 数据产品与纯数仓的人工 JD 依据、正式与转正实习同一报告排序、有正式岗和覆盖不足仍展示转正实习、每批默认次序及公司动作前确认、旧已投/回执与明确选择保护、旧偏好版本拒收、BU 未披露、缺前置逐岗报告拒收、实际执行/回执证据、同名不同 ID 与多城市同 ID、URL 业务参数保留。测试只验证声明与字段契约，不判断 JD 真伪或自然语言分类正确性。

## 完整包加载边界

完整仓库复制到临时安装目录后，从另一 cwd 运行上述测试/校验器。核对所有 SKILL frontmatter 的 name/description、Markdown 本地引用、Python 语法和各校验器 CLI 加载。该检查不等于在 Codex/Claude/Dots 等平台实际注册、远程安装或招聘站端到端验证；嵌套子技能必须随父包复制，偏好空模板必须留在私有工作区填充。旧快照缺新字段须保留原始证据、人工适配后再交付，不能把版本号改成新版本绕过校验。


真实表单操作/保存/提交、来源真实性、公司覆盖、前置报告时间、用户确认及历史快照完整性仍须现场核验；结构校验不是平台强制 hook 或真实 ATS 验证。

## 精确文件范围

- `.gitignore`
- `docs/validation/check_package.py`
- `docs/validation/preference-contract.md`
- `【校招岗位收集与投递】campus-job-application/SKILL.md`
- `【校招岗位收集与投递】campus-job-application/references/browser-and-submission-gates.md`
- `【校招岗位收集与投递】campus-job-application/references/preference-contract.md`
- `【校招岗位收集与投递】campus-job-application/references/preference-template.md`
- `【校招岗位收集与投递】campus-job-application/references/profile-template.md`
- `【校招岗位收集与投递】campus-job-application/references/research-progress.md`
- `【校招岗位收集与投递】campus-job-application/references/workbook-schema.md`
- `【校招岗位收集与投递】campus-job-application/scripts/preference_contract.py`
- `【校招岗位收集与投递】campus-job-application/scripts/test_preference_contract.py`
- `【校招岗位收集与投递】campus-job-application/scripts/validate_research_progress.py`
- `【校招岗位收集与投递】campus-job-application/skills/campus-job-parallel-search/SKILL.md`
- `【校招岗位收集与投递】campus-job-application/skills/campus-job-parallel-search/references/coordination-contract.md`
- `【校招岗位收集与投递】campus-job-application/skills/campus-job-parallel-search/scripts/validate_assignments.py`
- `【校招批次范围确认】campus-job-preference-intake/SKILL.md`
- `【校招计划盘点】campus-job-plan-inventory/SKILL.md`

## 独立复审反例回归

跨轴冲突执行必须同时有 selected:true 与 selection_ref；仅有引用而 selected:false 会拒收。uncertain 无任何动作证据会拒收；准备数量只计算候选中具备完成填写证据与材料身份的记录，部分填写不计已准备。搜索 manifest 及其 delta 岗位仅允许只读阶段，fill/save/submit 全部 not_started；通用申请入口仍可在授权内填写。合成回归覆盖原样反例与有效正例。

历史保留冲突回归：既有 submitted/回执/执行快照原样保留，并用 current_batch_execution 单独声明本批不执行；顶层和 delta 正例通过。新动作、改回执、重置执行、伪造历史或缺历史执行快照均拒收。历史只读例外不授权填写，也不放宽通用提交审核。

## 本轮候选规则修订

本批先初始化默认岗位次序，每家公司申请动作前用同一逐岗包确认相对默认次序有无变化，衔接最终目标与当前表单版本审核。私有 conversion_last 政策将明确转正的实习放同一报告末尾，不以无正式岗或覆盖充分为展示条件；旧 conditional/fallback 记录保留原件后按新确认版本适配。历史、明确选择、完整 JD 判断、身份与既有反例回归仍保留。README 和原仓库图保持基线内容，不在候选 PR 净差异中。

## 范围与候选状态

仅保留技能规则、使用参考、必要配套校验与合成测试。README 和 docs/flow 全部相对 main 零差异，不引入渲染生成文件；不安装绘图依赖、不重画。保持 Draft，不正式安装、不合并。通用 profile 结构校验未部署到本仓库，不声称已运行；使用本包实际校验器和合成招聘记录，不为不适用检查伪造字段。
