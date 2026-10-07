# 独立偏好与执行证据维护验证

基线 main：`3ab6cec960bef31e950d864cbbf97655a3cf3285`。仅新分支/Draft PR；不合并、不修改招聘网站或公共模板中的个人取值。保留上游布局、Excel 列/枚举、首轮/二轮范围、批量分阶段、自动化授权与提交审核。

## 可复现检查

```sh
python3 -m unittest discover -s '【校招岗位收集与投递】campus-job-application/scripts' -p 'test_*.py'
python3 docs/validation/check_package.py .
```

Python 合成测试：33 项通过，无跳过。包括不限城市不排除、硬城市范围、弱优先/并列/未知原样保留、跨轴不代选/不加权、AI 数据产品与纯数仓的人工 JD 依据、正式岗与实习分列、候补覆盖充分/不足、有正式岗矛盾、旧已投/回执与明确选择保护、旧偏好版本拒收、BU 未披露、缺前置逐岗报告拒收、实际执行/回执证据、同名不同 ID 与多城市同 ID、URL 业务参数保留。测试只验证声明与字段契约，不判断 JD 真伪或自然语言分类正确性。

## 完整包与流程图

完整仓库复制到临时安装目录后，从另一 cwd 运行上述测试/校验器。核对所有 SKILL frontmatter 的 name/description、Markdown 本地引用、Python 语法和各校验器 CLI 加载。该检查不等于在 Codex/Claude/Dots 等平台实际注册、远程安装或招聘站端到端验证；嵌套子技能必须随父包复制，偏好空模板必须留在私有工作区填充。旧快照缺新字段须保留原始证据、人工适配后再交付，不能把版本号改成新版本绕过校验。

本次改变的 6 张流程图均由对应 Excalidraw 源完整重绘，PNG 内嵌源 SHA256，check_package 核对一致性。其余图不改。`docs/flow/render_flow.py` 仅支持本仓库的 text/rectangle/ellipse/line/arrow 与无旋转元素，遇到其他形状拒绝；需要 Pillow 和中文字体（`FLOW_FONT` 可指定字体文件）。它保留流程形状/关系，使用确定性直线与字体替代手绘样式。例：

```sh
FLOW_FONT=<Chinese-font-file> python3 docs/flow/render_flow.py docs/flow/module-1.excalidraw
```

人工检查新 PNG 文本和流程逻辑；源变更必须同时重绘 PNG。真实表单操作/保存/提交、来源真实性、公司覆盖充分性、前置报告时间、用户确认及历史快照完整性仍须执行者现场核验。结构校验不是平台强制 hook、授权、账号锁或事实验收。

## 精确文件范围

- `.gitignore`
- `README.md`
- `docs/flow/module-1.excalidraw`
- `docs/flow/module-1.png`
- `docs/flow/module-2.excalidraw`
- `docs/flow/module-2.png`
- `docs/flow/module-3a.excalidraw`
- `docs/flow/module-3a.png`
- `docs/flow/module-3b.excalidraw`
- `docs/flow/module-3b.png`
- `docs/flow/overview.excalidraw`
- `docs/flow/overview.png`
- `docs/flow/render_flow.py`
- `docs/flow/sub-parallel.excalidraw`
- `docs/flow/sub-parallel.png`
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
