# campus-job-skills

校招季最耗时间的往往不是写简历，而是一家一家地翻官网：找到公司真正的校招入口，分清秋招、实习、专项计划，从几百个岗位里挑出和自己方向相关的，再把链接和 JD 整理进表格，最后才轮到填申请表。

这个仓库是一组给 AI 助手用的"技能包"（Skill）。装进 Claude Code、Codex 等能在你电脑上干活的 AI 工具后，AI 会按固定流程帮你逐公司收集校招信息、写入 Excel 求职账本，并在你选好岗位后协助填写申请表——但最终提交永远由你本人确认。

## 三个 Skill 分别做什么

| Skill | 负责的环节 | 一句话说明 |
|---|---|---|
| `campus-job-preference-intake` | 开批前的范围确认 | 每一批开始前，先和你确认本批看哪些地区、公司、岗位族和城市；有歧义就一次问一个问题，确认后才开始翻网页 |
| `campus-job-plan-inventory` | 校招计划盘点 | 对每家公司核实官方校招入口，盘点正式批、实习、专项计划，写入工作簿的计划表 |
| `campus-job-application` | 岗位收集与投递协作 | 逐公司筛选岗位、按"标题 + 完整 JD"判断是否纳入，把直达链接和 JD 写入地区表并读回核验；你选定岗位后协助填表，停在提交前 |

## 它遵守的几条规矩

- 只认公司官方校招页面；第三方招聘网站只用来找入口，不作为最终岗位来源。
- 岗位是否纳入看 JD 的实际职责，不只看标题里有没有"产品经理"。
- 每写入一行都要读回 Excel 核对；同一公司的岗位集中存放，按开始时间排序，岗位类型的先后顺序由你每批指定。
- 不保存密码、Cookie、验证码或证件信息；需要登录的地方由你本人完成；任何最终提交都要你单独批准。

## 怎么用

1. 准备一个求职工作目录，放入你的求职账本 Excel 和简历文件夹。
2. 把三个 Skill 目录复制到你的 AI 助手的 Skill 目录（例如 Codex 的 `.agents/skills/` 或 Claude Code 的 `.claude/skills/`）。
3. 按 `campus-job-application/references/profile-template.md` 在工作目录建立 `产品管理/求职配置.md`，填入工作簿、地区表、目标届别、岗位方向、简历位置等你自己的信息。Skill 本身不含任何个人取值，全部从这份配置读取；缺失时 AI 会先问你再创建。
4. 对 AI 说"开始新一批校招信息收集"，它会先调用 `campus-job-preference-intake` 和你确认范围。

工作簿的表结构见 `campus-job-application/references/workbook-schema.md`，浏览器与提交门禁见 `campus-job-application/references/browser-and-submission-gates.md`。

## 关于本仓库

本仓库与作者本地使用的 Skill 是同一份通用文件；所有个人与项目取值都放在各自工作区的求职配置中，不进入仓库。由 [SilentLake-Tech-SkillHub](https://github.com/SilentLake-Tech-SkillHub) 维护。
