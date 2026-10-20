# 工作区配置与接线（campus-job-workspace）

这个组件把三个 Skill（范围确认、计划盘点、岗位收集与投递）接成一个能运行的求职工作区：标签组、字段规则、求职目标、交接目录、Hook 与 MCP 接线，以及初始化脚本。Skill 目录里只有通用代码和规则；工作区要用的配置都在这里。

## 目录

```text
【工作区配置与接线】campus-job-workspace/
├── 初始化工作区.py              新建工作区：放入配置、建交接目录、生成空账本与主数据（已有文件不覆盖）
├── AGENTS.md                    工作区执行入口，放在工作区根目录
├── 产品管理/
│   ├── 投递状态标签.json         岗位状态标签组，同时是脚本主配置（主数据、账本、台账、名额规则、交接目录的位置）
│   ├── 公司状态标签.json         公司状态标签组（每家公司一行公司行）
│   ├── 表格字段规则.json         岗位名称、备注、投递截止的写入规则
│   ├── 求职目标.json             求职目标与选岗规则骨架，尖括号处由用户填写
│   ├── 组件维护.json             组件只读开关，默认关闭
│   └── 投递状态变更确认.json      状态变更确认台账（初始为空）
├── 岗位交接/
│   ├── 说明.md                   搜索端 → 投递端中间区域的固定文件与流程
│   └── 公司/                     每家公司一个文件夹（由脚本生成）
└── 接线/
    ├── claude-settings.json      Claude Code Hook（复制为 <WORKSPACE>/.claude/settings.json）
    ├── codex-hooks.json          Codex Hook（复制为 <WORKSPACE>/.codex/hooks.json）
    ├── antigravity-hooks.json    Antigravity Hook（合并进 <WORKSPACE>/.agents/hooks.json）
    └── codex-config.toml         MCP 注册：搜索 MCP 与投递 MCP（合并进 <WORKSPACE>/.codex/config.toml）
```

## 安装

1. 新建工作区目录，把三个 Skill 目录复制到 `<WORKSPACE>/.agents/skills/` 下，目录名去掉中文前缀：`campus-job-application`、`campus-job-plan-inventory`、`campus-job-preference-intake`。
2. 准备 Python 3.10 以上，安装 `openpyxl` 与 `【校招岗位收集与投递】campus-job-application/scripts/search-mcp-requirements.txt` 里的依赖（建议放在工作区的虚拟环境里）。
3. 运行初始化：`python3 初始化工作区.py --workspace <WORKSPACE> --skills <WORKSPACE>/.agents/skills`。
4. 把本组件的 `AGENTS.md` 复制到工作区根目录；按 `接线/` 里的文件配置 Hook 和 MCP，把 `<WORKSPACE>`、`<PYTHON>`、`<NODE>` 换成本机绝对路径。
5. 填写 `产品管理/求职目标.json` 的尖括号项，或让 Agent 在第一次会话时向你逐项确认后写入。
6. 自检：在工作区运行 `python3 .agents/skills/campus-job-application/scripts/json_excel_sync.py check`，`label_problems` 与 `field_problems` 应为 0。

投递组件在独立仓库 [job-application-assistant](https://github.com/SilentLake-Tech-SkillHub/job-application-assistant)，放到 `<WORKSPACE>/.agents/skills/job-application-assistant-v2/`；它的 Hook 按该仓库说明另行加入。两端只通过 `岗位交接/` 对接，见 [岗位交接/说明.md](岗位交接/说明.md)。

## 标签组

- 一组一个文件，每组用 `rowKind` 声明只适用于岗位行（role）还是公司行（company）。`投递状态标签.json` 是岗位组，其它组在它的 `labelGroups` 里登记。
- 所有脚本只经 `json_excel_sync.load_label_groups` 统一加载，写入前统一检查：岗位行用公司标签、公司行用岗位标签、跨组重名、组文件缺失都会被拒绝。
- 要增删或改名标签，只改对应组文件；任何状态变更都要先 `propose` 给用户看“当前状态是X，更新状态是Y”，确认后再 `confirm`。

## 隐私边界

本组件只含通用配置和占位路径，不含个人资料、账本内容、Cookie 或登录信息。`产品管理/网申个人信息主档.json`、简历、`.runtime/` 等私有内容只在用户本机工作区，不进入任何仓库。
