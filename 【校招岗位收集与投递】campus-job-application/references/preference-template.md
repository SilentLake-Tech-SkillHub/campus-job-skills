# 私有偏好空模板

复制到工作区，按用户原话填充；空值不是已确认。此模板没有默认城市、方向或就业路径。

```json
{
  "version": "",
  "confirmation_ref": "",
  "effective_at": "",
  "supersedes_version": null,
  "recheck_scope": [],
  "hard_scope": {"cities": []},
  "axes": {
    "employment": {"raw": "", "relations": []},
    "city": {"raw": "", "relations": []},
    "direction": {"raw": "", "relations": []}
  },
  "internship_policy": "",
  "conflict_policy": "user_decides"
}
```

另记录生效时间、替代版本、用户指定复查范围和首轮停点。关系对象、逐岗包、公司覆盖字段见 [偏好与证据契约](preference-contract.md)。`cities: []` 只表示不限城市，轴原话仍可记录排序；未确认字段保持待核。不得把实际个人取值写回公共模板。
