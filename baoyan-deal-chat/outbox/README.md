# outbox/ · 脚本产物目录

跑下面两条命令后会在这里生成文件：

```bash
"$PY" "$SK/parse_labels.py"    # 成交客户名单（读微信「已成交」标签）
"$PY" "$SK/refresh.py" --with-deals
```

| 产物 | 内容 |
| --- | --- |
| `deals_labeled.json` | 成交客户名单：wxid / 备注 / 昵称 / 标签 / 聊单笔记 |
| `deal_wxids.txt` | 成交客户 wxid 列表，供 `export_all.py --wxid-file` 用 |

⚠️ **这两个文件含你客户的真实身份信息，属于隐私数据。**
打包 / 分享技能前请删除本目录下的产物（本包分发时已清空）。
