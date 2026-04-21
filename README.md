# F:\scripts

Windows/WSL 运维脚本集合。

## 脚本列表

| 脚本 | 用途 |
|------|------|
| `obsidian-sync.sh` | Obsidian 知识库一键 git add + commit + push |
| `start-opencode.sh` | 启动双 OpenCode Web 实例（代码学习 + 知识库教师） |
| `port-fwd-win.py` | Python 端口转发（9080→8080, 9081→8081） |

## 用法

### 知识库同步
```bash
# 默认消息
wsl -e bash /mnt/f/scripts/obsidian-sync.sh

# 自定义消息
wsl -e bash /mnt/f/scripts/obsidian-sync.sh "feat: add new notes"
```

### 启动 OpenCode 双实例
```powershell
wsl -e bash /mnt/f/scripts/start-opencode.sh
```

### 启动端口转发
```powershell
python F:\scripts\port-fwd-win.py
```
