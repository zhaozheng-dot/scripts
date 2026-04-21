#!/bin/bash
# OpenCode 双实例启动脚本
# 用法: bash start-opencode.sh

OC_BIN="/home/alex/.opencode/bin/opencode"

echo "=== Starting OpenCode instances ==="

# 代码学习实例
echo "[1/2] Code learning (port 8080)..."
cd /mnt/f/projects/lc4j-agentic-tutorial/agentic-tutorial
nohup $OC_BIN web --hostname 0.0.0.0 --port 8080 > /tmp/oc-8080.log 2>&1 &

sleep 2

# 知识库教师实例
echo "[2/2] Knowledge base teacher (port 8081)..."
cd /mnt/f/obsidian_repository/scienc-project-repo
nohup $OC_BIN web --hostname 0.0.0.0 --port 8081 > /tmp/oc-8081.log 2>&1 &

sleep 2

echo ""
echo "=== Status ==="
echo "Code learning:  http://localhost:8080  (mobile: http://100.84.60.105:9080)"
echo "KB teacher:     http://localhost:8081  (mobile: http://100.84.60.105:9081)"
