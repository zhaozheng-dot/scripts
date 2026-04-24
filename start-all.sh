#!/bin/bash
# start-all.sh - OpenCode 全栈启动脚本
# 用法: bash /mnt/f/scripts/start-all.sh

set -e

OC_BIN="/home/alex/.opencode/bin/opencode"
LOG_DIR="/tmp/opencode"
mkdir -p $LOG_DIR

echo "=== OpenCode 全栈启动 ==="
echo "时间: $(date '+%Y-%m-%d %H:%M:%S')"

# 检查是否已在运行
check_port() {
    if ss -tlnp 2>/dev/null | grep -q ":$1 "; then
        echo "[WARN] 端口 $1 已被占用，跳过启动"
        return 1
    fi
    return 0
}

# 代码学习实例 (8080)
if check_port 8080; then
    echo "[1/2] 启动 Code Learning (port 8080)..."
    cd /mnt/f/projects/lc4j-agentic-tutorial/agentic-tutorial
    nohup $OC_BIN web --hostname 0.0.0.0 --port 8080 > $LOG_DIR/oc-8080.log 2>&1 &
    echo "  PID: $!"
else
    echo "[1/2] Code Learning (8080) 已运行"
fi

sleep 3

# 知识库教师实例 (8081)
if check_port 8081; then
    echo "[2/2] 启动 KB Teacher (port 8081)..."
    cd /mnt/f/obsidian_repository/scienc-project-repo
    nohup $OC_BIN web --hostname 0.0.0.0 --port 8081 > $LOG_DIR/oc-8081.log 2>&1 &
    echo "  PID: $!"
else
    echo "[2/2] KB Teacher (8081) 已运行"
fi

sleep 5

# 健康检查
echo ""
echo "=== 健康检查 ==="
for port in 8080 8081; do
    for i in 1 2 3 4 5; do
        if curl -sf http://127.0.0.1:$port/ > /dev/null 2>&1; then
            echo "[OK] 端口 $port 响应正常 (尝试 $i/5)"
            break
        elif [ $i -eq 5 ]; then
            echo "[FAIL] 端口 $port 无响应，请检查日志:"
            echo "  tail -50 $LOG_DIR/oc-$port.log"
        else
            echo "[WAIT] 端口 $port 等待响应... ($i/5)"
            sleep 3
        fi
    done
done

echo ""
echo "=== 启动完成 ==="
echo "Code Learning:  http://100.84.60.105:9080 (→8080)"
echo "KB Teacher:     http://100.84.60.105:9081 (→8081)"
