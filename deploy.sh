#!/usr/bin/env bash
# =============================================================================
# deploy.sh – QuoToCon 部署脚本
# 将项目同步到腾讯云 Lighthouse 服务器并以 Docker 方式启动
#
# 用法:
#   chmod +x deploy.sh
#   ./deploy.sh
#
# 前提条件:
#   1. 本地已安装 rsync、ssh
#   2. 已将本机 SSH 公钥复制到服务器 (~/.ssh/authorized_keys)
#      若尚未配置, 运行: ssh-copy-id ubuntu@124.222.92.176
#   3. 服务器已安装 Docker 和 Docker Compose (脚本会自动检查并安装)
# =============================================================================
set -euo pipefail

# ── 配置 ─────────────────────────────────────────────────────────────────────
REMOTE_USER="ubuntu"
REMOTE_HOST="124.222.92.176"
REMOTE_DIR="/home/ubuntu/QuoToCon"
APP_PORT=5001
SSH_OPTS="-o StrictHostKeyChecking=accept-new -o ConnectTimeout=15"

# ── 颜色输出 ──────────────────────────────────────────────────────────────────
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

info()  { echo -e "${GREEN}[INFO]${NC}  $*"; }
warn()  { echo -e "${YELLOW}[WARN]${NC}  $*"; }
error() { echo -e "${RED}[ERROR]${NC} $*" >&2; exit 1; }

# ── 检查本地依赖 ──────────────────────────────────────────────────────────────
command -v rsync >/dev/null || error "本地未找到 rsync，请先安装。"
command -v ssh   >/dev/null || error "本地未找到 ssh，请先安装。"

# ── 测试 SSH 连通性 ───────────────────────────────────────────────────────────
info "测试 SSH 连接 ${REMOTE_USER}@${REMOTE_HOST} ..."
ssh ${SSH_OPTS} "${REMOTE_USER}@${REMOTE_HOST}" "echo '连接成功'" \
    || error "SSH 连接失败，请检查密钥或网络。"

# ── 同步项目文件 ──────────────────────────────────────────────────────────────
info "同步项目文件到 ${REMOTE_HOST}:${REMOTE_DIR} ..."
rsync -avz --delete \
    --exclude='.git/' \
    --exclude='.venv/' \
    --exclude='__pycache__/' \
    --exclude='*.pyc' \
    --exclude='.env' \
    --exclude='uploads/' \
    --exclude='outputs/' \
    --exclude='logs/' \
    -e "ssh ${SSH_OPTS}" \
    ./ "${REMOTE_USER}@${REMOTE_HOST}:${REMOTE_DIR}/"

# ── 在服务器端执行部署 ────────────────────────────────────────────────────────
info "在服务器上执行部署 ..."
ssh ${SSH_OPTS} "${REMOTE_USER}@${REMOTE_HOST}" bash <<REMOTE_SCRIPT
set -euo pipefail

GREEN='\033[0;32m'; NC='\033[0m'
info() { echo -e "\${GREEN}[远端][\${NC} \$*"; }

# ── 安装 Docker（如未安装）────────────────────────────────────────────────────
if ! command -v docker &>/dev/null; then
    echo "[远端] 正在安装 Docker ..."
    curl -fsSL https://get.docker.com | sudo sh
    sudo usermod -aG docker ubuntu
    newgrp docker <<NEWGRP
    echo "[远端] Docker 安装完成"
NEWGRP
fi

# ── 安装 Docker Compose plugin（如未安装）────────────────────────────────────
if ! docker compose version &>/dev/null 2>&1; then
    echo "[远端] 正在安装 Docker Compose plugin ..."
    sudo apt-get update -qq
    sudo apt-get install -y -qq docker-compose-plugin
fi

cd ${REMOTE_DIR}

# ── 创建 .env（如不存在则从示例生成）────────────────────────────────────────
if [ ! -f .env ]; then
    echo "[远端] 生成 .env 文件 ..."
    cp .env.example .env
    SECRET=\$(python3 -c "import secrets; print(secrets.token_hex(32))")
    sed -i "s|change-me-in-production|\${SECRET}|g" .env
    echo "[远端] .env 已创建，SECRET_KEY 已随机生成。"
fi

# ── 创建运行时目录 ────────────────────────────────────────────────────────────
mkdir -p uploads outputs logs

# ── 构建并启动容器 ────────────────────────────────────────────────────────────
echo "[远端] 构建镜像 ..."
docker compose build --no-cache

echo "[远端] 启动容器 ..."
docker compose up -d

echo "[远端] 等待服务启动 ..."
sleep 5

# ── 健康检查 ──────────────────────────────────────────────────────────────────
if curl -sf http://localhost:${APP_PORT}/ >/dev/null; then
    echo "[远端] 服务已正常运行: http://${REMOTE_HOST}:${APP_PORT}"
else
    echo "[远端] 警告: 健康检查未通过，查看日志:"
    docker compose logs --tail=30
fi

echo "[远端] 容器状态:"
docker compose ps
REMOTE_SCRIPT

info "部署完成！访问地址: http://${REMOTE_HOST}:${APP_PORT}"
