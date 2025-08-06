#!/bin/bash

# Docker部署脚本
# 使用方法: ./scripts/docker-deploy.sh [dev|prod]

set -e

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 打印带颜色的消息
print_message() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

print_step() {
    echo -e "${BLUE}[STEP]${NC} $1"
}

# 检查Docker是否安装
check_docker() {
    if ! command -v docker &> /dev/null; then
        print_error "Docker未安装，请先安装Docker"
        exit 1
    fi
    
    if ! command -v docker-compose &> /dev/null; then
        print_error "Docker Compose未安装，请先安装Docker Compose"
        exit 1
    fi
    
    print_message "Docker环境检查通过"
}

# 停止并删除现有容器
cleanup_containers() {
    print_step "清理现有容器..."
    docker-compose down --remove-orphans || true
    print_message "容器清理完成"
}

# 构建镜像
build_image() {
    local env=$1
    print_step "构建Docker镜像 (环境: $env)..."
    
    if [ "$env" = "dev" ]; then
        docker-compose build easyfilepreview-dev
    else
        docker-compose build easyfilepreview
    fi
    
    print_message "镜像构建完成"
}

# 启动服务
start_service() {
    local env=$1
    print_step "启动服务 (环境: $env)..."
    
    if [ "$env" = "dev" ]; then
        docker-compose --profile dev up -d easyfilepreview-dev
    else
        docker-compose up -d easyfilepreview
    fi
    
    print_message "服务启动完成"
}

# 检查服务状态
check_service() {
    local env=$1
    print_step "检查服务状态..."
    
    sleep 5
    
    if [ "$env" = "dev" ]; then
        local port=3001
    else
        local port=3000
    fi
    
    # 等待服务启动
    local max_attempts=30
    local attempt=1
    
    while [ $attempt -le $max_attempts ]; do
        if curl -s "http://localhost:$port/api/filePreview/formats" > /dev/null; then
            print_message "服务启动成功！"
            print_message "访问地址: http://localhost:$port"
            print_message "演示页面: http://localhost:$port/filePreviewDemo.html"
            return 0
        fi
        
        print_warning "等待服务启动... (尝试 $attempt/$max_attempts)"
        sleep 2
        attempt=$((attempt + 1))
    done
    
    print_error "服务启动超时"
    return 1
}

# 显示容器日志
show_logs() {
    local env=$1
    print_step "显示容器日志..."
    
    if [ "$env" = "dev" ]; then
        docker-compose logs -f easyfilepreview-dev
    else
        docker-compose logs -f easyfilepreview
    fi
}

# 主函数
main() {
    local env=${1:-prod}
    
    print_message "开始部署 EasyFilePreview (环境: $env)"
    
    # 检查环境参数
    if [ "$env" != "dev" ] && [ "$env" != "prod" ]; then
        print_error "无效的环境参数: $env"
        print_error "使用方法: $0 [dev|prod]"
        exit 1
    fi
    
    # 检查Docker环境
    check_docker
    
    # 清理现有容器
    cleanup_containers
    
    # 构建镜像
    build_image $env
    
    # 启动服务
    start_service $env
    
    # 检查服务状态
    if check_service $env; then
        print_message "部署完成！"
        print_message "要查看日志，请运行: $0 logs $env"
    else
        print_error "部署失败"
        exit 1
    fi
}

# 处理日志命令
if [ "$1" = "logs" ]; then
    show_logs ${2:-prod}
    exit 0
fi

# 处理停止命令
if [ "$1" = "stop" ]; then
    print_step "停止服务..."
    docker-compose down
    print_message "服务已停止"
    exit 0
fi

# 处理重启命令
if [ "$1" = "restart" ]; then
    print_step "重启服务..."
    docker-compose restart
    print_message "服务已重启"
    exit 0
fi

# 处理清理命令
if [ "$1" = "clean" ]; then
    print_step "清理所有Docker资源..."
    docker-compose down --volumes --remove-orphans
    docker system prune -f
    print_message "清理完成"
    exit 0
fi

# 运行主函数
main "$@" 