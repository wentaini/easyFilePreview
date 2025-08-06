# Docker 部署指南

本文档介绍如何使用 Docker 部署 EasyFilePreview 文件预览组件。

## 前置要求

- Docker Engine 20.10+
- Docker Compose 2.0+
- 至少 1GB 可用内存
- 至少 2GB 可用磁盘空间

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/wentaini/easyFilePreview.git
cd easyFilePreview
```

### 2. 使用部署脚本（推荐）

```bash
# 生产环境部署
./scripts/docker-deploy.sh prod

# 开发环境部署
./scripts/docker-deploy.sh dev

# 查看日志
./scripts/docker-deploy.sh logs prod

# 停止服务
./scripts/docker-deploy.sh stop

# 重启服务
./scripts/docker-deploy.sh restart

# 清理资源
./scripts/docker-deploy.sh clean
```

### 3. 手动部署

#### 生产环境

```bash
# 构建并启动生产环境
docker-compose up -d easyfilepreview

# 查看日志
docker-compose logs -f easyfilepreview

# 停止服务
docker-compose down
```

#### 开发环境

```bash
# 构建并启动开发环境
docker-compose --profile dev up -d easyfilepreview-dev

# 查看日志
docker-compose logs -f easyfilepreview-dev

# 停止服务
docker-compose down
```

## 环境配置

### 环境变量

| 变量名 | 默认值 | 说明 |
|--------|--------|------|
| `NODE_ENV` | `production` | 运行环境 |
| `PORT` | `3000` | 服务端口 |
| `HOST` | `0.0.0.0` | 监听地址 |
| `CORS_ORIGIN` | `*` | 跨域配置 |
| `MAX_FILE_SIZE` | `52428800` | 最大文件大小（字节） |
| `UPLOAD_DIR` | `./uploads` | 上传目录 |
| `CACHE_DIR` | `./cache` | 缓存目录 |
| `MAX_CACHE_SIZE` | `104857600` | 最大缓存大小（字节） |
| `CACHE_EXPIRE` | `86400000` | 缓存过期时间（毫秒） |

### 自定义配置

创建 `.env` 文件来自定义配置：

```bash
# 复制环境变量模板
cp env.example .env

# 编辑配置
vim .env
```

## 端口映射

| 环境 | 容器端口 | 主机端口 | 说明 |
|------|----------|----------|------|
| 生产环境 | 3000 | 3000 | 主服务端口 |
| 开发环境 | 3000 | 3001 | 开发服务端口 |

## 数据持久化

以下目录会被挂载到主机：

- `./uploads` - 文件上传目录
- `./cache` - 缓存目录
- `./logs` - 日志目录

## 健康检查

Docker 容器包含健康检查，会定期检查服务状态：

```bash
# 查看健康状态
docker ps

# 查看健康检查日志
docker inspect easyfilepreview
```

## 监控和日志

### 查看容器状态

```bash
# 查看运行中的容器
docker ps

# 查看所有容器（包括停止的）
docker ps -a
```

### 查看日志

```bash
# 查看实时日志
docker-compose logs -f easyfilepreview

# 查看最近的日志
docker-compose logs --tail=100 easyfilepreview

# 查看错误日志
docker-compose logs easyfilepreview | grep ERROR
```

### 进入容器

```bash
# 进入生产环境容器
docker exec -it easyfilepreview sh

# 进入开发环境容器
docker exec -it easyfilepreview-dev sh
```

## 性能优化

### 资源限制

在 `docker-compose.yml` 中添加资源限制：

```yaml
services:
  easyfilepreview:
    # ... 其他配置
    deploy:
      resources:
        limits:
          cpus: '1.0'
          memory: 1G
        reservations:
          cpus: '0.5'
          memory: 512M
```

### 多实例部署

```yaml
services:
  easyfilepreview:
    # ... 其他配置
    deploy:
      replicas: 3
      update_config:
        parallelism: 1
        delay: 10s
      restart_policy:
        condition: on-failure
```

## 故障排除

### 常见问题

1. **端口被占用**
   ```bash
   # 检查端口占用
   lsof -i :3000
   
   # 修改端口映射
   # 在 docker-compose.yml 中修改 ports 配置
   ```

2. **权限问题**
   ```bash
   # 修复目录权限
   sudo chown -R $USER:$USER uploads cache logs
   ```

3. **内存不足**
   ```bash
   # 清理 Docker 资源
   docker system prune -a
   
   # 增加 swap 空间
   sudo fallocate -l 2G /swapfile
   sudo chmod 600 /swapfile
   sudo mkswap /swapfile
   sudo swapon /swapfile
   ```

4. **网络问题**
   ```bash
   # 检查网络连接
   docker network ls
   docker network inspect easyfilepreview_easyfilepreview-network
   ```

### 调试模式

```bash
# 以调试模式启动
docker-compose run --rm easyfilepreview node --inspect=0.0.0.0:9229 index.js

# 或者修改 Dockerfile 添加调试端口
EXPOSE 9229
```

## 安全建议

1. **使用非 root 用户**
   ```dockerfile
   # 在 Dockerfile 中添加
   RUN addgroup -g 1001 -S nodejs
   RUN adduser -S nodejs -u 1001
   USER nodejs
   ```

2. **限制文件系统访问**
   ```yaml
   # 在 docker-compose.yml 中添加
   security_opt:
     - no-new-privileges:true
   read_only: true
   tmpfs:
     - /tmp
     - /var/tmp
   ```

3. **使用 secrets 管理敏感信息**
   ```yaml
   # 在 docker-compose.yml 中添加
   secrets:
     - db_password
     - api_key
   ```

## 更新部署

```bash
# 拉取最新代码
git pull

# 重新构建并部署
./scripts/docker-deploy.sh prod

# 或者手动更新
docker-compose down
docker-compose build --no-cache
docker-compose up -d
```

## 备份和恢复

### 备份数据

```bash
# 备份上传文件
tar -czf uploads-backup-$(date +%Y%m%d).tar.gz uploads/

# 备份缓存
tar -czf cache-backup-$(date +%Y%m%d).tar.gz cache/

# 备份日志
tar -czf logs-backup-$(date +%Y%m%d).tar.gz logs/
```

### 恢复数据

```bash
# 恢复上传文件
tar -xzf uploads-backup-20250101.tar.gz

# 恢复缓存
tar -xzf cache-backup-20250101.tar.gz

# 恢复日志
tar -xzf logs-backup-20250101.tar.gz
```

## 生产环境建议

1. **使用反向代理**
   - Nginx
   - Traefik
   - HAProxy

2. **配置 SSL/TLS**
   - Let's Encrypt
   - 自签名证书

3. **监控和告警**
   - Prometheus + Grafana
   - ELK Stack
   - 自定义监控脚本

4. **日志管理**
   - 日志轮转
   - 日志聚合
   - 日志分析

5. **备份策略**
   - 定期备份
   - 异地备份
   - 备份验证

## 支持

如果遇到问题，请：

1. 查看日志文件
2. 检查 Docker 容器状态
3. 查看本文档的故障排除部分
4. 提交 Issue 到 GitHub 仓库 