# EasyFilePreview 📄

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/Node.js-18+-green.svg)](https://nodejs.org/)
[![Express](https://img.shields.io/badge/Express-4.19+-orange.svg)](https://expressjs.com/)
[![Vue.js](https://img.shields.io/badge/Vue.js-3.0+-green.svg)](https://vuejs.org/)

> 🚀 轻量级文件预览组件，支持 PDF、Office、CSV、Markdown 等多格式响应式预览，可 iframe 嵌入，提供极简界面与单元格一键复制。

## ✨ 功能特性

### 📄 多格式支持
- **PDF文档**: 原生PDF预览，支持缩放、翻页
- **Office文档**: Word、Excel、PowerPoint 文件预览
- **文本文件**: Markdown、XML、TXT 等格式
- **表格文件**: CSV、Excel 表格数据展示
- **图片文件**: 常见图片格式预览

### 📱 响应式设计
- **PC端**: 完整功能界面，支持复杂操作
- **平板端**: 适配中等屏幕，优化触控体验
- **移动端**: 精简界面，支持手势操作

### 🔗 集成特性
- **iframe嵌入**: 可轻松嵌入到任何网页
- **API接口**: RESTful API，支持第三方集成
- **跨域支持**: 完整的CORS配置
- **缓存机制**: 智能文件缓存，提升性能

### 🎨 用户体验
- **一键复制**: 表格单元格内容快速复制
- **极简界面**: 简洁美观的用户界面
- **加载动画**: 优雅的加载状态提示
- **错误处理**: 友好的错误信息展示

## 🛠️ 技术栈

### 前端技术
| 技术 | 版本 | 用途 |
|------|------|------|
| **TailwindCSS** | 最新 | UI框架，响应式样式 |
| **Vue.js** | 3.0+ | MVVM框架，数据绑定 |
| **Marked.js** | 最新 | Markdown渲染 |
| **Chart.js** | 最新 | 图表展示 |

### 后端技术
| 技术 | 版本 | 用途 |
|------|------|------|
| **Express.js** | 4.19+ | Web框架 |
| **axios** | 1.8+ | HTTP客户端 |
| **pdf-parse** | 1.1+ | PDF解析 |
| **xlsx** | 0.18+ | Excel处理 |
| **marked** | 16.1+ | Markdown解析 |
| **xml2js** | 0.6+ | XML解析 |

## 📁 项目结构

```
easyfilePreview/
├── src/                    # 主代码目录
│   ├── router.js           # Express路由配置
│   ├── public/             # 静态资源
│   │   ├── css/            # 样式文件
│   │   │   ├── common.css  # 通用样式
│   │   │   ├── demo.css    # 演示页面样式
│   │   │   ├── preview.css # 预览页面样式
│   │   │   └── simple.css  # 简化页面样式
│   │   ├── js/             # JavaScript库
│   │   ├── demoDoc/        # 演示文档
│   │   ├── filePreview.html        # 主预览页面
│   │   ├── filePreviewDemo.html    # 演示页面
│   │   └── filePreviewSimple.html  # 简化预览页面
│   └── utils/              # 工具函数
│       ├── filePreview.js          # 文件预览核心逻辑
│       ├── filePreviewHandler.js   # 预览处理器
│       └── excelImageExtractor.js  # Excel图片提取
├── config/                 # 配置文件
│   └── default.js          # 默认配置
├── docs/                   # 文档
│   └── docker-deployment.md # Docker部署文档
├── scripts/                # 脚本文件
│   └── docker-deploy.sh    # Docker部署脚本
├── Dockerfile              # 生产环境Docker配置
├── Dockerfile.dev          # 开发环境Docker配置
├── docker-compose.yml      # Docker编排配置
├── index.js                # 应用入口
├── package.json            # 项目配置
├── README.md               # 项目说明
└── LICENSE                 # 开源协议
```

## 🚀 快速开始

### 环境要求

- **Node.js**: 18.0 或更高版本
- **npm**: 8.0 或更高版本
- **操作系统**: Windows、macOS、Linux

### 安装部署

#### 方式一：本地安装

```bash
# 1. 克隆项目
git clone https://github.com/wentaini/easyFilePreview.git
cd easyFilePreview

# 2. 安装依赖
npm install
# 或使用cnpm（推荐国内用户）
cnpm install

# 3. 启动服务
npm start
```

#### 方式二：Docker部署

```bash
# 1. 构建并启动（生产环境）
docker-compose up --build

# 2. 开发环境
docker-compose -f docker-compose.yml up easyfilepreview-dev
```

#### 方式三：一键部署脚本

```bash
# 使用部署脚本
./scripts/docker-deploy.sh prod
```

### 访问地址

- **演示页面**: http://localhost:3000/
- **主预览页面**: http://localhost:3000/filePreview.html
- **简化预览页面**: http://localhost:3000/filePreviewSimple.html

## 📚 API文档

### 基础接口

#### 1. 文件预览接口

```http
GET /api/filePreview/preview?url={fileUrl}
```

**参数说明**
- `url`: 文件URL地址（必需）
- `simple`: 是否使用简化模式（可选，true/false）

**响应示例**
```json
{
  "success": true,
  "data": {
    "fileInfo": {
      "fileName": "example.pdf",
      "fileSize": "1.2MB",
      "fileType": "pdf"
    },
    "preview": {
      "content": "文件预览内容",
      "type": "pdf"
    }
  }
}
```

#### 2. 文件信息接口

```http
GET /api/filePreview/info?url={fileUrl}
```

**响应示例**
```json
{
  "success": true,
  "data": {
    "fileName": "example.xlsx",
    "fileSize": "2.5MB",
    "fileType": "excel",
    "lastModified": "2025-08-07T10:30:00Z"
  }
}
```

#### 3. 支持格式接口

```http
GET /api/filePreview/formats
```

**响应示例**
```json
{
  "success": true,
  "data": {
    "supportedFormats": [
      "pdf", "docx", "xlsx", "pptx", "csv", "md", "xml", "txt"
    ]
  }
}
```

#### 4. 文件流下载接口

```http
GET /api/filePreview/stream/{fileId}
```

### 使用示例

#### 基本预览

```html
<!-- 完整预览模式 -->
<iframe 
  src="http://localhost:3000/filePreview.html?url=https://example.com/document.pdf" 
  width="100%" 
  height="600px"
  frameborder="0">
</iframe>
```

#### 简化预览

```html
<!-- 简化预览模式 -->
<iframe 
  src="http://localhost:3000/filePreviewSimple.html?url=https://example.com/spreadsheet.xlsx" 
  width="100%" 
  height="500px"
  frameborder="0">
</iframe>
```

#### API调用

```javascript
// 获取文件预览
fetch('/api/filePreview/preview?url=https://example.com/file.pdf')
  .then(response => response.json())
  .then(data => {
    console.log('预览数据:', data);
  });

// 获取文件信息
fetch('/api/filePreview/info?url=https://example.com/file.xlsx')
  .then(response => response.json())
  .then(data => {
    console.log('文件信息:', data);
  });
```

## 🔧 配置说明

### 环境变量

创建 `.env` 文件（参考 `env.example`）：

```bash
# 服务器配置
PORT=3000
HOST=0.0.0.0
NODE_ENV=production

# CORS配置
CORS_ORIGIN=*
CORS_CREDENTIALS=true

# 日志配置
LOG_LEVEL=info

# 文件上传配置
MAX_FILE_SIZE=100mb
UPLOAD_PATH=./uploads

# 缓存配置
CACHE_ENABLED=true
CACHE_TIMEOUT=3600000
```

### 自定义配置

编辑 `config/default.js`：

```javascript
module.exports = {
  server: {
    port: process.env.PORT || 3000,
    host: process.env.HOST || '0.0.0.0'
  },
  cors: {
    origin: process.env.CORS_ORIGIN || '*',
    credentials: process.env.CORS_CREDENTIALS === 'true'
  },
  filePreview: {
    maxFileSize: process.env.MAX_FILE_SIZE || '100mb',
    cacheTimeout: process.env.CACHE_TIMEOUT || 3600000
  }
};
```

## 🐳 Docker部署

### 生产环境

```bash
# 构建生产镜像
docker build -t easyfilepreview .

# 运行容器
docker run -d -p 3000:3000 --name easyfilepreview easyfilepreview
```

### 开发环境

```bash
# 构建开发镜像
docker build -f Dockerfile.dev -t easyfilepreview-dev .

# 运行开发容器
docker run -d -p 3000:3000 -v $(pwd):/app --name easyfilepreview-dev easyfilepreview-dev
```

### Docker Compose

```bash
# 启动所有服务
docker-compose up -d

# 查看服务状态
docker-compose ps

# 查看日志
docker-compose logs -f
```

## 🧪 测试

### 单元测试

```bash
# 运行测试
npm test

# 测试覆盖率
npm run test:coverage
```

### 集成测试

```bash
# 启动测试服务器
npm run test:integration
```

## 📊 性能优化

### 缓存策略
- **文件缓存**: 智能缓存已下载文件
- **内存缓存**: 减少重复下载
- **CDN支持**: 支持CDN加速

### 加载优化
- **懒加载**: 按需加载文件内容
- **压缩传输**: Gzip压缩响应
- **静态资源**: 静态文件缓存

## 🔒 安全特性

- **SQL注入防护**: 参数过滤和验证
- **XSS防护**: 输出内容转义
- **CORS配置**: 跨域请求控制
- **文件类型验证**: 严格的文件类型检查
- **大小限制**: 文件大小限制

## 🤝 贡献指南

### 开发流程

1. **Fork项目**
   ```bash
   git clone https://github.com/your-username/easyFilePreview.git
   ```

2. **创建分支**
   ```bash
   git checkout -b feature/your-feature-name
   ```

3. **提交更改**
   ```bash
   git add .
   git commit -m "feat: 添加新功能"
   ```

4. **推送分支**
   ```bash
   git push origin feature/your-feature-name
   ```

5. **创建Pull Request**

### 代码规范

- 使用ESLint进行代码检查
- 遵循JavaScript标准规范
- 添加适当的注释和文档
- 编写单元测试

## 📄 许可证

本项目采用 [Apache License 2.0](LICENSE) 开源协议。

## 🙏 致谢

感谢以下开源项目的支持：

- [Express.js](https://expressjs.com/) - Web应用框架
- [Vue.js](https://vuejs.org/) - 前端框架
- [TailwindCSS](https://tailwindcss.com/) - CSS框架
- [Marked.js](https://marked.js.org/) - Markdown解析器

## 📞 联系方式

- **项目地址**: https://github.com/wentaini/easyFilePreview
- **问题反馈**: [Issues](https://github.com/wentaini/easyFilePreview/issues)
- **功能建议**: [Discussions](https://github.com/wentaini/easyFilePreview/discussions)

---

⭐ 如果这个项目对您有帮助，请给我们一个Star！ 