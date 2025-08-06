# EasyFilePreview 文件预览组件

> 轻量级文件预览组件，支持 PDF、Office、CSV、Markdown 等多格式响应式预览

## 📋 功能特性

### 多格式支持
- **PDF文档** (.pdf)
- **Excel文件** (.xls, .xlsx)
- **Word文档** (.doc, .docx)
- **PowerPoint** (.ppt, .pptx)
- **CSV文件** (.csv)
- **Markdown** (.md, .markdown)
- **文本文件** (.txt)
- **XML文档** (.xml)

### 核心功能
1. **响应式设计** - 完美适配PC、平板、手机
2. **iframe嵌入** - 可轻松嵌入到其他页面
3. **一键复制** - Excel和CSV支持单元格复制
4. **极简界面** - 简洁美观的用户界面
5. **快速加载** - 轻量级，性能优异

## 🚀 快速开始

### 安装依赖

```bash
# 使用cnpm安装
cnpm install

# 启动开发服务器
npm run dev
```

### 基本使用

```html
<!-- 完整预览模式 -->
<iframe src="http://localhost:3000/filePreview.html?url=YOUR_FILE_URL" 
        width="100%" height="600"></iframe>

<!-- 简洁预览模式 -->
<iframe src="http://localhost:3000/filePreviewSimple.html?url=YOUR_FILE_URL" 
        width="100%" height="600"></iframe>
```

## 📊 API接口

### 文件预览接口

```javascript
GET /api/filePreview/preview?url={fileUrl}

// 响应示例
{
  "success": true,
  "data": {
    "type": "pdf",
    "fileName": "document.pdf",
    "contentType": "application/pdf",
    "fileUrl": "https://example.com/document.pdf"
  }
}
```

### 获取文件信息

```javascript
GET /api/filePreview/info?url={fileUrl}

// 响应示例
{
  "success": true,
  "data": {
    "fileName": "document.pdf",
    "extension": "pdf",
    "contentType": "application/pdf",
    "isSupported": true
  }
}
```

### 支持格式列表

```javascript
GET /api/filePreview/formats

// 响应示例
{
  "success": true,
  "data": {
    "formats": [
      {
        "extension": "pdf",
        "mimeType": "application/pdf",
        "description": "PDF文档"
      }
    ],
    "count": 12
  }
}
```

## 🎨 技术栈

### 前端技术
| 技术 | 版本 | 用途 |
|------|------|------|
| **TailwindCSS** | 最新版 | UI框架 |
| **Vue.js** | 3.x | MVVM框架 |
| **Chart.js** | 最新版 | 图表展示 |
| **Marked.js** | 最新版 | Markdown渲染 |

### 后端技术
| 技术 | 版本 | 用途 |
|------|------|------|
| **Express.js** | 4.x | Web框架 |
| **Multer** | 最新版 | 文件上传 |
| **Joi** | 最新版 | 参数校验 |
| **Winston** | 最新版 | 日志管理 |

## 📱 响应式设计

### 断点设置
```css
/* 移动端 */
@media (max-width: 768px) {
  .container { padding: 10px; }
  .excel-table { font-size: 12px; }
}

/* 平板端 */
@media (min-width: 769px) and (max-width: 1024px) {
  .container { padding: 20px; }
}

/* 桌面端 */
@media (min-width: 1025px) {
  .container { padding: 40px; }
}
```

## 🔧 配置选项

### 环境变量
```bash
# 服务器配置
PORT=3000
HOST=localhost

# 跨域配置
CORS_ORIGIN=*

# 文件上传配置
MAX_FILE_SIZE=52428800
UPLOAD_DIR=./uploads

# 缓存配置
CACHE_DIR=./cache
MAX_CACHE_SIZE=104857600
CACHE_EXPIRE=86400000
```

### 安全配置
```javascript
// CORS配置
const corsOptions = {
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true
};

// 内容安全策略
const cspOptions = {
  directives: {
    defaultSrc: ["'self'"],
    styleSrc: ["'self'", "'unsafe-inline'"],
    scriptSrc: ["'self'"],
    imgSrc: ["'self'", "data:", "https:"]
  }
};
```

## 📈 性能优化

### 缓存策略
- **文件缓存** - 已预览的文件缓存24小时
- **CDN加速** - 静态资源使用CDN
- **压缩传输** - 启用gzip压缩
- **懒加载** - 大文件分片加载

### 内存管理
```javascript
// 缓存清理
setInterval(() => {
  const cacheSize = getCacheSize();
  if (cacheSize > MAX_CACHE_SIZE) {
    cleanExpiredCache();
  }
}, 60000); // 每分钟检查一次
```

## 🐛 常见问题

### Q: 如何解决跨域问题？
A: 确保服务器配置了正确的CORS头：
```javascript
app.use(cors({
  origin: '*',
  credentials: true
}));
```

### Q: 大文件加载慢怎么办？
A: 可以启用分片加载：
```javascript
// 分片加载配置
const chunkSize = 1024 * 1024; // 1MB
const chunks = Math.ceil(fileSize / chunkSize);
```

### Q: 如何自定义样式？
A: 可以通过CSS变量自定义：
```css
:root {
  --primary-color: #3b82f6;
  --border-color: #e5e7eb;
  --text-color: #374151;
}
```

## 📄 许可证

本项目采用 [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0) 许可证。

## 🤝 贡献指南

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 📞 联系方式

- **项目地址**: [https://github.com/wentaini/easyFilePreview](https://github.com/wentaini/easyFilePreview)
- **问题反馈**: [Issues](https://github.com/wentaini/easyFilePreview/issues)
- **邮箱**: support@example.com

---

> 💡 **提示**: 这是一个演示文档，展示了Markdown预览功能的各种特性，包括标题、列表、表格、代码块、引用等格式。

*最后更新时间: 2025年8月*
