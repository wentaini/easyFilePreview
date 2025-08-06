# EasyFilePreview

轻量级文件预览组件，支持 PDF、Office、CSV、Markdown 等多格式响应式预览，可 iframe 嵌入，提供极简界面与单元格一键复制。

## 功能特性

- 📄 **多格式支持**: PDF、Office文档、CSV、Markdown等
- 📱 **响应式设计**: 支持PC、PAD、MOBILE设备
- 🔗 **iframe嵌入**: 可轻松嵌入到其他页面
- 📋 **一键复制**: 支持单元格内容快速复制
- 🎨 **极简界面**: 简洁美观的用户界面
- ⚡ **轻量级**: 快速加载，性能优异

## 技术栈

### 前端
- **UI框架**: TailwindCSS
- **MVVM**: Vue.js
- **Markdown渲染**: Marked.js

### 后端
- **服务端**: Express.js
- **文件处理**: Multer
- **参数校验**: Joi
- **日志**: Winston

## 项目结构

```
easyfilePreview/
├── src/               # 主代码目录
│   ├── index.js       # 应用入口
│   ├── router.js      # 路由定义
│   ├── public/        # 静态资源
│   ├── middlewares/   # 中间件
│   └── utils/         # 工具函数
├── config/            # 配置文件
├── docs/              # 文档
├── index.js           # 主入口
└── package.json       # 项目配置
```

## 快速开始

### 安装依赖

```bash
# 使用cnpm安装
cnpm install
```

### 开发环境

```bash
# 启动开发服务器
npm run dev
```

### 生产环境

```bash
# 启动生产服务器
npm start
```

## API文档

### 文件预览接口

```
GET /api/filePreview/preview?url={fileUrl}
```

### 文件信息接口

```
GET /api/filePreview/info?url={fileUrl}
```

### 支持格式接口

```
GET /api/filePreview/supportedFormats
```

## 使用示例

### 基本预览

```html
<iframe src="http://localhost:3000/api/filePreview/preview?url=https://example.com/file.pdf" 
        width="100%" height="600px">
</iframe>
```

### 简单预览

```html
<iframe src="http://localhost:3000/api/filePreview/preview?url=https://example.com/file.xlsx&simple=true" 
        width="100%" height="600px">
</iframe>
```

## 许可证

Apache License 2.0

## 贡献

欢迎提交Issue和Pull Request！ 