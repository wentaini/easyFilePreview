# EasyFilePreview

一个强大的文件预览系统，支持多种文件格式的在线预览，包括PDF、Office文档、图片、文本文件等。

## 功能特性

### 📄 文件预览
- **PDF文件**: 在线预览，支持缩放、翻页
- **Office文档**: Word、Excel、PowerPoint文档预览
- **图片文件**: JPG、PNG、GIF等格式预览
- **文本文件**: TXT、MD、XML、CSV等格式预览
- **代码文件**: 语法高亮显示

### 🔍 PDF文本提取
- **文本提取**: 从PDF文件中提取纯文本内容
- **分页处理**: 智能分割PDF页面，返回每页的文本内容
- **Markdown转换**: 自动将文本转换为Markdown格式
- **通用表格识别**: 智能识别和转换各种表格格式

### 🎨 用户界面
- **响应式设计**: 支持PC、平板、手机等设备
- **现代化UI**: 使用Tailwind CSS构建的美观界面
- **交互友好**: 直观的操作体验

## 快速开始

### 环境要求
- Node.js 16+
- npm 或 cnpm

### 安装依赖
```bash
npm install
# 或使用cnpm
cnpm install
```

### 启动服务
```bash
node index.js
```

服务将在 `http://localhost:3000` 启动

## API文档

### 文件预览API

#### GET /api/filePreview/preview
获取文件预览信息

**参数:**
- `url`: 文件URL地址（必填）
- `includeHiddenSheets`: 是否包含隐藏的sheet（仅对Excel文件有效，可选，默认为false）

**示例:**
```bash
# 基本预览
GET /api/filePreview/preview?url=https://example.com/document.pdf

# Excel文件预览（包含隐藏sheet）
GET /api/filePreview/preview?url=https://example.com/document.xlsx&includeHiddenSheets=true
```

**响应格式:**
```json
{
  "success": true,
  "message": "文件预览成功",
  "data": {
    "fileInfo": {
      "url": "文件URL",
      "extension": "文件扩展名",
      "contentType": "MIME类型",
      "fileName": "文件名",
      "isSupported": true
    },
    "preview": {
      "type": "文件类型（excel/pdf/doc等）",
      "sheets": {},
      "sheetNames": [],
      "sheetInfo": {},
      "images": {},
      "hasHiddenSheets": false,
      "hiddenSheetNames": []
    }
  }
}
```

**Excel文件预览响应说明:**

当文件类型为Excel（.xls, .xlsx）时，响应包含以下字段：

| 字段 | 类型 | 描述 |
|------|------|------|
| type | string | 固定为 "excel" |
| sheets | object | 工作表数据对象，key为工作表名称，value为二维数组 |
| sheetNames | array | 返回的工作表名称列表（默认不包含隐藏的sheet） |
| sheetInfo | object | 工作表信息，包含每个工作表的行数、列数、图片数量等 |
| images | object | 工作表图片对象，key为工作表名称，value为图片数组 |
| hasHiddenSheets | boolean | 是否存在隐藏的工作表 |
| hiddenSheetNames | array | 隐藏的工作表名称列表 |

**sheetInfo 字段说明:**
```json
{
  "工作表名称": {
    "maxRow": 36,
    "maxCol": 18,
    "dataLength": 36,
    "imageCount": 0,
    "isHidden": false
  }
}
```

**includeHiddenSheets 参数说明:**
- 默认值: `false`
- 当设置为 `true` 时，返回结果中的 `sheets` 对象将包含所有工作表（包括隐藏的工作表）
- 隐藏的工作表信息会在 `sheetInfo` 中标记 `isHidden: true`
- 即使 `includeHiddenSheets=false`，响应中仍会包含 `hasHiddenSheets` 和 `hiddenSheetNames` 字段，用于告知客户端是否存在隐藏的工作表

### PDF文本提取API

#### GET /api/filePreview/pdfText
从PDF文件中提取文本内容

**参数:**
- `url`: PDF文件URL地址

**示例:**
```bash
GET /api/filePreview/pdfText?url=https://example.com/document.pdf
```

**响应格式:**
```json
{
  "success": true,
  "data": {
    "text": "完整的PDF文本内容",
    "markdown": "转换为Markdown格式的内容",
    "pages": [
      {
        "pageNumber": 1,
        "text": "第1页的文本内容"
      }
    ],
    "markdownPages": [
      {
        "pageNumber": 1,
        "markdown": "第1页的Markdown内容"
      }
    ]
  }
}
```

## 通用表格识别功能

系统支持智能识别和转换多种表格格式：

### 支持的表格类型

1. **标准表格格式**
   ```
   姓名  年龄  职业  城市
   张三  25   工程师  北京
   李四  30   设计师  上海
   ```

2. **键值对格式**
   ```
   姓名: 张三
   年龄: 25
   职业: 工程师
   ```

3. **对齐列格式**
   ```
   产品名称    价格    库存    状态
   苹果手机    6999    100     有货
   华为平板    3999    50      缺货
   ```

4. **单行表格格式**
   ```
   序号 产品 价格 库存 1 苹果 6999 100 2 华为 3999 50
   ```

### 表格转换示例

**输入（键值对格式）：**
```
户名: 张三
账号: 1234567890123456
金额: ¥1000.00
```

**输出（Markdown表格）：**
```markdown
| 字段 | 值 |
| --- | --- |
| 户名 | 张三 |
| 账号 | 1234567890123456 |
| 金额 | ¥1000.00 |
```

### 技术特点

- **通用性**: 不依赖特定文档格式，支持多种表格结构
- **准确性**: 基于格式特征而非内容关键词，避免误判
- **灵活性**: 支持单行和多行表格，不同分隔符
- **可扩展性**: 模块化设计，易于添加新的表格类型

## 项目结构

```
easyfilePreview/
├── src/
│   ├── public/           # 静态资源
│   │   ├── css/         # 样式文件
│   │   ├── js/          # JavaScript文件
│   │   └── demoDoc/     # 示例文档
│   ├── router.js        # 路由配置
│   └── utils/           # 工具函数
│       ├── filePreview.js           # 文件预览核心逻辑
│       ├── filePreviewHandler.js    # 预览处理器
│       └── excelImageExtractor.js   # Excel图片提取
├── config/              # 配置文件
├── docs/                # 文档
├── scripts/             # 脚本文件
├── index.js             # 应用入口
└── package.json         # 项目配置
```

## 技术栈

### 前端
- **Tailwind CSS**: 现代化CSS框架 (本地: ./js/tailwindcss.js)
- **Vue.js**: 渐进式JavaScript框架 (本地: ./js/vue.global.js)
- **Chart.js**: 图表库 (本地: ./js/chart.js)
- **Marked**: Markdown解析器 (本地: ./js/marked.min.js)

### 后端
- **Node.js**: JavaScript运行时
- **Express**: Web应用框架
- **pdf-parse**: PDF解析库
- **mammoth**: Word文档解析
- **xlsx**: Excel文件处理

## 配置说明

### 环境变量
复制 `env.example` 为 `.env` 并配置：

```bash
# 服务器配置
PORT=3000
NODE_ENV=development

# 文件缓存配置
CACHE_DURATION=3600000
MAX_FILE_SIZE=52428800
```

### 文件格式支持

| 格式 | 扩展名 | 预览方式 |
|------|--------|----------|
| PDF | .pdf | 在线预览 |
| Word | .doc, .docx | 转换为HTML |
| Excel | .xls, .xlsx | 转换为HTML |
| PowerPoint | .ppt, .pptx | 转换为HTML |
| 图片 | .jpg, .jpeg, .png, .gif | 直接显示 |
| 文本 | .txt, .md, .xml, .csv | 语法高亮 |

## 使用示例

### 基本预览
```javascript
// 获取文件预览信息
const response = await fetch('/api/filePreview/preview?url=https://example.com/document.pdf');
const data = await response.json();

if (data.success) {
    console.log('文件类型:', data.data.preview.type);
    console.log('文件信息:', data.data.fileInfo);
}
```

### Excel文件预览
```javascript
// 预览Excel文件（不包含隐藏sheet）
const response = await fetch('/api/filePreview/preview?url=https://example.com/document.xlsx');
const data = await response.json();

if (data.success && data.data.preview.type === 'excel') {
    const preview = data.data.preview;
    console.log('工作表列表:', preview.sheetNames);
    console.log('是否存在隐藏sheet:', preview.hasHiddenSheets);
    console.log('隐藏的sheet名称:', preview.hiddenSheetNames);
    
    // 访问工作表数据
    preview.sheetNames.forEach(sheetName => {
        const sheetData = preview.sheets[sheetName];
        console.log(`${sheetName} 数据:`, sheetData);
    });
}

// 预览Excel文件（包含隐藏sheet）
const responseWithHidden = await fetch(
    '/api/filePreview/preview?url=https://example.com/document.xlsx&includeHiddenSheets=true'
);
const dataWithHidden = await responseWithHidden.json();

if (dataWithHidden.success && dataWithHidden.data.preview.type === 'excel') {
    const preview = dataWithHidden.data.preview;
    // 现在 sheets 对象包含所有工作表，包括隐藏的
    console.log('所有工作表（包括隐藏的）:', Object.keys(preview.sheets));
    
    // 检查哪些是隐藏的
    Object.keys(preview.sheetInfo).forEach(sheetName => {
        if (preview.sheetInfo[sheetName].isHidden) {
            console.log(`隐藏的工作表: ${sheetName}`);
        }
    });
}
```

### PDF文本提取
```javascript
// 提取PDF文本
const response = await fetch('/api/filePreview/pdfText?url=https://example.com/document.pdf');
const data = await response.json();

if (data.success) {
    // 检查是否包含表格
    const hasTable = data.data.markdown.includes('|');
    
    if (hasTable) {
        console.log('PDF包含表格数据');
        console.log('Markdown表格:', data.data.markdown);
    }
    
    // 分页处理
    data.data.markdownPages.forEach(page => {
        const hasTable = page.markdown.includes('|');
        if (hasTable) {
            console.log(`第${page.pageNumber}页包含表格`);
        }
    });
}
```

## 部署

### Docker部署
```bash
# 构建镜像
docker build -t easyfilepreview .

# 运行容器
docker run -p 3000:3000 easyfilepreview
```

### 生产环境
```bash
# 使用PM2管理进程
npm install -g pm2
pm2 start index.js --name easyfilepreview

# 设置开机自启
pm2 startup
pm2 save
```

## 开发

### 代码规范
- 使用ESLint + Prettier
- 遵循Node.js项目目录规范
- 使用async/await避免回调地狱

### 测试
```bash
# 运行测试
npm test

# 代码检查
npm run lint
```

## 贡献

欢迎提交Issue和Pull Request！

### 开发流程
1. Fork项目
2. 创建功能分支
3. 提交更改
4. 创建Pull Request

## 许可证

MIT License

## 更新日志

### v1.2.0
- ✨ Excel文件支持隐藏sheet检测
- ✨ 新增 `includeHiddenSheets` 参数，支持返回隐藏sheet内容
- 🔧 优化Excel文件处理逻辑
- 📚 更新API文档，添加Excel预览详细说明

### v1.1.0
- ✨ 添加通用表格识别功能
- ✨ 支持多种表格类型转换
- 🔧 改进表格检测算法
- 🐛 修复占位符保护机制
- 📚 完善API文档

### v1.0.0
- 🎉 初始版本发布
- ✨ 基础文件预览功能
- ✨ PDF文本提取功能
- ✨ Markdown转换功能
- 📱 响应式界面设计

## 支持

如有问题，请提交Issue或联系开发团队。 