# 功能特性

## 核心功能

### 📄 文件预览
- **PDF文件**: 在线预览，支持缩放、翻页
- **Office文档**: Word、Excel、PowerPoint文档预览
  - **Excel增强功能**: 
    - 支持检测隐藏的工作表
    - 可通过参数控制是否返回隐藏工作表内容
    - 自动识别工作表可见性状态
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

## 通用表格识别

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

### 技术特点

- **通用性**: 不依赖特定文档格式，支持多种表格结构
- **准确性**: 基于格式特征而非内容关键词，避免误判
- **灵活性**: 支持单行和多行表格，不同分隔符
- **可扩展性**: 模块化设计，易于添加新的表格类型

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

## 文件格式支持

| 格式 | 扩展名 | 预览方式 |
|------|--------|----------|
| PDF | .pdf | 在线预览 |
| Word | .doc, .docx | 转换为HTML |
| Excel | .xls, .xlsx | 转换为HTML |
| PowerPoint | .ppt, .pptx | 转换为HTML |
| 图片 | .jpg, .jpeg, .png, .gif | 直接显示 |
| 文本 | .txt, .md, .xml, .csv | 语法高亮 |

## Excel文件预览增强

### 隐藏工作表支持
- **自动检测**: 系统会自动检测Excel文件中的隐藏工作表
- **灵活控制**: 通过 `includeHiddenSheets` 参数控制是否返回隐藏工作表内容
- **状态标识**: 响应中包含 `hasHiddenSheets` 和 `hiddenSheetNames` 字段，方便客户端判断
- **详细信息**: 每个工作表的信息中包含 `isHidden` 字段，标识工作表是否隐藏

### 使用示例
```javascript
// 默认不包含隐藏sheet
GET /api/filePreview/preview?url=https://example.com/file.xlsx

// 包含隐藏sheet
GET /api/filePreview/preview?url=https://example.com/file.xlsx&includeHiddenSheets=true
```

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
