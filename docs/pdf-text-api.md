# PDF文本提取API文档

## 概述

PDF文本提取API提供了强大的PDF文档内容提取功能，支持文本提取、分页处理、Markdown转换和通用表格识别。

## API端点

### GET /api/filePreview/pdfText

从PDF文件中提取文本内容，支持分页处理和Markdown转换。

#### 请求参数

| 参数 | 类型 | 必填 | 描述 |
|------|------|------|------|
| url | string | 是 | PDF文件的URL地址 |

#### 请求示例

```bash
GET /api/filePreview/pdfText?url=https://example.com/document.pdf
```

#### 响应格式

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

#### 响应字段说明

| 字段 | 类型 | 描述 |
|------|------|------|
| text | string | 完整的PDF文本内容 |
| markdown | string | 转换为Markdown格式的内容 |
| pages | array | 分页的文本内容数组 |
| markdownPages | array | 分页的Markdown内容数组 |

## 功能特性

### 1. 文本提取
- 支持各种PDF格式的文本提取
- 保持原始文本格式和结构
- 自动处理编码问题

### 2. 分页处理
- 智能分页算法
- 保持页面边界
- 支持多页PDF文档

### 3. Markdown转换
- 自动转换为Markdown格式
- 支持标题、列表、强调等格式
- 保持文档结构

### 4. 通用表格识别

#### 支持的表格类型

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

#### 表格检测算法

系统通过以下特征识别表格：
- 列数一致性检测
- 编号行检测
- 表格结构检测
- 键值对格式检测
- 表格分隔符检测
- 对齐数据检测

#### 表格转换示例

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

## 错误处理

### 常见错误

| 状态码 | 错误信息 | 描述 |
|--------|----------|------|
| 400 | 缺少URL参数 | 请求中缺少必需的url参数 |
| 400 | 无效的URL格式 | 提供的URL格式不正确 |
| 500 | 文件下载失败 | 无法下载指定的PDF文件 |
| 500 | PDF解析失败 | PDF文件解析过程中出现错误 |

### 错误响应格式

```json
{
  "success": false,
  "message": "错误描述信息"
}
```

## 使用示例

### JavaScript示例

```javascript
// 获取PDF文本内容
const response = await fetch('/api/filePreview/pdfText?url=https://example.com/file.pdf');
const data = await response.json();

if (data.success) {
    // 检查是否包含表格
    const hasTable = data.data.markdown.includes('|');
    
    if (hasTable) {
        console.log('PDF包含表格数据');
        console.log('Markdown表格:');
        console.log(data.data.markdown);
    }
    
    // 检查分页表格
    data.data.markdownPages.forEach(page => {
        const hasTable = page.markdown.includes('|');
        if (hasTable) {
            console.log(`第${page.pageNumber}页包含表格`);
            console.log(page.markdown);
        }
    });
}
```

### cURL示例

```bash
curl -X GET "http://localhost:3000/api/filePreview/pdfText?url=https://example.com/document.pdf"
```

## 技术特点

### 1. 通用性
- 不依赖特定文档格式
- 支持多种表格结构
- 自动识别表格类型

### 2. 准确性
- 基于格式特征而非内容关键词
- 严格的检测条件避免误判
- 智能的表格类型识别

### 3. 灵活性
- 支持单行和多行表格
- 支持不同分隔符
- 支持对齐和非对齐数据

### 4. 可扩展性
- 模块化的检测算法
- 易于添加新的表格类型
- 可配置的检测参数

## 注意事项

1. **文件大小限制**：建议PDF文件大小不超过50MB
2. **网络超时**：下载大文件时可能需要较长时间
3. **表格识别**：表格识别基于格式特征，可能不适用于所有复杂表格
4. **编码支持**：支持UTF-8编码的PDF文件

## 更新日志

### v1.0.0
- 基础PDF文本提取功能
- 分页处理支持
- Markdown转换功能

### v1.1.0
- 添加通用表格识别功能
- 支持多种表格类型
- 改进表格转换算法
- 修复占位符保护机制
