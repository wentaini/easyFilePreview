const express = require('express')
const cors = require('cors')
const bodyParser = require('body-parser')
const path = require('path');
const app = express()

// 导入文件预览模块
const { filePreviewHandler, fileInfoHandler, supportedFormatsHandler, pdfTextHandler } = require('./utils/filePreviewHandler')
const filePreview = require('./utils/filePreview')


// 移除全局body-parser，改为为每个路由单独配置
// app.use(bodyParser.json({ limit: '10mb' }))

// SQL 注入防御中间件
function sqlInjectionDefense(req, res, next) {
  const sanitize = (value) => {
      if (typeof value === 'string') {
          return value.replace(/['";]/g, '');
      }
      return value;
  };

  const sanitizeObject = (obj) => {
      for (let key in obj) {
          if (obj.hasOwnProperty(key)) {
              obj[key] = sanitize(obj[key]);
          }
      }
  };

  sanitizeObject(req.body);
  sanitizeObject(req.query);
  sanitizeObject(req.params);

  next();
}

// 使用 SQL 注入防御中间件
app.use(sqlInjectionDefense);

// 跨域配置
// 1. 定义 CORS 配置
const corsOptions = {
    // 【1】指定允许的源
    // 开发环境调试可以设为 '*'
    // 生产环境建议只允许特定域名，如：origin: ['https://www.****.com']
    origin: '*',
  
    // 【2】指定允许的请求方法
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  
    // 【3】显式列出前端会发送的自定义请求头
    // 如果你的请求中包含 X-TI-Nonce、X-TI-Signature、X-TI-Timestamp 等，就要全部写上
    // 注意：authorization != authentication，如果你用的是 "authentication" 头，则要写对
    allowedHeaders: [
      'Accept',
      'Authorization',      // 如果用 'Authorization' 头
      'Authentication',     // 如果用 'Authentication' 头
      'Content-Type',
      'Origin',
      'X-Requested-With',
      'Application-Key',
      'Cache-Control',
      'Expires',
      'Pragma',
      'X-Real-Hostname',
      'X-TI-Nonce',
      'X-TI-Signature',
      'X-TI-Timestamp'
    ],
  
    // 【4】是否允许客户端携带 Cookie、凭证等
    credentials: true,
  
    // 【5】遇到 OPTIONS 请求时，如果想让 cors 处理完成后继续到下一个中间件，可设为 true
    // 如果不需要，默认 false 即可
    preflightContinue: false,
  
    // 【6】某些老浏览器对 204 状态不兼容时可改成 200
    optionsSuccessStatus: 204
  };

app.use(cors(corsOptions));

// 确保options请求被正确响应
app.options('*', cors(corsOptions));

// 启动静态资源服务
app.use(express.static(path.join(__dirname, 'public')))

// 根路径 - 重定向到demo页面
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'filePreviewDemo.html'));
});

// 文件预览相关路由
/**
 * @api {get} /api/filePreview/preview 文件预览
 * @apiName FilePreview
 * @apiGroup 文件预览
 * @apiVersion 1.0.0
 * 
 * @apiParam {String} url 文件URL
 * 
 * @apiSuccess {Object} data 预览结果
 * @apiSuccess {Object} data.fileInfo 文件信息
 * @apiSuccess {Object} data.preview 预览内容
 * 
 * @apiExample {curl} 示例:
 *     curl -X GET "http://localhost/api/filePreview/preview?url=https://example.com/file.pdf"
 */
app.get('/api/filePreview/preview', filePreviewHandler);

/**
 * @api {get} /api/filePreview/info 获取文件信息
 * @apiName FileInfo
 * @apiGroup 文件预览
 * @apiVersion 1.0.0
 * 
 * @apiParam {String} url 文件URL
 * 
 * @apiSuccess {Object} data 文件信息
 * 
 * @apiExample {curl} 示例:
 *     curl -X GET "http://localhost/api/filePreview/info?url=https://example.com/file.pdf"
 */
app.get('/api/filePreview/info', fileInfoHandler);

/**
 * @api {get} /api/filePreview/formats 获取支持的文件格式
 * @apiName SupportedFormats
 * @apiGroup 文件预览
 * @apiVersion 1.0.0
 * 
 * @apiSuccess {Object} data 支持的文件格式列表
 * 
 * @apiExample {curl} 示例:
 *     curl -X GET "http://localhost/api/filePreview/formats"
 */
app.get('/api/filePreview/formats', supportedFormatsHandler);

/**
 * @api {get} /api/filePreview/pdfText 获取PDF文本内容
 * @apiName PdfText
 * @apiGroup 文件预览
 * @apiVersion 1.0.0
 * 
 * @apiParam {String} url 文件URL
 * 
 * @apiSuccess {Object} data PDF文本内容
 * @apiSuccess {String} data.text 提取的文本内容
 * @apiSuccess {Number} data.pageCount 页面数量
 * @apiSuccess {Boolean} data.hasText 是否包含文本
 * @apiSuccess {String} data.fileSize 文件大小
 * 
 * @apiExample {curl} 示例:
 *     curl -X GET "http://localhost/api/filePreview/pdfText?url=https://example.com/file.pdf"
 */
app.get('/api/filePreview/pdfText', pdfTextHandler);

/**
 * @api {get} /api/filePreview/stream/:fileId 文件流下载
 * @apiName FileStream
 * @apiGroup FilePreview
 * @apiParam {String} fileId 文件ID
 * @apiSuccess {Buffer} 文件内容
 * @apiExample {curl} Example usage:
 *     curl -X GET "http://localhost/api/filePreview/stream/file_1234567890"
 */
app.get('/api/filePreview/stream/:fileId', (req, res) => {
    const fileId = req.params.fileId;
    const fileData = filePreview.getCachedFile(fileId);
    
    if (!fileData) {
        return res.status(404).json({
            success: false,
            message: '文件不存在或已过期'
        });
    }
    
    // 设置响应头
    res.setHeader('Content-Type', fileData.type === 'pdf' ? 'application/pdf' : 'application/octet-stream');
    res.setHeader('Content-Disposition', `inline; filename="${fileData.fileName}"`);
    res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');
    
    // 发送文件内容
    res.send(fileData.buffer);
});

/**
 * 运行名称
 * node index.js prod | dev 
 */
const port = 3000
const server = app.listen(port, () => {
    console.log(`app listening on port ${port}!`)
})

// 设置服务器超时时间为600秒
server.timeout = 600000; // 600秒 = 600000毫秒
console.log('服务器超时时间设置为: 600秒');

module.exports = app
