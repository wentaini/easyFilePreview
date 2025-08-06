const bodyParser = require('body-parser');

/**
 * 为AI聊天接口设置更大的请求体限制
 * 因为AI聊天可能需要处理长文本对话
 */
const aiChatBodyParser = bodyParser.json({ 
    limit: '50mb',
    verify: (req, res, buf) => {
        // 可以在这里添加额外的验证逻辑
        req.rawBody = buf;
    }
});

/**
 * 为普通接口设置标准请求体限制
 */
const standardBodyParser = bodyParser.json({ 
    limit: '10mb' 
});

/**
 * URL编码解析器
 */
const urlencodedParser = bodyParser.urlencoded({ 
    limit: '10mb', 
    extended: true 
});

/**
 * 为文件上传接口设置更大的限制
 */
const uploadBodyParser = bodyParser.json({ 
    limit: '100mb' 
});

module.exports = {
    aiChatBodyParser,
    standardBodyParser,
    urlencodedParser,
    uploadBodyParser
}; 