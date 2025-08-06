const filePreview = require('./filePreview');

// 简单的日志函数
function logError(message, data = {}) {
    console.error(`[ERROR] ${message}`, data);
}

function logInfo(message, data = {}) {
    console.log(`[INFO] ${message}`, data);
}

/**
 * 文件预览API处理器
 * @param {Object} req 请求对象
 * @param {Object} res 响应对象
 */
async function filePreviewHandler(req, res) {
    console.log('🔍 [DEBUG] filePreviewHandler 开始执行');
    console.log('🔍 [DEBUG] 请求查询参数:', req.query);
    console.log('🔍 [DEBUG] 请求头:', req.headers);
    
    try {
        // 重构URL参数处理，解决Express.js将&误认为参数分隔符的问题
        let url = req.query.url;
        console.log('🔍 [DEBUG] 原始URL参数:', url);
        
        if (!url) {
            console.log('🔍 [DEBUG] 缺少URL参数');
            return res.status(400).json({
                success: false,
                message: '缺少文件URL参数',
                data: null
            });
        }
        
        // 检查是否有额外的AWS S3参数被Express.js误解析
        const awsParams = [];
        if (req.query['X-Amz-Algorithm']) awsParams.push(`X-Amz-Algorithm=${req.query['X-Amz-Algorithm']}`);
        if (req.query['X-Amz-Credential']) awsParams.push(`X-Amz-Credential=${req.query['X-Amz-Credential']}`);
        if (req.query['X-Amz-Date']) awsParams.push(`X-Amz-Date=${req.query['X-Amz-Date']}`);
        if (req.query['X-Amz-Expires']) awsParams.push(`X-Amz-Expires=${req.query['X-Amz-Expires']}`);
        if (req.query['X-Amz-SignedHeaders']) awsParams.push(`X-Amz-SignedHeaders=${req.query['X-Amz-SignedHeaders']}`);
        if (req.query['X-Amz-Signature']) awsParams.push(`X-Amz-Signature=${req.query['X-Amz-Signature']}`);
        
        // 如果检测到AWS参数被分离，重新组合完整URL
        if (awsParams.length > 0) {
            console.log('🔍 [DEBUG] 检测到AWS参数被分离，重新组合URL');
            console.log('🔍 [DEBUG] 分离的AWS参数:', awsParams);
            
            // 检查URL是否已经包含查询参数
            const separator = url.includes('?') ? '&' : '?';
            url = url + separator + awsParams.join('&');
            console.log('🔍 [DEBUG] 重新组合后的URL:', url);
        }
        
        // 检查URL是否需要解码
        if (url.includes('%3A') || url.includes('%2F') || url.includes('%26')) {
            console.log('🔍 [DEBUG] 检测到编码的URL，进行解码');
            try {
                url = decodeURIComponent(url);
                console.log('🔍 [DEBUG] 解码后的URL:', url);
            } catch (decodeError) {
                console.log('🔍 [DEBUG] URL解码失败:', decodeError.message);
                // 解码失败时继续使用原始URL
            }
        }
        
        console.log('🔍 [DEBUG] 最终使用的URL:', url);

        // 验证URL格式
        console.log('🔍 [DEBUG] 验证URL格式');
        try {
            new URL(url);
            console.log('🔍 [DEBUG] URL格式验证通过');
        } catch (error) {
            console.log('🔍 [DEBUG] URL格式验证失败:', error.message);
            return res.status(400).json({
                success: false,
                message: '无效的文件URL格式',
                data: null
            });
        }

        // 检查文件格式是否支持
        console.log('🔍 [DEBUG] 检查文件格式支持');
        if (!filePreview.isSupportedFormat(url)) {
            console.log('🔍 [DEBUG] 文件格式不支持');
            return res.status(400).json({
                success: false,
                message: '不支持的文件格式',
                data: {
                    fileInfo: filePreview.getFileInfo(url)
                }
            });
        }
        console.log('🔍 [DEBUG] 文件格式支持');

        // 获取文件信息
        console.log('🔍 [DEBUG] 获取文件信息');
        const fileInfo = filePreview.getFileInfo(url);
        console.log('🔍 [DEBUG] 文件信息:', fileInfo);
        
        // 预览文件
        console.log('🔍 [DEBUG] 开始预览文件');
        const previewResult = await filePreview.previewFile(url);
        
        console.log('🔍 [DEBUG] 文件预览成功');
        logInfo(`文件预览成功: ${url}`, {
            fileName: fileInfo.fileName,
            fileType: fileInfo.extension,
            userAgent: req.get('User-Agent'),
            ip: req.ip
        });

        res.json({
            success: true,
            message: '文件预览成功',
            data: {
                fileInfo: fileInfo,
                preview: previewResult
            }
        });

    } catch (error) {
        console.log('🔍 [DEBUG] filePreviewHandler 发生错误:');
        console.log('🔍 [DEBUG] 错误类型:', error.constructor.name);
        console.log('🔍 [DEBUG] 错误消息:', error.message);
        console.log('🔍 [DEBUG] 错误堆栈:', error.stack);
        
        logError(`文件预览失败: ${error.message}`, {
            url: req.query.url,
            error: error.stack,
            userAgent: req.get('User-Agent'),
            ip: req.ip
        });

        res.status(500).json({
            success: false,
            message: `文件预览失败: ${error.message}`,
            data: null
        });
    }
}

/**
 * 文件信息获取处理器
 * @param {Object} req 请求对象
 * @param {Object} res 响应对象
 */
function fileInfoHandler(req, res) {
    try {
        // 重构URL参数处理，解决Express.js将&误认为参数分隔符的问题
        let url = req.query.url;
        
        if (!url) {
            return res.status(400).json({
                success: false,
                message: '缺少文件URL参数',
                data: null
            });
        }
        
        // 检查是否有额外的AWS S3参数被Express.js误解析
        const awsParams = [];
        if (req.query['X-Amz-Algorithm']) awsParams.push(`X-Amz-Algorithm=${req.query['X-Amz-Algorithm']}`);
        if (req.query['X-Amz-Credential']) awsParams.push(`X-Amz-Credential=${req.query['X-Amz-Credential']}`);
        if (req.query['X-Amz-Date']) awsParams.push(`X-Amz-Date=${req.query['X-Amz-Date']}`);
        if (req.query['X-Amz-Expires']) awsParams.push(`X-Amz-Expires=${req.query['X-Amz-Expires']}`);
        if (req.query['X-Amz-SignedHeaders']) awsParams.push(`X-Amz-SignedHeaders=${req.query['X-Amz-SignedHeaders']}`);
        if (req.query['X-Amz-Signature']) awsParams.push(`X-Amz-Signature=${req.query['X-Amz-Signature']}`);
        
        // 如果检测到AWS参数被分离，重新组合完整URL
        if (awsParams.length > 0) {
            // 检查URL是否已经包含查询参数
            const separator = url.includes('?') ? '&' : '?';
            url = url + separator + awsParams.join('&');
        }
        
        // 检查URL是否需要解码
        if (url.includes('%3A') || url.includes('%2F') || url.includes('%26')) {
            try {
                url = decodeURIComponent(url);
            } catch (decodeError) {
                // 解码失败时继续使用原始URL
            }
        }

        // 验证URL格式
        try {
            new URL(url);
        } catch (error) {
            return res.status(400).json({
                success: false,
                message: '无效的文件URL格式',
                data: null
            });
        }

        const fileInfo = filePreview.getFileInfo(url);
        
        res.json({
            success: true,
            message: '获取文件信息成功',
            data: fileInfo
        });

    } catch (error) {
        logError(`获取文件信息失败: ${error.message}`, {
            url: req.query.url,
            error: error.stack
        });

        res.status(500).json({
            success: false,
            message: `获取文件信息失败: ${error.message}`,
            data: null
        });
    }
}

/**
 * 获取支持的文件格式列表
 * @param {Object} req 请求对象
 * @param {Object} res 响应对象
 */
function supportedFormatsHandler(req, res) {
    try {
        const formats = Object.keys(filePreview.supportedFormats).map(ext => ({
            extension: ext,
            mimeType: filePreview.supportedFormats[ext],
            description: getFormatDescription(ext)
        }));

        res.json({
            success: true,
            message: '获取支持的文件格式成功',
            data: {
                formats: formats,
                count: formats.length
            }
        });

    } catch (error) {
        logError(`获取支持的文件格式失败: ${error.message}`, {
            error: error.stack
        });

        res.status(500).json({
            success: false,
            message: `获取支持的文件格式失败: ${error.message}`,
            data: null
        });
    }
}

/**
 * 获取文件格式描述
 * @param {string} ext 文件扩展名
 * @returns {string} 格式描述
 */
function getFormatDescription(ext) {
    const descriptions = {
        'pdf': 'PDF文档',
        'xls': 'Excel 97-2003工作簿',
        'xlsx': 'Excel工作簿',
        'csv': 'CSV文件',
        'doc': 'Word 97-2003文档',
        'docx': 'Word文档',
        'ppt': 'PowerPoint 97-2003演示文稿',
        'pptx': 'PowerPoint演示文稿',
        'markdown': 'Markdown文档',
        'md': 'Markdown文档',
        'txt': '文本文件',
        'xml': 'XML文档'
    };
    
    return descriptions[ext] || '未知格式';
}

module.exports = {
    filePreviewHandler,
    fileInfoHandler,
    supportedFormatsHandler
}; 