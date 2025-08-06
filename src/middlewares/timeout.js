/**
 * 超时中间件
 * 为特定接口设置自定义超时时间
 */

/**
 * 为AI聊天接口设置600秒超时的中间件
 * @param {number} timeout - 超时时间（毫秒），默认600秒
 * @returns {Function} Express中间件函数
 */
function aiChatTimeout(timeout = 600000) {
    return (req, res, next) => {
        // 设置请求超时
        req.setTimeout(timeout, () => {
            console.error('AI聊天请求超时:', {
                path: req.path,
                method: req.method,
                timeout: timeout / 1000 + '秒',
                timestamp: new Date().toISOString()
            });
            
            // 如果响应还没有发送，则发送超时错误
            if (!res.headersSent) {
                res.status(408).json({
                    error: {
                        message: '请求超时，AI服务响应时间过长',
                        type: 'timeout_error',
                        code: 'request_timeout',
                        timeout: timeout / 1000
                    }
                });
            }
        });

        // 设置响应超时
        res.setTimeout(timeout, () => {
            console.error('AI聊天响应超时:', {
                path: req.path,
                method: req.method,
                timeout: timeout / 1000 + '秒',
                timestamp: new Date().toISOString()
            });
            
            // 如果响应还没有发送，则发送超时错误
            if (!res.headersSent) {
                res.status(408).json({
                    error: {
                        message: '响应超时，AI服务处理时间过长',
                        type: 'timeout_error',
                        code: 'response_timeout',
                        timeout: timeout / 1000
                    }
                });
            }
        });

        next();
    };
}

/**
 * 通用超时中间件
 * @param {number} timeout - 超时时间（毫秒）
 * @returns {Function} Express中间件函数
 */
function customTimeout(timeout) {
    return (req, res, next) => {
        req.setTimeout(timeout, () => {
            console.error('请求超时:', {
                path: req.path,
                method: req.method,
                timeout: timeout / 1000 + '秒',
                timestamp: new Date().toISOString()
            });
            
            if (!res.headersSent) {
                res.status(408).json({
                    error: {
                        message: '请求超时',
                        type: 'timeout_error',
                        code: 'request_timeout',
                        timeout: timeout / 1000
                    }
                });
            }
        });

        res.setTimeout(timeout, () => {
            console.error('响应超时:', {
                path: req.path,
                method: req.method,
                timeout: timeout / 1000 + '秒',
                timestamp: new Date().toISOString()
            });
            
            if (!res.headersSent) {
                res.status(408).json({
                    error: {
                        message: '响应超时',
                        type: 'timeout_error',
                        code: 'response_timeout',
                        timeout: timeout / 1000
                    }
                });
            }
        });

        next();
    };
}

module.exports = {
    aiChatTimeout,
    customTimeout
}; 