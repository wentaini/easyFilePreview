const axios = require('axios');
const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const xml2js = require('xml2js');
const iconv = require('iconv-lite');
const ExcelImageExtractor = require('./excelImageExtractor');

// 声明marked变量
let marked = null;

/**
 * 文件预览工具类
 */
class FilePreview {
    constructor() {
        this.supportedFormats = {
            'pdf': 'application/pdf',
            'xls': 'application/vnd.ms-excel',
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'csv': 'text/csv',
            'doc': 'application/msword',
            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'ppt': 'application/vnd.ms-powerpoint',
            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'markdown': 'text/markdown',
            'md': 'text/markdown',
            'txt': 'text/plain',
            'xml': 'application/xml'
        };
        
        // 文件缓存系统
        this.fileCache = new Map();
        this.cacheTimeout = 60 * 60 * 1000; // 1小时缓存
    }

    /**
     * 获取文件扩展名
     * @param {string} url 文件URL
     * @returns {string} 文件扩展名
     */
    getFileExtension(url) {
        try {
            const urlObj = new URL(url);
            const pathname = urlObj.pathname;
            const ext = path.extname(pathname).toLowerCase().substring(1);
            return ext;
        } catch (error) {
            // 如果不是有效的URL，尝试从路径中提取扩展名
            const ext = path.extname(url).toLowerCase().substring(1);
            return ext;
        }
    }

    /**
     * 检查文件格式是否支持
     * @param {string} url 文件URL
     * @returns {boolean} 是否支持
     */
    isSupportedFormat(url) {
        const ext = this.getFileExtension(url);
        return this.supportedFormats.hasOwnProperty(ext);
    }

    /**
     * 获取文件内容类型
     * @param {string} url 文件URL
     * @returns {string} 内容类型
     */
    getContentType(url) {
        const ext = this.getFileExtension(url);
        return this.supportedFormats[ext] || 'application/octet-stream';
    }

    /**
     * 下载文件
     * @param {string} url 文件URL
     * @returns {Promise<Buffer>} 文件内容
     */
    async downloadFile(url) {
        console.log('🔍 [DEBUG] downloadFile 开始执行');
        console.log('🔍 [DEBUG] URL:', url);
        console.log('🔍 [DEBUG] URL长度:', url.length);
        
        try {
            // 检查是否为本地文件路径
            if (url.startsWith('file://')) {
                console.log('🔍 [DEBUG] 检测到本地文件路径');
                const fs = require('fs');
                const path = require('path');
                const filePath = url.replace('file://', '');
                
                // 安全检查：确保文件路径在项目目录内
                const projectRoot = path.resolve(__dirname, '..');
                const resolvedPath = path.resolve(filePath);
                
                if (!resolvedPath.startsWith(projectRoot)) {
                    throw new Error('访问被拒绝：文件路径超出项目范围');
                }
                
                if (!fs.existsSync(resolvedPath)) {
                    throw new Error('文件不存在');
                }
                
                console.log('🔍 [DEBUG] 本地文件读取成功');
                return fs.readFileSync(resolvedPath);
            }
            
            console.log('🔍 [DEBUG] 开始远程文件下载');
            console.log('🔍 [DEBUG] 请求配置:', {
                method: 'GET',
                url: url,
                responseType: 'arraybuffer',
                timeout: 30000,
                headers: {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                }
            });
            
            // 远程文件下载
            const response = await axios({
                method: 'GET',
                url: url,
                responseType: 'arraybuffer',
                timeout: 30000, // 30秒超时
                headers: {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                }
            });
            
            console.log('🔍 [DEBUG] 下载成功');
            console.log('🔍 [DEBUG] 响应状态码:', response.status);
            console.log('🔍 [DEBUG] 响应头:', response.headers);
            console.log('🔍 [DEBUG] 文件大小:', response.data.length, '字节');
            
            return response.data;
        } catch (error) {
            console.log('🔍 [DEBUG] 下载失败，错误详情:');
            console.log('🔍 [DEBUG] 错误类型:', error.constructor.name);
            console.log('🔍 [DEBUG] 错误消息:', error.message);
            
            if (error.response) {
                console.log('🔍 [DEBUG] 响应状态码:', error.response.status);
                console.log('🔍 [DEBUG] 响应状态文本:', error.response.statusText);
                console.log('🔍 [DEBUG] 响应头:', error.response.headers);
                console.log('🔍 [DEBUG] 响应数据:', error.response.data);
            }
            
            if (error.request) {
                console.log('🔍 [DEBUG] 请求对象存在，但没有收到响应');
            }
            
            if (error.code) {
                console.log('🔍 [DEBUG] 错误代码:', error.code);
            }
            
            if (error.config) {
                console.log('🔍 [DEBUG] 请求配置:', {
                    url: error.config.url,
                    method: error.config.method,
                    timeout: error.config.timeout,
                    headers: error.config.headers
                });
            }
            
            throw new Error(`下载文件失败: ${error.message}`);
        }
    }

    /**
     * 生成唯一的文件ID
     * @returns {string} 文件ID
     */
    generateFileId() {
        return `file_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    }

    /**
     * 清理过期的缓存文件
     */
    cleanupFileCache() {
        const now = Date.now();
        for (const [fileId, fileData] of this.fileCache.entries()) {
            if (now - fileData.timestamp > this.cacheTimeout) {
                this.fileCache.delete(fileId);
                console.log(`清理过期缓存文件: ${fileId}`);
            }
        }
    }

    /**
     * 获取缓存的文件
     * @param {string} fileId 文件ID
     * @returns {Object|null} 文件数据
     */
    getCachedFile(fileId) {
        const fileData = this.fileCache.get(fileId);
        if (fileData) {
            // 更新访问时间
            fileData.timestamp = Date.now();
            return fileData;
        }
        return null;
    }

    /**
     * 提取PDF文本内容（按页返回）
     * @param {string} url 文件URL
     * @returns {Object} PDF文本信息
     */
    async extractPdfText(url) {
        try {
            console.log('🔍 [DEBUG] extractPdfText 开始执行');
            console.log('🔍 [DEBUG] URL:', url);
            
            // 下载文件
            console.log('🔍 [DEBUG] 开始下载PDF文件');
            const buffer = await this.downloadFile(url);
            console.log('🔍 [DEBUG] PDF文件下载完成，大小:', buffer.length, '字节');
            
            const fileSize = buffer.length;
            const fileSizeMB = (fileSize / 1024 / 1024).toFixed(1);
            
            // 使用pdf-parse提取文本内容
            console.log('🔍 [DEBUG] 开始解析PDF文本');
            const pdfParse = require('pdf-parse');
            const pdfData = await pdfParse(buffer);
            
            console.log('🔍 [DEBUG] PDF解析完成');
            console.log('🔍 [DEBUG] 页面数量:', pdfData.numpages);
            console.log('🔍 [DEBUG] 文本长度:', pdfData.text ? pdfData.text.length : 0);
            
            // 按页分割文本内容
            const pages = this.splitPdfTextByPages(pdfData.text, pdfData.numpages);
            
            // 转换为Markdown格式
            const markdownText = this.convertTextToMarkdown(pdfData.text);
            const markdownPages = pages.map(page => ({
                ...page,
                markdown: this.convertTextToMarkdown(page.text)
            }));

            const result = {
                text: pdfData.text || '', // 保留完整文本
                markdown: markdownText, // Markdown格式的完整文本
                pages: pages, // 按页分割的文本数组
                markdownPages: markdownPages, // 按页分割的Markdown数组
                pageCount: pdfData.numpages || 0,
                hasText: pdfData.text && pdfData.text.trim().length > 0,
                fileSize: fileSizeMB,
                info: pdfData.info || {},
                metadata: pdfData.metadata || {},
                version: pdfData.version || '',
                textLength: pdfData.text ? pdfData.text.length : 0
            };
            
            console.log('🔍 [DEBUG] PDF文本提取成功，页面数:', pages.length);
            return result;
            
        } catch (error) {
            console.log('🔍 [DEBUG] PDF文本提取失败:', error.message);
            throw new Error(`PDF文本提取失败: ${error.message}`);
        }
    }

    /**
     * 按页分割PDF文本内容
     * @param {string} text 完整文本内容
     * @param {number} pageCount 页面数量
     * @returns {Array} 按页分割的文本数组
     */
    splitPdfTextByPages(text, pageCount) {
        if (!text || !pageCount || pageCount <= 0) {
            return [];
        }

        const pages = [];
        const lines = text.split('\n');
        let currentPage = [];
        let pageIndex = 0;

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // 检测页面分隔符（常见的PDF页面分隔模式）
            if (this.isPageSeparator(line, i, lines)) {
                if (currentPage.length > 0) {
                    pages.push({
                        pageNumber: pageIndex + 1,
                        text: currentPage.join('\n').trim(),
                        lineCount: currentPage.length
                    });
                    pageIndex++;
                    currentPage = [];
                }
            } else {
                currentPage.push(line);
            }
        }

        // 添加最后一页
        if (currentPage.length > 0) {
            pages.push({
                pageNumber: pageIndex + 1,
                text: currentPage.join('\n').trim(),
                lineCount: currentPage.length
            });
        }

        // 如果检测到的页面数与实际页面数不匹配，使用简单分割
        if (pages.length !== pageCount) {
            console.log('🔍 [DEBUG] 页面分割不匹配，使用简单分割方法');
            return this.simpleSplitByPages(text, pageCount);
        }

        return pages;
    }

    /**
     * 检测是否为页面分隔符
     * @param {string} line 当前行
     * @param {number} lineIndex 行索引
     * @param {Array} allLines 所有行
     * @returns {boolean} 是否为页面分隔符
     */
    isPageSeparator(line, lineIndex, allLines) {
        // 检测页码模式
        const pageNumberPatterns = [
            /^\d+$/, // 纯数字
            /^第\s*\d+\s*页$/, // 第X页
            /^Page\s*\d+$/i, // Page X
            /^\d+\s*\/\s*\d+$/, // X/Y
            /^-\s*\d+\s*-$/, // - X -
            /^第\s*\d+\s*页\s*共\s*\d+\s*页$/, // 第X页共Y页
        ];

        // 检查当前行是否匹配页码模式
        for (const pattern of pageNumberPatterns) {
            if (pattern.test(line)) {
                return true;
            }
        }

        // 检测页面分隔特征
        // 1. 连续的空行
        if (line === '' && lineIndex > 0 && lineIndex < allLines.length - 1) {
            const prevLine = allLines[lineIndex - 1].trim();
            const nextLine = allLines[lineIndex + 1].trim();
            
            // 如果前后都有内容，且当前行为空，可能是页面分隔
            if (prevLine !== '' && nextLine !== '') {
                return true;
            }
        }

        // 2. 检测特殊的页面分隔符
        const separatorPatterns = [
            /^={3,}$/, // 连续等号
            /^-{3,}$/, // 连续横线
            /^\*{3,}$/, // 连续星号
            /^_+$/, // 连续下划线
        ];

        for (const pattern of separatorPatterns) {
            if (pattern.test(line)) {
                return true;
            }
        }

        return false;
    }

    /**
     * 简单按页分割文本（备用方法）
     * @param {string} text 完整文本
     * @param {number} pageCount 页面数量
     * @returns {Array} 按页分割的文本数组
     */
    simpleSplitByPages(text, pageCount) {
        if (!text || !pageCount || pageCount <= 0) {
            return [];
        }

        const pages = [];
        const lines = text.split('\n').filter(line => line.trim() !== '');
        const linesPerPage = Math.ceil(lines.length / pageCount);

        for (let i = 0; i < pageCount; i++) {
            const startIndex = i * linesPerPage;
            const endIndex = Math.min((i + 1) * linesPerPage, lines.length);
            const pageLines = lines.slice(startIndex, endIndex);

            pages.push({
                pageNumber: i + 1,
                text: pageLines.join('\n').trim(),
                lineCount: pageLines.length
            });
        }

        return pages;
    }

    /**
     * 将文本转换为Markdown格式
     * @param {string} text 原始文本
     * @returns {string} Markdown格式的文本
     */
    convertTextToMarkdown(text) {
        if (!text || typeof text !== 'string') {
            return '';
        }

        let markdown = text;
        const lines = text.split('\n');

        // 处理标题
        markdown = this.processMarkdownHeadings(markdown);

        // 处理列表
        markdown = this.processMarkdownLists(markdown);

        // 处理表格
        markdown = this.processMarkdownTables(markdown);

        // 处理强调文本
        markdown = this.processMarkdownEmphasis(markdown);

        // 处理段落和换行
        markdown = this.processMarkdownParagraphs(markdown);

        // 处理特殊字符
        markdown = this.processMarkdownSpecialChars(markdown);

        return markdown;
    }

    /**
     * 处理Markdown标题
     * @param {string} text 文本内容
     * @returns {string} 处理后的文本
     */
    processMarkdownHeadings(text) {
        const lines = text.split('\n');
        const processedLines = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // 检测标题（基于长度、大写字母比例、位置等）
            if (this.isHeading(line, i, lines)) {
                // 根据标题级别添加不同数量的#
                const level = this.getHeadingLevel(line);
                const headingText = line.replace(/^[0-9]+\.?\s*/, ''); // 移除数字编号
                processedLines.push(`${'#'.repeat(level)} ${headingText}`);
            } else {
                processedLines.push(line);
            }
        }

        return processedLines.join('\n');
    }

    /**
     * 获取标题级别
     * @param {string} line 标题行
     * @returns {number} 标题级别 (1-6)
     */
    getHeadingLevel(line) {
        const length = line.length;
        const upperCaseCount = (line.match(/[A-Z]/g) || []).length;
        const upperCaseRatio = upperCaseCount / length;

        // 根据特征判断标题级别
        if (length < 30 && upperCaseRatio > 0.5) {
            return 1; // 主标题
        } else if (length < 50 && upperCaseRatio > 0.3) {
            return 2; // 二级标题
        } else if (length < 80 && upperCaseRatio > 0.2) {
            return 3; // 三级标题
        } else if (/^\d+\./.test(line)) {
            return 4; // 编号标题
        } else {
            return 5; // 小标题
        }
    }

    /**
     * 处理Markdown列表
     * @param {string} text 文本内容
     * @returns {string} 处理后的文本
     */
    processMarkdownLists(text) {
        const lines = text.split('\n');
        const processedLines = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // 处理无序列表
            if (/^[-*•]\s/.test(line)) {
                processedLines.push(line.replace(/^[-*•]\s/, '- '));
            }
            // 处理有序列表
            else if (/^\d+\.\s/.test(line)) {
                processedLines.push(line); // 保持原样
            }
            // 处理缩进列表
            else if (/^\s+[-*•]\s/.test(line)) {
                const indentLevel = line.match(/^\s+/)[0].length;
                const indent = '  '.repeat(Math.floor(indentLevel / 2));
                processedLines.push(indent + line.replace(/^\s+[-*•]\s/, '- '));
            }
            else {
                processedLines.push(line);
            }
        }

        return processedLines.join('\n');
    }

    /**
     * 处理Markdown表格
     * @param {string} text 文本内容
     * @returns {string} 处理后的文本
     */
    processMarkdownTables(text) {
        // 检测表格结构
        if (this.isTableData(text)) {
            return this.convertTableToMarkdown(text);
        }
        return text;
    }

    /**
     * 将表格转换为Markdown格式
     * @param {string} text 表格文本
     * @returns {string} Markdown表格
     */
    convertTableToMarkdown(text) {
        const lines = text.split('\n').filter(line => line.trim());
        if (lines.length < 2) return text;

        // 检测表格类型并选择合适的转换方法
        const tableType = this.detectTableType(lines);
        
        switch (tableType) {
            case 'keyValue':
                return this.convertKeyValueToMarkdown(lines);
            case 'structured':
                return this.convertStructuredToMarkdown(lines);
            case 'aligned':
                return this.convertAlignedToMarkdown(lines);
            case 'standard':
            default:
                return this.convertStandardToMarkdown(lines);
        }
    }

    /**
     * 检测表格类型
     * @param {Array} lines 文本行数组
     * @returns {string} 表格类型
     */
    detectTableType(lines) {
        // 检测键值对格式
        const hasKeyValuePairs = lines.some(line => line.match(/^(.+?):\s*(.+)$/));
        if (hasKeyValuePairs) {
            return 'keyValue';
        }
        
        // 检测结构化数据（有规律的列分隔）
        const hasStructuredData = this.detectAlignedColumns(lines);
        if (hasStructuredData) {
            return 'aligned';
        }
        
        // 检测标准表格格式
        const hasStandardFormat = lines.some(line => {
            const parts = line.trim().split(/\s{2,}/);
            return parts.length >= 2;
        });
        if (hasStandardFormat) {
            return 'standard';
        }
        
        return 'structured';
    }

    /**
     * 检测是否为银行回单格式
     * @param {Array} lines 文本行数组
     * @returns {boolean} 是否为银行回单格式
     */
    isBankReceiptFormat(lines) {
        return lines.some(line => 
            line.includes('银行') && (line.includes('回单') || line.includes('收付款'))
        );
    }

    /**
     * 转换键值对格式为Markdown表格
     * @param {Array} lines 文本行数组
     * @returns {string} Markdown表格
     */
    convertKeyValueToMarkdown(lines) {
        let markdown = '';
        let currentSection = null;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // 检测新的段落或标题
            if (line.length > 0 && !line.includes(':') && !line.match(/^\d+/) && !line.includes('¥') && !line.includes('元')) {
                if (currentSection) {
                    markdown += '\n';
                }
                currentSection = line;
                markdown += `### ${line}\n\n`;
                markdown += '| 字段 | 值 |\n';
                markdown += '| --- | --- |\n';
                continue;
            }
            
            // 处理键值对
            const fieldMatch = line.match(/^(.+?):\s*(.+)$/);
            if (fieldMatch) {
                const key = fieldMatch[1].trim();
                const value = fieldMatch[2].trim();
                markdown += `| ${key} | ${value} |\n`;
            } else if (line.includes('¥') || line.includes('元')) {
                // 金额字段
                markdown += `| 金额 | ${line} |\n`;
            } else if (line.match(/^\d{4}-\d{2}-\d{2}/)) {
                // 日期字段
                markdown += `| 日期 | ${line} |\n`;
            } else if (line.match(/^\d{16,19}$/)) {
                // 账号字段
                markdown += `| 账号 | ${line} |\n`;
            } else if (line.trim().length > 0) {
                // 其他数据字段
                markdown += `| 数据 | ${line} |\n`;
            }
        }
        
        return markdown;
    }

    /**
     * 转换对齐格式为Markdown表格
     * @param {Array} lines 文本行数组
     * @returns {string} Markdown表格
     */
    convertAlignedToMarkdown(lines) {
        let markdown = '';
        
        // 处理表头
        const headerLine = lines[0];
        const headers = this.splitTableRow(headerLine);
        markdown += '| ' + headers.join(' | ') + ' |\n';
        markdown += '| ' + headers.map(() => '---').join(' | ') + ' |\n';

        // 处理数据行
        for (let i = 1; i < lines.length; i++) {
            const row = this.splitTableRow(lines[i]);
            markdown += '| ' + row.join(' | ') + ' |\n';
        }

        return markdown;
    }

    /**
     * 转换标准格式为Markdown表格
     * @param {Array} lines 文本行数组
     * @returns {string} Markdown表格
     */
    convertStandardToMarkdown(lines) {
        let markdown = '';
        
        // 处理表头
        const headerLine = lines[0];
        const headers = this.splitTableRow(headerLine);
        markdown += '| ' + headers.join(' | ') + ' |\n';
        markdown += '| ' + headers.map(() => '---').join(' | ') + ' |\n';

        // 处理数据行
        for (let i = 1; i < lines.length; i++) {
            const row = this.splitTableRow(lines[i]);
            markdown += '| ' + row.join(' | ') + ' |\n';
        }

        return markdown;
    }

    /**
     * 转换结构化数据为Markdown表格
     * @param {Array} lines 文本行数组
     * @returns {string} Markdown表格
     */
    convertStructuredToMarkdown(lines) {
        let markdown = '';
        
        // 尝试识别表头和数据
        const columnCounts = lines.map(line => {
            const parts = line.trim().split(/\s{2,}/);
            return parts.length;
        });
        
        const mostCommonColumnCount = this.getMostCommonValue(columnCounts);
        
        // 如果列数一致，按标准表格处理
        if (columnCounts.filter(count => count === mostCommonColumnCount).length >= Math.floor(lines.length * 0.6)) {
            return this.convertStandardToMarkdown(lines);
        }
        
        // 否则按键值对处理
        return this.convertKeyValueToMarkdown(lines);
    }

    /**
     * 将银行回单转换为Markdown表格
     * @param {Array} lines 文本行数组
     * @returns {string} Markdown表格
     */
    convertBankReceiptToMarkdown(lines) {
        const receipts = [];
        let currentReceipt = null;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            // 检测回单标题
            if (line.includes('银行') && (line.includes('回单') || line.includes('收付款'))) {
                if (currentReceipt) {
                    receipts.push(currentReceipt);
                }
                currentReceipt = {
                    title: line,
                    fields: []
                };
                continue;
            }
            
            if (currentReceipt) {
                // 检测字段对（键值对）
                const fieldMatch = line.match(/^(.+?):\s*(.+)$/);
                if (fieldMatch) {
                    currentReceipt.fields.push({
                        key: fieldMatch[1].trim(),
                        value: fieldMatch[2].trim()
                    });
                } else if (line.includes('¥') || line.includes('元')) {
                    // 金额字段
                    currentReceipt.fields.push({
                        key: '金额',
                        value: line
                    });
                } else if (line.match(/^\d{4}-\d{2}-\d{2}/)) {
                    // 日期字段
                    currentReceipt.fields.push({
                        key: '日期',
                        value: line
                    });
                } else if (line.match(/^\d{16,19}$/)) {
                    // 账号字段
                    currentReceipt.fields.push({
                        key: '账号',
                        value: line
                    });
                }
            }
        }
        
        if (currentReceipt) {
            receipts.push(currentReceipt);
        }
        
        // 转换为Markdown表格
        let markdown = '';
        
        receipts.forEach((receipt, index) => {
            markdown += `### ${receipt.title}\n\n`;
            markdown += '| 字段 | 值 |\n';
            markdown += '| --- | --- |\n';
            
            receipt.fields.forEach(field => {
                markdown += `| ${field.key} | ${field.value} |\n`;
            });
            
            markdown += '\n';
        });
        
        return markdown;
    }

    /**
     * 处理Markdown强调文本
     * @param {string} text 文本内容
     * @returns {string} 处理后的文本
     */
    processMarkdownEmphasis(text) {
        // 处理粗体文本（用**包围的文本）
        text = text.replace(/\*\*(.*?)\*\*/g, '**$1**');
        
        // 处理斜体文本（用*包围的文本）
        text = text.replace(/\*(.*?)\*/g, '*$1*');
        
        // 处理下划线文本（用__包围的文本）
        text = text.replace(/__(.*?)__/g, '**$1**');
        
        // 处理重要文本（全大写或包含关键词）
        text = text.replace(/\b(重要|注意|警告|错误|成功)\b/g, '**$1**');
        
        return text;
    }

    /**
     * 处理Markdown段落和换行
     * @param {string} text 文本内容
     * @returns {string} 处理后的文本
     */
    processMarkdownParagraphs(text) {
        const lines = text.split('\n');
        const processedLines = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            if (line === '') {
                // 空行保持空行
                processedLines.push('');
            } else if (line.startsWith('#')) {
                // 标题行前添加空行
                if (i > 0 && lines[i - 1].trim() !== '') {
                    processedLines.push('');
                }
                processedLines.push(line);
            } else if (line.startsWith('|')) {
                // 表格行
                processedLines.push(line);
            } else if (line.startsWith('-') || /^\d+\./.test(line)) {
                // 列表项
                processedLines.push(line);
            } else {
                // 普通段落
                processedLines.push(line);
            }
        }

        return processedLines.join('\n');
    }

    /**
     * 处理Markdown特殊字符
     * @param {string} text 文本内容
     * @returns {string} 处理后的文本
     */
    processMarkdownSpecialChars(text) {
        // 保护表格行，避免转义表格分隔符
        const tableLines = [];
        const lines = text.split('\n');
        const processedLines = [];
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            if (line.startsWith('|') || line.includes('| --- |')) {
                // 保护表格行，不进行特殊字符转义
                tableLines.push(line);
                // 使用不会被转义的占位符格式
                processedLines.push(`TABLE_PLACEHOLDER_${tableLines.length - 1}`);
            } else {
                processedLines.push(line);
            }
        }
        
        // 转义Markdown特殊字符（排除表格行）
        let processedText = processedLines.join('\n');
        processedText = processedText.replace(/([\\`*_{}\[\]()#+\-!])/g, '\\$1');
        
        // 处理URL
        processedText = processedText.replace(/(https?:\/\/[^\s]+)/g, '<$1>');
        
        // 处理邮箱
        processedText = processedText.replace(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g, '<$1>');
        
        // 恢复表格行
        tableLines.forEach((tableLine, index) => {
            const placeholder = `TABLE_PLACEHOLDER_${index}`;
            const escapedPlaceholder = `TABLE\\_PLACEHOLDER\\_${index}`;
            processedText = processedText.replace(escapedPlaceholder, tableLine);
        });
        
        return processedText;
    }

    /**
     * 预览PDF文件
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewPdf(buffer) {
        try {
            const fileSize = buffer.length;
            const fileSizeMB = (fileSize / 1024 / 1024).toFixed(1);
            
            // 对于PDF文件，我们使用文件流URL而不是base64
            // 生成唯一的文件ID
            const fileId = this.generateFileId();
            
            // 将文件存储到临时缓存中
            this.fileCache.set(fileId, {
                buffer: buffer,
                type: 'pdf',
                fileName: `document_${Date.now()}.pdf`,
                timestamp: Date.now()
            });
            
            // 清理过期的缓存文件（超过1小时）
            this.cleanupFileCache();
            
            // 尝试获取PDF信息
            let pageCount = null;
            let hasText = null;
            
            try {
                const pdfParse = require('pdf-parse');
                const pdfData = await pdfParse(buffer);
                pageCount = pdfData.numpages;
                hasText = pdfData.text && pdfData.text.trim().length > 0;
            } catch (error) {
                console.log('PDF解析失败，使用基础信息:', error.message);
            }
            
            return {
                type: 'pdf',
                fileId: fileId,
                fileUrl: `/api/filePreview/stream/${fileId}`,
                contentType: 'application/pdf',
                fileSize: fileSizeMB,
                pageCount: pageCount,
                hasText: hasText
            };
        } catch (error) {
            console.log('PDF处理失败:', error.message);
            return {
                type: 'pdf',
                error: error.message
            };
        }
    }

    /**
     * 预览Excel文件
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewExcel(buffer) {
        try {
            const workbook = XLSX.read(buffer, { type: 'buffer' });
            const sheets = {};
            const sheetInfo = {};
            const images = {};
            
            // 使用新的图片提取器
            const imageExtractor = new ExcelImageExtractor();
            const extractedImages = await imageExtractor.extractImages(buffer);
            
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                
                // 获取工作表范围
                const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
                const maxRow = range.e.r;
                const maxCol = range.e.c;
                
                // 提取数据，保持原始格式
                const data = [];
                for (let row = 0; row <= maxRow; row++) {
                    const rowData = [];
                    for (let col = 0; col <= maxCol; col++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                        const cell = worksheet[cellAddress];
                        rowData.push(cell ? cell.v : '');
                    }
                    data.push(rowData);
                }
                
                // 获取该工作表的图片并过滤无效图片
                let sheetImages = extractedImages.images[sheetName] || [];
                
                // 过滤掉没有内容的图片
                sheetImages = sheetImages.filter(image => {
                    // 检查是否有有效的数据
                    if (!image.data || image.data.length === 0) {
                        console.log(`⚠️  过滤空数据图片: ${image.name}`);
                        return false;
                    }
                    
                    // 检查是否有有效的base64数据
                    if (!image.base64 || image.base64.length === 0) {
                        console.log(`⚠️  过滤无效base64图片: ${image.name}`);
                        return false;
                    }
                    
                    // 检查src是否有效
                    if (!image.src || image.src === `data:${image.type};base64,`) {
                        console.log(`⚠️  过滤无效src图片: ${image.name}`);
                        return false;
                    }
                    
                    // 验证是否为真正的图片数据
                    const isValidImage = this.validateImageData(image.data, image.type);
                    if (!isValidImage) {
                        console.log(`⚠️  过滤非图片数据: ${image.name} (检测到非图片内容)`);
                        return false;
                    }
                    
                    return true;
                });
                
                // 为图片添加位置信息（简化处理，均匀分布）
                sheetImages.forEach((image, index) => {
                    // 计算图片位置（简化算法）
                    const row = Math.floor(index / 2) + 1; // 每行2张图片
                    const col = (index % 2) + 1;
                    
                    image.position = {
                        row: row,
                        col: col,
                        width: 2, // 默认2列宽
                        height: 1  // 默认1行高
                    };
                });
                
                sheets[sheetName] = data;
                images[sheetName] = sheetImages;
                sheetInfo[sheetName] = {
                    maxRow: maxRow + 1,
                    maxCol: maxCol + 1,
                    dataLength: data.length,
                    imageCount: sheetImages.length
                };
            });

            return {
                type: 'excel',
                sheets: sheets,
                sheetNames: workbook.SheetNames,
                sheetInfo: sheetInfo,
                images: images,
                contentType: 'application/json'
            };
        } catch (error) {
            throw new Error(`解析Excel文件失败: ${error.message}`);
        }
    }

    /**
     * 预览Word文档
     * @param {Buffer} buffer 文件内容
     * @param {string} fileExtension 文件扩展名
     * @returns {Object} 预览信息
     */
    async previewWord(buffer, fileExtension = 'docx') {
        try {
            // 根据文件扩展名选择不同的处理方法
            if (fileExtension.toLowerCase() === 'doc') {
                return await this.previewDocFile(buffer);
            } else {
                return await this.previewDocxFile(buffer);
            }
        } catch (error) {
            throw new Error(`解析Word文档失败: ${error.message}`);
        }
    }

    /**
     * 预览DOCX文件（新版Word格式）
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewDocxFile(buffer) {
        try {
            // 使用mammoth提取HTML格式，保留格式
            const result = await mammoth.convertToHtml({ 
                buffer,
                styleMap: [
                    "p[style-name='Heading 1'] => h1:fresh",
                    "p[style-name='Heading 2'] => h2:fresh",
                    "p[style-name='Heading 3'] => h3:fresh",
                    "p[style-name='Heading 4'] => h4:fresh",
                    "p[style-name='Heading 5'] => h5:fresh",
                    "p[style-name='Heading 6'] => h6:fresh",
                    "p[style-name='Title'] => h1.title:fresh",
                    "p[style-name='Subtitle'] => h2.subtitle:fresh",
                    "p[style-name='Quote'] => blockquote:fresh",
                    "p[style-name='Intense Quote'] => blockquote.intense:fresh",
                    "p[style-name='List Paragraph'] => li:fresh",
                    "r[style-name='Strong'] => strong",
                    "r[style-name='Emphasis'] => em",
                    "r[style-name='Code'] => code",
                    "table => table.table",
                    "tr => tr",
                    "td => td",
                    "th => th"
                ],
                ignoreEmptyParagraphs: false,
                preserveEmptyParagraphs: true,
                includeDefaultStyleMap: true
            });
            
            // 如果HTML转换失败，回退到纯文本
            if (!result.value) {
                const textResult = await mammoth.extractRawText({ buffer });
                return {
                    type: 'word',
                    content: textResult.value,
                    messages: textResult.messages,
                    contentType: 'text/plain'
                };
            }
            
            // 处理HTML内容，增强格式显示
            let processedContent = result.value;
            
            // 添加缩进样式
            processedContent = processedContent.replace(
                /<p([^>]*)>/g, 
                '<p$1 style="margin: 0.5em 0; line-height: 1.6;">'
            );
            
            // 处理列表缩进
            processedContent = processedContent.replace(
                /<ul([^>]*)>/g,
                '<ul$1 style="margin: 0.5em 0; padding-left: 2em;">'
            );
            
            processedContent = processedContent.replace(
                /<ol([^>]*)>/g,
                '<ol$1 style="margin: 0.5em 0; padding-left: 2em;">'
            );
            
            // 处理表格样式
            processedContent = processedContent.replace(
                /<table([^>]*)>/g,
                '<table$1 style="border-collapse: collapse; width: 100%; margin: 1em 0;">'
            );
            
            processedContent = processedContent.replace(
                /<td([^>]*)>/g,
                '<td$1 style="border: 1px solid #ddd; padding: 8px; text-align: left;">'
            );
            
            processedContent = processedContent.replace(
                /<th([^>]*)>/g,
                '<th$1 style="border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f5f5f5; font-weight: bold;">'
            );
            
            // 处理标题样式
            processedContent = processedContent.replace(
                /<h1([^>]*)>/g,
                '<h1$1 style="color: #2c3e50; margin: 1em 0 0.5em 0; font-size: 1.8em; font-weight: bold;">'
            );
            
            processedContent = processedContent.replace(
                /<h2([^>]*)>/g,
                '<h2$1 style="color: #34495e; margin: 1em 0 0.5em 0; font-size: 1.5em; font-weight: bold;">'
            );
            
            processedContent = processedContent.replace(
                /<h3([^>]*)>/g,
                '<h3$1 style="color: #34495e; margin: 1em 0 0.5em 0; font-size: 1.3em; font-weight: bold;">'
            );
            
            // 处理引用样式
            processedContent = processedContent.replace(
                /<blockquote([^>]*)>/g,
                '<blockquote$1 style="border-left: 4px solid #3498db; margin: 1em 0; padding-left: 1em; color: #555; font-style: italic;">'
            );
            
            // 处理代码样式
            processedContent = processedContent.replace(
                /<code([^>]*)>/g,
                '<code$1 style="background-color: #f8f9fa; padding: 2px 4px; border-radius: 3px; font-family: monospace; color: #e74c3c;">'
            );
            
            return {
                type: 'word',
                content: processedContent,
                messages: result.messages,
                contentType: 'text/html'
            };
        } catch (error) {
            throw new Error(`解析Word文档失败: ${error.message}`);
        }
    }

    /**
     * 预览DOC文件（旧版Word格式）
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewDocFile(buffer) {
        try {
                        // 尝试使用mammoth直接转换为HTML（如果支持）
            try {
                const result = await mammoth.convertToHtml({ buffer });
                if (result.value) {
                    console.log('🔍 [DEBUG] mammoth成功转换为HTML，长度:', result.value.length);
                    console.log('🔍 [DEBUG] mammoth转换结果预览:', result.value.substring(0, 200));
                    return {
                        type: 'word',
                        content: result.value,
                        messages: result.messages,
                        contentType: 'text/html'
                    };
                }
            } catch (mammothError) {
                console.log('🔍 [DEBUG] mammoth转换失败，回退到textract:', mammothError.message);
            }
            
            // 使用textract提取文本内容
            const textract = require('textract');
            
            return new Promise((resolve, reject) => {
                textract.fromBufferWithMime('application/msword', buffer, {
                    preserveLineBreaks: true,
                    preserveOnlyMultipleLineBreaks: true
                }, (error, text) => {
                    if (error) {
                        console.log('🔍 [DEBUG] textract提取失败，尝试使用officeparser:', error.message);
                        // 如果textract失败，尝试使用officeparser
                        this.previewDocWithOfficeParser(buffer)
                            .then(resolve)
                            .catch(reject);
                    } else {
                        console.log('🔍 [DEBUG] 原始提取的文本:', text);
                        console.log('🔍 [DEBUG] 文本长度:', text.length);
                        
                        // mammoth转换失败时，直接显示原始文本
                        const htmlContent = this.convertRawTextToHtml(text);
                        console.log('🔍 [DEBUG] 转换后的HTML长度:', htmlContent.length);
                        console.log('🔍 [DEBUG] 转换后的HTML预览:', htmlContent);
                        resolve({
                            type: 'word',
                            content: htmlContent,
                            messages: [],
                            contentType: 'text/html'
                        });
                    }
                });
            });
        } catch (error) {
            throw new Error(`解析DOC文件失败: ${error.message}`);
        }
    }

    /**
     * 使用officeparser处理DOC文件
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewDocWithOfficeParser(buffer) {
        try {
            const officeParser = require('officeparser');
            
            const result = await officeParser.parseBuffer(buffer);
            
            if (result && result.text) {
                const htmlContent = this.convertTextToHtml(result.text);
                return {
                    type: 'word',
                    content: htmlContent,
                    messages: [],
                    contentType: 'text/html'
                };
            } else {
                throw new Error('无法提取文档内容');
            }
        } catch (error) {
            throw new Error(`officeparser解析失败: ${error.message}`);
        }
    }

    /**
     * 将纯文本转换为格式化的HTML
     * @param {string} text 纯文本内容
     * @returns {string} 格式化的HTML内容
     */
    convertTextToHtml(text) {
        if (!text || typeof text !== 'string') {
            return '<p>无法提取文档内容</p>';
        }

        // 清理文本
        let cleanText = text.trim();
        
        // 处理表格数据 - 检测表格结构
        if (this.isTableData(cleanText)) {
            return this.convertTableToHtml(cleanText);
        }
        
        // 处理段落分隔 - 更智能的分段
        let paragraphs = this.splitIntelligentParagraphs(cleanText);
        
        // 转换为HTML
        let htmlContent = '<div class="word-document" style="font-family: Arial, sans-serif; line-height: 1.6;">';
        
        paragraphs.forEach(paragraph => {
            const trimmedParagraph = paragraph.trim();
            if (trimmedParagraph.length === 0) return;
            
            // 检测标题（基于长度、大写字母比例等）
            if (this.isHeading(trimmedParagraph)) {
                htmlContent += `<h2 style="color: #2c3e50; margin: 1em 0 0.5em 0; font-size: 1.5em; font-weight: bold; border-bottom: 2px solid #3498db; padding-bottom: 0.3em;">${this.escapeHtml(trimmedParagraph)}</h2>`;
            } else if (this.isListItem(trimmedParagraph)) {
                // 处理列表项
                let listContent = this.processListItem(trimmedParagraph);
                htmlContent += `<div style="margin: 0.5em 0; padding-left: 1.5em;">${listContent}</div>`;
            } else {
                // 处理普通段落
                let processedParagraph = this.processParagraph(trimmedParagraph);
                htmlContent += `<p style="margin: 0.8em 0; line-height: 1.8; text-align: justify; text-indent: 2em;">${processedParagraph}</p>`;
            }
        });
        
        htmlContent += '</div>';
        
        return htmlContent;
    }

    /**
     * 将原始文本直接转换为HTML（mammoth转换失败时使用）
     * @param {string} text 原始文本内容
     * @returns {string} 格式化的HTML内容
     */
    convertRawTextToHtml(text) {
        if (!text || typeof text !== 'string') {
            return '<p>无法提取文档内容</p>';
        }

        // 清理文本
        let cleanText = text.trim();
        
        // 直接按换行符分割，保持原始格式
        let lines = cleanText.split(/\n/).filter(line => line.trim().length > 0);
        
        // 转换为HTML，保持原始格式
        let htmlContent = '<div class="word-document" style="font-family: Arial, sans-serif; line-height: 1.6; white-space: pre-wrap;">';
        
        if (lines.length === 0) {
            htmlContent += '<p>文档内容为空</p>';
        } else {
            lines.forEach(line => {
                const trimmedLine = line.trim();
                if (trimmedLine.length === 0) {
                    htmlContent += '<br>';
                } else {
                    htmlContent += `<p style="margin: 0.5em 0; line-height: 1.6;">${this.escapeHtml(trimmedLine)}</p>`;
                }
            });
        }
        
        htmlContent += '</div>';
        
        return htmlContent;
    }

    /**
     * 判断是否为标题
     * @param {string} text 文本内容
     * @returns {boolean} 是否为标题
     */
    isHeading(text) {
        // 标题通常较短，且包含较多大写字母
        const length = text.length;
        const upperCaseCount = (text.match(/[A-Z]/g) || []).length;
        const upperCaseRatio = upperCaseCount / length;
        
        // 标题特征：长度小于100，大写字母比例大于0.3，或者以数字开头
        return (length < 100 && upperCaseRatio > 0.3) || /^\d+\./.test(text);
    }

    /**
     * 处理段落内容
     * @param {string} paragraph 段落文本
     * @returns {string} 处理后的HTML
     */
    processParagraph(paragraph) {
        let processed = this.escapeHtml(paragraph);
        
        // 处理粗体文本（用**包围的文本）
        processed = processed.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
        
        // 处理斜体文本（用*包围的文本）
        processed = processed.replace(/\*(.*?)\*/g, '<em>$1</em>');
        
        // 处理下划线文本（用__包围的文本）
        processed = processed.replace(/__(.*?)__/g, '<u>$1</u>');
        
        // 处理列表项
        if (/^[\s]*[-*•]\s/.test(processed)) {
            processed = processed.replace(/^[\s]*[-*•]\s/, '• ');
        }
        
        // 处理编号列表
        if (/^[\s]*\d+\.\s/.test(processed)) {
            // 保持原样，因为已经在段落中
        }
        
        return processed;
    }

    /**
     * 检测是否为表格数据
     * @param {string} text 文本内容
     * @returns {boolean} 是否为表格数据
     */
    isTableData(text) {
        // 通用的表格检测方法，基于格式特征而非内容
        const lines = text.split(/\n/).filter(line => line.trim().length > 0);
        
        // 如果只有一行，尝试检测是否包含表格特征
        if (lines.length === 1) {
            return this.isSingleLineTable(lines[0]);
        }
        
        if (lines.length < 2) return false;
        
        // 检测格式特征
        const formatFeatures = {
            hasConsistentColumns: false,
            hasNumberedRows: false,
            hasTabularStructure: false,
            hasMultipleDataRows: false,
            hasKeyValuePairs: false,
            hasStructuredData: false,
            hasTableSeparators: false,
            hasAlignedData: false
        };
        
        // 1. 检测是否有编号行
        formatFeatures.hasNumberedRows = lines.some(line => /^\d+[\.\s]/.test(line.trim()));
        
        // 2. 检测是否有表格结构（多列数据）
        const columnCounts = lines.map(line => {
            // 计算每行的列数（基于空格分隔）
            const parts = line.trim().split(/\s{2,}/);
            return parts.length;
        });
        
        // 如果大部分行都有相同的列数，说明有表格结构
        const mostCommonColumnCount = this.getMostCommonValue(columnCounts);
        formatFeatures.hasConsistentColumns = columnCounts.filter(count => count === mostCommonColumnCount).length >= Math.floor(lines.length * 0.6);
        formatFeatures.hasTabularStructure = mostCommonColumnCount >= 2; // 降低要求，支持2列以上的表格
        
        // 3. 检测是否有多个数据行
        formatFeatures.hasMultipleDataRows = lines.length >= 2;
        
        // 4. 检测是否有键值对格式（通用结构化数据）
        formatFeatures.hasKeyValuePairs = lines.some(line => line.includes(':'));
        formatFeatures.hasStructuredData = lines.some(line => line.match(/^(.+?):\s*(.+)$/));
        
        // 5. 检测表格分隔符（制表符、多个空格等）
        formatFeatures.hasTableSeparators = lines.some(line => 
            line.includes('\t') || line.match(/\s{3,}/) || line.includes('|')
        );
        
        // 6. 检测对齐的数据（通过空格对齐的列）
        formatFeatures.hasAlignedData = this.detectAlignedColumns(lines);
        
        // 综合判断是否为表格（更严格的条件）
        return (formatFeatures.hasTabularStructure && formatFeatures.hasMultipleDataRows && formatFeatures.hasConsistentColumns) ||
               (formatFeatures.hasStructuredData && formatFeatures.hasKeyValuePairs) ||
               (formatFeatures.hasTableSeparators && formatFeatures.hasMultipleDataRows) ||
               (formatFeatures.hasAlignedData && formatFeatures.hasMultipleDataRows && formatFeatures.hasConsistentColumns) ||
               (formatFeatures.hasNumberedRows && formatFeatures.hasMultipleDataRows && formatFeatures.hasConsistentColumns);
    }

    /**
     * 检测列是否对齐
     * @param {Array} lines 文本行数组
     * @returns {boolean} 列是否对齐
     */
    detectAlignedColumns(lines) {
        if (lines.length < 2) return false;
        
        // 分析每行的空格模式
        const spacePatterns = lines.map(line => {
            const matches = line.match(/\s{2,}/g);
            return matches ? matches.length : 0;
        });
        
        // 如果大部分行都有相似的空格模式，说明列对齐
        const avgSpaces = spacePatterns.reduce((sum, count) => sum + count, 0) / spacePatterns.length;
        const consistentSpaces = spacePatterns.filter(count => count >= avgSpaces * 0.5).length;
        
        return consistentSpaces >= Math.floor(lines.length * 0.6);
    }

    /**
     * 检测单行是否为表格
     * @param {string} line 单行文本
     * @returns {boolean} 是否为表格
     */
    isSingleLineTable(line) {
        // 通用的单行表格检测
        const trimmedLine = line.trim();
        
        // 检测是否包含多个数字序号
        const numberMatches = trimmedLine.match(/\d+[\.\s]+/g);
        if (numberMatches && numberMatches.length >= 2) {
            // 检测是否包含多个数据项
            const dataItems = trimmedLine.split(/\d+[\.\s]+/).filter(item => item.trim().length > 0);
            return dataItems.length >= 2;
        }
        
        // 检测是否包含表格分隔符
        if (trimmedLine.includes('\t') || trimmedLine.match(/\s{3,}/) || trimmedLine.includes('|')) {
            const parts = trimmedLine.split(/\t|\s{3,}|\|/).filter(part => part.trim().length > 0);
            return parts.length >= 2;
        }
        
        // 检测是否包含键值对格式
        if (trimmedLine.includes(':')) {
            const keyValuePairs = trimmedLine.split(/\s*:\s*/).filter(part => part.trim().length > 0);
            return keyValuePairs.length >= 2;
        }
        
        return false;
    }

    /**
     * 将表格数据转换为HTML表格
     * @param {string} text 表格文本
     * @returns {string} HTML表格
     */
    convertTableToHtml(text) {
        // 按行分割
        const lines = text.split(/\n/).filter(line => line.trim().length > 0);
        
        let htmlContent = '<div class="word-document" style="font-family: Arial, sans-serif;">';
        
        // 处理标题行
        if (lines.length > 0) {
            const titleLine = lines[0];
            // 提取文档标题
            const title = this.extractDocumentTitle(titleLine);
            if (title) {
                htmlContent += `<h2 style="color: #2c3e50; margin: 1em 0 0.5em 0; font-size: 1.5em; font-weight: bold; text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 0.3em;">${this.escapeHtml(title)}</h2>`;
            }
        }
        
        // 智能解析表格结构
        const tableStructure = this.analyzeTableStructure(text);
        
        // 处理表格内容
        htmlContent += '<div style="overflow-x: auto; margin: 1em 0;">';
        htmlContent += '<table style="width: 100%; border-collapse: collapse; border: 1px solid #ddd; font-size: 14px;">';
        
        // 动态生成表头
        if (tableStructure.headers.length > 0) {
            htmlContent += '<thead>';
            htmlContent += '<tr style="background-color: #f8f9fa;">';
            tableStructure.headers.forEach(header => {
                htmlContent += `<th style="border: 1px solid #ddd; padding: 12px; text-align: center; font-weight: bold;">${this.escapeHtml(header)}</th>`;
            });
            htmlContent += '</tr>';
            htmlContent += '</thead>';
        }
        
        // 表格内容
        htmlContent += '<tbody>';
        
        // 解析数据行
        const dataRows = this.parseTableData(text, tableStructure);
        dataRows.forEach(row => {
            htmlContent += '<tr>';
            tableStructure.headers.forEach(header => {
                const value = row[header] || '';
                const align = this.getColumnAlignment(header);
                htmlContent += `<td style="border: 1px solid #ddd; padding: 8px; text-align: ${align};">${this.escapeHtml(value)}</td>`;
            });
            htmlContent += '</tr>';
        });
        
        htmlContent += '</tbody>';
        htmlContent += '</table>';
        htmlContent += '</div>';
        htmlContent += '</div>';
        
        return htmlContent;
    }

    /**
     * 解析表格数据
     * @param {string} text 表格文本
     * @param {Object} tableStructure 表格结构信息
     * @returns {Array} 解析后的数据行
     */
    parseTableData(text, tableStructure) {
        const rows = [];
        const lines = text.split(/\n/).filter(line => line.trim().length > 0);
        
        if (lines.length === 1) {
            // 处理单行表格
            return this.parseSingleLineTable(text, tableStructure);
        }
        
        // 跳过标题行
        const dataLines = lines.slice(1);
        
        dataLines.forEach(line => {
            const row = this.parseTableRow(line, tableStructure);
            if (row && Object.keys(row).length > 0) {
                rows.push(row);
            }
        });
        
        return rows;
    }

    /**
     * 解析单行表格
     * @param {string} text 表格文本
     * @param {Object} tableStructure 表格结构信息
     * @returns {Array} 解析后的数据行
     */
    parseSingleLineTable(text, tableStructure) {
        const rows = [];
        
        console.log('🔍 [DEBUG] 解析单行表格，文本长度:', text.length);
        console.log('🔍 [DEBUG] 表头:', tableStructure.headers);
        
        // 使用更通用的正则表达式匹配数据行
        // 匹配模式：数字 + 空格 + 任意内容（直到下一个数字或结尾）
        const dataMatches = text.match(/\d+\s+[^0-9]+?(?=\d+\s|$)/g);
        
        console.log('🔍 [DEBUG] 找到的数据行数:', dataMatches ? dataMatches.length : 0);
        
        if (dataMatches) {
            dataMatches.forEach((match, index) => {
                console.log('🔍 [DEBUG] 处理数据行', index + 1, ':', match);
                
                const row = {};
                
                // 提取序号
                const numberMatch = match.match(/^(\d+)/);
                const number = numberMatch ? numberMatch[1] : (index + 1).toString();
                
                // 提取数据部分
                const dataPart = match.replace(/^\d+\s+/, '');
                
                // 智能分割数据
                const parts = this.splitTableDataIntelligently(dataPart);
                console.log('🔍 [DEBUG] 分割后的数据:', parts);
                
                // 动态映射到表头
                if (tableStructure.headers.length > 0) {
                    row[tableStructure.headers[0]] = number; // 序号
                    
                    for (let i = 1; i < tableStructure.headers.length && i - 1 < parts.length; i++) {
                        row[tableStructure.headers[i]] = parts[i - 1] || '';
                    }
                }
                
                console.log('🔍 [DEBUG] 解析后的行数据:', row);
                
                if (Object.keys(row).length > 0) {
                    rows.push(row);
                }
            });
        }
        
        console.log('🔍 [DEBUG] 最终解析的行数:', rows.length);
        return rows;
    }

    /**
     * 智能分割表格数据
     * @param {string} dataPart 数据部分
     * @returns {Array} 分割后的数据
     */
    splitTableDataIntelligently(dataPart) {
        const parts = [];
        let currentPart = '';
        let i = 0;
        
        while (i < dataPart.length) {
            const char = dataPart[i];
            
            // 检测列分隔符（多个空格）
            if (char === ' ' && i + 1 < dataPart.length && dataPart[i + 1] === ' ') {
                if (currentPart.trim()) {
                    parts.push(currentPart.trim());
                    currentPart = '';
                }
                // 跳过多个空格
                while (i < dataPart.length && dataPart[i] === ' ') {
                    i++;
                }
                continue;
            }
            
            currentPart += char;
            i++;
        }
        
        // 添加最后一个部分
        if (currentPart.trim()) {
            parts.push(currentPart.trim());
        }
        
        return parts;
    }

    /**
     * 解析单行表格数据
     * @param {string} line 单行文本
     * @param {Object} tableStructure 表格结构信息
     * @returns {Object} 解析后的行数据
     */
    parseTableRow(line, tableStructure) {
        const row = {};
        
        // 移除行首的数字序号（如果存在）
        const cleanLine = line.replace(/^\d+[\.\s]+/, '');
        
        // 根据分隔符分割数据
        const parts = this.splitTableRow(cleanLine);
        
        // 将数据映射到表头
        tableStructure.headers.forEach((header, index) => {
            row[header] = parts[index] || '';
        });
        
        return row;
    }

    /**
     * 智能分割表格行
     * @param {string} line 表格行文本
     * @returns {Array} 分割后的数据
     */
    splitTableRow(line) {
        // 尝试多种分割方式
        let parts = [];
        
        // 1. 尝试按多个空格分割
        parts = line.split(/\s{2,}/);
        if (parts.length > 1) {
            return parts.map(part => part.trim());
        }
        
        // 2. 尝试按制表符分割
        parts = line.split(/\t/);
        if (parts.length > 1) {
            return parts.map(part => part.trim());
        }
        
        // 3. 尝试按单个空格分割，但合并短词
        parts = line.split(/\s+/);
        const mergedParts = [];
        let currentPart = '';
        
        for (let i = 0; i < parts.length; i++) {
            const part = parts[i];
            if (part.length <= 2 && i < parts.length - 1) {
                // 短词可能与下一个词合并
                currentPart += (currentPart ? ' ' : '') + part;
            } else {
                currentPart += (currentPart ? ' ' : '') + part;
                mergedParts.push(currentPart.trim());
                currentPart = '';
            }
        }
        
        if (currentPart) {
            mergedParts.push(currentPart.trim());
        }
        
        return mergedParts;
    }

    /**
     * 智能分段处理
     * @param {string} text 文本内容
     * @returns {Array} 分段后的数组
     */
    splitIntelligentParagraphs(text) {
        // 首先按双换行分割
        let paragraphs = text.split(/\n\s*\n/);
        
        // 如果没有明显的段落分隔，尝试按单换行分割
        if (paragraphs.length <= 1) {
            paragraphs = text.split(/\n/);
        }
        
        // 如果还是没有分段，尝试按数字序号分割
        if (paragraphs.length <= 1) {
            paragraphs = this.splitByNumberedItems(text);
        }
        
        // 过滤空段落并清理
        return paragraphs
            .map(p => p.trim())
            .filter(p => p.length > 0);
    }

    /**
     * 按数字序号分割文本
     * @param {string} text 文本内容
     * @returns {Array} 分割后的段落
     */
    splitByNumberedItems(text) {
        const paragraphs = [];
        const matches = text.match(/\d+[\.\s]+[^0-9]+/g);
        
        if (matches) {
            matches.forEach(match => {
                paragraphs.push(match.trim());
            });
        } else {
            // 如果没有找到数字序号，按空格分割长文本
            const words = text.split(/\s+/);
            const chunkSize = Math.ceil(words.length / 3);
            
            for (let i = 0; i < words.length; i += chunkSize) {
                const chunk = words.slice(i, i + chunkSize).join(' ');
                if (chunk.trim()) {
                    paragraphs.push(chunk);
                }
            }
        }
        
        return paragraphs;
    }

    /**
     * 检测是否为列表项
     * @param {string} text 文本内容
     * @returns {boolean} 是否为列表项
     */
    isListItem(text) {
        return /^[\s]*[-*•]\s/.test(text) || /^[\s]*\d+\.\s/.test(text);
    }

    /**
     * 处理列表项
     * @param {string} text 列表项文本
     * @returns {string} 处理后的HTML
     */
    processListItem(text) {
        let processed = this.escapeHtml(text);
        
        // 处理无序列表
        if (/^[\s]*[-*•]\s/.test(processed)) {
            processed = processed.replace(/^[\s]*[-*•]\s/, '• ');
            return `<div style="margin: 0.3em 0;">${processed}</div>`;
        }
        
        // 处理有序列表
        if (/^[\s]*\d+\.\s/.test(processed)) {
            return `<div style="margin: 0.3em 0;">${processed}</div>`;
        }
        
        return processed;
    }

    /**
     * 提取文档标题
     * @param {string} titleLine 标题行文本
     * @returns {string} 提取的标题
     */
    extractDocumentTitle(titleLine) {
        // 常见的标题模式
        const titlePatterns = [
            /^(.+?)(?:报检单号|单号|编号|NO\.|No\.|no\.)/i,
            /^(.+?)(?:清单|列表|表格|表|报告|报表)/,
            /^(.+?)(?:\s*$)/  // 如果前面都不匹配，取整行作为标题
        ];
        
        for (const pattern of titlePatterns) {
            const match = titleLine.match(pattern);
            if (match && match[1]) {
                return match[1].trim();
            }
        }
        
        return titleLine.trim();
    }

    /**
     * 分析表格结构
     * @param {string} text 表格文本
     * @returns {Object} 表格结构信息
     */
    analyzeTableStructure(text) {
        const lines = text.split(/\n/).filter(line => line.trim().length > 0);
        const headers = [];
        
        if (lines.length > 1) {
            // 基于格式分析表头，而不是内容
            const headerLine = this.findHeaderLine(lines);
            const headerParts = this.splitTableRow(headerLine);
            
            // 根据列的位置和格式特征生成表头
            headerParts.forEach((part, index) => {
                const header = this.generateHeaderByPosition(part, index, headerParts.length);
                headers.push(header);
            });
        } else if (lines.length === 1) {
            // 处理单行表格 - 从文本中提取表头
            const extractedHeaders = this.extractHeadersFromSingleLine(text);
            if (extractedHeaders.length > 0) {
                headers.push(...extractedHeaders);
            } else {
                // 动态生成表头，基于数据列数
                const columnCount = this.estimateColumnCount(text);
                headers.push('序号');
                for (let i = 1; i < columnCount; i++) {
                    headers.push(`列${i}`);
                }
            }
        }
        
        return { headers };
    }

    /**
     * 估算列数
     * @param {string} text 文本内容
     * @returns {number} 估算的列数
     */
    estimateColumnCount(text) {
        // 基于数据行估算列数
        const dataMatches = text.match(/\d+\s+[^0-9]+?(?=\d+\s|$)/g);
        if (dataMatches && dataMatches.length > 0) {
            // 分析第一行数据来估算列数
            const firstDataRow = dataMatches[0];
            const dataPart = firstDataRow.replace(/^\d+\s+/, '');
            const parts = this.splitTableDataIntelligently(dataPart);
            return parts.length + 1; // +1 for 序号列
        }
        return 4; // 默认4列
    }

    /**
     * 从单行文本中提取表头
     * @param {string} text 文本内容
     * @returns {Array} 提取的表头
     */
    extractHeadersFromSingleLine(text) {
        const headers = [];
        
        // 查找表头行（通常在数字序号之前）
        const headerMatch = text.match(/序号\s+品名\s+原产地\/地区\s+规格\s+报检数\/重量\s+生产日期\s+保质期/);
        if (headerMatch) {
            return ['序号', '品名', '原产地/地区', '规格', '报检数/重量', '生产日期', '保质期'];
        }
        
        // 查找其他常见的表头模式
        const patterns = [
            /序号\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)/,
            /序号\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)/,
            /序号\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)\s+([^\s]+)/
        ];
        
        for (const pattern of patterns) {
            const match = text.match(pattern);
            if (match) {
                headers.push('序号');
                for (let i = 1; i < match.length; i++) {
                    headers.push(match[i]);
                }
                return headers;
            }
        }
        
        return headers;
    }

    /**
     * 查找表头行
     * @param {Array} lines 所有行
     * @returns {string} 表头行
     */
    findHeaderLine(lines) {
        // 跳过第一行（通常是标题）
        const dataLines = lines.slice(1);
        
        // 查找最可能的表头行
        for (let i = 0; i < Math.min(3, dataLines.length); i++) {
            const line = dataLines[i];
            const parts = this.splitTableRow(line);
            
            // 表头行通常具有以下特征：
            // 1. 不包含数字序号
            // 2. 列数适中（3-8列）
            // 3. 每列内容较短
            const hasNoNumbers = !/\d+/.test(line);
            const hasReasonableColumns = parts.length >= 3 && parts.length <= 8;
            const hasShortContent = parts.every(part => part.length <= 10);
            
            if (hasNoNumbers && hasReasonableColumns && hasShortContent) {
                return line;
            }
        }
        
        // 如果没找到合适的表头行，返回第二行
        return dataLines[0] || '';
    }

    /**
     * 根据位置和格式生成表头
     * @param {string} content 列内容
     * @param {number} index 列索引
     * @param {number} totalColumns 总列数
     * @returns {string} 生成的表头
     */
    generateHeaderByPosition(content, index, totalColumns) {
        // 基于位置和内容特征生成表头
        const trimmedContent = content.trim();
        
        // 如果内容看起来像表头，直接使用
        if (trimmedContent && trimmedContent.length <= 10) {
            return trimmedContent || `列${index + 1}`;
        }
        
        // 根据位置推断表头类型
        if (index === 0) {
            return '序号';
        } else if (index === 1) {
            return '名称';
        } else if (index === totalColumns - 1) {
            return '备注';
        } else {
            return `列${index + 1}`;
        }
    }

    /**
     * 获取最常见的值
     * @param {Array} array 数组
     * @returns {*} 最常见的值
     */
    getMostCommonValue(array) {
        const counts = {};
        let maxCount = 0;
        let mostCommon = array[0];
        
        array.forEach(item => {
            counts[item] = (counts[item] || 0) + 1;
            if (counts[item] > maxCount) {
                maxCount = counts[item];
                mostCommon = item;
            }
        });
        
        return mostCommon;
    }

    /**
     * 获取列对齐方式
     * @param {string} header 表头
     * @returns {string} 对齐方式
     */
    getColumnAlignment(header) {
        // 基于表头内容特征判断对齐方式
        const headerLower = header.toLowerCase();
        
        // 数字类列居中对齐
        if (/^\d+$/.test(header) || headerLower.includes('序号') || headerLower.includes('编号') || 
            headerLower.includes('数量') || headerLower.includes('重量') || headerLower.includes('金额') ||
            headerLower.includes('价格') || headerLower.includes('日期') || headerLower.includes('时间')) {
            return 'center';
        }
        
        // 文本类列左对齐
        if (headerLower.includes('名称') || headerLower.includes('品名') || headerLower.includes('内容') ||
            headerLower.includes('规格') || headerLower.includes('备注') || headerLower.includes('说明')) {
            return 'left';
        }
        
        // 默认居中对齐
        return 'center';
    }

    /**
     * HTML转义
     * @param {string} text 原始文本
     * @returns {string} 转义后的文本
     */
    escapeHtml(text) {
        const map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return text.replace(/[&<>"']/g, function(m) { return map[m]; });
    }

    /**
     * 提取幻灯片中的样式化内容
     * @param {string} slideXml 幻灯片XML内容
     * @returns {string} 带样式的HTML内容
     */
    extractStyledContent(slideXml) {
        try {
            let htmlContent = '';
            
            // 提取形状和文本框，保持布局结构
            const shapeMatches = slideXml.match(/<p:sp[^>]*>[\s\S]*?<\/p:sp>/g);
            if (shapeMatches) {
                shapeMatches.forEach((shape, shapeIndex) => {
                    // 检查是否是文本框
                    if (shape.includes('<p:txBody>')) {
                        const txBodyMatch = shape.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
                        if (txBodyMatch) {
                            const txBody = txBodyMatch[1];
                            
                            // 提取文本框的位置信息
                            const spPrMatch = shape.match(/<p:spPr>([\s\S]*?)<\/p:spPr>/);
                            let position = '';
                            if (spPrMatch) {
                                const xfrmMatch = spPrMatch[1].match(/<a:xfrm>([\s\S]*?)<\/a:xfrm>/);
                                if (xfrmMatch) {
                                    const offMatch = xfrmMatch[1].match(/<a:off[^>]*x="([^"]*)"[^>]*y="([^"]*)"[^>]*>/);
                                    if (offMatch) {
                                        const x = parseInt(offMatch[1]) / 914400; // EMU to points
                                        const y = parseInt(offMatch[2]) / 914400;
                                        position = `style="position: relative; left: ${x}pt; top: ${y}pt;"`;
                                    }
                                }
                            }
                            
                            // 开始文本框容器
                            htmlContent += `<div class="text-box mb-4" ${position}>`;
                            
                            // 提取段落内容
                            const paragraphMatches = txBody.match(/<a:p[^>]*>[\s\S]*?<\/a:p>/g);
                            if (paragraphMatches) {
                                paragraphMatches.forEach(paragraph => {
                                    // 检查段落级别和样式
                                    const levelMatch = paragraph.match(/<a:lvl[^>]*val="(\d+)"/);
                                    const level = levelMatch ? parseInt(levelMatch[1]) : 0;
                                    
                                    // 检查字体大小
                                    const fontSizeMatch = paragraph.match(/<a:sz[^>]*val="(\d+)"/);
                                    const fontSize = fontSizeMatch ? parseInt(fontSizeMatch[1]) : 18;
                                    
                                    // 检查字体样式
                                    const isBold = paragraph.includes('<a:b/>');
                                    const isItalic = paragraph.includes('<a:i/>');
                                    const isUnderline = paragraph.includes('<a:u/>');
                                    
                                    // 检查对齐方式
                                    let alignment = 'left';
                                    if (paragraph.includes('<a:jc val="ctr"/>')) alignment = 'center';
                                    else if (paragraph.includes('<a:jc val="r"/>')) alignment = 'right';
                                    else if (paragraph.includes('<a:jc val="just"/>')) alignment = 'justify';
                                    
                                    // 提取文本内容
                                    const textMatches = paragraph.match(/<a:t[^>]*>([^<]*)<\/a:t>/g);
                                    if (textMatches) {
                                        let text = textMatches.map(t => {
                                            const content = t.match(/<a:t[^>]*>([^<]*)<\/a:t>/);
                                            return content ? content[1] : '';
                                        }).join('');
                                        
                                        if (text.trim()) {
                                            // 根据级别和样式应用CSS
                                            let cssClass = '';
                                            let tag = 'p';
                                            
                                            if (level === 0) {
                                                tag = 'h2';
                                                cssClass = 'text-2xl font-bold text-gray-800 mb-4';
                                            } else if (level === 1) {
                                                tag = 'h3';
                                                cssClass = 'text-xl font-semibold text-gray-700 mb-3';
                                            } else if (level === 2) {
                                                tag = 'h4';
                                                cssClass = 'text-lg font-medium text-gray-600 mb-2';
                                            } else {
                                                tag = 'p';
                                                cssClass = 'text-base text-gray-600 mb-2 leading-relaxed';
                                            }
                                            
                                            // 应用字体样式
                                            if (isBold) cssClass += ' font-bold';
                                            if (isItalic) cssClass += ' italic';
                                            if (isUnderline) cssClass += ' underline';
                                            
                                            // 根据字体大小调整
                                            if (fontSize >= 32) cssClass = cssClass.replace('text-base', 'text-3xl');
                                            else if (fontSize >= 28) cssClass = cssClass.replace('text-base', 'text-2xl');
                                            else if (fontSize >= 24) cssClass = cssClass.replace('text-base', 'text-xl');
                                            else if (fontSize >= 20) cssClass = cssClass.replace('text-base', 'text-lg');
                                            
                                            // 应用对齐方式
                                            cssClass += ` text-${alignment}`;
                                            
                                            htmlContent += `<${tag} class="${cssClass}">${text}</${tag}>`;
                                        }
                                    }
                                });
                            }
                            
                            // 结束文本框容器
                            htmlContent += `</div>`;
                        }
                    }
                });
            }
            
            // 如果没有找到文本框，尝试提取普通段落
            if (!htmlContent) {
                const paragraphMatches = slideXml.match(/<a:p[^>]*>[\s\S]*?<\/a:p>/g);
                if (paragraphMatches) {
                    paragraphMatches.forEach(paragraph => {
                        // 检查段落级别样式
                        const levelMatch = paragraph.match(/<a:lvl[^>]*val="(\d+)"/);
                        const level = levelMatch ? parseInt(levelMatch[1]) : 0;
                        
                        // 提取文本内容
                        const textMatches = paragraph.match(/<a:t[^>]*>([^<]*)<\/a:t>/g);
                        if (textMatches) {
                            const text = textMatches.map(match => {
                                const textMatch = match.match(/<a:t[^>]*>([^<]*)<\/a:t>/);
                                return textMatch ? textMatch[1] : '';
                            }).join('');
                            
                            if (text.trim()) {
                                // 应用样式
                                if (level === 0) {
                                    htmlContent += `<h2 class="text-xl font-bold text-gray-800 mb-3">${text}</h2>`;
                                } else if (level === 1) {
                                    htmlContent += `<h3 class="text-lg font-semibold text-gray-700 mb-2">${text}</h3>`;
                                } else {
                                    htmlContent += `<p class="text-gray-600 mb-2">${text}</p>`;
                                }
                            }
                        }
                    });
                }
            }
            
            return htmlContent || '<p class="text-gray-500 italic">无文本内容</p>';
        } catch (error) {
            console.log('提取样式内容失败:', error.message);
            return '<p class="text-gray-500 italic">样式解析失败</p>';
        }
    }

    /**
     * 提取幻灯片中的图片
     * @param {string} slideXml 幻灯片XML内容
     * @param {number} slideIndex 幻灯片索引
     * @returns {Array} 图片信息数组
     */
    extractSlideImages(slideXml, slideIndex) {
        try {
            const images = [];
            
            // 查找图片引用 - 支持多种格式
            const pictureMatches = slideXml.match(/<a:pic[^>]*>([\s\S]*?)<\/a:pic>/g);
            if (pictureMatches) {
                pictureMatches.forEach((picture, index) => {
                    // 提取图片ID - 支持多种属性
                    let imageId = null;
                    
                    // 尝试不同的属性名
                    const embedMatch = picture.match(/r:embed="([^"]*)"/);
                    const linkMatch = picture.match(/r:link="([^"]*)"/);
                    
                    if (embedMatch) {
                        imageId = embedMatch[1];
                    } else if (linkMatch) {
                        imageId = linkMatch[1];
                    }
                    
                    if (imageId) {
                        images.push({
                            id: imageId,
                            slideIndex: slideIndex,
                            index: index + 1,
                            type: 'embedded',
                            description: `幻灯片${slideIndex}的图片${index + 1}`
                        });
                    }
                });
            }
            
            // 查找形状中的图片
            const shapeMatches = slideXml.match(/<p:pic[^>]*>([\s\S]*?)<\/p:pic>/g);
            if (shapeMatches) {
                shapeMatches.forEach((shape, index) => {
                    const blipMatch = shape.match(/<a:blip[^>]*r:embed="([^"]*)"[^>]*>/);
                    if (blipMatch) {
                        const imageId = blipMatch[1];
                        images.push({
                            id: imageId,
                            slideIndex: slideIndex,
                            index: index + 1,
                            type: 'shape',
                            description: `幻灯片${slideIndex}的形状图片${index + 1}`
                        });
                    }
                });
            }
            
            return images;
        } catch (error) {
            console.log('提取图片信息失败:', error.message);
            return [];
        }
    }

    /**
     * 提取PowerPoint文件中的图片资源
     * @param {Object} zip ZIP文件对象
     * @returns {Object} 图片资源映射
     */
    async extractPowerPointImages(zip) {
        try {
            const images = {};
            
            // 查找所有可能的图片文件
            const imageExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tiff', '.svg'];
            const imageFiles = zip.files.filter(file => {
                const extension = file.path.toLowerCase().split('.').pop();
                return imageExtensions.includes('.' + extension) || 
                       file.path.includes('/media/') || 
                       file.path.includes('/images/');
            });
            
            console.log(`找到 ${imageFiles.length} 个图片文件`);
            
            for (const file of imageFiles) {
                try {
                    const buffer = await file.buffer();
                    const imageType = this.detectImageFormat(buffer);
                    
                    if (imageType) {
                        const base64 = buffer.toString('base64');
                        const mimeType = this.getImageMimeType(imageType);
                        const fileName = file.path.split('/').pop();
                        
                        images[file.path] = {
                            data: `data:${mimeType};base64,${base64}`,
                            type: imageType,
                            size: buffer.length,
                            name: fileName,
                            path: file.path
                        };
                        
                        console.log(`处理图片: ${file.path} (${imageType}, ${buffer.length} bytes)`);
                    }
                } catch (error) {
                    console.log(`处理图片文件失败 ${file.path}:`, error.message);
                }
            }
            
            return images;
        } catch (error) {
            console.log('提取PowerPoint图片失败:', error.message);
            return {};
        }
    }

    /**
     * 获取图片MIME类型
     * @param {string} imageType 图片类型
     * @returns {string} MIME类型
     */
    getImageMimeType(imageType) {
        const mimeTypes = {
            'png': 'image/png',
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'webp': 'image/webp'
        };
        return mimeTypes[imageType] || 'image/png';
    }

    /**
     * 建立图片ID到文件路径的映射关系
     * @param {Object} zip ZIP文件对象
     * @param {Object} imageResources 图片资源映射
     * @returns {Object} 图片ID映射
     */
    async buildImageIdMapping(zip, imageResources) {
        try {
            const mapping = {};
            
            // 查找关系文件
            const relationshipFiles = zip.files.filter(file => 
                file.path.includes('_rels/') && file.path.endsWith('.rels')
            );
            
            for (const relFile of relationshipFiles) {
                try {
                    const content = await relFile.buffer();
                    const xmlContent = content.toString('utf8');
                    
                    // 解析关系XML
                    const relationshipMatches = xmlContent.match(/<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*>/g);
                    if (relationshipMatches) {
                        relationshipMatches.forEach(rel => {
                            const idMatch = rel.match(/Id="([^"]*)"/);
                            const targetMatch = rel.match(/Target="([^"]*)"/);
                            
                            if (idMatch && targetMatch) {
                                const id = idMatch[1];
                                let target = targetMatch[1];
                                
                                // 处理相对路径
                                if (target.startsWith('../')) {
                                    const basePath = relFile.path.replace('/_rels/' + relFile.path.split('/').pop(), '');
                                    target = target.replace('../', basePath + '/');
                                }
                                
                                // 只处理图片关系
                                if (target.includes('/media/') || target.includes('.jpg') || target.includes('.png') || 
                                    target.includes('.jpeg') || target.includes('.gif')) {
                                    mapping[id] = target;
                                }
                            }
                        });
                    }
                } catch (error) {
                    console.log(`处理关系文件失败 ${relFile.path}:`, error.message);
                }
            }
            
            console.log('图片ID映射关系:', mapping);
            return mapping;
        } catch (error) {
            console.log('建立图片ID映射失败:', error.message);
            return {};
        }
    }

    /**
     * 渲染幻灯片图片
     * @param {Array} images 图片信息数组
     * @param {Object} imageResources 图片资源映射
     * @param {Object} imageIdMapping 图片ID映射
     * @returns {string} HTML内容
     */
    renderSlideImages(images, imageResources, imageIdMapping) {
        if (!images || images.length === 0) {
            return '';
        }

        let html = '<div class="slide-images mt-4">';
        html += '<h4 class="text-md font-medium text-gray-700 mb-3">图片内容</h4>';
        html += '<div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">';
        
        images.forEach(image => {
            // 查找对应的图片资源 - 使用映射关系
            let imageResource = null;
            
            // 方法1: 通过ID映射查找
            if (imageIdMapping[image.id]) {
                const mappedPath = imageIdMapping[image.id];
                imageResource = imageResources[mappedPath];
            }
            
            // 方法2: 通过ID直接匹配（备用方法）
            if (!imageResource) {
                imageResource = imageResources[`ppt/media/image${image.id.replace('rId', '')}.jpg`] ||
                               imageResources[`ppt/media/image${image.id.replace('rId', '')}.png`] ||
                               imageResources[`ppt/media/image${image.id.replace('rId', '')}.jpeg`] ||
                               imageResources[`ppt/media/image${image.id.replace('rId', '')}.gif`];
            }
            
            // 方法3: 遍历所有资源查找匹配（最后备用）
            if (!imageResource) {
                const imageId = image.id.replace('rId', '');
                imageResource = Object.values(imageResources).find(resource => {
                    const fileName = resource.name || resource.path || '';
                    return fileName.includes(`image${imageId}`) || 
                           fileName.includes(image.id);
                });
            }
            
            if (imageResource) {
                html += `<div class="image-container">
                    <img src="${imageResource.data}" 
                         alt="${image.description}" 
                         class="w-full h-auto rounded-lg border border-gray-200 shadow-sm"
                         loading="lazy">
                    <p class="text-xs text-gray-500 mt-1">${image.description}</p>
                </div>`;
            } else {
                html += `<div class="image-container">
                    <div class="w-full h-32 bg-gray-100 rounded-lg border border-gray-200 flex items-center justify-center">
                        <span class="text-gray-400 text-sm">图片 ${image.id}</span>
                    </div>
                    <p class="text-xs text-gray-500 mt-1">${image.description} (未找到资源)</p>
                </div>`;
            }
        });
        
        html += '</div></div>';
        return html;
    }
    async previewPowerPoint(buffer) {
        try {
            // 检查文件格式
            const isPptx = buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4B;
            
            if (isPptx) {
                // .pptx文件（ZIP格式）
                return await this.previewPptx(buffer);
            } else {
                // .ppt文件（二进制格式）
                return await this.previewPpt(buffer);
            }
        } catch (error) {
            console.log('PowerPoint解析失败:', error.message);
            return this.previewPowerPointFallback(buffer);
        }
    }

    /**
     * 预览.pptx文件（PowerPoint 2007+）
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewPptx(buffer) {
        try {
            const unzipper = require('unzipper');
            
            // 解析ZIP文件
            const zip = await unzipper.Open.buffer(buffer);
            
            // 提取图片资源
            const imageResources = await this.extractPowerPointImages(zip);
            
            // 建立图片ID到文件路径的映射关系
            const imageIdMapping = await this.buildImageIdMapping(zip, imageResources);
            
            // 提取幻灯片信息
            const slides = [];
            let slideCount = 0;
            
            // 查找presentation.xml获取幻灯片数量
            const presentationEntry = zip.files.find(file => file.path === 'ppt/presentation.xml');
            if (presentationEntry) {
                const presentationXml = await presentationEntry.buffer();
                const presentationText = presentationXml.toString('utf8');
                const presentationMatch = presentationText.match(/<p:sldIdLst[^>]*>([\s\S]*?)<\/p:sldIdLst>/);
                if (presentationMatch) {
                    const slideMatches = presentationMatch[1].match(/<p:sldId[^>]*>/g);
                    slideCount = slideMatches ? slideMatches.length : 0;
                }
            }
            
            // 提取每个幻灯片的内容
            for (let i = 1; i <= slideCount; i++) {
                const slideEntry = zip.files.find(file => file.path === `ppt/slides/slide${i}.xml`);
                if (slideEntry) {
                    try {
                        const slideXml = await slideEntry.buffer();
                        const slideText = slideXml.toString('utf8');
                        
                        // 提取文本内容和样式
                        const textMatches = slideText.match(/<a:t[^>]*>([^<]*)<\/a:t>/g);
                        let slideContent = '';
                        let styledContent = '';
                        
                        if (textMatches) {
                            slideContent = textMatches.map(match => {
                                const textMatch = match.match(/<a:t[^>]*>([^<]*)<\/a:t>/);
                                return textMatch ? textMatch[1] : '';
                            }).filter(text => text.trim()).join(' ');
                            
                            // 构建带样式的HTML内容
                            styledContent = this.extractStyledContent(slideText);
                        }
                        
                        // 提取图片
                        const images = this.extractSlideImages(slideText, i);
                        
                        slides.push({
                            index: i,
                            content: slideContent,
                            styledContent: styledContent,
                            images: images,
                            rawXml: slideText
                        });
                    } catch (slideError) {
                        console.log(`幻灯片${i}解析失败:`, slideError.message);
                        slides.push({
                            index: i,
                            content: '解析失败',
                            rawXml: ''
                        });
                    }
                }
            }
            
            // 构建HTML内容
            let htmlContent = `
                <div class="powerpoint-preview bg-white rounded-lg shadow-lg p-6 max-w-6xl mx-auto">
                    <div class="mb-6">
                        <h1 class="text-3xl font-bold text-gray-800 mb-2">PowerPoint 演示文稿</h1>
                        <p class="text-gray-600">文件大小: ${(buffer.length / 1024 / 1024).toFixed(2)} MB</p>
                        <p class="text-gray-600">幻灯片数量: ${slideCount}</p>
                    </div>
                    
                    <div class="slides-container space-y-8">
                        ${slides.map(slide => `
                            <div class="slide bg-white rounded-lg p-8 border border-gray-200 shadow-sm" style="min-height: 400px; position: relative;">
                                <div class="slide-header mb-6 pb-4 border-b border-gray-200">
                                    <div class="flex items-center">
                                        <div class="w-10 h-10 bg-gradient-to-r from-blue-500 to-purple-600 text-white rounded-full flex items-center justify-center text-sm font-bold mr-4">
                                            ${slide.index}
                                        </div>
                                        <h2 class="text-2xl font-bold text-gray-800">幻灯片 ${slide.index}</h2>
                                    </div>
                                </div>
                                
                                <div class="slide-content relative" style="min-height: 300px;">
                                    <!-- 文本内容 -->
                                    <div class="text-content mb-6">
                                        ${slide.styledContent || '<p class="text-gray-500 italic">无文本内容</p>'}
                                    </div>
                                    
                                    <!-- 图片内容 -->
                                    ${this.renderSlideImages(slide.images, imageResources, imageIdMapping)}
                                </div>
                            </div>
                        `).join('')}
                    </div>
                    
                    <style>
                        .powerpoint-preview .text-box {
                            margin-bottom: 1rem;
                            padding: 0.75rem;
                            border-radius: 0.5rem;
                            background: linear-gradient(135deg, #f8fafc 0%, #ffffff 100%);
                            border: 1px solid #e5e7eb;
                            transition: all 0.3s ease;
                        }
                        .powerpoint-preview .text-box:hover {
                            background: linear-gradient(135deg, #eff6ff 0%, #ffffff 100%);
                            border-color: #3b82f6;
                            transform: translateY(-2px);
                            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.15);
                        }
                        .powerpoint-preview h2 {
                            color: #1f2937;
                            font-weight: 700;
                            margin-bottom: 1rem;
                            font-size: 1.875rem;
                            line-height: 2.25rem;
                        }
                        .powerpoint-preview h3 {
                            color: #374151;
                            font-weight: 600;
                            margin-bottom: 0.75rem;
                            font-size: 1.5rem;
                            line-height: 2rem;
                        }
                        .powerpoint-preview h4 {
                            color: #4b5563;
                            font-weight: 500;
                            margin-bottom: 0.5rem;
                            font-size: 1.25rem;
                            line-height: 1.75rem;
                        }
                        .powerpoint-preview p {
                            color: #6b7280;
                            line-height: 1.6;
                            margin-bottom: 0.75rem;
                            font-size: 1rem;
                        }
                        .powerpoint-preview .slide {
                            background: linear-gradient(135deg, #f8fafc 0%, #ffffff 100%);
                            transition: all 0.3s ease;
                        }
                        .powerpoint-preview .slide:hover {
                            transform: translateY(-4px);
                            box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
                        }
                        .powerpoint-preview .slide-content {
                            display: flex;
                            flex-direction: column;
                            gap: 1.5rem;
                        }
                        .powerpoint-preview .text-content {
                            flex: 1;
                        }
                        .powerpoint-preview .slide-images {
                            flex-shrink: 0;
                        }
                        .powerpoint-preview .slide-images .grid {
                            gap: 1rem;
                        }
                        .powerpoint-preview .image-container {
                            transition: all 0.3s ease;
                        }
                        .powerpoint-preview .image-container:hover {
                            transform: scale(1.02);
                        }
                        .powerpoint-preview .image-container img {
                            border-radius: 0.5rem;
                            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
                        }
                        @media (max-width: 768px) {
                            .powerpoint-preview .slide {
                                padding: 1rem;
                            }
                            .powerpoint-preview .slide-content {
                                flex-direction: column;
                            }
                            .powerpoint-preview h2 {
                                font-size: 1.5rem;
                            }
                            .powerpoint-preview h3 {
                                font-size: 1.25rem;
                            }
                        }
                    </style>
                </div>
            `;
            
            return {
                type: 'powerpoint',
                content: htmlContent,
                slides: slides,
                slideCount: slideCount,
                contentType: 'text/html'
            };
            
        } catch (error) {
            throw new Error(`解析.pptx文件失败: ${error.message}`);
        }
    }

    /**
     * 预览.ppt文件（PowerPoint 97-2003）
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewPpt(buffer) {
        try {
            // .ppt文件是二进制格式，需要特殊处理
            // 这里提供一个基本的解析框架
            
            let htmlContent = '<div class="powerpoint-preview">';
            htmlContent += '<h1 class="text-2xl font-bold mb-6">PowerPoint演示文稿 (.ppt)</h1>';
            
            // 文件信息
            htmlContent += `<div class="mb-4 p-4 bg-blue-50 rounded-lg">
                <h2 class="text-lg font-semibold mb-2">文件信息</h2>
                <p><strong>文件大小:</strong> ${(buffer.length / 1024 / 1024).toFixed(2)} MB</p>
                <p><strong>文件格式:</strong> PowerPoint 97-2003 (.ppt)</p>
                <p><strong>状态:</strong> 二进制格式解析开发中</p>
            </div>`;
            
            // 尝试提取基本信息
            const fileHeader = buffer.toString('ascii', 0, 8);
            htmlContent += `<div class="mb-4 p-4 bg-gray-50 rounded-lg">
                <h3 class="text-md font-semibold mb-2">文件头信息</h3>
                <p><strong>文件头:</strong> ${fileHeader}</p>
            </div>`;
            
            htmlContent += `<div class="p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
                <p class="text-yellow-800"><strong>注意:</strong> .ppt格式（PowerPoint 97-2003）的完整解析功能正在开发中。目前只能显示文件基本信息。</p>
            </div>`;
            
            htmlContent += '</div>';
            
            return {
                type: 'powerpoint',
                content: htmlContent,
                slides: [],
                slideCount: 0,
                contentType: 'text/html'
            };
            
        } catch (error) {
            throw new Error(`解析.ppt文件失败: ${error.message}`);
        }
    }

    /**
     * PowerPoint预览备用方法
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    previewPowerPointFallback(buffer) {
        try {
            // 尝试使用unzipper解析.pptx文件（PowerPoint 2007+格式）
            if (buffer.length > 4 && buffer[0] === 0x50 && buffer[1] === 0x4B) {
                // 这是一个ZIP文件（.pptx格式）
                return {
                    type: 'powerpoint',
                    content: `<div class="powerpoint-preview">
                        <h1 class="text-2xl font-bold mb-6">PowerPoint演示文稿 (.pptx)</h1>
                        <div class="mb-4 p-4 bg-blue-50 rounded-lg">
                            <h2 class="text-lg font-semibold mb-2">文件信息</h2>
                            <p><strong>文件大小:</strong> ${(buffer.length / 1024 / 1024).toFixed(2)} MB</p>
                            <p><strong>文件格式:</strong> PowerPoint 2007+ (.pptx)</p>
                            <p><strong>状态:</strong> 文件格式已识别，内容解析需要进一步开发</p>
                        </div>
                        <div class="p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
                            <p class="text-yellow-800"><strong>注意:</strong> 此PowerPoint文件的内容解析功能正在开发中。目前只能显示文件基本信息。</p>
                        </div>
                    </div>`,
                    contentType: 'text/html'
                };
            } else {
                // 可能是.ppt格式（PowerPoint 97-2003）
                return {
                    type: 'powerpoint',
                    content: `<div class="powerpoint-preview">
                        <h1 class="text-2xl font-bold mb-6">PowerPoint演示文稿 (.ppt)</h1>
                        <div class="mb-4 p-4 bg-blue-50 rounded-lg">
                            <h2 class="text-lg font-semibold mb-2">文件信息</h2>
                            <p><strong>文件大小:</strong> ${(buffer.length / 1024 / 1024).toFixed(2)} MB</p>
                            <p><strong>文件格式:</strong> PowerPoint 97-2003 (.ppt)</p>
                            <p><strong>状态:</strong> 文件格式已识别，内容解析需要进一步开发</p>
                        </div>
                        <div class="p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
                            <p class="text-yellow-800"><strong>注意:</strong> 此PowerPoint文件的内容解析功能正在开发中。目前只能显示文件基本信息。</p>
                        </div>
                    </div>`,
                    contentType: 'text/html'
                };
            }
        } catch (error) {
            return {
                type: 'powerpoint',
                content: `<div class="powerpoint-preview">
                    <h1 class="text-2xl font-bold mb-6">PowerPoint演示文稿</h1>
                    <div class="mb-4 p-4 bg-red-50 border border-red-200 rounded-lg">
                        <h2 class="text-lg font-semibold mb-2">解析失败</h2>
                        <p><strong>文件大小:</strong> ${(buffer.length / 1024 / 1024).toFixed(2)} MB</p>
                        <p><strong>错误信息:</strong> ${error.message}</p>
                    </div>
                    <div class="p-4 bg-yellow-50 border border-yellow-200 rounded-lg">
                        <p class="text-yellow-800"><strong>注意:</strong> PowerPoint文件预览功能正在开发中，目前只能显示文件基本信息。</p>
                    </div>
                </div>`,
                contentType: 'text/html'
            };
        }
    }

    /**
     * 预览Markdown文件
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewMarkdown(buffer) {
        try {
            const text = buffer.toString('utf8');
            
            // 确保marked已加载
            if (!marked) {
                const markedModule = await import('marked');
                marked = markedModule.marked;
            }
            
            const html = marked.parse(text);
            return {
                type: 'markdown',
                content: html,
                rawContent: text,
                contentType: 'text/html'
            };
        } catch (error) {
            throw new Error(`解析Markdown文件失败: ${error.message}`);
        }
    }

    /**
     * 预览文本文件
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewText(buffer) {
        try {
            // 尝试检测编码
            let text;
            try {
                text = buffer.toString('utf8');
            } catch (error) {
                // 如果UTF-8失败，尝试GBK编码
                text = iconv.decode(buffer, 'gbk');
            }
            
            return {
                type: 'text',
                content: text,
                contentType: 'text/plain'
            };
        } catch (error) {
            throw new Error(`解析文本文件失败: ${error.message}`);
        }
    }

    /**
     * 预览CSV文件
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewCsv(buffer) {
        try {
            // 尝试不同的编码格式
            let csvString;
            const encodings = ['utf8', 'gbk', 'gb2312', 'big5', 'utf16le'];
            
            for (const encoding of encodings) {
                try {
                    csvString = iconv.decode(buffer, encoding);
                    // 检查是否包含乱码字符
                    if (!/[\uFFFD]/.test(csvString)) {
                        break;
                    }
                } catch (e) {
                    continue;
                }
            }
            
            if (!csvString) {
                csvString = buffer.toString('utf8');
            }
            
            // 解析CSV数据
            const lines = csvString.split(/\r?\n/).filter(line => line.trim());
            const data = [];
            
            for (const line of lines) {
                const row = this.parseCsvLine(line);
                data.push(row);
            }
            
            // 获取最大列数
            const maxCols = Math.max(...data.map(row => row.length));
            
            // 确保所有行都有相同的列数
            for (let i = 0; i < data.length; i++) {
                while (data[i].length < maxCols) {
                    data[i].push('');
                }
            }
            
            return {
                type: 'csv',
                data: data,
                maxRows: data.length,
                maxCols: maxCols,
                contentType: 'application/json'
            };
        } catch (error) {
            throw new Error(`解析CSV文件失败: ${error.message}`);
        }
    }

    /**
     * 解析CSV行数据
     * @param {string} line CSV行
     * @returns {Array} 解析后的数据
     */
    parseCsvLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    // 转义的双引号
                    current += '"';
                    i++;
                } else {
                    // 开始或结束引号
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                // 字段分隔符
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        // 添加最后一个字段
        result.push(current.trim());
        
        return result;
    }

    /**
     * 预览XML文件
     * @param {Buffer} buffer 文件内容
     * @returns {Object} 预览信息
     */
    async previewXml(buffer) {
        try {
            const text = buffer.toString('utf8');
            const parser = new xml2js.Parser();
            const result = await parser.parseStringPromise(text);
            
            return {
                type: 'xml',
                content: JSON.stringify(result, null, 2),
                rawContent: text,
                contentType: 'application/json'
            };
        } catch (error) {
            throw new Error(`解析XML文件失败: ${error.message}`);
        }
    }

    /**
     * 预览文件
     * @param {string} url 文件URL
     * @returns {Promise<Object>} 预览结果
     */
    async previewFile(url) {
        console.log('🔍 [DEBUG] previewFile 开始执行');
        console.log('🔍 [DEBUG] 输入URL:', url);
        
        try {
            // 检查文件格式是否支持
            console.log('🔍 [DEBUG] 检查文件格式支持');
            if (!this.isSupportedFormat(url)) {
                console.log('🔍 [DEBUG] 文件格式不支持');
                throw new Error('不支持的文件格式');
            }
            console.log('🔍 [DEBUG] 文件格式支持');

            // 下载文件
            console.log('🔍 [DEBUG] 开始下载文件');
            const buffer = await this.downloadFile(url);
            console.log('🔍 [DEBUG] 文件下载完成，大小:', buffer.length, '字节');
            
            const ext = this.getFileExtension(url);
            console.log('🔍 [DEBUG] 文件扩展名:', ext);

            // 根据文件类型进行预览
            console.log('🔍 [DEBUG] 开始处理文件类型:', ext);
            switch (ext) {
                case 'pdf':
                    console.log('🔍 [DEBUG] 处理PDF文件');
                    return await this.previewPdf(buffer);
                case 'xls':
                case 'xlsx':
                    console.log('🔍 [DEBUG] 处理Excel文件');
                    return await this.previewExcel(buffer);
                case 'csv':
                    console.log('🔍 [DEBUG] 处理CSV文件');
                    return await this.previewCsv(buffer);
                case 'doc':
                case 'docx':
                    console.log('🔍 [DEBUG] 处理Word文件');
                    return await this.previewWord(buffer, ext);
                case 'ppt':
                case 'pptx':
                    console.log('🔍 [DEBUG] 处理PowerPoint文件');
                    return await this.previewPowerPoint(buffer);
                case 'markdown':
                case 'md':
                    console.log('🔍 [DEBUG] 处理Markdown文件');
                    return await this.previewMarkdown(buffer);
                case 'txt':
                    console.log('🔍 [DEBUG] 处理文本文件');
                    return await this.previewText(buffer);
                case 'xml':
                    console.log('🔍 [DEBUG] 处理XML文件');
                    return await this.previewXml(buffer);
                default:
                    console.log('🔍 [DEBUG] 不支持的文件格式:', ext);
                    throw new Error('不支持的文件格式');
            }
        } catch (error) {
            console.log('🔍 [DEBUG] previewFile 发生错误:', error.message);
            throw error;
        }
    }

    /**
     * 获取文件信息
     * @param {string} url 文件URL
     * @returns {Object} 文件信息
     */
    getFileInfo(url) {
        const ext = this.getFileExtension(url);
        const contentType = this.getContentType(url);
        const isSupported = this.isSupportedFormat(url);

        return {
            url: url,
            extension: ext,
            contentType: contentType,
            isSupported: isSupported,
            fileName: this.getFileName(url)
        };
    }

    /**
     * 获取文件名
     * @param {string} url 文件URL
     * @returns {string} 文件名
     */
    getFileName(url) {
        try {
            const urlObj = new URL(url);
            const pathname = urlObj.pathname;
            return path.basename(pathname);
        } catch (error) {
            return path.basename(url);
        }
    }
    
    /**
     * 验证数据是否为真正的图片
     * @param {Buffer} data 图片数据
     * @param {string} imageType 图片类型
     * @returns {boolean} 是否为有效图片
     */
    validateImageData(data, imageType) {
        try {
            if (!data || data.length < 8) {
                return false;
            }
            
            // 检查文件头部签名
            const header = data.slice(0, 8);
            
            switch (imageType.toLowerCase()) {
                case 'image/png':
                    // PNG签名: 89 50 4E 47 0D 0A 1A 0A
                    return header[0] === 0x89 && 
                           header[1] === 0x50 && 
                           header[2] === 0x4E && 
                           header[3] === 0x47 && 
                           header[4] === 0x0D && 
                           header[5] === 0x0A && 
                           header[6] === 0x1A && 
                           header[7] === 0x0A;
                
                case 'image/jpeg':
                case 'image/jpg':
                    // JPEG签名: FF D8 FF
                    return header[0] === 0xFF && 
                           header[1] === 0xD8 && 
                           header[2] === 0xFF;
                
                case 'image/gif':
                    // GIF签名: 47 49 46 38 (GIF8)
                    return (header[0] === 0x47 && 
                            header[1] === 0x49 && 
                            header[2] === 0x46 && 
                            header[3] === 0x38) ||
                           // 或者 GIF87a/GIF89a
                           (header[0] === 0x47 && 
                            header[1] === 0x49 && 
                            header[2] === 0x46 && 
                            header[3] === 0x38 && 
                            (header[4] === 0x37 || header[4] === 0x39) && 
                            header[5] === 0x61);
                
                case 'image/bmp':
                    // BMP签名: 42 4D (BM)
                    return header[0] === 0x42 && header[1] === 0x4D;
                
                default:
                    // 对于未知类型，尝试检测常见图片格式
                    return this.detectImageFormat(header);
            }
        } catch (error) {
            console.warn(`图片验证失败: ${error.message}`);
            return false;
        }
    }
    
    /**
     * 检测图片格式
     * @param {Buffer} header 文件头部
     * @returns {boolean} 是否为图片格式
     */
    detectImageFormat(header) {
        // PNG
        if (header[0] === 0x89 && header[1] === 0x50 && header[2] === 0x4E && header[3] === 0x47) {
            return true;
        }
        
        // JPEG
        if (header[0] === 0xFF && header[1] === 0xD8 && header[2] === 0xFF) {
            return true;
        }
        
        // GIF
        if (header[0] === 0x47 && header[1] === 0x49 && header[2] === 0x46 && header[3] === 0x38) {
            return true;
        }
        
        // BMP
        if (header[0] === 0x42 && header[1] === 0x4D) {
            return true;
        }
        
        // WebP
        if (header[0] === 0x52 && header[1] === 0x49 && header[2] === 0x46 && header[3] === 0x46 &&
            header[8] === 0x57 && header[9] === 0x45 && header[10] === 0x42 && header[11] === 0x50) {
            return true;
        }
        
        // TIFF
        if ((header[0] === 0x49 && header[1] === 0x49 && header[2] === 0x2A && header[3] === 0x00) ||
            (header[0] === 0x4D && header[1] === 0x4D && header[2] === 0x00 && header[3] === 0x2A)) {
            return true;
        }
        
        return false;
    }
}

module.exports = new FilePreview(); 