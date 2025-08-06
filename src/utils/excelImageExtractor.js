const unzipper = require('unzipper');
const path = require('path');

class ExcelImageExtractor {
    /**
     * 从Excel文件中提取图片
     * @param {Buffer} buffer Excel文件内容
     * @returns {Object} 包含图片信息的对象
     */
    async extractImages(buffer) {
        try {
            const images = {};
            const imageCounts = {};
            
            // 检查文件头部，判断是否为ZIP格式（.xlsx）
            const isZipFormat = buffer.length >= 4 && 
                               buffer[0] === 0x50 && 
                               buffer[1] === 0x4B && 
                               buffer[2] === 0x03 && 
                               buffer[3] === 0x04;
            
            if (!isZipFormat) {
                console.log('⚠️  文件不是ZIP格式，可能是.xls文件，无法提取图片');
                return {
                    images: {},
                    imageCounts: {},
                    totalImages: 0
                };
            }
            
            // 使用unzipper解析Excel文件（本质上是ZIP文件）
            const directory = await unzipper.Open.buffer(buffer);
            
            // 查找所有图片文件
            const imageFiles = directory.files.filter(file => {
                const fileName = file.path.toLowerCase();
                return fileName.includes('media/') || 
                       fileName.includes('drawings/') ||
                       fileName.endsWith('.png') ||
                       fileName.endsWith('.jpg') ||
                       fileName.endsWith('.jpeg') ||
                       fileName.endsWith('.gif') ||
                       fileName.endsWith('.bmp');
            });
            
            console.log(`🔍 找到 ${imageFiles.length} 个图片文件`);
            
            // 处理每个图片文件
            for (const file of imageFiles) {
                try {
                    const fileName = path.basename(file.path);
                    const fileExtension = path.extname(fileName).toLowerCase();
                    
                    // 确定图片类型
                    let imageType = 'image/png';
                    if (fileExtension === '.jpg' || fileExtension === '.jpeg') {
                        imageType = 'image/jpeg';
                    } else if (fileExtension === '.gif') {
                        imageType = 'image/gif';
                    } else if (fileExtension === '.bmp') {
                        imageType = 'image/bmp';
                    }
                    
                    // 读取图片数据
                    const imageData = await file.buffer();
                    
                    // 检查图片数据是否有效
                    if (!imageData || imageData.length === 0) {
                        console.log(`⚠️  跳过空图片: ${fileName}`);
                        continue;
                    }
                    
                    // 生成base64数据
                    const base64Data = imageData.toString('base64');
                    
                    // 检查base64数据是否有效
                    if (!base64Data || base64Data.length === 0) {
                        console.log(`⚠️  跳过无效base64图片: ${fileName}`);
                        continue;
                    }
                    
                    // 验证base64数据是否为真正的图片
                    const isValidImage = this.validateImageData(imageData, imageType);
                    if (!isValidImage) {
                        console.log(`⚠️  跳过非图片数据: ${fileName} (检测到非图片内容)`);
                        continue;
                    }
                    
                    // 生成图片信息
                    const imageInfo = {
                        id: `image_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
                        name: fileName,
                        type: imageType,
                        data: imageData,
                        base64: base64Data,
                        src: `data:${imageType};base64,${base64Data}`,
                        size: imageData.length,
                        path: file.path
                    };
                    
                    // 根据文件路径确定工作表
                    let sheetName = 'Sheet1'; // 默认工作表
                    
                    // 尝试从路径中提取工作表信息
                    if (file.path.includes('drawings/')) {
                        // 从drawing关系文件中提取工作表信息
                        const drawingMatch = file.path.match(/drawings\/(drawing\d+)\.xml/);
                        if (drawingMatch) {
                            // 这里可以进一步解析drawing文件来确定工作表
                            // 简化处理，暂时使用默认工作表
                        }
                    }
                    
                    // 将图片添加到对应工作表
                    if (!images[sheetName]) {
                        images[sheetName] = [];
                    }
                    if (!imageCounts[sheetName]) {
                        imageCounts[sheetName] = 0;
                    }
                    
                    images[sheetName].push(imageInfo);
                    imageCounts[sheetName]++;
                    
                    console.log(`✅ 提取图片: ${fileName} (${imageType}, ${imageData.length} bytes)`);
                    
                } catch (error) {
                    console.warn(`⚠️  处理图片文件失败: ${file.path} - ${error.message}`);
                }
            }
            
            return {
                images,
                imageCounts,
                totalImages: imageFiles.length
            };
            
        } catch (error) {
            console.error('❌ 提取Excel图片失败:', error.message);
            return {
                images: {},
                imageCounts: {},
                totalImages: 0
            };
        }
    }
    
    /**
     * 解析Excel文件中的drawing关系，确定图片在工作表中的位置
     * @param {Buffer} buffer Excel文件内容
     * @returns {Object} 图片位置信息
     */
    async parseDrawingRelations(buffer) {
        try {
            const directory = await unzipper.Open.buffer(buffer);
            const positions = {};
            
            // 查找drawing关系文件
            const drawingRels = directory.files.filter(file => 
                file.path.includes('drawings/_rels/') && file.path.endsWith('.xml.rels')
            );
            
            for (const relFile of drawingRels) {
                try {
                    const content = await relFile.buffer();
                    const xmlContent = content.toString('utf8');
                    
                    // 解析XML关系文件
                    // 这里可以进一步解析来确定图片与工作表的对应关系
                    console.log(`📄 找到drawing关系文件: ${relFile.path}`);
                    
                } catch (error) {
                    console.warn(`⚠️  解析drawing关系文件失败: ${relFile.path}`);
                }
            }
            
            return positions;
            
        } catch (error) {
            console.error('❌ 解析drawing关系失败:', error.message);
            return {};
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

module.exports = ExcelImageExtractor; 