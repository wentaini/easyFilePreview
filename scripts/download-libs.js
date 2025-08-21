// 下载前端库文件到本地
const https = require('https');
const fs = require('fs');
const path = require('path');

// 库文件配置
const libraries = [
    {
        name: 'tailwindcss',
        url: 'https://cdn.tailwindcss.com',
        localPath: 'src/public/js/tailwindcss.js'
    },
    {
        name: 'vue',
        url: 'https://unpkg.com/vue@3/dist/vue.global.js',
        localPath: 'src/public/js/vue.global.js'
    },
    {
        name: 'marked',
        url: 'https://cdn.jsdelivr.net/npm/marked@12.0.0/marked.min.js',
        localPath: 'src/public/js/marked.min.js'
    },
    {
        name: 'chartjs',
        url: 'https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js',
        localPath: 'src/public/js/chart.js'
    }
];

// 下载文件的函数
function downloadFile(url, filePath) {
    return new Promise((resolve, reject) => {
        console.log(`正在下载 ${url}...`);
        
        const file = fs.createWriteStream(filePath);
        
        https.get(url, (response) => {
            if (response.statusCode !== 200) {
                reject(new Error(`下载失败: ${response.statusCode}`));
                return;
            }
            
            response.pipe(file);
            
            file.on('finish', () => {
                file.close();
                console.log(`✅ 下载完成: ${filePath}`);
                resolve();
            });
            
            file.on('error', (err) => {
                fs.unlink(filePath, () => {}); // 删除不完整的文件
                reject(err);
            });
        }).on('error', (err) => {
            reject(err);
        });
    });
}

// 确保目录存在
function ensureDirectoryExists(filePath) {
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
}

// 主函数
async function downloadLibraries() {
    console.log('🚀 开始下载前端库文件...\n');
    
    for (const lib of libraries) {
        try {
            ensureDirectoryExists(lib.localPath);
            await downloadFile(lib.url, lib.localPath);
        } catch (error) {
            console.error(`❌ 下载 ${lib.name} 失败:`, error.message);
        }
    }
    
    console.log('\n🎉 所有库文件下载完成！');
}

// 运行下载
downloadLibraries().catch(console.error);
