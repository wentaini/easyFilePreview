#!/bin/bash

# 下载前端库文件到本地
echo "🚀 开始下载前端库文件..."

# 确保目录存在
mkdir -p src/public/js

# 下载 TailwindCSS
echo "正在下载 TailwindCSS..."
curl -L "https://cdn.tailwindcss.com" -o "src/public/js/tailwindcss.js"
if [ $? -eq 0 ]; then
    echo "✅ TailwindCSS 下载完成"
else
    echo "❌ TailwindCSS 下载失败"
fi

# 下载 Vue.js
echo "正在下载 Vue.js..."
curl -L "https://unpkg.com/vue@3/dist/vue.global.js" -o "src/public/js/vue.global.js"
if [ $? -eq 0 ]; then
    echo "✅ Vue.js 下载完成"
else
    echo "❌ Vue.js 下载失败"
fi

# 下载 Marked.js
echo "正在下载 Marked.js..."
curl -L "https://cdn.jsdelivr.net/npm/marked@12.0.0/marked.min.js" -o "src/public/js/marked.min.js"
if [ $? -eq 0 ]; then
    echo "✅ Marked.js 下载完成"
else
    echo "❌ Marked.js 下载失败"
fi

# 下载 Chart.js
echo "正在下载 Chart.js..."
curl -L "https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.js" -o "src/public/js/chart.js"
if [ $? -eq 0 ]; then
    echo "✅ Chart.js 下载完成"
else
    echo "❌ Chart.js 下载失败"
fi

echo "🎉 所有库文件下载完成！"
