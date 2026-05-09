#!/bin/bash
# 图片文字提取工具 - 启动脚本

cd "$(dirname "$0")"

echo "🚀 启动图片文字提取工具..."
echo ""

# 检查node_modules
if [ ! -d "node_modules" ]; then
    echo "📦 首次运行，正在安装依赖..."
    npm install
fi

# 启动应用
npm start