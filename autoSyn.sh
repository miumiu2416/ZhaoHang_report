#!/bin/bash
# 每次运行该文件将自动提交代码到远程仓库

# 如果需要放弃本地工作区已暂存和未暂存的所有更改，首先执行
# git reset --hard HEAD
# 然后再运行此文件 bash autoSyn.sh，拉取最新代码

set -e

# 定义样式
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
BOLD='\033[1m'
RESET='\033[0m'

echo -e "${BOLD}${GREEN}🚀 开始自动同步流程...${RESET}"

# 1. 恢复 stash (如果有)
if git stash list | grep -q .; then
    echo -e "${YELLOW}💾 检测到暂存内容，正在恢复...${RESET}"
    git stash pop
fi

# 2. 添加所有更改
echo -e "\n${BLUE}➕ 添加项目下所有文件的更改...${RESET}"
git add .

# 3. 智能获取服务器标识 (Linux=IP, macOS=主机名)
get_server_id() {
    OS_TYPE=$(uname -s)
    # Linux系统使用内网IP
    if [ "$OS_TYPE" = "Linux" ]; then
        # 优先尝试 hostname -I
        ID=$(hostname -I 2>/dev/null | awk '{print $1}')
        # 备用方案 ip addr
        if [ -z "$ID" ]; then
            ID=$(ip addr show 2>/dev/null | grep "inet " | grep -v 127.0.0.1 | head -n 1 | awk '{print $2}' | cut -d/ -f1)
        fi
        # 保底方案，使用主机名
        if [ -z "$ID" ]; then
            ID=$(hostname)
        fi
    
    # macOS使用主机名
    elif [ "$OS_TYPE" = "Darwin" ]; then
        ID=$(hostname)
    else
        # 其他系统用主机名
        ID=$(hostname)
    fi
    
    echo "$ID"
}

SERVER_ID=$(get_server_id)
echo -e "${BOLD}🆔 当前标识: $SERVER_ID (系统: $(uname -s))${RESET}"

# 4. 提交 (如果有变化)
if ! git diff --cached --quiet; then
    echo -e "\n${BLUE}📝 正在提交代码...${RESET}"
    git commit -m "auto sync: updates from [$SERVER_ID]"
else
    echo -e "\n${GREEN}✅ 没有检测到新的更改，跳过 commit。${RESET}"
fi

# 5. 拉取 (强制 rebase)
echo -e "\n${BLUE}⬇️ 拉取远程更新 (Rebase 模式)...${RESET}"
git pull --rebase origin main

# 6. 推送
# 检查本地是否有比远程新的提交
if [ -n "$(git log origin/main..HEAD --oneline)" ]; then
    echo -e "\n${BLUE}⬆️ 推送代码...${RESET}"
    git push origin main
else
    echo -e "\n${GREEN}✅ 本地无待推送更改，跳过 push。${RESET}"
fi

echo -e "\n${BOLD}${GREEN}🎉 同步完成！${RESET}"