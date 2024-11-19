#!/bin/bash

# 載入配置
if [ -f "config.sh" ]; then
    source config.sh
else
    echo "找不到配置檔，請輸入 GitHub 資訊..."
    echo "請輸入 GitHub 使用者名稱："
    read GITHUB_USERNAME
    echo "請輸入 GitHub 個人訪問令牌(Personal Access Token)："
    read -s GITHUB_TOKEN
    echo
    echo "請輸入 GitHub 倉庫地址："
    read GITHUB_REPO
    
    # 儲存配置
    echo "#!/bin/bash" > config.sh
    echo "GITHUB_USERNAME=\"$GITHUB_USERNAME\"" >> config.sh
    echo "GITHUB_TOKEN=\"$GITHUB_TOKEN\"" >> config.sh
    echo "GITHUB_REPO=\"$GITHUB_REPO\"" >> config.sh
    chmod +x config.sh
fi

echo "開始清空 GitHub 倉庫..."

# 建立一個臨時分支，不影響本地檔案
git checkout --orphan temp_branch_remote
git commit --allow-empty -m "Clean repository"

# 設定遠端倉庫
repo_url_with_auth="https://$GITHUB_USERNAME:$GITHUB_TOKEN@${GITHUB_REPO#https://}"
git remote remove origin 2>/dev/null || true
git remote add origin "$repo_url_with_auth"

# 強制推送空分支
git push -f origin temp_branch_remote:main

# 切回原本的分支
git checkout -f main
git branch -D temp_branch_remote

echo "GitHub 倉庫已清空！本地檔案未受影響。"
