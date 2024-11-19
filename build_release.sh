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

echo "開始建置流程..."

# 建立 .gitignore
echo "建立 .gitignore..."
cat > .gitignore << EOL
.DS_Store
.venv/
__pycache__/
*.pyc
config.sh
build/
dist/
backup/
EOL

# Git 操作
echo "準備 Git 提交..."
git add main.py
git add teacher_doc_generator.py
git add requirements.txt
git add .github/workflows/build.yml
git add courses.json
git add teachers.json
git add 使用說明.md
git add .gitignore

# 提交更改
echo "提交更改..."
git commit -m "Update project files"

# 設定遠端倉庫
repo_url_with_auth="https://$GITHUB_USERNAME:$GITHUB_TOKEN@${GITHUB_REPO#https://}"
git remote remove origin 2>/dev/null || true
git remote add origin "$repo_url_with_auth"

echo "推送到 GitHub..."
git push -f origin main

echo "完成！" 