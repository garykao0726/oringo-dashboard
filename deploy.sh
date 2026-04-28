#!/usr/bin/env bash
# 一鍵部署：git commit + push 到 GitHub（GitHub Pages 自動更新）
# 用法：./deploy.sh "更新說明"   ← 省略則自動用時間戳

set -e
MSG="${1:-chore: auto-deploy $(date '+%Y-%m-%d %H:%M')}"

echo "▶ 狀態確認..."
git status --short

echo ""
echo "▶ 暫存所有前端變更（.html / .css / docs/）"
git add index.html marketing.html seo.html finance.html docs/ 2>/dev/null || true
# 排除 GAS / secrets（已在 .gitignore）

if git diff --cached --quiet; then
  echo "⚠️  沒有需要提交的變更，略過 commit。"
else
  echo "▶ Commit: $MSG"
  git commit -m "$MSG"
fi

echo "▶ Push 到 GitHub main..."
git push origin HEAD:main

echo ""
echo "✅ 部署完成！GitHub Pages 約 1 分鐘後更新。"
