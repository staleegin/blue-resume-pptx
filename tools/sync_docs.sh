#!/bin/sh
# 把网站（docs/）用到的文件从源文件同步过去：排版核心、Python 脚本、示例。改完源文件后运行一次。
set -e
cd "$(dirname "$0")/.."
cp skill/blue-resume-pptx/scripts/resume_core.js docs/lib/resume_core.js
cp resume_pptx.py docs/resume_pptx.py
for f in zh_campus.json zh_campus.pptx zh_campus.pdf en_marketing.json en_marketing.pptx en_marketing.pdf \
         preview_zh_campus.png preview_en_marketing.png avatar.png; do
  cp "examples/$f" "docs/examples/$f"
done
cp tests/platform_runs/chatgpt.png tests/platform_runs/doubao.png docs/img/
echo "docs/ 已同步"
