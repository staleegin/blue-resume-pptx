# ChatGPT 实测流程（2026-10-24）

截图：[chatgpt.png](chatgpt.png)

## 提示词

上传 `blue-resume-pptx-main.zip`，然后发送：

```text
我上传了简历 PPT 技能的压缩包。请：
1. 解压后阅读 SKILL.md 和 resume_pptx.py 开头的使用说明；
2. 按 SKILL.md 里的填写模板，自己编一份虚构的中文简历（数据全部虚构，邮箱用 example.com），不用问我；
3. 整理成 data.json，运行：python resume_pptx.py data.json 简历.pptx；
4. 把 简历.pptx 给我下载，并贴出脚本实际打印的那一行 JSON。

必须真实运行这个脚本，不要自己另画版式；运行失败就贴出报错，不要编造结果。
```

## 结果

- 思考 1 分 8 秒，运行了 `python resume_pptx.py data.json 简历.pptx`（Python 版）。
- 虚构了一份中文运营岗简历（教育、实习、项目、校园实践、核心技能、个人优势），邮箱用 example.com。
- 贴出了脚本打印的 JSON（截图中可见部分）：

```json
{"out": "简历.pptx", "level": 0, "fontSize": 9, "lineHeight": 15.1, "bottom": 661, "pageBottom": 813.85, "overflowPt": 0, "movedToRight": []}
```

- 自检表：PPTX 约 38 KB、1 页、A4 纵向、level 0、overflowPt 0、movedToRight 为空，并转成 PDF 检查了预览图。
- 提供 简历.pptx 下载。

## 观察

- 贴出了脚本输出的 JSON，可以确认真的运行了脚本，这是和豆包流程最大的区别。
- 提示词里写明「必须真实运行」「失败就贴报错」，能有效避免 AI 自己另画版式或编造结果。
