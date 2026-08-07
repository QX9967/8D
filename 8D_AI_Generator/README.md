# Adayo 8D AI 报告生成器

程序读取 D0–D8 材料文件夹中的 TXT 和图片，调用 AI 生成结构化内容，并填入可编辑的 PowerPoint 8D 模板。

## 使用

1. 将 `dist\Adayo8D-AI.exe` 放入材料根目录。
2. 建立 `D0`–`D8`（或 `00`–`08`）文件夹，放入对应 TXT 与图片。
3. 双击 EXE，选择“生成 PPT”。首次使用时输入 Packy API Key。
4. 输出文件为同目录的 `8D报告_AI生成.pptx` 和同名 JSON 草稿。

菜单中的“清空 D0-D8 文件夹”会递归删除材料文件，并在删除前显示数量、要求确认。命令行可使用 `--clear --yes` 跳过确认。

## 模板

默认模板为仓库根目录的 `D:\Project\Adayo\8D\8D模板.pptx`，打包时会嵌入 EXE。

程序按页面标题定位 D0–D8、D4 发生/流出 5WHY、D5–D8，而不是依赖固定页码。模板中应保留这些标题文字。

`{{content}}` 只会被替换为 AI 内容；没有内容时仅移除占位符，保留“编制：”“背景：”等模板标签。

## 本地开发与打包

```powershell
python -m venv .build_venv
.\.build_venv\Scripts\python.exe -m pip install -r requirements.txt
.\.build_venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --name Adayo8D-AI --add-data "..\8D模板.pptx;." ai8d.py
```

成品位于 `dist\Adayo8D-AI.exe`。图片在上传前会自动缩放至最长边 1600 px，并转换为压缩 JPEG，以降低请求大小。

## 注意

- AI 仅根据材料生成；证据不足时会保留空白字段并提示人工补充。
- API 请求超时为 90 秒，客户端自动重试最多两次。
- EXE、构建目录、虚拟环境、草稿和测试材料均已被 Git 忽略。
