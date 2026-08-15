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

> **Win7 兼容性要求**：客户运行环境为 Win7 时必须用 **Python 3.8.10** 打包。Python 3.9+ 编译产物依赖
> `api-ms-win-core-path-l1-1-0.dll`（Win8+ 才有的 API set），在 Win7 上会报“无法启动此程序，因为计算机中丢失
> api-ms-win-core-path-l1-1-0.dll”。请使用 `build_exe.bat`（自动创建 `.build_venv38` 的 Python 3.8 环境）。
>
> Win7 还须安装过 Universal C Runtime 更新（KB2999226），否则会报 `api-ms-win-crt-*` 丢失。

```powershell
# 手动方式（需要已安装 Python 3.8.10）
python -m venv .build_venv38
.\.build_venv38\Scripts\python.exe -m pip install -r requirements.txt
.\.build_venv38\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --name Adayo8D-AI --add-data "..\8D模板.pptx;." ai8d.py
```

或直接双击 `build_exe.bat`（推荐）。成品位于 `dist\Adayo8D-AI.exe`。图片在上传前会自动缩放至最长边 1600 px，并转换为压缩 JPEG，以降低请求大小。

## 注意

- AI 仅根据材料生成；证据不足时会保留空白字段并提示人工补充。
- 图片始终按每批最多 4 张识别，报告正文按 D0–D3、D4、D5–D7、D8 四段生成；单次请求限制为 105 秒，避免超过上游网关的 120 秒限制。
- 仅重试当前失败的小批次/小段（最多 1 次）；某批图片识别超时时使用保守证据说明继续生成，不会让整份报告作废。
- EXE、构建目录、虚拟环境、草稿和测试材料均已被 Git 忽略。
