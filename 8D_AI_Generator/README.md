# Adayo 8D AI 报告生成器

程序读取 D0--D8 材料文件夹中的文字说明和图片，调用 MiniMax-M3 多模态模型生成结构化 8D 内容，然后以可编辑元素填充 PowerPoint 模板。

## 最简单的使用方式

1. 将 `Adayo8D-AI.exe` 放进一个报告材料目录，例如 `D:\Project\Adayo\8D\Test`。
2. 在该目录创建 `00`--`08` 文件夹（也兼容现有的 `D0`--`D8` 文件夹），把对应阶段的 TXT 与图片放进去。
3. 双击 EXE。首次输入 Packy API Key 后，程序自动使用内置模板并在同一目录输出 `8D报告_AI生成.pptx` 和 AI 草稿 JSON。
4. 程序完成或报错时会停留在窗口中，按 Enter 才关闭。

## 命令行方式

1. 安装 Python 3.10 或更新版本。
2. 双击 `build_exe.bat` 生成 `dist\Adayo8D-AI.exe`，或执行：

   ```powershell
   python -m pip install -r requirements.txt
   python ai8d.py --template "D:\Project\Adayo\8D\Test\8D模板.pptx" --materials "D:\Project\Adayo\8D\Test"
   ```

3. 首次运行时，程序会在控制台要求输入 Packy API Key。该值保存到当前 Windows 用户的凭据管理器；不会写入 PPT、日志或项目文件。需要更换密钥时运行：

   ```powershell
   python ai8d.py --reset-key
   ```

4. 输出文件默认为 `out\8D报告_AI生成.pptx`。模板与原始材料均保持不变。

## 材料要求

- 资料根目录内有 `D0 相关信息` 至 `D8 结案总结` 文件夹。
- 每个文件夹可放 PNG/JPG/JPEG/BMP/WEBP 图片以及 TXT 说明。
- TXT 支持 UTF-8、GB18030/GBK 与 UTF-16；可直接保存现场记录，不需要严格字段格式。
- 模型只根据材料生成内容。证据不足时会写入“待确认”，不会自行编造数据、责任人、日期或验证结论。

## 当前模板映射

- 第 1--11 页分别对应封面、目录、D0--D8。
- 图片页紧跟对应的 D 阶段正文页插入；AI 按图片主题决定单页或组合页。
- 图片完整等比显示，不裁剪，不显示图片文件名。
- 原模板中的表格仍为原生 PowerPoint 表格，插入照片仍可在 PowerPoint 中裁剪、移动和替换。

## 限制

- 本版针对 `8D模板.pptx` 的 12 页结构；更换客户模板需要调整 `fill_template()` 中的映射。
- 生成前仍应对事实、责任人、日期和验证结果进行审核。
