# PowerPoint 可编辑转换工具

这个小工具通过 Microsoft PowerPoint 的 COM 接口，把演示文稿中的幻灯片复制到一个全新的演示文稿中，再保存为 Open XML 格式。这样可以去掉源文件的文档级只读/保护状态，方便继续编辑。

它不修改源文件，也不是破解密码工具。需要密码、IRM 权限或其他访问授权的文件，仍然必须先由 PowerPoint 正常打开；如果源文件中的内容本来就是一张图片，转换后也仍然是一张图片。

## 支持的输入格式

`.ppt`、`.pptx`、`.pptm`、`.pps`、`.ppsx`、`.ppsm`、`.pot`、`.potx`、`.potm`（扩展名大小写不敏感）。

默认输出为 `.pptx`。也可以用 `--output-format pptm` 保存成宏启用容器，但本工具只复制幻灯片，不会复制 VBA 工程。

## 环境要求

- Windows
- 已安装 Microsoft PowerPoint（桌面版）
- Python 3.10 或更高版本
- `pywin32`

建议在虚拟环境中安装：

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
```

## 使用示例

转换一个文件：

```powershell
ppt-editable-convert "D:\课件\第1章.ppt"
```

转换一个文件夹中的所有支持格式：

```powershell
ppt-editable-convert "D:\课件" --output-dir "D:\课件\可编辑" --overwrite
```

递归扫描子文件夹：

```powershell
ppt-editable-convert "D:\课件" -r -o "D:\输出"
```

转换前只查看文件映射：

```powershell
ppt-editable-convert "D:\课件" -r --dry-run
```

需要观察 PowerPoint 窗口或手动处理提示时，可以加 `--visible`。默认不会覆盖已有输出文件；确认后才使用 `--overwrite`。

也可以不安装命令行入口，直接从项目根目录运行：

```powershell
python -m pip install -e .
python -m ppt_editable_converter "D:\课件" -r
```

## Python API

```python
from ppt_editable_converter import PowerPointConverter

with PowerPointConverter(visible=False) as converter:
    result = converter.convert(
        "D:\\课件\\第1章.pptx",
        "D:\\输出\\第1章-editable.pptx",
    )
    print(result.slide_count)
```

## Git

这个目录本身是独立 Git 项目。初始化后可按普通项目方式提交：

```powershell
git add .
git commit -m "整理为可复用的 PowerPoint 转换工具"
```

