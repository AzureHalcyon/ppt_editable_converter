# PowerPoint Editable Converter

简体中文 | [English](README_EN.md)

一个可直接复制使用的 Windows PowerPoint 只读文件转换脚本。

只需把 `convert_editable.py` 复制到需要处理的演示文稿所在文件夹，然后运行脚本。

程序会列出该文件夹中的 PowerPoint 文件。选择需要处理的文件后，转换结果会保存在源文件所在目录。

脚本通过 Microsoft PowerPoint 的 COM 自动化接口，以只读方式打开源文件，再使用 PowerPoint 自身的幻灯片导入能力，将幻灯片插入一个新建的演示文稿并保存。因此它会修改源文件属性。

重新生成可能无法完整保留母版、主题、动画、外部链接等内容，转换后请检查版式和播放效果。脚本不会修改源文件，也不能破解打开密码、IRM 或其他权限控制；源文件必须能够被本机安装的 Microsoft PowerPoint 正常打开。

## 支持格式

输入支持：

`.ppt`、`.pptx`、`.pptm`、`.pps`、`.ppsx`、`.ppsm`、`.pot`、`.potx`、`.potm`。

默认输出为新的 `.pptx` 文件，文件名会增加 `-editable` 后缀。例如：

```text
presentation.ppt  ->  presentation-editable.pptx
```

## 环境要求

- Windows
- Microsoft PowerPoint 桌面版
- Python 3.10 或更高版本
- `pywin32`

安装依赖：

```powershell
python -m pip install pywin32
```

## 使用方法

将脚本复制到演示文稿所在目录，然后运行：

```powershell
python convert_editable.py
```

程序会显示类似这样的选择菜单：

```text
发现以下 PowerPoint 文件：
  [1] presentation.ppt
  [2] report.pptx
  [a] 全部转换
  [q] 退出
请输入编号（可用逗号选择多个）：
```

可以输入单个编号、逗号分隔的多个编号、`a` 转换全部文件，或输入 `q` 退出。

如果输出文件已经存在，程序会询问是否覆盖。转换完成后，输出文件会与源文件保存在同一个目录中。名称已经以 `-editable` 结尾的文件不会再次出现在默认选择列表中。

## 可选命令

```powershell
# 递归扫描脚本目录下的子目录
python convert_editable.py --recursive

# 不显示选择菜单，直接转换所有找到的文件
python convert_editable.py --all

# 允许自动覆盖已有输出文件
python convert_editable.py --all --overwrite

# 显示 PowerPoint 窗口
python convert_editable.py --visible

# 只查看将要转换的文件和输出路径
python convert_editable.py --dry-run
```

也可以把文件或文件夹作为参数传入：

```powershell
python convert_editable.py "path\to\presentation.pptx"
```

## 限制

- 如果源文件中的内容本来就是图片，转换后仍然是图片，不会自动还原成文本或图形。
- 母版、主题、动画、外部链接或其他高级效果可能无法完整保留。
- 输出为 `.pptm` 时只复制幻灯片，不复制 VBA 宏工程。
- 需要密码、IRM 或其他访问权限的文件，必须先能被 PowerPoint 正常打开。
- 工具依赖桌面版 PowerPoint，不适用于只有在线版 Office 的环境。

## 项目结构

```text
convert_editable.py       可复制运行的完整脚本
tests/                    自动化测试
pyproject.toml            项目依赖和配置
```
