# AGENT.md

## 目录用途

这个目录只放桌面程序用到的图标资源：

- `app_icon.ico`
  Windows 程序和任务栏图标
- `app_icon.png`
  应用图标 PNG 版本
- `icons\*.png`
  窗口内按钮、菜单、功能入口图标

## 来源

这些图标由根目录的 `generate_icons.py` 生成或更新。

如果要统一图标风格，不要只手改单个 PNG，优先回到：

- `..\generate_icons.py`

统一生成后再检查：

- `assets\app_icon.ico`
- `assets\app_icon.png`
- `assets\icons\*.png`

## 修改建议

- 程序图标和界面功能图标要保持同一套风格
- 不要随意替换成系统默认图标
- 如果用户只说“任务栏图标”，优先检查 `app_icon.ico`
- 如果用户说“页面里的按钮图标”，优先检查 `icons\*.png`

## 打包关系

打包脚本会把整个 `assets` 目录带进最终程序：

- `..\build_application_template_exe.ps1`

所以资源文件名和目录结构不要随意改，否则需要同步改打包脚本和加载路径。

