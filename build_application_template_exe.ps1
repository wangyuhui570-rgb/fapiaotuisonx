$ErrorActionPreference = "Stop"
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
chcp 65001 > $null

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$PyInstallerExe = Join-Path $ProjectRoot ".venv\Scripts\pyinstaller.exe"
$IconScript = Join-Path $ProjectRoot "generate_icons.py"
$AppIcon = Join-Path $ProjectRoot "assets\app_icon.ico"
$GuideFile = Join-Path $ProjectRoot "APPLICATION_TEMPLATE_GUIDE.txt"

if (-not (Test-Path $PythonExe)) {
    throw "Missing Python: $PythonExe"
}

if (-not (Test-Path $PyInstallerExe)) {
    throw "Missing PyInstaller: $PyInstallerExe"
}

if (-not (Test-Path $GuideFile)) {
    throw "Missing guide file: $GuideFile"
}

Set-Location $ProjectRoot
& $PythonExe $IconScript

if (-not (Test-Path $AppIcon)) {
    throw "Missing app icon: $AppIcon"
}

& $PyInstallerExe `
  --noconfirm `
  --clean `
  --windowed `
  --name "InvoiceRequestTemplateTool" `
  --icon "$AppIcon" `
  --hidden-import PySide6.QtCore `
  --hidden-import PySide6.QtGui `
  --hidden-import PySide6.QtWidgets `
  --hidden-import wecom_aibot_sdk `
  --hidden-import httpx `
  --hidden-import httpcore `
  --hidden-import anyio `
  --hidden-import websockets `
  --add-data "assets;assets" `
  --add-data "APPLICATION_TEMPLATE_GUIDE.txt;." `
  application_template_desktop.py

$DistDir = Join-Path $ProjectRoot "dist\InvoiceRequestTemplateTool"
Copy-Item -Force $GuideFile (Join-Path $DistDir "APPLICATION_TEMPLATE_GUIDE.txt")

@'
import os
import shutil
import zipfile
from pathlib import Path

project_root = Path(r"__PROJECT_ROOT__")
dist_dir = project_root / "dist" / "InvoiceRequestTemplateTool"
release_dir = project_root / "release" / "\u53d1\u7968\u7533\u8bf7\u6a21\u7248\u4ea7\u51fa\u5de5\u5177_\u7eff\u8272\u7248"
zip_path = project_root / "release" / "\u53d1\u7968\u7533\u8bf7\u6a21\u7248\u4ea7\u51fa\u5de5\u5177_\u7eff\u8272\u7248.zip"

if release_dir.exists():
    shutil.rmtree(release_dir)
release_dir.mkdir(parents=True, exist_ok=True)

for item in dist_dir.iterdir():
    target = release_dir / item.name
    if item.is_dir():
        shutil.copytree(item, target)
    else:
        shutil.copy2(item, target)

exe_path = release_dir / "InvoiceRequestTemplateTool.exe"
renamed_exe_path = release_dir / "\u53d1\u7968\u7533\u8bf7\u6a21\u7248\u4ea7\u51fa\u5de5\u5177.exe"
if renamed_exe_path.exists():
    renamed_exe_path.unlink()
exe_path.rename(renamed_exe_path)

guide_src = project_root / "APPLICATION_TEMPLATE_GUIDE.txt"
shutil.copy2(guide_src, release_dir / "\u4f7f\u7528\u8bf4\u660e.txt")

(release_dir / "\u5148\u770b\u8fd9\u91cc.txt").write_text(
    "先看这里\\n\\n"
    "1. 双击“发票申请模版产出工具.exe”。\\n"
    "2. 选择存放 CSV 的文件夹。\\n"
    "3. 确认店铺名、所属店铺、货物名称、申请人。\\n"
    "4. 点击“开始生成申请模板”。\\n"
    "5. 在结果框中直接复制文本。\\n\\n"
    "提示：\\n"
    "- 程序会按固定申请模板自动生成文字。\\n"
    "- 不再输出 txt 文件，结果直接显示在窗口里。\\n",
    encoding="utf-8",
)

sample_csv_dir = release_dir / "\u793a\u4f8b_CSV\u6587\u4ef6\u5939"
sample_csv_dir.mkdir(exist_ok=True)

(sample_csv_dir / "\u793a\u4f8b\u6570\u636e.csv").write_text(
    "公司名称,税号,金额\\n"
    "示例科技有限公司,91320000123456789X,1280.50\\n"
    "星河商贸有限公司,91320000987654321Y,560.00\\n",
    encoding="utf-8",
)

(release_dir / "\u793a\u4f8b_\u7533\u8bf7\u6587\u5b57\u6a21\u677f.txt").write_text(
    "申请内容：\\n\\n"
    "申请单位：{{公司名称}}\\n"
    "统一社会信用代码：{{税号}}\\n"
    "申请金额：{{金额}}\\n"
    "申请日期：{{当前日期}}\\n\\n"
    "请按实际业务需要调整文字内容。\\n",
    encoding="utf-8",
)

if zip_path.exists():
    zip_path.unlink()
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
    for path in release_dir.rglob("*"):
        zf.write(path, path.relative_to(release_dir.parent))

print()
print(f"Build complete.")
print(f"Release dir: {release_dir}")
print(f"Zip path: {zip_path}")
'@.Replace("__PROJECT_ROOT__", $ProjectRoot.Replace("\", "\\")) | & $PythonExe -
