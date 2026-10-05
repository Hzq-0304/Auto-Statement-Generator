# 开发环境与安装包

## Codegraph

本地 Codegraph 0.9.9 已于 2026-10-05 初始化。首次索引18个文件、272个节点、623条关系。`.codegraph/` 是机器本地索引，已加入 Git 忽略。

在项目目录执行：

```powershell
codegraph init -i
codegraph sync
codegraph status
codegraph query export_statement
```

MCP 可通过 `codegraph_status`、`codegraph_explore` 查询本项目；其他项目环境需要将 `projectPath` 指向当前项目根目录。索引用于代码定位，不能替代测试。

## 构建安装包

安装包版本从 `app/__init__.py` 读取，当前为1.1.0。目标系统为Windows 10/11 x64。

安装到当前用户目录，无需管理员权限，包含完整Python运行时、Tk界面、Excel处理依赖、脱敏模板和帮助文档。安装向导为简体中文，提供开始菜单、可选桌面快捷方式和卸载入口。

准备Python环境并安装 `requirements-build.txt`，另从 [Inno Setup 官方发行页](https://github.com/jrsoftware/issrc/releases) 安装编译器。本机采用7.1.0。

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_installer.ps1 -Iscc "C:\路径\Inno Setup\ISCC.exe"
```

已有最新目录版程序时可加 `-SkipFreeze`。脚本生成：

- `release/AutoStatementGenerator-1.1.0-Setup-x64.exe`
- 同名 `.sha256` 校验文件。

构建输入仅为 `dist/自动对账工具`，不包含项目原始表格、输出账单、工作进度、Codegraph索引、编译器或开发环境。卸载程序只移除安装文件，用户另存的业务文件仍由用户保管。

安装参数与无管理员权限安装依据：[官方安装参数](https://jrsoftware.org/ishelp/topic_setupcmdline.htm)、[PrivilegesRequired](https://jrsoftware.org/ishelp/topic_setup_privilegesrequired.htm)。

## 安装包验收

在项目 `build/installer-smoke` 临时目录内执行静默安装，验证安装文件与目录版一致，运行窗口和样例导出，然后调用该目录内的卸载程序。此测试不会卸载用户已安装的其他版本。

安装日志及样例输出保留在 `analysis/` 与 `outputs/`，均不上传仓库。安装包位于本地 `release/` 并作为 GitHub Release 附件发布，源代码仓库提交可重复构建脚本。

2026-10-05 已实际通过：安装成功、全部安装文件SHA-256一致、安装后的GUI启动、99条样例导出及8,969.50已确认小计核验、卸载成功、卸载后保留模拟用户数据。验收脚本为 `tools/test_installer.ps1`。

## 1.1.0 导入编辑与显示适配

`parser.import_excel(..., preview=True)`保留可修正坏记录；`review.py`实时验证并独立提交快照，结构问题不能因改单行而消失。严格CLI入口兼容旧行为。进度格式升级为v2以保存原值和修改历史，仍能读取v1。

打包必须带 `--manifest packaging/app.manifest`，在进程启动时声明PerMonitorV2；源码运行在创建Tk窗口前设置DPI感知，中文字体及表格行高同步缩放。依据[微软DPI文档](https://learn.microsoft.com/en-us/windows/win32/hidpi/setting-the-default-dpi-awareness-for-a-process)。不修改用户系统缩放或字体设置。

```powershell
.venv/Scripts/python -m unittest discover -s tests -v
.venv/Scripts/python tests/gui_review_smoke.py
.venv/Scripts/python tests/gui_smoke.py
powershell -ExecutionPolicy Bypass -File tools/test_installer.ps1
```

GUI编辑测试使用临时合成Excel；样例GUI测试需本地业务样例。安装验收目录包含版本和随机后缀；若发现当前用户已安装正式版本会停止，避免覆盖用户安装。
