# waterRPA

一个基于 `pyautogui + Excel(cmd.xls)` 的桌面自动化脚本项目，通过识图、点击、输入、循环与条件判断来执行 RPA 流程。

## 功能概览

- 支持识图点击：左键、双击、右键
- 支持坐标点击/移动鼠标
- 支持文本输入（剪贴板粘贴）
- 支持等待、滚轮滚动
- 支持流程控制：
  - `if_exists / if_not_exists / else / endif`
  - `while_exists / while_not_exists / endwhile`
  - `for_loop / endfor`
  - `stop_if_exists / stop`
- 支持运行热键：
  - `F10` 暂停
  - `F9` 继续
  - `F8` 停止

## 环境要求

建议 Python 3.7.x（说明书中提到 3.9 可能有兼容坑）。

安装依赖：

```bash
pip install pyperclip
pip install xlrd
pip install pyautogui==0.9.50
pip install opencv-python -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install pillow
# 如需热键功能
pip install keyboard
```

## 项目结构

- `waterRPA.py`：主程序
- `cmd.xls`：流程指令表
- `*.png`：识图模板（示例图片）
- `使用说明书1.docx`：原始使用说明

## 使用方法

1. 准备好 `cmd.xls` 指令表（默认读取第一个 sheet）。
2. 将需要识别的图片模板放在项目目录下。
3. 运行脚本：

```bash
python waterRPA.py
```

启动后会循环执行流程，直到触发停止条件或手动停止。

## cmd.xls 基本说明（与代码一致）

- 第 1 列：指令（可用数字、英文别名、中文别名）
- 第 2 列：参数1（通常是图片名、文本、坐标等）
- 第 3 列：参数2（通常是重试次数、检查次数等）

> 示例：
> - `left_click` / `1`
> - `double_click` / `2`
> - `right_click` / `3`
> - `input` / `4`
> - `wait` / `5`
> - `scroll` / `6`

## 重要提示

- 使用识图自动化时，请保持屏幕缩放、分辨率、窗口状态稳定。
- **进行网易云自动化点击时，请把网易云窗口放在屏幕左下角**，避免识图偏移导致点击失败。
- `pyautogui.FAILSAFE = True` 已启用，鼠标移动到屏幕角落可触发保护中断。

## 常见问题

- 识图失败：检查模板图是否清晰、界面缩放是否变化。
- 点击错位：优先使用坐标或设置偏移参数。
- 热键无效：确认已安装 `keyboard` 并以合适权限运行终端。