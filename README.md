# Excel整理

本项目是一个本地运行的小工具，用来整理投票订单 Excel：

- 按 `款号` 分组
- 相同款号只保留一张图片
- 输出为 `图片 / 款号 / 名称 / 颜色 / 尺码数量 / 订货数`
- 自动跳过 `00000`、`抹零` 等非商品行
- 支持 `.xlsx` 和同版式 `销售订货明细.pdf`
- 数量为 `0` 的单元格留空不显示
- 最多成功生成 `50` 次，达到次数后不再允许生成

## 本地运行

```bash
python vote_excel_formatter.py
```

也可以命令行直接处理：

```bash
python vote_excel_formatter.py "D:\投票.xlsx"
python vote_excel_formatter.py "D:\销售订货明细.pdf"
```

## Windows 打包

在 Windows 电脑上双击：

```text
build_windows.bat
```

生成文件：

```text
dist\Excel整理.exe
```

如果打包失败，请查看或发送：

```text
build_windows.log
```

如果只是想先运行调试，不生成 exe，可以双击：

```text
run_windows_debug.bat
```

## Mac 打包

在 Mac 上双击或运行：

```bash
./build_mac.command
```

生成的是 Mac 可执行程序，不是 Windows exe。

## Mac 能不能直接打包 Windows exe？

不建议，也基本不能靠 PyInstaller 直接稳定完成。PyInstaller 通常要求在目标系统上打包：

- Windows exe：在 Windows 上打包
- Mac app：在 Mac 上打包
- Linux 程序：在 Linux 上打包

如果你只有 Mac，又想产出 Windows exe，推荐两种方式：

1. 使用 Windows 虚拟机打包
2. 使用 GitHub Actions 的 Windows runner 自动打包

本项目已经预留 Windows 打包脚本，后续也可以加 CI 自动生成 exe。

## PDF 支持

当前版本已支持这类“销售订货明细”电子 PDF。其他版式 PDF 需要先看真实样例：

- 电子 PDF：可以解析文字和表格
- 扫描/图片 PDF：需要 OCR

拿到稳定样例后，可以把 PDF 解析结果转成同一套数据结构，再复用当前 Excel 输出逻辑。
