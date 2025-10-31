# Excel-Linear-Fit-GUI
一键线性拟合 Excel 数据的小工具，**无需安装 Python，双击即用**！

## 下载使用（推荐）
1. 进入 [Releases 页面](https://github.com/Solitude669/Excel-Linear-Fit-GUI/releases)
2. 下载 `main.exe`
3. 双击运行 → **浏览** 选择 Excel → **开始拟合** → **保存图片**

## 功能亮点
- 自动读取 Excel 第 1、2 列作为 X、Y 数据
- 第一行若为文字，自动设为坐标轴名称
- 实时显示拟合方程与 R²
- 支持中文路径、中文轴名
- 可导出高清 PNG / PDF 图片

## Excel 格式要求
| 列 | 内容 |
|----|----|
| A 列 | X 数据（或第一行是轴名） |
| B 列 | Y 数据（或第一行是轴名） |
| 从第 2 行开始 | 纯数字即可 |

示例文件：`demo.xlsx`（ Releases 同页可下载）

## 截图
![screenshot](screenshot.png)

## 本地开发 / 二次打包
1. 克隆或下载源码
2. 安装依赖  
   ```bash
   pip install -r requirements.txt
   ```
3. 打包
   ```bash
   pyinstaller -F -w main.py
   ```
   生成文件在 dist\main.exe


