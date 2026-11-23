---
name: steel-beam-section-takeoff
description: 从钢结构梁翻样图 DWG 批量提取梁编号与梁截面（BH 型钢规格），并与附图构件清单 xlsx 核对，生成「梁编号-截面对照表」Excel。当用户需要从深化图纸/翻样图中提取钢梁截面、制作梁编号-截面对照清单、或处理「筒内地上XX层梁布置图及翻样图」这类图纸目录时使用。支持多层批量处理。
agent_created: true
---

# 钢结构梁截面对照表提取

## 用途

从一批钢结构「梁翻样图」DWG 中程序化提取每根梁的截面规格（材料表里的 BH 型钢规格，如 `BH500*400*20*35`），与附图构件清单 xlsx 的梁编号一一对应，输出对照表。适用于深化设计审核、工程量核对、模型校核等场景。

处理链路：**DWG → (accoreconsole) → DXF → (文字提取) → BH 截面 → 与 xlsx 清单核对 → Excel 对照表**。

## 何时使用

- 用户给出一批含 CAD/ 子目录的楼层图纸文件夹，要提取梁编号与截面
- 用户要「梁编号-截面对照表」「梁截面清单」之类的成果
- 图纸目录形如 `终审单-钢结构-91S-187(筒内地上四十八层梁布置图及翻样图)...`，内含 `CAD/*.dwg` 和 `*附图构件清单.xlsx`

## 环境依赖（先跑自检）

```bash
PYTHONUTF8=1 python "C:/Users/Tong/.workbuddy/skills/steel-beam-section-takeoff/scripts/env_check.py"
```

依赖三项：
1. **Python + openpyxl**（读清单、写对照表）
2. **AutoCAD accoreconsole.exe**（DWG→DXF 转换，自动扫描 `C:/Program Files/Autodesk/AutoCAD 20xx/`；找不到时用 `--accoreconsole` 指定）
3. **dwg-offline-extract 技能** 的 `scripts/extract_text.py`（展开块内文字；自动探测 `~/.workbuddy/skills/dwg-offline-extract/scripts/extract_text.py`）

缺哪个补哪个，不要跳过自检直接跑全流程。

## 工作流

### 1. 先摸清目录结构（只读）

列出图纸根目录，确认：
- 楼层文件夹数量与命名（决定 `--dir-glob` / `--floor-re`）
- 每个文件夹是否有 `CAD/*.dwg`、`*附图构件清单.xlsx`
- 清单里构件编号在第几列（决定 `--code-col`，默认第 4 列 / index 3）

**注意**：附图构件清单里通常只有梁编号、数量、单重，**没有截面**；截面必须从翻样图 DWG 的材料表（MATERIAL LIST，右上角白色表头表格）第一行 BH-xxx 规格列读取。

### 2. 跑四阶段流水线

一条命令跑全部（推荐）：

```bash
PYTHONUTF8=1 python "C:/Users/Tong/.workbuddy/skills/steel-beam-section-takeoff/scripts/takeoff.py" \
  --base "<图纸根目录>" \
  --work "<工作目录，纯ASCII路径>" \
  --out "<输出对照表.xlsx>" \
  --excluded "I-48GLL3-2,I-48GLL4-1"
```

四阶段可单独执行（`--stage prepare|convert|extract|report`），用于断点续跑或调试：

| 阶段 | 做什么 | 产物 |
|---|---|---|
| prepare | 复制所有 DWG 到 `--work/dwg`，改成 0001.dwg 序号名，生成 mapping.json | mapping.json |
| convert | 逐个调 accoreconsole 转 DXF | work/dxf/*.dxf、conv.log |
| extract | 调 extract_text.py 提取文字，正则匹配 BH 截面 | sections.json、work/txt/*.txt |
| report | 读构件清单 xlsx，按梁编号合并截面，写对照表 | --out 指定的 xlsx |

**关键路径要求**：`--work` 必须是纯 ASCII 路径。中文路径下 accoreconsole 转换不稳定。脚本已自动把中文原图复制成序号文件名，但工作目录本身仍是中文会出问题。

### 3. 关键参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `--base` | 无 | 图纸根目录（内含各层终审单文件夹）|
| `--work` | 必填 | 工作目录，纯 ASCII |
| `--out` | 无 | 对照表输出路径（report 阶段必需）|
| `--dir-glob` | `*梁布置图及翻样图*` | 楼层文件夹匹配模式 |
| `--floor-re` | `筒内地上(.+?)层梁` | 从文件夹名提取楼层 |
| `--code-re` | `(I-[^\s]+)` | 从 DWG 文件名提取梁编号 |
| `--manifest-glob` | `*附图构件清单.xlsx` | 清单文件匹配 |
| `--code-col` / `--qty-col` / `--weight-col` | 3 / 4 / 5 | 清单中构件编号/数量/单重列（0-based）|
| `--bh-re` | `^BH\d+[\*xX×]\d+` | 截面规格正则 |
| `--excluded` | 空 | 逗号分隔，已确认非标准 H 型钢、要标记为「已确认排除」的梁编号 |
| `--accoreconsole` | 自动探测 | accoreconsole.exe 路径 |
| `--extractor` | 自动探测 | extract_text.py 路径 |

若图纸命名规则不同，用参数覆盖正则；仍需调整时直接改脚本顶部的 `DEFAULT_*` 常量。

### 4. 核对与交付

跑完必须核对，不能直接进正式流程（用户铁律）：
1. **编号对齐**：xlsx 清单里的梁编号应与 sections.json 一一对应，无遗漏无多余（注意清单末尾「合计」行要排除）
2. **双路径互验**：抽 1~2 根，用 GUI（CAD快速看图）打开翻样图看材料表，与程序提取的截面比对
3. **重量反算**：按截面理论线密度（BH h*b*tw*tf → 面积 `2*b*tf+(h-2*tf)*tw` mm² × 7850 kg/m³）反推梁长，与图纸及清单单重是否量级吻合
4. 对照表先给用户审阅确认，再决定是否进正式流程

对照表含两个 sheet：主表（梁编号-截面对照）、「未提取到截面」（需人工复核的梁）。被 `--excluded` 标记的梁在主表截面列写「—」、备注「已确认排除」，不进待复核 sheet。

## 已踩过的坑（务必遵守）

1. **绝不用 git-bash 调 accoreconsole**：`/i`、`/s` 参数会被 MSYS 路径转换破坏，命令静默不执行（表现为日志里只有空命令提示符、连 QUIT 都不生效）。必须用 Python subprocess 或 PowerShell 直接调用。脚本已用 subprocess，勿改回 bash。
2. **accoreconsole 一次只开一个 DWG**：必须 `/i` 传 DWG，脚本里只写 `DXFOUT` + `QUIT`；试图在脚本里连续 OPEN 多个文件会失败。
3. **ezdxf 直接读不到文字**：图纸文字在块引用 INSERT 里，`doc.modelspace()` 直接遍历 TEXT/MTEXT 得到 0 条。必须调用 `dwg-offline-extract` 的 `extract_text.py` 展开块。它输出 TAB 分隔、内容在最后一列。
4. **非标准 H 型钢**：部分翻样图材料表第一行是连接板（PL10*90），没有 BH 主材行——不是提取失败，是截面非标准 H 型钢。这类梁经用户确认后用 `--excluded` 标记，不要反复补查。
5. **PDF 兜底不通**：同目录的 PDF 图纸是纯图形无文字层，无法提取文字，别浪费时间。
6. **GUI 会话是独立桌面**：computer use 的独立桌面是全机单例，删 GUI 任务会释放会话导致用户自己的看图窗口受影响。需要 GUI 复核时用「继续任务」而非删任务重建；用完如需让出名额再显式关闭。

## 详细排错

遇到转换失败、提取为空、编号对不上等问题，见 `references/troubleshooting.md`。
