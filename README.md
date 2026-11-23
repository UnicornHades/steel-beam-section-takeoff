# steel-beam-section-takeoff

WorkBuddy Skill：从钢结构梁翻样图 DWG 批量提取梁编号与 BH 型钢截面，与附图构件清单 xlsx 核对，生成「梁编号-截面对照表」Excel。

## 处理链路

DWG →（AutoCAD accoreconsole）→ DXF →（文字提取）→ BH 截面 → 与 xlsx 清单核对 → Excel 对照表

## 目录结构

```
.
├── SKILL.md                        # Skill 定义：触发条件、工作流、踩坑清单
├── scripts/
│   ├── takeoff.py                  # 四阶段流水线（prepare/convert/extract/report）
│   └── env_check.py                # 环境自检
└── references/
    └── troubleshooting.md          # 排错指南
```

## 环境依赖

1. Python 3 + openpyxl
2. AutoCAD（提供 `accoreconsole.exe`，用于 DWG→DXF 无界面转换）
3. WorkBuddy 技能 `dwg-offline-extract` 的 `scripts/extract_text.py`（展开块内文字）

先跑自检：

```bash
python scripts/env_check.py
```

## 用法

```bash
python scripts/takeoff.py \
  --base "<图纸根目录>" \
  --work "<工作目录，纯ASCII路径>" \
  --out "<输出对照表.xlsx>" \
  --excluded "<已确认非标准H型钢的梁编号，逗号分隔>"
```

四阶段可单独跑：`--stage prepare|convert|extract|report`，用于断点续跑或调试。

完整参数与工作流见 [SKILL.md](SKILL.md)，问题排查见 [排错指南](references/troubleshooting.md)。

## 关键约束

- DWG→DXF 必须用 Python subprocess 直调 accoreconsole，不能经 git-bash（`/i`、`/s` 参数会被 MSYS 破坏）。
- 中文路径下转换不稳，脚本先把原图复制成 ASCII 序号文件名。
- 文字在块引用里，ezdxf 直接读不到，必须走 `extract_text.py`。
- 无 BH 主材行的梁可能是非标准截面，用 `--excluded` 标记，不当作提取失败。
