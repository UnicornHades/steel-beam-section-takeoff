# -*- coding: utf-8 -*-
"""环境自检：检查 DWG 梁截面提取所需的工具是否齐备。"""
import os
import sys
import importlib

OK = "[OK]"
NO = "[NO]"


def check(label, cond, detail=""):
    print("%s %-28s %s" % (OK if cond else NO, label, detail))
    return cond


def main():
    print("=== steel-beam-section-takeoff 环境自检 ===")
    all_ok = True

    # Python 依赖
    for mod in ("openpyxl",):
        try:
            m = importlib.import_module(mod)
            all_ok &= check("python: " + mod, True, getattr(m, "__version__", ""))
        except ImportError:
            all_ok &= check("python: " + mod, False, "未安装，需 pip install")

    # 把 takeoff.py 里的探测器复用过来
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from takeoff import find_accoreconsole, find_extractor

    exe = find_accoreconsole(None)
    all_ok &= check("AutoCAD accoreconsole", bool(exe), exe or "未找到，需 --accoreconsole 指定")

    ext = find_extractor(None)
    all_ok &= check("dwg-offline-extract 技能", bool(ext), ext or "未安装该技能或路径不对")

    print("\n结论：%s" % ("环境齐备，可运行" if all_ok else "有缺失项，按上面提示处理"))
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
