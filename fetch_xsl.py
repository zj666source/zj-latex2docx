# -*- coding: utf-8 -*-
"""从本机 Microsoft Office 安装目录查找并复制 MML2OMML.XSL 到当前目录。

MML2OMML.XSL 是微软 Office 自带组件（MathML -> OMML 官方转换样式表），
版权归微软所有，不随本仓库分发。运行前请先执行本脚本。
"""
import os
import shutil
import sys

CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office15\MML2OMML.XSL",
]


def find_xsl():
    for p in CANDIDATES:
        if os.path.isfile(p):
            return p
    # 兜底：全盘搜索常见盘符的 Office 目录（较慢，仅当上方没找到时）
    for drive in ("C:", "D:"):
        for root in (os.path.join(drive, os.sep, "Program Files"),
                     os.path.join(drive, os.sep, "Program Files (x86)")):
            for base, dirs, files in os.walk(root):
                if "Office16" in base or "Office15" in base:
                    if "MML2OMML.XSL" in files:
                        return os.path.join(base, "MML2OMML.XSL")
                dirs[:] = [d for d in dirs if d.lower().startswith("microsoft")]
    return None


def main():
    src = find_xsl()
    if not src:
        print("[错误] 未找到 MML2OMML.XSL，请确认已安装 Microsoft Office。")
        print("      也可以手动把该文件复制到当前目录后重试。")
        sys.exit(1)
    dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MML2OMML.XSL")
    shutil.copyfile(src, dst)
    print(f"[完成] 已复制：{src}")
    print(f"      -> {dst}")


if __name__ == "__main__":
    main()
