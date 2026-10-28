#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""版本链管理：定位基线、算出下一个版本文件名、基于基线开出新版本工作副本。

命名规则：<名称>_V<N>.<扩展名>，原文件视为 V0（永不改动）。
脚本只读文件系统做判断，不修改除 snap 目标以外的东西。

用法:
  python version_file.py chain --file <path>          # 看版本链与下一个版本名
  python version_file.py snap  --file <path> [--label "说明"]   # 以该文件为蓝本开新版本
  python version_file.py list  --file <path>          # 列出同链所有版本（含原版）

所有命令向 stdout 输出 JSON。
"""

import argparse
import json
import os
import re
import shutil
import sys
from datetime import datetime

try:  # Windows 控制台默认 GBK，强制 UTF-8 避免中文结果乱码
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # pragma: no cover
    pass

VER_RE = re.compile(r"^(?P<stem>.+)_V(?P<num>\d+)$")


def emit(obj, json_out=""):
    if json_out:
        with open(json_out, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
    print(json.dumps(obj, ensure_ascii=False, indent=2))


def split_version(path):
    """返回 (目录, 名称干, 扩展名, 版本号)。原文件版本号为 0。"""
    abs_path = os.path.abspath(path)
    folder, name = os.path.split(abs_path)
    stem, ext = os.path.splitext(name)
    m = VER_RE.match(stem)
    if m:
        return folder, m.group("stem"), ext, int(m.group("num"))
    return folder, stem, ext, 0


def scan_chain(folder, stem, ext):
    """列出同链已存在的版本号（不含原版 0）。"""
    nums = []
    try:
        entries = os.listdir(folder)
    except OSError:
        return nums
    for entry in entries:
        s, e = os.path.splitext(entry)
        m = VER_RE.match(s)
        if m and m.group("stem") == stem and e == ext:
            nums.append(int(m.group("num")))
    return sorted(set(nums))


def next_free(folder, stem, ext, start):
    """从 start 开始找第一个未被占用的版本号。"""
    n = start
    while os.path.exists(os.path.join(folder, f"{stem}_V{n}{ext}")):
        n += 1
    return n


def build_chain(path):
    folder, stem, ext, num = split_version(path)
    existing = scan_chain(folder, stem, ext)
    latest = max(existing) if existing else 0
    base_num = num if num else latest
    nxt = next_free(folder, stem, ext, base_num + 1)
    base_path = os.path.join(folder, f"{stem}{ext}")
    return {
        "folder": folder,
        "stem": stem,
        "ext": ext,
        "base_path": base_path,
        "base_number": base_num,
        "existing_versions": existing,
        "latest_number": latest,
        "next_number": nxt,
        "next_path": os.path.join(folder, f"{stem}_V{nxt}{ext}"),
        "input_path": os.path.abspath(path),
    }


def cmd_chain(args):
    emit(build_chain(args.file), args.json_out)
    return 0


def cmd_list(args):
    info = build_chain(args.file)
    rows = [{
        "version": "V0",
        "number": 0,
        "path": info["base_path"],
        "exists": os.path.exists(info["base_path"]),
        "is_original": True,
    }]
    for n in info["existing_versions"]:
        p = os.path.join(info["folder"], f"{info['stem']}_V{n}{info['ext']}")
        rows.append({
            "version": f"V{n}",
            "number": n,
            "path": p,
            "exists": os.path.exists(p),
            "is_original": False,
        })
    emit({"chain": rows}, args.json_out)
    return 0


def cmd_snap(args):
    info = build_chain(args.file)
    src = os.path.abspath(args.file)
    if not os.path.isfile(src):
        emit({
            "ok": False,
            "error": "源文件不存在",
            "source": src,
        }, args.json_out)
        return 1
    target = info["next_path"]
    shutil.copy2(src, target)
    emit({
        "ok": True,
        "source": src,
        "source_version": f"V{info['base_number']}",
        "target": target,
        "target_version": f"V{info['next_number']}",
        "label": args.label or "",
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "note": "原版与所有旧版本保持不变；后续编辑请直接改这个新版本文件。",
    }, args.json_out)
    return 0


def main():
    ap = argparse.ArgumentParser(description="版本链管理")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_chain = sub.add_parser("chain", help="查看版本链与下一个版本名")
    p_chain.add_argument("--file", required=True)
    p_chain.add_argument("--json-out", default="")
    p_chain.set_defaults(func=cmd_chain)

    p_list = sub.add_parser("list", help="列出同链所有版本")
    p_list.add_argument("--file", required=True)
    p_list.add_argument("--json-out", default="")
    p_list.set_defaults(func=cmd_list)

    p_snap = sub.add_parser("snap", help="以该文件为蓝本开出下一个版本")
    p_snap.add_argument("--file", required=True)
    p_snap.add_argument("--label", default="")
    p_snap.add_argument("--json-out", default="")
    p_snap.set_defaults(func=cmd_snap)

    args = ap.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
