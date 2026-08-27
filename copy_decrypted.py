# coding=utf-8
"""把企业文件保护透明解密后的文件内容复制到临时构建传输文件。"""
import shutil
import sys
from pathlib import Path


def main() -> int:
    if len(sys.argv) != 3:
        return 2
    source = Path(sys.argv[1])
    destination = Path(sys.argv[2])
    if not source.is_file():
        return 3
    shutil.copyfile(source, destination)
    return 0 if destination.stat().st_size else 4


if __name__ == "__main__":
    raise SystemExit(main())
