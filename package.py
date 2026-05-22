#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
尽调小助手 - 打包脚本
用于将skill打包成可分享的压缩包
"""

import os
import sys
import zipfile
import shutil
from pathlib import Path
from datetime import datetime


def create_package(version="1.0.0", output_dir=".."):
    """创建打包文件"""
    
    script_dir = Path(__file__).parent
    package_name = f"尽调小助手-v{version}"
    temp_dir = script_dir / "package_temp"
    zip_path = Path(output_dir) / f"{package_name}.zip"
    
    print("=" * 40)
    print("  尽调小助手 - 打包工具")
    print("=" * 40)
    print()
    
    # 清理旧的临时目录
    if temp_dir.exists():
        print("清理旧的临时目录...")
        shutil.rmtree(temp_dir)
    
    # 创建临时目录结构
    print("创建临时目录...")
    temp_dir.mkdir(exist_ok=True)
    (temp_dir / ".claude" / "skills").mkdir(parents=True, exist_ok=True)
    (temp_dir / "tools").mkdir(exist_ok=True)
    (temp_dir / "templates").mkdir(exist_ok=True)
    (temp_dir / "input").mkdir(exist_ok=True)
    (temp_dir / "output").mkdir(exist_ok=True)
    (temp_dir / "temp").mkdir(exist_ok=True)
    
    # 定义要复制的文件
    files_to_copy = [
        (".claude/skills/尽调小助手.md", ".claude/skills/"),
        ("tools/file_reader.py", "tools/"),
        ("tools/excel_writer.py", "tools/"),
        ("tools/report_writer.py", "tools/"),
        ("templates/报告模板.txt", "templates/"),
        ("templates/项目表示例.xlsx", "templates/"),
        ("CLAUDE.md", ""),
        ("README.md", ""),
        ("INSTALL.md", ""),
        ("requirements.txt", ""),
        ("skill.json", "")
    ]
    
    # 复制文件
    print("复制文件...")
    for src, dest in files_to_copy:
        src_path = script_dir / src
        dest_path = temp_dir / dest
        
        if src_path.exists():
            if dest_path.is_dir() or dest_path == "":
                shutil.copy2(src_path, temp_dir / dest)
            else:
                shutil.copy2(src_path, dest_path)
            print(f"  已复制: {src}")
        else:
            print(f"  警告: 文件不存在 - {src}")
    
    # 创建占位文件
    (temp_dir / "input" / "README.txt").write_text("# 在此目录放入待处理的尽调文件\n", encoding="utf-8")
    (temp_dir / "output" / "README.txt").write_text("# 生成的文件会保存在此目录\n", encoding="utf-8")
    (temp_dir / "temp" / "README.txt").write_text("# 临时文件目录\n", encoding="utf-8")
    
    print()
    print("创建压缩包...")
    
    # 删除旧的压缩包
    if zip_path.exists():
        zip_path.unlink()
    
    # 创建压缩包
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = Path(root) / file
                arcname = file_path.relative_to(temp_dir)
                zipf.write(file_path, arcname)
    
    # 清理临时目录
    print("清理临时目录...")
    shutil.rmtree(temp_dir)
    
    print()
    print("=" * 40)
    print("  打包完成!")
    print("=" * 40)
    print()
    print(f"压缩包位置: {zip_path.resolve()}")
    print(f"版本: v{version}")
    print()
    print("分享说明:")
    print(f"1. 将 {package_name}.zip 发送给其他用户")
    print("2. 用户解压后在Trae中打开目录")
    print("3. 运行 'pip install -r requirements.txt' 安装依赖")
    print("4. 输入 '/尽调小助手' 开始使用")
    print()


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="尽调小助手打包工具")
    parser.add_argument("--version", "-v", default="1.0.0", help="版本号")
    parser.add_argument("--output", "-o", default="..", help="输出目录")
    
    args = parser.parse_args()
    
    try:
        create_package(args.version, args.output)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)
