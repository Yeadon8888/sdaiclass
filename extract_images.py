#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从Word文档中提取图片的脚本
需要安装: pip install python-docx pillow
"""

import os
import zipfile
from pathlib import Path

def extract_images_from_docx(docx_path, output_dir='images'):
    """
    从Word文档中提取所有图片
    
    Args:
        docx_path: Word文档路径
        output_dir: 输出目录
    """
    # 创建输出目录
    output_path = Path(output_dir)
    output_path.mkdir(exist_ok=True)
    
    # Word文档实际上是一个zip文件
    docx_file = zipfile.ZipFile(docx_path, 'r')
    
    # 获取所有媒体文件（图片通常在word/media/目录下）
    image_files = [f for f in docx_file.namelist() if f.startswith('word/media/')]
    
    print(f"找到 {len(image_files)} 张图片")
    
    # 提取图片
    extracted_files = []
    for i, image_path in enumerate(image_files, 1):
        # 获取文件扩展名
        ext = os.path.splitext(image_path)[1]
        if not ext:
            ext = '.png'  # 默认使用png
        
        # 生成输出文件名
        output_filename = f"image_{i:03d}{ext}"
        output_filepath = output_path / output_filename
        
        # 提取文件
        image_data = docx_file.read(image_path)
        with open(output_filepath, 'wb') as f:
            f.write(image_data)
        
        extracted_files.append({
            'original': image_path,
            'saved': str(output_filepath),
            'name': output_filename
        })
        
        print(f"[{i}/{len(image_files)}] 已提取: {output_filename}")
    
    docx_file.close()
    
    print(f"\n所有图片已提取到: {output_path.absolute()}")
    return extracted_files


if __name__ == '__main__':
    # 获取当前目录下的Word文档
    current_dir = Path(__file__).parent
    docx_file = current_dir / '伊登.docx'
    
    if not docx_file.exists():
        print(f"错误: 找不到文件 {docx_file}")
        exit(1)
    
    print(f"正在从 {docx_file.name} 提取图片...")
    print("-" * 50)
    
    try:
        images = extract_images_from_docx(docx_file, 'images')
        print("-" * 50)
        print(f"成功提取 {len(images)} 张图片!")
        print("\n提取的图片列表:")
        for img in images:
            print(f"  - {img['name']}")
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()

