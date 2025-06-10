"""
文档创建模块 - 用于创建空白的TXT和Word文档
"""

import os
from typing import Optional

# 检查python-docx库是否可用
try:
    from docx import Document
    docx_installed = True
except ImportError:
    docx_installed = False

def create_empty_txt(filename: str, output_path: Optional[str] = None) -> str:
    """
    在指定路径上创建一个空白的TXT文件。
    
    Args:
        filename: 要创建的文件名 (不需要包含.txt扩展名)
        output_path: 输出路径，如果为None则从环境变量获取
    
    Returns:
        包含操作结果的消息
    """
    # 确保文件名有.txt扩展名
    if not filename.lower().endswith('.txt'):
        filename += '.txt'
    
    # 从环境变量获取输出路径，如果未设置则使用默认桌面路径
    if output_path is None:
        output_path = os.environ.get('OFFICE_EDIT_PATH')
        if not output_path:
            output_path = os.path.join(os.path.expanduser('~'), '桌面')
    
    # 创建完整的文件路径
    file_path = os.path.join(output_path, filename)
    
    try:
        # 创建输出目录（如果不存在）
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # 创建空白文件
        with open(file_path, 'w', encoding='utf-8') as f:
            pass
        return f"成功在 {output_path} 创建了空白文件: {filename}"
    except Exception as e:
        return f"创建文件时出错: {str(e)}"

def create_word_document(filename: str, output_path: Optional[str] = None) -> str:
    """
    创建一个新的Word文档。
    
    Args:
        filename: 要创建的文件名 (不需要包含.docx扩展名)
        output_path: 输出路径，如果为None则从环境变量获取
    
    Returns:
        包含操作结果的消息
    """
    # 检查是否安装了必要的库
    if not docx_installed:
        return "错误: 无法创建Word文档，请先安装python-docx库: pip install python-docx"
    
    # 确保文件名有.docx扩展名
    if not filename.lower().endswith('.docx'):
        filename += '.docx'
    
    # 从环境变量获取输出路径，如果未设置则使用默认桌面路径
    if output_path is None:
        output_path = os.environ.get('OFFICE_EDIT_PATH')
        if not output_path:
            output_path = os.path.join(os.path.expanduser('~'), '桌面')
    
    # 创建完整的文件路径
    file_path = os.path.join(output_path, filename)
    
    try:
        # 创建输出目录（如果不存在）
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # 创建新的Word文档
        doc = Document()
        
        # 保存文档
        doc.save(file_path)
        
        return f"成功在 {output_path} 创建了Word文档: {filename}"
    except Exception as e:
        return f"创建Word文档时出错: {str(e)}"
