"""
MCP Server for Word Document Operations

This server provides tools to create, edit and manage Word documents.
It's implemented using the Model Context Protocol (MCP) Python SDK.
"""
from utils.createWordorTxt import create_empty_txt as create_txt_file, create_word_document as create_word_file
import os
import sys
import io
from mcp.server.fastmcp import FastMCP
from typing import Optional, List, Dict, Any, Union, Tuple
from utils.batch_paragraph_operations import (
    batch_add_formatted_paragraphs, batch_format_paragraphs,batch_set_paragraph_spacing)
from utils.media_table_operations import (
    batch_insert_images, batch_insert_tables, batch_edit_table_cells,insert_table_of_contents as insert_toc_func
)
from utils.saveMethod import (save_document_as_pdf as save_pdf,
                                save_document_as as save_as)
from utils.document_operations import (
    open_and_read_word_document as read_document,
    close_document as close_doc
)
from utils.edit_operations import (
    edit_paragraph_in_document as edit_paragraph_func,
    find_and_replace_text as find_replace_func,
    delete_paragraph as delete_paragraph_func
)
from utils.document_formatting import (
    add_header_footer as add_header_footer_func,
    set_page_layout as set_page_layout_func,
    merge_documents as merge_documents_func,
    apply_consistent_formatting as apply_consistent_formatting_func
)
from utils.advanced_formatting import (
    add_text_box, add_drop_cap, add_word_art, add_custom_bullets
)
from utils.style_management import (
    create_custom_style, apply_style, export_document_styles,
    import_document_styles, copy_style_between_documents
)


# 标记库是否已安装
docx_installed = True

# 尝试导入python-docx库，如果没有安装则标记为未安装但不退出
try:
    import docx
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.section import WD_ORIENTATION, WD_SECTION
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import docx.opc.constants
except ImportError:
    print("警告: 未检测到python-docx库，Word文档功能将不可用")
    print("请使用以下命令安装: pip install python-docx")
    docx_installed = False

# 尝试导入Pillow库，用于图片处理
pillow_installed = True
try:
    from PIL import Image
except ImportError:
    print("警告: 未检测到Pillow库，图片处理功能将受限")
    print("请使用以下命令安装: pip install Pillow")
    pillow_installed = False

# 创建一个MCP服务器，保持名称与配置文件一致
mcp = FastMCP("office editor")

# 添加回原始的TXT文件创建功能，确保基本功能可用
@mcp.tool()
def create_empty_txt(filename: str) -> str:
    """
    在指定路径上创建一个空白的TXT文件。
    
    Args:
        filename: 要创建的文件名 (不需要包含.txt扩展名)
    
    Returns:
        包含操作结果的消息
    """
    return create_txt_file(filename)


@mcp.tool()
def create_word_document(filename: str) -> str:
    """
    创建一个新的Word文档。
    
    Args:
        filename: 要创建的文件名 (不需要包含.docx扩展名)
    
    Returns:
        包含操作结果的消息
    """
    return create_word_file(filename)

@mcp.tool()
def open_and_read_word_document(file_path: str) -> str:
    """
    打开并读取Word文档的完整内容。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
    
    Returns:
        文档的完整内容
    """
    return read_document(file_path)



@mcp.tool()
def batch_add_formatted_text(
    file_path: str, 
    paragraphs_data: List[Dict[str, Any]]
) -> str:
    """
    批量添加格式化段落到Word文档。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        paragraphs_data: 段落数据列表，每个元素可包含：
            - text: 文本内容（必需）
            - is_heading: 是否作为标题添加（可选，默认False）
            - heading_level: 标题级别1-9（可选，默认1）
            - alignment: 对齐方式（可选，left/center/right/justify）
            - insert_position: 插入位置段落索引（可选，默认-1表示末尾）
            - font_name: 字体名称（可选）
            - font_size: 字体大小（可选）
            - bold: 是否加粗（可选）
            - italic: 是否斜体（可选）
            - underline: 是否下划线（可选）
            - font_color: 字体颜色十六进制格式（可选）
            - before_spacing: 段前间距磅值（可选）
            - after_spacing: 段后间距磅值（可选）
            - line_spacing: 行间距值（可选）
    
    Returns:
        操作结果信息
    """
    return batch_add_formatted_paragraphs(file_path, paragraphs_data)



@mcp.tool()
def batch_format_document_text(
    file_path: str,
    format_operations: List[Dict[str, Any]]
) -> str:
    """
    批量设置Word文档中多个段落的文本格式。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        format_operations: 格式化操作列表，每个元素包含：
            - paragraph_indices: 段落索引列表（必需，从0开始计数）
            - font_name: 字体名称（可选）
            - font_size: 字体大小点数（可选）
            - bold: 是否加粗（可选，默认False）
            - italic: 是否斜体（可选，默认False）
            - underline: 是否下划线（可选，默认False）
            - font_color: 字体颜色十六进制RGB格式如"#FF0000"（可选）
            - highlight_color: 突出显示颜色（可选，支持yellow/green/blue/red等）
    
    Returns:
        操作结果信息
    """
    return batch_format_paragraphs(file_path, format_operations)


@mcp.tool()
def batch_set_document_spacing(
    file_path: str,
    spacing_operations: List[Dict[str, Any]]
) -> str:
    """
    批量设置Word文档中多个段落的间距。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        spacing_operations: 间距设置操作列表，每个元素包含：
            - paragraph_indices: 段落索引列表（必需，从0开始计数）
            - before_spacing: 段前间距磅值（可选）
            - after_spacing: 段后间距磅值（可选）
            - line_spacing: 行间距值（可选）
            - line_spacing_rule: 行间距规则（可选，multiple/exact/atLeast，默认multiple）
    
    Returns:
        操作结果信息
    """
    return batch_set_paragraph_spacing(file_path, spacing_operations)


@mcp.tool()
def batch_insert_document_images(
    file_path: str,
    images_data: List[Dict[str, Any]]
) -> str:
    """
    批量在Word文档中插入图片。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        images_data: 图片数据列表，每个元素包含：
            - image_path: 图片文件路径（必需）
            - width: 图片宽度厘米（可选）
            - height: 图片高度厘米（可选）
            - after_paragraph: 插入位置段落索引（可选，默认-1表示文档末尾）
    
    Returns:
        操作结果信息
    """
    return batch_insert_images(file_path, images_data)


@mcp.tool()
def batch_insert_document_tables(
    file_path: str,
    tables_data: List[Dict[str, Any]]
) -> str:
    """
    批量在Word文档中插入表格。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        tables_data: 表格数据列表，每个元素包含：
            - rows: 表格行数（必需）
            - cols: 表格列数（必需）
            - data: 表格内容二维数组（可选）
            - after_paragraph: 插入位置段落索引（可选，默认-1表示文档末尾）
            - style: 表格样式（可选，默认"Table Grid"）
    
    Returns:
        操作结果信息
    """
    return batch_insert_tables(file_path, tables_data)


@mcp.tool()
def batch_edit_document_table_cells(
    file_path: str,
    edit_operations: List[Dict[str, Any]]
) -> str:
    """
    批量编辑Word文档中表格的单元格内容。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        edit_operations: 编辑操作列表，每个元素包含：
            - table_index: 表格索引（必需，从0开始计数）
            - cell_edits: 单元格编辑列表，每个元素包含：
                - row: 行索引（从0开始）
                - col: 列索引（从0开始）
                - text: 单元格内容
    
    Returns:
        操作结果信息
    """
    return batch_edit_table_cells(file_path, edit_operations)


@mcp.tool()
def save_document_as_pdf(file_path: str) -> str:
    """
    将Word文档保存为PDF格式。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
    
    Returns:
        操作结果信息
    """
    return save_pdf(file_path)


@mcp.tool()
def save_document_as(file_path: str, output_format: str = "docx", new_filename: str = None) -> str:
    """
    将Word文档保存为指定格式。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        output_format: 输出格式，可选值: "docx", "doc", "pdf", "txt", "html"
        new_filename: 新文件名(不含扩展名)，如果不提供则使用原文件名
    
    Returns:
        操作结果信息
    """
    from .utils.saveMethod import save_document_as as save_as
    return save_as(file_path, output_format, new_filename)


@mcp.tool()
def close_document(file_path: str, save_changes: bool = True) -> str:
    """
    关闭Word文档，可选是否保存更改。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        save_changes: 是否保存更改，默认为True
    
    Returns:
        操作结果信息
    """
    return close_doc(file_path, save_changes)


@mcp.tool()
def edit_paragraph_in_document(
    file_path: str,
    paragraph_index: int,
    new_text: str,
    save: bool = True,
    end_index: int = None,
    replacement_texts: list = None
) -> str:
    """
    编辑Word文档中指定段落或段落范围的文本内容。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        paragraph_index: 起始段落索引 (从0开始计数)
        new_text: 新的文本内容（用于单段落编辑）
        save: 是否保存更改，默认为True
        end_index: 结束段落索引（包含），如果为None则只编辑单个段落
        replacement_texts: 替换文本列表，用于批量编辑多个段落
    
    Returns:
        操作结果信息
    """
    return edit_paragraph_func(file_path, paragraph_index, new_text, save, end_index, replacement_texts)



@mcp.tool()
def find_and_replace_text(
    file_path: str,
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_whole_word: bool = False,
    save: bool = True
) -> str:
    """
    在Word文档中查找并替换文本。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        find_text: 要查找的文本
        replace_text: 替换为的文本
        match_case: 是否区分大小写，默认为False
        match_whole_word: 是否匹配整个单词，默认为False
        save: 是否保存更改，默认为True
    
    Returns:
        操作结果信息
    """
    return find_replace_func(file_path, find_text, replace_text, match_case, match_whole_word, save)


@mcp.tool()
def delete_paragraph(
    file_path: str,
    paragraph_index: Union[int, List[int]],
    save: bool = True
) -> str:
    """
    删除Word文档中指定的段落或多个段落。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        paragraph_index: 要删除的段落索引 (从0开始计数) 或索引列表
        save: 是否保存更改，默认为True
    
    Returns:
        操作结果信息
    """
    return delete_paragraph_func(file_path, paragraph_index, save)



@mcp.tool()
def insert_table_of_contents(
    file_path: str,
    title: str = "目录",
    levels: int = 3,
    after_paragraph: int = 0
) -> str:
    """
    在Word文档中插入目录。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        title: 目录标题
        levels: 目录级别数 (1-9)
        after_paragraph: 在指定段落后插入目录，默认为文档开头第一段后
    
    Returns:
        操作结果信息
    """
    return insert_toc_func(file_path, title, levels, after_paragraph)


@mcp.tool()
def add_header_footer(
    file_path: str,
    header_text: str = None,
    footer_text: str = None,
    page_numbers: bool = False
) -> str:
    """
    为Word文档添加页眉和页脚。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        header_text: 页眉文本（可选）
        footer_text: 页脚文本（可选）
        page_numbers: 是否在页脚添加页码
    
    Returns:
        操作结果信息
    """
    return add_header_footer_func(file_path, header_text, footer_text, page_numbers)


@mcp.tool()
def set_page_layout(
    file_path: str,
    orientation: str = None,
    page_width: float = None,
    page_height: float = None,
    left_margin: float = None,
    right_margin: float = None,
    top_margin: float = None,
    bottom_margin: float = None,
    section_indices: List[int] = None,
    apply_to_all: bool = False
) -> str:
    """
    设置Word文档的页面布局，支持单个或多个节的批量设置。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        orientation: 页面方向，可选值: "portrait"(纵向), "landscape"(横向)
        page_width: 页面宽度（厘米，自定义纸张尺寸时使用）
        page_height: 页面高度（厘米，自定义纸张尺寸时使用）
        left_margin: 左边距（厘米）
        right_margin: 右边距（厘米）
        top_margin: 上边距（厘米）
        bottom_margin: 下边距（厘米）
        section_indices: 要设置的节索引列表（从0开始），如果为None且apply_to_all为False则只设置第一节
        apply_to_all: 是否应用到所有节，默认为False
    
    Returns:
        操作结果信息
    """
    return set_page_layout_func(
        file_path, orientation, page_width, page_height,
        left_margin, right_margin, top_margin, bottom_margin,
        section_indices, apply_to_all
    )


@mcp.tool()
def merge_documents(
    main_file_path: str,
    files_to_merge: List[str]
) -> str:
    """
    合并多个Word文档。
    
    Args:
        main_file_path: 主文档的完整路径或相对于输出目录的路径（合并后的文档将保存为该文件）
        files_to_merge: 要合并的文档路径列表
    
    Returns:
        操作结果信息
    """
    return merge_documents_func(main_file_path, files_to_merge)


@mcp.tool()
def apply_consistent_style(
    file_path: str,
    content_type: str = "heading",
    level: int = 1,
    font_name: str = None,
    font_size: float = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    font_color: str = None,
    before_spacing: float = None,
    after_spacing: float = None,
    line_spacing: float = None,
    line_spacing_rule: str = "multiple"
) -> str:
    """
    将指定的格式应用到所有同级标题或正文，确保格式一致性。
    
    Args:
        file_path: Word文档的完整路径或相对于输出目录的路径
        content_type: 内容类型，可选值为"heading"（标题）, "title"（文档标题）, "normal"（正文）
        level: 标题级别（仅当content_type为"heading"时有效），1-9
        font_name: 字体名称
        font_size: 字体大小（磅值）
        bold: 是否加粗
        italic: 是否斜体
        underline: 是否下划线
        font_color: 字体颜色（十六进制格式，如"#FF0000"）
        before_spacing: 段前间距（磅值）
        after_spacing: 段后间距（磅值）
        line_spacing: 行间距值
        line_spacing_rule: 行间距规则（"multiple"/"exact"/"atLeast"）
    
    Returns:
        操作结果信息
    """
    return apply_consistent_formatting_func(
        file_path, content_type, level, font_name, font_size, 
        bold, italic, underline, font_color, 
        before_spacing, after_spacing, line_spacing, line_spacing_rule
    )


@mcp.tool()
def add_document_text_box(
    file_path: str,
    text: str,
    width: float = 10.0,
    height: float = 5.0,
    position: str = "center",
    border_style: str = "single",
    border_color: str = None,
    fill_color: str = None,
    font_name: str = None,
    font_size: float = None,
    font_bold: bool = False,
    font_italic: bool = False,
    font_color: str = None,
    paragraph_index: int = -1
) -> str:
    """
    在Word文档中添加文本框。
    
    Args:
        file_path: Word文档路径
        text: 文本框内容
        width: 文本框宽度（厘米）
        height: 文本框高度（厘米）
        position: 位置 (center/left/right)
        border_style: 边框样式 (single/double/none)
        border_color: 边框颜色 (十六进制格式，如"#FF0000")
        fill_color: 填充颜色 (十六进制格式，如"#FFFFFF")
        font_name: 字体名称
        font_size: 字体大小（磅值）
        font_bold: 是否加粗
        font_italic: 是否斜体
        font_color: 字体颜色 (十六进制格式，如"#000000")
        paragraph_index: 插入位置段落索引
    
    Returns:
        操作结果信息
    """
    return add_text_box(
        file_path, text, width, height, position, border_style,
        border_color, fill_color, font_name, font_size, 
        font_bold, font_italic, font_color, paragraph_index
    )

@mcp.tool()
def add_paragraph_drop_cap(
    file_path: str,
    paragraph_index: int,
    dropped_lines: int = 2,
    font_name: str = None,
    font_color: str = None
) -> str:
    """
    为Word文档中的段落添加首字下沉效果。
    
    Args:
        file_path: Word文档路径
        paragraph_index: 要添加首字下沉效果的段落索引
        dropped_lines: 下沉的行数（1-10）
        font_name: 字体名称
        font_color: 字体颜色（十六进制格式，如"#FF0000"）
    
    Returns:
        操作结果信息
    """
    return add_drop_cap(file_path, paragraph_index, dropped_lines, font_name, font_color)

@mcp.tool()
def add_document_word_art(
    file_path: str,
    text: str,
    style: int = 1,
    size: float = 36.0,
    fill_color: str = None,
    outline_color: str = None,
    paragraph_index: int = -1
) -> str:
    """
    在Word文档中添加艺术字。
    
    Args:
        file_path: Word文档路径
        text: 艺术字文本内容
        style: 艺术字样式编号（1-47）
        size: 艺术字大小（磅值）
        fill_color: 填充颜色（十六进制格式，如"#FF0000"）
        outline_color: 轮廓颜色（十六进制格式，如"#000000"）
        paragraph_index: 插入位置段落索引
    
    Returns:
        操作结果信息
    """
    return add_word_art(
        file_path, text, style, size, fill_color, outline_color, paragraph_index
    )

@mcp.tool()
def add_paragraph_bullets(
    file_path: str,
    paragraph_indices: List[int],
    bullet_style: str = "disc",
    custom_symbol: str = None,
    font_name: str = None,
    font_color: str = None
) -> str:
    """
    为Word文档中的段落添加自定义项目符号。
    
    Args:
        file_path: Word文档路径
        paragraph_indices: 要添加项目符号的段落索引列表
        bullet_style: 项目符号样式 (disc/circle/square/number/custom)
        custom_symbol: 自定义符号（仅当bullet_style为custom时有效）
        font_name: 符号字体名称
        font_color: 符号颜色（十六进制格式，如"#FF0000"）
    
    Returns:
        操作结果信息
    """
    return add_custom_bullets(
        file_path, paragraph_indices, bullet_style, custom_symbol, font_name, font_color
    )

@mcp.tool()
def create_document_style(
    file_path: str,
    style_name: str,
    style_type: str = "paragraph",
    based_on: str = None,
    font_name: str = None,
    font_size: float = None,
    font_bold: bool = None,
    font_italic: bool = None,
    font_underline: bool = None,
    font_color: str = None,
    alignment: str = None,
    line_spacing: float = None,
    space_before: float = None,
    space_after: float = None,
    first_line_indent: float = None,
    left_indent: float = None,
    right_indent: float = None
) -> str:
    """
    在Word文档中创建自定义样式。
    
    Args:
        file_path: Word文档路径
        style_name: 样式名称
        style_type: 样式类型 (paragraph/character/table/list)
        based_on: 基于哪个样式（可选）
        font_name: 字体名称
        font_size: 字体大小（磅值）
        font_bold: 是否加粗
        font_italic: 是否斜体
        font_underline: 是否下划线
        font_color: 字体颜色（十六进制格式，如"#FF0000"）
        alignment: 对齐方式 (left/right/center/justify)
        line_spacing: 行间距
        space_before: 段前间距（磅值）
        space_after: 段后间距（磅值）
        first_line_indent: 首行缩进（厘米）
        left_indent: 左缩进（厘米）
        right_indent: 右缩进（厘米）
    
    Returns:
        操作结果信息
    """
    return create_custom_style(
        file_path, style_name, style_type, based_on, font_name, font_size,
        font_bold, font_italic, font_underline, font_color, alignment,
        line_spacing, space_before, space_after, first_line_indent,
        left_indent, right_indent
    )

@mcp.tool()
def apply_document_style(
    file_path: str,
    paragraph_indices: List[int],
    style_name: str,
    create_if_not_exists: bool = False,
    style_properties: dict = None
) -> str:
    """
    将指定样式应用到文档中的段落。
    
    Args:
        file_path: Word文档路径
        paragraph_indices: 段落索引列表
        style_name: 要应用的样式名称
        create_if_not_exists: 如果样式不存在是否创建
        style_properties: 创建新样式的属性（仅在create_if_not_exists为True时有效）
    
    Returns:
        操作结果信息
    """
    return apply_style(
        file_path, paragraph_indices, style_name, create_if_not_exists, style_properties
    )

@mcp.tool()
def export_styles_to_file(
    file_path: str,
    output_path: str = None,
    style_names: List[str] = None
) -> str:
    """
    将文档中的样式导出到JSON文件。
    
    Args:
        file_path: Word文档路径
        output_path: 导出的JSON文件路径（不含扩展名），默认与文档同名
        style_names: 要导出的样式名称列表，None表示导出所有
    
    Returns:
        操作结果信息
    """
    return export_document_styles(file_path, output_path, style_names)

@mcp.tool()
def import_styles_from_file(
    file_path: str,
    style_file_path: str,
    style_names: List[str] = None,
    overwrite_existing: bool = False
) -> str:
    """
    从JSON文件导入样式到Word文档。
    
    Args:
        file_path: Word文档路径
        style_file_path: 样式JSON文件路径
        style_names: 要导入的样式名称列表，None表示导入所有
        overwrite_existing: 是否覆盖现有样式
    
    Returns:
        操作结果信息
    """
    return import_document_styles(
        file_path, style_file_path, style_names, overwrite_existing
    )

@mcp.tool()
def copy_styles_between_documents(
    source_file_path: str,
    target_file_path: str,
    style_names: List[str],
    overwrite_existing: bool = False
) -> str:
    """
    在文档之间复制样式。
    
    Args:
        source_file_path: 源文档路径
        target_file_path: 目标文档路径
        style_names: 要复制的样式名称列表
        overwrite_existing: 是否覆盖目标文档中的现有样式
    
    Returns:
        操作结果信息
    """
    return copy_style_between_documents(
        source_file_path, target_file_path, style_names, overwrite_existing
    )

if __name__ == "__main__":
    # 运行MCP服务器
    print("启动OFFICE EDITOR服务器...")
    mcp.run() 