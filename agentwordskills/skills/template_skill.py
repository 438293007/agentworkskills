import os
import json

def load_template(template_name):
    path = os.path.join(os.path.dirname(__file__), '..', 'templates', f'{template_name}.json')
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def apply_template(doc, template):
    """应用全局页面与正文模板"""
    # 1. 设置页边距 (上2.5cm, 下2.0cm, 左右3.0cm)
    for section in doc.Sections:
        section.PageSetup.TopMargin = 2.5 * 28.35
        section.PageSetup.BottomMargin = 2.0 * 28.35
        section.PageSetup.LeftMargin = 3.0 * 28.35
        section.PageSetup.RightMargin = 3.0 * 28.35

    # 2. 应用正文默认格式
    for para in doc.Paragraphs:
        # 【关键修复 1】：扩大识别范围，加入你截图中的“论文正文”样式
        if para.Style.NameLocal in ["正文", "Normal", "论文正文", "Body Text"]:
            font = para.Range.Font
            
            # 【关键修复 2】：双重绑定英文字体，强制将纯字母/数字转为 Times New Roman
            font.Name = template["body"]["font_en"]
            font.NameAscii = template["body"]["font_en"]
            
            # 绑定中文字体为宋体
            font.NameFarEast = template["body"]["font_cn"]
            
            # 统一设置字号 (小四 = 12磅)
            font.Size = template["body"]["size"]
            
            # 固定行距 23 磅
            para.Format.LineSpacingRule = 4 # wdLineSpaceExactly
            para.Format.LineSpacing = template["body"]["line_spacing_exact"]
            
            # 首行缩进 2 字符
            para.Format.CharacterUnitFirstLineIndent = 2
            
            # 强制清除段前段后间距
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
            para.Format.LineUnitBefore = 0
            para.Format.LineUnitAfter = 0