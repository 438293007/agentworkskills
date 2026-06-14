import re

def set_heading_style(paragraph, config, indent=0, default_before=0.5, default_after=0.5):
    """底层标题格式刷"""
    font = paragraph.Range.Font
    
    en_font = config.get("font_en", "Times New Roman")
    cn_font = config.get("font_cn", "黑体")
    
    font.Name = en_font
    font.NameAscii = en_font
    font.NameFarEast = cn_font
    
    font.Size = config.get("size", 14)
    font.Bold = config.get("bold", True)
    font.ColorIndex = 1 
    
    if "alignment" in config:
        paragraph.Format.Alignment = config["alignment"]
    elif indent == 0:
        paragraph.Format.Alignment = 1 
    else:
        paragraph.Format.Alignment = 0 
        
    paragraph.Format.CharacterUnitFirstLineIndent = indent
    paragraph.Format.FirstLineIndent = 0 
    paragraph.Format.SpaceBefore = 0
    paragraph.Format.SpaceAfter = 0
    paragraph.Format.LineSpacingRule = 0 
    paragraph.Format.LineUnitBefore = config.get("space_before", default_before)
    paragraph.Format.LineUnitAfter = config.get("space_after", default_after)

def process_headings(doc, template):
    """识别章节标题并精准应用多级格式，包含双字标题自动加空格逻辑"""
    pattern_h4 = re.compile(r'^\d+\.\d+\.\d+\.\d+[\s\.\、]*.*') 
    pattern_h3 = re.compile(r'^\d+\.\d+\.\d+[\s\.\、]*.*')      
    pattern_h2 = re.compile(r'^\d+\.\d+[\s\.\、]*.*')           
    # 放宽正则，兼容用户原稿中可能已经存在空格的情况
    pattern_h1 = re.compile(r'^(第[一二三四五六七八九十]+章|\d+)[\s\.\、]+.*|^致\s*谢$|^摘\s*要$|^ABSTRACT$|^结\s*论$|^参\s*考\s*文\s*献$|^附\s*录.*$')
    
    wdWithInTable = 12 
    
    for para in doc.Paragraphs:
        if para.Range.Information(wdWithInTable):
            continue
            
        raw_text = para.Range.Text
        text = raw_text.strip()
        if not text:
            continue
            
        # 提取纯净无空格的文本用于核心逻辑判断
        clean_text = text.replace(" ", "").replace("\u3000", "")

        if pattern_h4.match(text):
            para.Style = -5 # wdStyleHeading4
            set_heading_style(para, template.get("heading3", {}), indent=2, default_before=0.5, default_after=0.5) 
            
        elif pattern_h3.match(text):
            para.Style = -4 # wdStyleHeading3
            set_heading_style(para, template.get("heading3", {}), indent=2, default_before=0.5, default_after=0.5)
            
        elif pattern_h2.match(text):
            para.Style = -3 # wdStyleHeading2
            set_heading_style(para, template.get("heading2", {}), indent=2, default_before=0.5, default_after=0.5)
            
        elif pattern_h1.match(text):
            para.Style = -2 # wdStyleHeading1
            
            # =========================================================
            # 【核心修正】：给正文大标题“致谢”和“摘要”强行注入两个半角空格
            # =========================================================
            if clean_text in ["致谢", "摘要"]:
                # 极其安全的替换方式：绝不碰段落末尾的回车符，防止“黑洞吞噬Bug”重演！
                end_pos = para.Range.End - 1
                rng = doc.Range(para.Range.Start, end_pos)
                if clean_text == "致谢":
                    rng.Text = "致  谢"
                elif clean_text == "摘要":
                    rng.Text = "摘  要"
                    
            set_heading_style(para, template.get("heading1", {}), indent=0, default_before=0.8, default_after=0.5)