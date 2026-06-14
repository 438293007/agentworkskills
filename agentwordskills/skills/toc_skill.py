import re

def rebuild_toc(doc, levels=3, template=None):
    """重建目录域结构（隔离正文，生成纯净目录）"""
    for toc in doc.TablesOfContents:
        toc.Delete()
        
    toc_title_para = None
    for para in doc.Paragraphs:
        clean_text = para.Range.Text.replace(" ", "").replace("\u3000", "").replace("\r", "").replace("\n", "").replace("\x07", "").strip()
        if clean_text == "目录":
            toc_title_para = para
            break
            
    if not toc_title_para:
        for para in doc.Paragraphs:
            if para.Format.OutlineLevel == 1 and ("绪论" in para.Range.Text or "引言" in para.Range.Text or para.Range.Text.strip().startswith("1")):
                para.Range.InsertBefore("目  录\r")
                toc_title_para = para.Previous()
                break
                
    if not toc_title_para:
        doc.Range(0, 0).InsertBefore("目  录\r")
        toc_title_para = doc.Paragraphs(1)
        
    clean_text = toc_title_para.Range.Text.replace(" ", "").strip()
    if clean_text == "目录":
        end_pos = toc_title_para.Range.End - 1
        rng = doc.Range(toc_title_para.Range.Start, end_pos)
        rng.Text = "目  录"
    
    toc_title_para.Range.ListFormat.RemoveNumbers()
    toc_title_para.Style = doc.Styles(-1) 
    
    font = toc_title_para.Range.Font
    font.NameFarEast = "黑体"
    font.NameAscii = "Times New Roman"
    font.Name = "黑体"
    font.Size = 18       
    font.Bold = True
    font.ColorIndex = 1  
    
    toc_title_para.Format.Alignment = 1 
    toc_title_para.Format.OutlineLevel = 10 
    toc_title_para.Format.SpaceBefore = 0
    toc_title_para.Format.SpaceAfter = 12 
    toc_title_para.Format.PageBreakBefore = False 
    
    toc_title_para.Range.InsertParagraphAfter()
    toc_content_para = toc_title_para.Next()
    
    toc_content_para.Style = doc.Styles(-1) 
    toc_content_para.Format.OutlineLevel = 10 
    toc_content_para.Range.ListFormat.RemoveNumbers()
    toc_content_para.Format.SpaceBefore = 0
    toc_content_para.Format.SpaceAfter = 0
    toc_content_para.Format.CharacterUnitFirstLineIndent = 0
    toc_content_para.Format.FirstLineIndent = 0
    
    toc_rng = toc_content_para.Range
    toc_rng.Collapse(1) 
    
    doc.TablesOfContents.Add(
        Range=toc_rng, 
        RightAlignPageNumbers=True, 
        UseHeadingStyles=True, 
        UpperHeadingLevel=1, 
        LowerHeadingLevel=levels,
        IncludePageNumbers=True,
        UseHyperlinks=True 
    )
    
    try:
        toc1 = doc.Styles(-19)
        toc1.Font.NameFarEast = "黑体"
        toc1.Font.NameAscii = "Times New Roman"
        toc1.Font.Size = 14  
        toc1.Font.Bold = True
        
        toc2 = doc.Styles(-20)
        toc2.Font.NameFarEast = "宋体"
        toc2.Font.NameAscii = "Times New Roman"
        toc2.Font.Size = 12  
        toc2.Font.Bold = False
        # 提前在底层抹除二级目录的缩进
        toc2.ParagraphFormat.CharacterUnitLeftIndent = 0
        toc2.ParagraphFormat.LeftIndent = 0
        toc2.ParagraphFormat.FirstLineIndent = 0
        
        toc3 = doc.Styles(-21)
        toc3.Font.NameFarEast = "宋体"
        toc3.Font.NameAscii = "Times New Roman"
        toc3.Font.Size = 12
        toc3.Font.Bold = False
        toc3.ParagraphFormat.CharacterUnitLeftIndent = 0
        toc3.ParagraphFormat.LeftIndent = 0
        toc3.ParagraphFormat.FirstLineIndent = 0
    except Exception:
        pass
        
    if doc.TablesOfContents.Count > 0:
        doc.TablesOfContents(1).Update()

def fix_toc_format(doc):
    """【终极修复】：暴力清零二级标题缩进，并强制对齐截图中的 23 磅固定行距"""
    if doc.TablesOfContents.Count == 0:
        return
        
    toc = doc.TablesOfContents(1)
    for para in toc.Range.Paragraphs:
        txt = para.Range.Text
        clean_txt = txt.strip()
        if not clean_txt:
            continue
        
        level = 1
        match = re.match(r'^(\d+(\.\d+)*)', clean_txt)
        if match:
            level = match.group(1).count('.') + 1
            
        p_font = para.Range.Font
        p_format = para.Range.ParagraphFormat
        
        # ==========================================
        # 【核心修正】：完美复刻你发来的段落设置截图
        # ==========================================
        # 1. 左侧缩进清零（0 字符），无特殊悬挂格式
        p_format.CharacterUnitLeftIndent = 0
        p_format.LeftIndent = 0
        p_format.CharacterUnitFirstLineIndent = 0
        p_format.FirstLineIndent = 0
        
        # 2. 设置行距为固定值 23 磅 (4 = wdLineSpaceExactly)
        p_format.LineSpacingRule = 4 
        p_format.LineSpacing = 23
        p_format.SpaceBefore = 0
        p_format.SpaceAfter = 0
        # ==========================================
        
        if level == 1:
            p_font.NameFarEast = "黑体"
            p_font.NameAscii = "Times New Roman"
            p_font.Size = 14
            p_font.Bold = True
        else:
            p_font.NameFarEast = "宋体"
            p_font.NameAscii = "Times New Roman"
            p_font.Size = 12
            p_font.Bold = False