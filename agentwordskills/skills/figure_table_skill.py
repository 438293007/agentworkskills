import re

def format_figures_and_tables(doc):
    """格式化图表：修复跨页问题，规范表格内容、表头重复与独立标题格式"""
    
    # 1. 强制“浮动型”转为“嵌入型”
    for i in range(doc.Shapes.Count, 0, -1):
        try:
            doc.Shapes(i).ConvertToInlineShape()
        except Exception:
            pass

    # 2. 嵌入型图片排版
    for inline_shape in doc.InlineShapes:
        if inline_shape.Type in [3, 4]: 
            para = inline_shape.Range.Paragraphs(1)
            para.Format.Alignment = 1
            para.Format.CharacterUnitFirstLineIndent = 0
            para.Format.FirstLineIndent = 0 
            para.Format.LineSpacingRule = 0 
            para.Format.SpaceBefore = 6 
            para.Format.SpaceAfter = 6
            para.Format.KeepWithNext = True 

    # 3. 处理表格本身：居中对齐，内容宋体五号，【跨页表头重复】
    for table in doc.Tables:
        table.Rows.Alignment = 1 
        
        t_font = table.Range.Font
        t_font.Name = "Times New Roman"
        t_font.NameAscii = "Times New Roman"
        t_font.NameFarEast = "宋体"
        t_font.Size = 10.5
        t_font.Bold = False
        
        for row in table.Rows:
            # 【核心修改】：判断是否为第一行，第一行开启跨页重复表头 (HeadingFormat = True)
            if row.Index == 1:
                row.HeadingFormat = True
            else:
                row.HeadingFormat = False 
                
            row.AllowBreakAcrossPages = False 
            row.Range.ParagraphFormat.KeepWithNext = False 
            row.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0
            row.Range.ParagraphFormat.FirstLineIndent = 0

    # 4. 图表标题精准分治格式化
    pattern_fig = re.compile(r'^图[\s]*\d+\.\d+.*')
    pattern_tbl = re.compile(r'^表[\s]*\d+\.\d+.*')
    
    for para in doc.Paragraphs:
        raw_text = para.Range.Text 
        clean_text = raw_text.strip()
        
        # 【情况 A】：表标题（黑体小四加粗，防跨页）
        if pattern_tbl.match(clean_text):
            para.Format.CharacterUnitFirstLineIndent = 0
            para.Format.FirstLineIndent = 0 
            para.Format.Alignment = 1 
            para.Format.KeepWithNext = True 
            para.Format.LineSpacingRule = 4 
            para.Format.LineSpacing = 23
            para.Format.SpaceBefore = 0 
            para.Format.SpaceAfter = 0  
            
            font = para.Range.Font
            font.Name = "Times New Roman"
            font.NameAscii = "Times New Roman"
            font.NameFarEast = "黑体"
            font.Size = 12
            font.Bold = True
            
        # 【情况 B】：图标题（宋体小四，仅数字加粗）
        elif pattern_fig.match(clean_text):
            para.Format.CharacterUnitFirstLineIndent = 0
            para.Format.FirstLineIndent = 0 
            para.Format.Alignment = 1 
            para.Format.KeepWithNext = False 
            para.Format.LineSpacingRule = 4 
            para.Format.LineSpacing = 23
            para.Format.SpaceBefore = 0 
            para.Format.SpaceAfter = 0  
            
            font = para.Range.Font
            font.Name = "Times New Roman"
            font.NameAscii = "Times New Roman"
            font.NameFarEast = "宋体"
            font.Size = 12
            font.Bold = False
            
            match = re.search(r'\d+\.\d+', raw_text)
            if match:
                start_pos = para.Range.Start + match.start()
                end_pos = para.Range.Start + match.end()
                num_range = doc.Range(start_pos, end_pos)
                num_range.Font.Bold = True
                num_range.Font.Name = "Times New Roman"
                num_range.Font.NameAscii = "Times New Roman"