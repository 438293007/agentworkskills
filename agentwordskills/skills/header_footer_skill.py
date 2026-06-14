def setup_headers_footers(doc):
    """动态设置前置、正文与后置部分的页眉，并修复标题带空格的Bug"""
    doc.PageSetup.OddAndEvenPagesHeaderFooter = True
    
    current_phase = "front" 
    
    for section in doc.Sections:
        section.Headers(1).LinkToPrevious = False
        section.Headers(3).LinkToPrevious = False 
        
        first_text = ""
        for para in section.Range.Paragraphs: 
            txt = para.Range.Text.strip()
            if txt:
                first_text = txt
                break
                
        # 清理获取到的第一行文本，用于精准的状态判定
        clean_first_text = first_text.replace(" ", "").replace("\u3000", "")
        
        if first_text.startswith("1") and "绪论" in first_text: 
            current_phase = "body"
        elif clean_first_text in ["参考文献", "致谢", "结论"] or first_text.startswith("附录"):
            current_phase = "back"
            
        odd_header = section.Headers(1).Range
        even_header = section.Headers(3).Range
        
        if current_phase == "front" or current_phase == "back":
            chapter_title = first_text if len(first_text) < 20 else ""
            
            # 【核心修复】：如果是目录、摘要等前置/后置大标题，强行抹除中间的空格
            clean_title = chapter_title.replace(" ", "").replace("\u3000", "")
            if clean_title in ["目录", "摘要", "致谢", "结论", "参考文献"]:
                chapter_title = clean_title # 页眉采用无空格版本
                
            odd_header.Text = chapter_title
            even_header.Text = chapter_title
            
        elif current_phase == "body":
            odd_header.Text = "2026届XXXX专业毕业设计(论文)"
            even_header.Text = "学生姓名:学生课题名称"
            
        for hdr in [odd_header, even_header]:
            if hdr.Text.strip() not in ["", "\r", "\n"]: 
                hdr.Font.Name = "宋体"                   
                hdr.Font.NameFarEast = "宋体"           
                hdr.Font.NameAscii = "Times New Roman"  
                hdr.Font.Size = 10.5 
                hdr.ParagraphFormat.Alignment = 1
                hdr.ParagraphFormat.SpaceAfter = 0
                
                try:
                    bottom_border = hdr.ParagraphFormat.Borders(-3)
                    bottom_border.LineStyle = 1  
                    bottom_border.LineWidth = 6  
                    bottom_border.ColorIndex = 1 
                except Exception:
                    pass