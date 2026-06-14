def insert_odd_section_breaks(doc):
    """清理历史残留并精准插入奇数页分节符（带防重叠与空段落过滤）"""
    
    # 1. 暴力清场：彻底抹除所有手动分页(^m)和分节符(^b)
    find_obj = doc.Content.Find
    find_obj.ClearFormatting()
    find_obj.Execute(FindText="^m", ReplaceWith="", Replace=2) # 2 = wdReplaceAll
    find_obj.Execute(FindText="^b", ReplaceWith="", Replace=2)
    
    WD_SECTION_BREAK_ODD_PAGE = 5  
    
    # 记录上一次插入位置，防止在同一处重复插入
    last_inserted_pos = -1

    # 倒序遍历段落
    for i in range(doc.Paragraphs.Count, 0, -1):
        para = doc.Paragraphs(i)
        
        # 安全清理文本，去除所有空白和控制字符
        clean_text = para.Range.Text.replace(" ", "").replace("\u3000", "").replace("\r", "").replace("\n", "").replace("\x07", "").strip()
        
        # 【核心拦截器】：如果这一行是空的（比如多敲的回车），直接跳过！绝对不给空行插分节符！
        if not clean_text:
            continue
            
        # 识别逻辑：1级大纲标题 或 文本内容仅为"目录"
        is_heading1 = (para.Format.OutlineLevel == 1)
        is_toc = (clean_text == "目录" or clean_text == "目录") 
        
        if is_heading1 or is_toc:
            current_pos = para.Range.Start
            
            # 距离防重叠保护（防止在同一个标题上插两次）
            if last_inserted_pos != -1 and abs(last_inserted_pos - current_pos) < 10:
                continue
                
            if current_pos > 0:
                try:
                    rng = para.Range
                    rng.Collapse(1) 
                    rng.InsertBreak(Type=WD_SECTION_BREAK_ODD_PAGE)
                    last_inserted_pos = current_pos 
                except Exception:
                    pass