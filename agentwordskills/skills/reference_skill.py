def format_references(doc):
    """精准定位参考文献区域，挂载自动编号，并设置小四字体、半角空格与 1.5 倍行距"""
    
    # 1. 抓取并改造 Word 系统的自动编号库
    list_template = doc.Application.ListGalleries(2).ListTemplates(1)
    level = list_template.ListLevels(1)
    
    # 强制将编号格式定义为 [1], [2], [3]...
    level.NumberFormat = "[%1]"
    level.TrailingCharacter = 1  # 1 = 半角空格
    level.NumberStyle = 0        # 0 = 阿拉伯数字
    level.NumberPosition = 0     # 编号整体靠左对齐
    level.TextPosition = 21      # 悬挂缩进
    level.TabPosition = 21
    
    in_reference = False
    
    for para in doc.Paragraphs:
        text = para.Range.Text.strip()
        if not text:
            continue
            
        # 2. 状态机：精准侦测“参考文献”大标题
        if para.Format.OutlineLevel == 1:
            if "参考文献" in text:
                in_reference = True
                continue 
            else:
                in_reference = False 
        
        # 3. 在区域内给每一条文献施加魔法
        if in_reference:
            # 清零缩进
            para.Format.CharacterUnitFirstLineIndent = 0
            para.Format.FirstLineIndent = 0
            
            # 赋予自动编号属性
            para.Range.ListFormat.ApplyListTemplate(list_template)
            
            # 统一字体：西文/数字新罗马，中文宋体，小四(12磅)
            font = para.Range.Font
            font.Name = "Times New Roman"
            font.NameAscii = "Times New Roman"
            font.NameFarEast = "宋体"
            font.Size = 12  
            
            # 【核心修改】：彻底抛弃 23 磅固定值，改用 1.5 倍行距
            # 1 代表 wdLineSpace1pt5 (1.5倍行距)
            para.Format.LineSpacingRule = 1 
            
            para.Format.SpaceBefore = 0
            para.Format.SpaceAfter = 0
            para.Format.Alignment = 3       # 两端对齐