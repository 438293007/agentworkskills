def setup_page_numbers(doc):
    """精准配置罗马/阿拉伯数字页码，并强制格式化字体"""
    wdPageNumberStyleUppercaseRoman = 1 # I, II, III
    wdPageNumberStyleArabic = 0 # 1, 2, 3

    body_started = False
    front_restarted = False
    body_restarted = False

    for section in doc.Sections:
        section.Footers(1).LinkToPrevious = False
        section.Footers(3).LinkToPrevious = False

        first_text = ""
        for para in section.Range.Paragraphs:
            txt = para.Range.Text.strip()
            if txt:
                first_text = txt
                break

        if first_text.startswith("1") and "绪论" in first_text:
            body_started = True

        for footer_idx in [1, 3]:
            page_numbers = section.Footers(footer_idx).PageNumbers

            if not body_started:
                page_numbers.NumberStyle = wdPageNumberStyleUppercaseRoman
                if not front_restarted:
                    page_numbers.RestartNumberingAtSection = True
                    page_numbers.StartingNumber = 1
                else:
                    page_numbers.RestartNumberingAtSection = False
            else:
                page_numbers.NumberStyle = wdPageNumberStyleArabic
                if not body_restarted:
                    page_numbers.RestartNumberingAtSection = True
                    page_numbers.StartingNumber = 1
                else:
                    page_numbers.RestartNumberingAtSection = False

            try:
                page_numbers.Add(PageNumberAlignment=1, FirstPage=True)
            except Exception:
                pass
            
            # 【关键修复】：页码插入后，强制把整个页脚的字体刷成正确的格式
            footer_range = section.Footers(footer_idx).Range
            footer_range.Font.Name = "Times New Roman"
            footer_range.Font.NameAscii = "Times New Roman"
            footer_range.Font.NameFarEast = "宋体"
            footer_range.Font.Size = 10.5 # 五号

        if not body_started:
            front_restarted = True
        else:
            body_restarted = True