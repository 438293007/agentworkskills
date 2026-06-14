def process_cover(doc):
    """处理封面、原创性声明和版权授权书"""
    # 封面标题：1号宋体居中，副标题2号黑体加粗
    for para in doc.Paragraphs:
        text = para.Range.Text.strip()
        if "毕业设计(论文)" in text:
            para.Range.Font.Size = 26 # 1号
            para.Range.Font.NameFarEast = "宋体"
            # 修改这里：使用 ParagraphFormat
            para.ParagraphFormat.Alignment = 1 
        elif "学位论文原创性声明" in text or "学位论文版权使用授权书" in text:
            para.Range.Font.Size = 22 # 2号
            para.Range.Font.NameFarEast = "宋体"
            para.Range.Font.Bold = True
            # 修改这里：使用 ParagraphFormat
            para.ParagraphFormat.Alignment = 1