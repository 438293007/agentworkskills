def check_document(doc, template_config):
    """
    反向读取文档属性，验证排版正确性
    """
    print("\n--- 开始自动化格式断言 ---")
    errors = []

    # 1. 验证正文格式
    for para in doc.Paragraphs:
        # 只检查非标题的正文部分
        if para.Style.NameLocal in ["正文", "Normal"]:
            font_cn = para.Range.Font.NameFarEast
            size = para.Range.Font.Size
            # 模板要求：宋体，小四(12磅)
            if font_cn != template_config["body"]["font_cn"]:
                errors.append(f"正文字体错误: 期望 {template_config['body']['font_cn']}，实际为 {font_cn}")
            if size != template_config["body"]["size"]:
                errors.append(f"正文字号错误: 期望 {template_config['body']['size']}，实际为 {size}")
            break # 抽查第一段正文即可

    # 2. 验证一级标题 (如 "1 绪论")
    for para in doc.Paragraphs:
        if para.Style.NameLocal in ["Heading 1", "标题 1"]:
            font_cn = para.Range.Font.NameFarEast
            size = para.Range.Font.Size
            # 模板要求：黑体，小二(18磅)
            if font_cn != template_config["heading1"]["font_cn"]:
                errors.append(f"一级标题字体错误: 实际为 {font_cn}")
            if size != template_config["heading1"]["size"]:
                errors.append(f"一级标题字号错误: 实际为 {size}")
            break

    # 3. 输出测试结果
    if not errors:
        print("[ok] 核心格式断言通过！排版属性与 JSON 模板完全一致。")
    else:
        print("[error] 发现排版偏差：")
        for err in errors:
            print(f"  - {err}")
    print("--------------------------\n")