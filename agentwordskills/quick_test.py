# -*- coding: utf-8 -*-
"""
快速测试脚本 - 用于日常开发调试
仅验证核心功能，运行速度快
"""

import os
import sys
import json
import win32com.client as win32

sys.path.insert(0, os.path.dirname(__file__))

from skills import template_skill, heading_skill, section_break_skill


def quick_test():
    """快速测试核心功能"""
    print("="*50)
    print("快速测试 - 核心功能验证")
    print("="*50)
    
    word = None
    try:
        # 1. 加载模板
        print("\n[1/4] 加载模板...")
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        template = template_skill.load_template(config.get("default_template", "xaut"))
        print(f"✓ 模板加载成功: {template['school']}")
        print(f"  - 正文字体: {template['body']['font_cn']} {template['body']['size']}pt")
        print(f"  - 一级标题: {template['heading1']['font_cn']} {template['heading1']['size']}pt")
        
        # 2. 连接 Word
        print("\n[2/4] 连接 Word...")
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        
        # 3. 打开测试文档
        test_doc = os.path.join(os.path.dirname(__file__), 'input', 'thesis.docx')
        if not os.path.exists(test_doc):
            print(f"❌ 测试文档不存在: {test_doc}")
            return False
        
        print(f"✓ 打开文档: {test_doc}")
        doc = word.Documents.Open(os.path.abspath(test_doc))
        
        # 4. 执行核心技能测试
        print("\n[3/4] 执行核心排版...")
        
        # 测试标题识别
        print("  → 标题识别...")
        heading_skill.process_headings(doc, template)
        h1_count = len([p for p in doc.Paragraphs if p.Format.OutlineLevel == 1])
        print(f"    ✓ 识别到 {h1_count} 个一级标题")
        
        # 测试分节符
        print("  → 分节符插入...")
        initial_sections = doc.Sections.Count
        section_break_skill.insert_odd_section_breaks(doc)
        final_sections = doc.Sections.Count
        print(f"    ✓ 从 {initial_sections} 节增加到 {final_sections} 节")
        
        # 5. 保存结果
        print("\n[4/4] 保存测试结果...")
        output_dir = os.path.join(os.path.dirname(__file__), 'test_output')
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, 'quick_test_result.docx')
        doc.SaveAs(output_path)
        doc.Close()
        
        print(f"✓ 测试结果已保存: {output_path}")
        print("\n" + "="*50)
        print("✅ 快速测试通过！")
        print("="*50)
        
        return True
        
    except Exception as e:
        print(f"\n❌ 测试失败: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
        
    finally:
        if word:
            try:
                word.Quit()
            except:
                pass


if __name__ == "__main__":
    success = quick_test()
    sys.exit(0 if success else 1)
