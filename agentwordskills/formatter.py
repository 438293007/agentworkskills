import os
import json
import shutil
import win32com.client as win32
from skills import (
    template_skill, cover_skill, heading_skill, toc_skill, 
    section_break_skill, page_number_skill, header_footer_skill, 
    figure_table_skill, reference_skill, check_skill
)

def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def main(doc_path):
    config = load_config()
    template_name = config.get("default_template", "xaut")
    
    print(f"[*] 初始化 论文排版助手 V2 - 应用模板: {template_name}")
    template_config = template_skill.load_template(template_name)
    
    if config.get("backup_original"):
        backup_path = doc_path.replace('.docx', '_backup.docx')
        shutil.copy(doc_path, backup_path)
        print(f"[*] 已备份原始文档至: {backup_path}")

    print("[*] 正在连接 Word COM 接口...")
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True 
    
    abs_path = os.path.abspath(doc_path)
    doc = word.Documents.Open(abs_path)
    
    try:
        print("[1/9] 应用全局页面与正文模板...")
        template_skill.apply_template(doc, template_config)
        
        # print("[2/9] 处理封面与声明页...")
        # cover_skill.process_cover(doc)
        
        print("[3/9] 识别并设置标题层级...")
        heading_skill.process_headings(doc, template_config)
        
        print("[4/9] 重建并格式化目录 (分离正文)...")
        toc_skill.rebuild_toc(doc, levels=3, template=template_config)
        
        print("[5/9] 插入章节奇数页分节符...")
        section_break_skill.insert_odd_section_breaks(doc)
        
        print("[6/9] 断开链接并配置罗马/阿拉伯数字页码...")
        page_number_skill.setup_page_numbers(doc)
        
        print("[7/9] 设置页眉页脚 (奇偶页不同)...")
        header_footer_skill.setup_headers_footers(doc)
        
        print("[8/9] 处理图表规范...")
        figure_table_skill.format_figures_and_tables(doc)
        # 【新增这一段】：在第8步之后，激活参考文献技能！
        print("[9/9] 格式化参考文献 (生成自动编号)...")
        reference_skill.format_references(doc)
       
        # ======================================================
        print("[10/10] 执行最终排版与内容审查...")
        check_skill.check_document(doc, template_config)
         # ===== 请将这几行加在主程序即将结束、准备保存之前 =====
        print("[*] 正在进行最终的文档分页计算与目录同步...")
        doc.Repaginate() 
        
        if doc.TablesOfContents.Count > 0:
            doc.TablesOfContents(1).Update() # Word 在这里会搞破坏
            
            # 【核心反制】：在 Word 搞完破坏后，立刻呼叫修复函数重新锁定格式！
            from skills.toc_skill import fix_toc_format
            fix_toc_format(doc)
        print("[*] 请打开排版后的文档进行检查并手动更新目录...")
        # 保存并输出
        output_doc = os.path.abspath(doc_path.replace('.docx', '_final.docx'))
        doc.SaveAs(output_doc)
        print(f"\n[ok] 排版全部完成！已保存为: {output_doc}")
        
    except Exception as e:
        print(f"\n[error] 排版过程中发生异常: {str(e)}")
    finally:
        # doc.Close() 调试阶段保持文档打开状态以便观察
        pass

if __name__ == "__main__":
    # 请将此处替换为你的待排版 Word 文档路径
   main(r"input\test.docx")