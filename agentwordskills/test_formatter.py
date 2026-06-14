import os
import time
import subprocess
import shutil

# --- 配置测试路径 ---
# 请根据你的实际文件名修改
INPUT_RAW_DOC = r"input\test.docx"           # 准备用于测试的原稿
OUTPUT_DOC = r"input\test_final.docx"        # 👉 严格对齐主程序的输出路径！
MAIN_SCRIPT = "formatter.py"                # 主控制程序

# 终端颜色输出配置（让测试结果更直观）
class Colors:
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    RESET = '\033[0m'

def print_step(msg):
    print(f"\n{Colors.YELLOW}[TEST STEP]{Colors.RESET} {msg}")

def run_tests():
    print(f"{"="*50}")
    print("🚀 启动排版引擎自动化测试流程")
    print(f"{"="*50}")
    
    # ---------------------------------------------------------
    # 步骤 1：环境与物料检查
    # ---------------------------------------------------------
    print_step("检查输入源文件...")
    if not os.path.exists(INPUT_RAW_DOC):
        print(f"{Colors.RED}❌ 致命错误：未找到测试原稿 {INPUT_RAW_DOC}{Colors.RESET}")
        print("请确保 input 文件夹下有一份干净的未排版原稿！")
        return
    print(f"{Colors.GREEN}✓ 测试原稿就绪 ({os.path.getsize(INPUT_RAW_DOC) / 1024:.2f} KB){Colors.RESET}")

    # 清理上一次的旧输出，防止假阳性测试结果
    if os.path.exists(OUTPUT_DOC):
        try:
            os.remove(OUTPUT_DOC)
            print(f"{Colors.GREEN}✓ 已清理历史输出文件{Colors.RESET}")
        except Exception as e:
            print(f"{Colors.RED}❌ 无法删除旧输出文件，请检查是否被 Word 占用。报错: {e}{Colors.RESET}")
            return

    # ---------------------------------------------------------
    # 步骤 2：启动主排版程序 (子进程拦截)
    # ---------------------------------------------------------
    print_step(f"开始执行主程序: {MAIN_SCRIPT}")
    start_time = time.time()
    
    try:
        # 使用 subprocess 运行，并捕获所有的终端输出和报错
        process = subprocess.run(
            ["python", MAIN_SCRIPT],
            capture_output=True,
            text=True,
            check=True # 如果主程序崩溃返回非 0 状态码，直接抛出 CalledProcessError
        )
        print(f"{Colors.GREEN}✓ 主程序执行完毕，未发生崩溃！{Colors.RESET}")
        
    except subprocess.CalledProcessError as e:
        # 如果排版代码中途报错（像之前报 TypeError 那样），这里会精准拦截
        elapsed_time = time.time() - start_time
        print(f"{Colors.RED}❌ 排版引擎在运行时崩溃！(耗时: {elapsed_time:.2f}秒){Colors.RESET}")
        print(f"\n{Colors.RED}--- 崩溃日志抓取 ---{Colors.RESET}")
        print(e.stderr)
        return

    elapsed_time = time.time() - start_time
    print(f"{Colors.GREEN}✓ 引擎耗时: {elapsed_time:.2f} 秒{Colors.RESET}")

    # ---------------------------------------------------------
    # 步骤 3：输出产物验证 (物理验证)
    # ---------------------------------------------------------
    print_step("验证产物生成状态...")
    if not os.path.exists(OUTPUT_DOC):
        print(f"{Colors.RED}❌ 测试失败：主程序运行成功，但未生成目标文件！{Colors.RESET}")
        return
        
    file_size_kb = os.path.getsize(OUTPUT_DOC) / 1024
    if file_size_kb < 10: # Word 文档一般不可能小于 10KB
        print(f"{Colors.RED}❌ 测试失败：生成的文档过小 ({file_size_kb:.2f} KB)，可能为空文件！{Colors.RESET}")
        return
        
    print(f"{Colors.GREEN}✓ 成功生成规范文档 ({file_size_kb:.2f} KB){Colors.RESET}")

    # ---------------------------------------------------------
    # 测试总结
    # ---------------------------------------------------------
    print(f"\n{"="*50}")
    print(f"{Colors.GREEN}🎉 E2E 自动化测试全流程通过！排版引擎运行极其健康！{Colors.RESET}")
    print(f"{"="*50}")

if __name__ == "__main__":
    run_tests()