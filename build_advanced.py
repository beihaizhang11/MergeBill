"""
高级打包脚本 - 使用PyInstaller打包Excel账单合并工具
支持更多自定义选项
"""
import os
import sys
import shutil
import subprocess
from pathlib import Path


def check_pyinstaller():
    """检查PyInstaller是否已安装"""
    try:
        import PyInstaller
        print(f"✓ 检测到PyInstaller版本: {PyInstaller.__version__}")
        return True
    except ImportError:
        print("✗ 未检测到PyInstaller")
        response = input("是否现在安装PyInstaller? (y/n): ")
        if response.lower() == 'y':
            print("正在安装PyInstaller...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            return True
        return False


def clean_build_files():
    """清理旧的构建文件"""
    print("\n[清理] 删除旧的构建文件...")
    dirs_to_remove = ["build", "dist", "__pycache__"]
    files_to_remove = ["*.spec"]
    
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  - 已删除: {dir_name}/")
    
    for pattern in files_to_remove:
        for file in Path(".").glob(pattern):
            file.unlink()
            print(f"  - 已删除: {file}")
    
    print("✓ 清理完成")


def build_exe(console=False, onefile=True):
    """
    构建可执行文件
    
    Args:
        console: 是否显示控制台窗口
        onefile: 是否打包成单个文件
    """
    print("\n[构建] 开始打包程序...")
    print(f"  - 模式: {'单文件' if onefile else '多文件'}")
    print(f"  - 控制台: {'显示' if console else '隐藏'}")
    
    app_name = "Excel账单合并工具" + ("_调试版" if console else "")
    
    # 构建PyInstaller命令
    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--clean",
        "--name", app_name,
    ]
    
    # 添加打包模式
    if onefile:
        cmd.append("--onefile")
    else:
        cmd.append("--onedir")
    
    # 添加窗口模式
    if console:
        cmd.append("--console")
    else:
        cmd.append("--windowed")
    
    # 添加隐藏导入
    hidden_imports = [
        "openpyxl",
        "openpyxl.cell",
        "openpyxl.styles",
        "openpyxl.utils",
    ]
    
    for module in hidden_imports:
        cmd.extend(["--hidden-import", module])
    
    # 如果config.json存在，添加到打包中
    if os.path.exists("config.json"):
        cmd.extend(["--add-data", "config.json;."])
    
    # 添加主程序
    cmd.append("main.py")
    
    # 执行打包
    print(f"\n执行命令: {' '.join(cmd)}\n")
    
    try:
        subprocess.check_call(cmd)
        print("\n✓ 打包成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n✗ 打包失败: {e}")
        return False


def show_result():
    """显示构建结果"""
    print("\n" + "="*50)
    print("构建结果")
    print("="*50)
    
    dist_dir = Path("dist")
    if dist_dir.exists():
        files = list(dist_dir.glob("*.exe"))
        if files:
            for file in files:
                size_mb = file.stat().st_size / (1024 * 1024)
                print(f"\n文件名: {file.name}")
                print(f"大小: {size_mb:.2f} MB")
                print(f"路径: {file.absolute()}")
        else:
            print("\n未找到exe文件")
    else:
        print("\ndist目录不存在")
    
    print("\n" + "="*50)


def main():
    """主函数"""
    print("="*50)
    print("Excel账单合并工具 - 高级打包脚本")
    print("="*50)
    
    # 检查PyInstaller
    if not check_pyinstaller():
        print("\n需要安装PyInstaller才能继续")
        return
    
    # 显示选项
    print("\n请选择打包选项：")
    print("1. 标准版（无控制台，单文件）")
    print("2. 调试版（带控制台，单文件）")
    print("3. 标准版（无控制台，多文件）")
    print("4. 调试版（带控制台，多文件）")
    print("5. 全部打包")
    
    choice = input("\n请输入选项 (1-5): ").strip()
    
    if choice not in ['1', '2', '3', '4', '5']:
        print("无效的选项")
        return
    
    # 清理旧文件
    clean_build_files()
    
    # 根据选择打包
    configs = []
    if choice == '1':
        configs = [(False, True)]
    elif choice == '2':
        configs = [(True, True)]
    elif choice == '3':
        configs = [(False, False)]
    elif choice == '4':
        configs = [(True, False)]
    elif choice == '5':
        configs = [(False, True), (True, True)]
    
    success_count = 0
    for console, onefile in configs:
        if build_exe(console=console, onefile=onefile):
            success_count += 1
        
        # 清理中间文件
        if os.path.exists("build"):
            shutil.rmtree("build")
        for spec_file in Path(".").glob("*.spec"):
            spec_file.unlink()
    
    # 显示结果
    show_result()
    
    print(f"\n完成！成功打包 {success_count}/{len(configs)} 个版本")
    print("\n使用说明：")
    print("1. 将exe文件复制到目标位置")
    print("2. 首次运行会自动生成config.json配置文件")
    print("3. 如需备份配置，保存config.json文件即可")
    

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n操作已取消")
    except Exception as e:
        print(f"\n发生错误: {e}")
    finally:
        input("\n按任意键退出...")

