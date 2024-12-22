import os
import sys
import shutil
from subprocess import run

def build_exe():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 读取requirements.txt
    req_file = os.path.join(current_dir, 'requirements.txt')
    with open(req_file, encoding='utf-8') as f:
        requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]
    
    # 确保环境正确
    print("检查环境...")
    for req in requirements:
        try:
            package = req.split('==')[0]  # 处理版本号
            __import__(package)
        except ImportError:
            print(f"安装依赖: {req}")
            run([sys.executable, '-m', 'pip', 'install', req], check=True)
    
    # 清理之前的构建文件
    build_dir = os.path.join(current_dir, 'build')
    dist_dir = os.path.join(current_dir, 'dist')
    for dir_path in [build_dir, dist_dir]:
        if os.path.exists(dir_path):
            print(f"清理目录: {dir_path}")
            shutil.rmtree(dir_path)

    print("\n开始打包...")
    main_path = os.path.join(current_dir, 'src', 'ExcelTrans', 'main.py')
    
    if not os.path.exists(main_path):
        print(f"错误: 找不到主程序文件 {main_path}")
        sys.exit(1)

    try:
        args = [
            'pyinstaller',
            '--clean',
            '--windowed',
            '--onefile',
            '--noconfirm',
            f'--distpath={dist_dir}',
            f'--workpath={build_dir}',
            '--name=考勤数据处理工具',
            '--add-data=README.md;.',
            '--add-data=LICENSE;.',
            '--add-data=requirements.txt;.',
        ]

        # 添加图标（如果存在）
        icon_path = os.path.join(current_dir, 'resources', 'icon.ico')
        if os.path.exists(icon_path):
            args.append(f'--icon={icon_path}')

        # 添加主程序
        args.append(main_path)

        # 执行打包
        run(args, check=True)
        
        # 检查生成的文件
        exe_path = os.path.join(dist_dir, '考勤数据处理工具.exe')
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"\n打包成功!")
            print(f"生成文件: {exe_path}")
            print(f"文件大小: {size_mb:.1f}MB")
        else:
            print("\n警告: 未找到生成的exe文件")
    except Exception as e:
        print(f"\n打包失败: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    build_exe() 