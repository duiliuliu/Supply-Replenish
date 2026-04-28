# PyInstaller打包脚本 - 支持平台区分
import PyInstaller.__main__
import sys
import os
import platform

# 设置编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

def get_platform_name():
    system = platform.system().lower()
    if system == 'windows':
        return 'Windows'
    elif system == 'darwin':
        # 检测 Mac 架构，优先使用环境变量
        target_arch = os.environ.get('PYINSTALLER_TARGET_ARCH')
        if target_arch == 'x86_64':
            return 'Mac-Intel'
        elif target_arch == 'arm64':
            return 'Mac-ARM'
        # 自动检测
        machine = platform.machine().lower()
        if machine in ['arm64', 'aarch64']:
            return 'Mac-ARM'
        else:
            return 'Mac-Intel'
    elif system == 'linux':
        return 'Linux'
    return system.capitalize()

def build_app():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 确定平台名称
    platform_name = get_platform_name()
    app_name = f'加单分配系统_v3_{platform_name}'
    
    print(f'Current platform: {platform_name}')
    print(f'App name: {app_name}')
    
    # 打包参数
    args = [
        'allocation_app.py',
        '--name', app_name,
        '--windowed',  # 无控制台窗口
        '--onefile',   # 打包成单个文件
        '--clean',
        '--noconfirm',
        '--hidden-import=allocation_core',
    ]
    
    print('Starting build process...')
    print('Using arguments:', args)
    
    try:
        PyInstaller.__main__.run(args)
        
        print('\n✅ Build completed!')
        print(f'Executable located in: {os.path.join(current_dir, "dist")} directory')
        
        if platform_name == 'Windows':
            print(f'\nWindows version: {app_name}.exe')
        elif platform_name == 'Mac':
            print(f'\nMac version: {app_name}.app')
        else:
            print(f'\nLinux version: {app_name}')
        
    except Exception as e:
        print(f'\n❌ Build failed: {e}')

if __name__ == '__main__':
    build_app()
