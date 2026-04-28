# PyInstaller打包脚本 - 支持平台区分和版本号
import PyInstaller.__main__
import sys
import os
import platform
import toml

# 设置编码为UTF-8
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

def get_version():
    """从pyproject.toml获取版本号"""
    try:
        with open('pyproject.toml', 'r', encoding='utf-8') as f:
            config = toml.load(f)
            return config['project']['version']
    except Exception as e:
        print(f'Warning: Could not read version from pyproject.toml: {e}')
        return 'unknown'

def get_platform_name():
    system = platform.system().lower()
    if system == 'windows':
        return 'Windows'
    elif system == 'darwin':
        target_arch = os.environ.get('PYINSTALLER_TARGET_ARCH')
        if target_arch == 'x86_64':
            return 'Mac-Intel'
        elif target_arch == 'arm64':
            return 'Mac-ARM'
        machine = platform.machine().lower()
        if machine in ['arm64', 'aarch64']:
            return 'Mac-ARM'
        else:
            return 'Mac-Intel'
    elif system == 'linux':
        return 'Linux'
    return system.capitalize()

def build_app():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    platform_name = get_platform_name()
    version = get_version()
    
    app_name = f'加单分配系统_v3_v{version}_{platform_name}'
    
    print(f'Current platform: {platform_name}')
    print(f'App version: {version}')
    print(f'App name: {app_name}')
    
    args = [
        'allocation_app.py',
        '--name', app_name,
        '--windowed',
        '--clean',
        '--noconfirm',
        '--hidden-import=allocation_core',
        '--add-data=allocation_config.json:.',
    ]
    
    if platform_name.startswith('Mac'):
        args.append('--onedir')
    else:
        args.append('--onefile')
    
    print('Starting build process...')
    print('Using arguments:', args)
    
    try:
        PyInstaller.__main__.run(args)
        
        print('\n✅ Build completed!')
        
        dist_path = os.path.join(current_dir, 'dist')
        
        if platform_name.startswith('Mac'):
            app_path = os.path.join(dist_path, f'{app_name}.app')
            print(f'\nMac version: {app_path}')
            if os.path.exists(app_path):
                print(f'App bundle created successfully!')
            else:
                print(f'Warning: App bundle not found at {app_path}')
                print(f'Contents of dist: {os.listdir(dist_path)}')
        elif platform_name == 'Windows':
            exe_path = os.path.join(dist_path, f'{app_name}.exe')
            print(f'\nWindows version: {exe_path}')
        else:
            bin_path = os.path.join(dist_path, app_name)
            print(f'\nLinux version: {bin_path}')
        
    except Exception as e:
        print(f'\n❌ Build failed: {e}')
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    build_app()
