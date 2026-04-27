# PyInstaller打包脚本 - 支持平台区分
import PyInstaller.__main__
import sys
import os
import platform

def get_platform_name():
    system = platform.system().lower()
    if system == 'windows':
        return 'Windows'
    elif system == 'darwin':
        return 'Mac'
    elif system == 'linux':
        return 'Linux'
    return system.capitalize()

def build_app():
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 确定平台名称
    platform_name = get_platform_name()
    app_name = f'加单分配系统_v3_{platform_name}'
    
    print(f'当前平台：{platform_name}')
    print(f'应用名称：{app_name}')
    
    # 打包参数
    args = [
        'allocation_app.py',
        '--name', app_name,
        '--windowed',  # 无控制台窗口
        '--onefile',   # 打包成单个文件
        '--clean',
        '--noconfirm',
    ]
    
    # 平台特定配置
    if platform_name == 'Windows':
        args.append('--add-data=allocation_core.py;.')
        args.append('--icon=NONE')
    else:
        args.append('--add-data=allocation_core.py:.')
    
    print('开始打包应用程序...')
    print('使用参数:', args)
    
    try:
        PyInstaller.__main__.run(args)
        
        print('\n✅ 打包完成！')
        print(f'可执行文件位于：{os.path.join(current_dir, "dist")} 目录下')
        
        if platform_name == 'Windows':
            print(f'\nWindows版本：{app_name}.exe')
        elif platform_name == 'Mac':
            print(f'\nMac版本：{app_name}.app')
        else:
            print(f'\nLinux版本：{app_name}')
        
    except Exception as e:
        print(f'\n❌ 打包失败：{e}')

if __name__ == '__main__':
    build_app()
