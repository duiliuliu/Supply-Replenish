# 简单测试程序，看看问题在哪里
import sys
print(f'Python version: {sys.version}')
print(f'Working directory: {sys.path}')

# 测试基本的导入
try:
    import tkinter as tk
    print('✅ tkinter imported')
except Exception as e:
    print(f'❌ tkinter failed: {e}')
    import traceback
    traceback.print_exc()

try:
    import pandas as pd
    print('✅ pandas imported')
except Exception as e:
    print(f'❌ pandas failed: {e}')
    import traceback
    traceback.print_exc()

try:
    from allocation_core import load_config
    print('✅ allocation_core imported')
except Exception as e:
    print(f'❌ allocation_core failed: {e}')
    import traceback
    traceback.print_exc()

print('\n测试完成')
