"""
install.py — 注册 Windows 右键菜单
运行方式（需要管理员权限）：python install.py
卸载：python install.py --uninstall
"""

import sys
import os
import winreg

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(SCRIPT_DIR, 'mdx_convert.py')
PYTHON_EXE = sys.executable  # 当前 Python 解释器路径


def get_python():
    """返回 pythonw.exe（无控制台窗口），fallback 到 python.exe。"""
    pythonw = os.path.join(os.path.dirname(PYTHON_EXE), 'pythonw.exe')
    return pythonw if os.path.exists(pythonw) else PYTHON_EXE


def register():
    pythonw = get_python()

    entries = [
        # .md 文件 → To Excel
        {
            'key':     r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToExcel',
            'label':   'To Excel (.xlsx)',
            'cmd_key': r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToExcel\command',
            'cmd':     f'"{pythonw}" "{SCRIPT_PATH}" to-xl "%1"',
        },
        # .xlsx 文件 → To Markdown
        {
            'key':     r'SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\MdxToMarkdown',
            'label':   'To Markdown (.md)',
            'cmd_key': r'SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\MdxToMarkdown\command',
            'cmd':     f'"{pythonw}" "{SCRIPT_PATH}" to-md "%1"',
        },
        # .md 文件 → To CSV
        {
            'key':     r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToCsv',
            'label':   'To CSV (.csv)',
            'cmd_key': r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToCsv\command',
            'cmd':     f'"{pythonw}" "{SCRIPT_PATH}" to-csv "%1"',
        },
        # .csv 文件 → To Markdown
        {
            'key':     r'SOFTWARE\Classes\SystemFileAssociations\.csv\shell\MdxToMarkdown',
            'label':   'To Markdown (.md)',
            'cmd_key': r'SOFTWARE\Classes\SystemFileAssociations\.csv\shell\MdxToMarkdown\command',
            'cmd':     f'"{pythonw}" "{SCRIPT_PATH}" to-md "%1"',
        },
    ]

    for entry in entries:
        # 创建菜单项键，设置显示名称
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, entry['key'],
                                0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, entry['label'])
            # 添加图标（使用 Python 图标）
            icon_path = os.path.join(os.path.dirname(PYTHON_EXE), 'DLLs', 'py.ico')
            if os.path.exists(icon_path):
                winreg.SetValueEx(key, 'Icon', 0, winreg.REG_SZ, icon_path)

        # 创建 command 键，设置执行命令
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, entry['cmd_key'],
                                0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, entry['cmd'])

        print(f"[已注册] {entry['label']}")

    print('\n右键菜单已安装。如不生效请重启资源管理器（taskkill /f /im explorer.exe & start explorer）')


def unregister():
    keys_to_delete = [
        r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToExcel\command',
        r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToExcel',
        r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToCsv\command',
        r'SOFTWARE\Classes\SystemFileAssociations\.md\shell\MdxToCsv',
        r'SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\MdxToMarkdown\command',
        r'SOFTWARE\Classes\SystemFileAssociations\.xlsx\shell\MdxToMarkdown',
        r'SOFTWARE\Classes\SystemFileAssociations\.csv\shell\MdxToMarkdown\command',
        r'SOFTWARE\Classes\SystemFileAssociations\.csv\shell\MdxToMarkdown',
    ]
    for key_path in keys_to_delete:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            print(f"[已删除] {key_path}")
        except FileNotFoundError:
            print(f"[跳过] 不存在: {key_path}")

    print('\n右键菜单已卸载。')


if __name__ == '__main__':
    if '--uninstall' in sys.argv:
        unregister()
    else:
        register()
