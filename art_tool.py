# -*- coding: utf-8 -*-
"""美术资源配置工具 - 入口。

结构：core.py（纯逻辑：扫描/匹配/写入/校验）+ gui.py（tkinter 界面）。
运行: python art_tool.py
"""
from gui import App

if __name__ == "__main__":
    App().mainloop()
