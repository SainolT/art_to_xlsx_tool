' 美术资源配置工具启动器：双击运行 art_tool.py，不弹黑色控制台窗口
' 需要和 art_tool.py 放在同一目录，且系统已安装 Python（含 tkinter）和 openpyxl、pypinyin
Dim shell, fso, dir
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Wscript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = dir
shell.Run "pythonw.exe """ & dir & "\art_tool.py""", 0, False
