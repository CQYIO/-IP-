# 局域网下获取全部IP地址
局域网下获取全部IP地址

需要注意，只能在Win10和Win11上进行运行，Win7版本不支持

# 安装依赖包：
pip install openpyxl \n
pip install pyinstaller

# 打开本地文件地址，进入CMD中执行下面的语句，生成.EXE可执行文件在同级目录下的\dist中
pyinstaller --noconsole --onefile ip_scanner_gui.py
