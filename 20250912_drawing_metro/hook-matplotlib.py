# hook-matplotlib.py
from PyInstaller.utils.hooks import collect_all, collect_data_files

# 方法1：使用 collect_all
datas, binaries, hiddenimports = collect_all('matplotlib')
