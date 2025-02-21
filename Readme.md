# 简介
这是一个使用Python和Tkinter库编写的投标辅助工具。    
目前功能包括
- 暗标格式刷
- pdf转图片

计划加入：
- 多标书查重
- 招标要求与投标响应批量比对    
# 使用方法
仅支持Windows且需win10以上系统。
在[此处](https://github.com/libin47/toubiao_zhushou/releases)下载最新的zip文件，解压后运行其中的main.exe。

# 开发
## 环境需求
需python==3.13（低些应该也行，未经测试）
```bash
pip install -r requirements.txt
```
## 文件结构
```
toubiao_zhushou/
│
├── main.py             # 程序入口
│
└─── app/
    ├── __init__.py             # 主界面窗口
    │
    ├── hidden_clean.py         # 暗标格式刷窗口   
    ├── hidden_clean_config.py  # 暗标格式刷配置
    ├── hidden_clean_fun.py     # 暗标格式刷功能
    │
    ├ ── pdf2img.py             # pdf转图片窗口
    ├─ ─ pdf2img_config.py      # pdf转图片配置
    ├──  pdf2img_fun.py         # pdf转图片功能
    │
    ├── utils.py                 # 工具函数
```

## 运行
```bash
python main.py
```
## 打包
如果使用-F则会打包成一个文件，对启动速度有较大影响。
```
pyinstaller -F -w main.py
或
pyinstaller -w main.py
```

或者安装nuitka使用nuitka打包（按说比pyinstaller的快些且文件更小，但实际上文件大了好几倍且启动和运行速度没有明显变化，不知道是不是配置哪儿不对）。
```commandline
python -m nuitka --standalone --windows-disable-console --mingw64 --nofollow-imports --show-progress --enable-plugin=tk-inter --windows-icon-from-ico=ico.ico --onefile main.py
```
