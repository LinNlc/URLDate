推荐的程序框架目录
url_tool/
├── run.py                     # 程序入口（python run.py / 打包exe指向这里）
├── requirements.txt           # Pillow / openpyxl / requests / PySide6 ...
├── microdrama/                # 主包
│   ├── __init__.py
│   ├── app.py                 # 启动GUI、全局初始化
│   ├── gui/                   # 图形界面相关
│   │   ├── main_window.py     # 主窗口：选文件、模式、按钮、日志框、进度条
│   │   ├── staff_dialog.py    # 人员库管理弹窗
│   │   └── resources/         # 图标、样式表
│   ├── core/                  # 业务核心（你现在的大部分逻辑）
│   │   ├── excel_processor.py # “处理Excel”这件事都放这
│   │   ├── image_fetcher.py   # 下载并缩放图片的独立模块（支持并发）
│   │   ├── staff_db.py        # 现在的 staff_database.json 的读写和匹配
│   │   ├── config_store.py    # 现在的 config.ini 读写
│   │   └── version_checker.py # 版本检查 / 更新地址（可选）
│   └── utils/
│       ├── logger.py          # 统一日志输出，GUI也能接收
│       └── workers.py         # 后台线程/任务封装，避免GUI卡死
└── data/
    └── staff_database.json    # 人员库文件（运行时生成/用户维护）
