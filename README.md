# 微剧 URL 转换工具

全新的图形化版本，基于 PySide6 构建，支持一键选择 Excel 文件、后台下载剧照并插入到 H 列，同时保留原有的内容质检与审核人员匹配能力。项目结构已经模块化，方便维护与扩展。

## 目录结构

```
url_tool/
├── run.py
├── requirements.txt
├── data/
│   └── staff_database.json
└── microdrama/
    ├── __init__.py
    ├── app.py
    ├── gui/
    │   ├── __init__.py
    │   ├── main_window.py
    │   └── staff_dialog.py
    ├── core/
    │   ├── __init__.py
    │   ├── config_store.py
    │   ├── excel_processor.py
    │   ├── image_fetcher.py
    │   ├── staff_db.py
    │   ├── text_utils.py
    │   └── version_checker.py
    └── utils/
        ├── __init__.py
        ├── logger.py
        └── workers.py
```

## 功能亮点

- **PySide6 图形界面**：支持文件选择、模式切换、进度条显示与日志查看。
- **并发图片下载与压缩**：使用线程池快速抓取剧照，自动调整尺寸后插入 Excel。
- **内容检测与人员匹配**：继续对内容摘要、演员字段进行校验，并自动填充身份证后四位。
- **人员库管理弹窗**：可在 GUI 中新增、删除审核人员。
- **模式配置持久化**：处理模式写入 `config.ini`，下次启动自动加载。

## 快速开始

1. 安装依赖：
   ```bash
   pip install -r url_tool/requirements.txt
   ```
2. 运行程序：
   ```bash
   python url_tool/run.py
   ```
3. 在界面中选择 Excel 文件，确认处理模式后点击“开始处理”。

## 注意事项

- 输出文件默认保存在与原始 Excel 相同的目录下：
  - 模式 2：`*_converted.xlsx`
  - 模式 1：每 50 条拆分为多个 `*_partN.xlsx`
- 图片处理完成后，请在 Excel 中手动设置图片属性为“移动并调整大小”。
- 人员库保存在 `url_tool/data/staff_database.json`，也可在 GUI 中维护。

