# Excel-Trans

考勤数据处理工具 - 一个用于处理和统计Excel格式考勤数据的桌面应用。

## 功能特点

- 简单易用的图形界面
- 支持Excel文件导入导出
- 自动处理考勤数据
- 生成统计报表

## 开发环境准备

1. 安装Python环境（推荐Python 3.8及以上版本）
```bash
# 检查Python版本
python --version
```

2. 克隆项目
```bash
git clone https://github.com/yourusername/excel-trans.git
cd excel-trans
```

3. 创建虚拟环境
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux/Mac
python -m venv venv
source venv/bin/activate
```

4. 安装依赖
```bash
pip install -r requirements.txt
```

## 项目结构

```
excel-trans/
├── src/
│ └── ExcelTrans/
│ ├── init.py
│ └── main.py # 主程序
├── resources/
│ └── icon.ico # 程序图标
├── README.md # 项目说明
├── requirements.txt # 依赖列表
├── LICENSE # 许可证
└── build_exe.py # 打包脚本
```


## 开发运行

1. 确保在虚拟环境中
   ```bash
   # Windows
   venv\Scripts\activate

   # Linux/Mac
   source venv/bin/activate
   ```

2. 运行程序
   ```bash
   python src/ExcelTrans/main.py
   ```

## 打包发布

1. 安装打包工具
```bash
pip install pyinstaller
```

2. 运行打包脚本
```bash
python build_exe.py
```

3. 打包完成后，可执行文件位于：
   - `dist/考勤数据处理工具.exe`

## 使用说明

1. 启动程序
   - 双击运行 `考勤数据处理工具.exe`
   - 或在开发环境中运行 `python src/ExcelTrans/main.py`

2. 选择输入文件
   - 点击"选择文件"按钮
   - 选择需要处理的Excel文件

3. 选择输出位置
   - 默认在原文件同目录下生成结果文件
   - 可以点击"选择位置"更改输出路径

4. 处理数据
   - 点击"开始处理"按钮
   - 等待处理完成
   - 查看处理日志
   - 检查输出文件

## 常见问题

1. 找不到依赖库
```bash
# 重新安装依赖
pip install -r requirements.txt
```

2. 打包失败
```bash
# 确保在虚拟环境中
# 清理之前的构建文件
rm -rf build dist
# 重新运行打包脚本
python build_exe.py
```

## 贡献指南

1. Fork 项目
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

## 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 联系方式

作者 - [@yourusername](https://github.com/yourusername)

项目链接: [https://github.com/yourusername/excel-trans](https://github.com/yourusername/excel-trans)