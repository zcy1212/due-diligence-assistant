# 尽调小助手 - 安装使用指南

## 方法一：完整导入（推荐）

### 1. 解压文件
将压缩包解压到你想要的目录。

### 2. 打开项目
在Trae中打开解压后的目录。

### 3. 安装依赖
打开终端，运行：
```bash
pip install -r requirements.txt
```

### 4. 开始使用
在Trae的对话栏中输入：
```
/尽调小助手
```

## 方法二：手动复制Skill

### 1. 复制Skill文件
将 `.claude/skills/尽调小助手.md` 复制到你的项目的 `.claude/skills/` 目录下。

### 2. 复制工具脚本
将 `tools/` 目录下的所有Python文件复制到你的项目的 `tools/` 目录。

### 3. 复制模板
将 `templates/` 目录下的模板文件复制到你的项目的 `templates/` 目录。

### 4. 创建必要目录
确保你的项目有以下目录：
- `input/` - 放待处理文件
- `output/` - 生成的文件
- `temp/` - 临时文件

### 5. 安装依赖
```bash
pip install -r requirements.txt
```

## 目录结构（安装后）

```
你的项目/
├── .claude/
│   └── skills/
│       └── 尽调小助手.md     ← Skill配置
├── tools/
│   ├── file_reader.py         ← 文件读取
│   ├── excel_writer.py        ← Excel生成
│   └── report_writer.py       ← Word生成
├── templates/
│   ├── 报告模板.txt
│   └── 项目表示例.xlsx
├── input/                     ← 放待处理文件
├── output/                    ← 生成的文件
├── temp/                      ← 临时文件
└── requirements.txt           ← 依赖列表
```

## 使用说明

### 基本使用
1. 将待处理的文件夹或压缩包放入 `input/` 目录
2. 在Trae中输入 `/尽调小助手`
3. 按照提示选择操作

### 迭代补充材料
1. 拿到《待补充材料清单》
2. 让目标公司补充材料
3. 将新材料放入 `input/` 目录
4. 重新运行 `/尽调小助手`
5. 选择迭代流程选项

## 常见问题

### Q: 提示缺少依赖怎么办？
A: 运行 `pip install -r requirements.txt` 安装所有依赖。

### Q: PDF文件读取有问题？
A: 工具会尝试用PyMuPDF和PyPDF2两种方式读取，如果都不行会标注[无法读取]。

### Q: 可以自定义报告模板吗？
A: 可以，将你的模板放入 `templates/` 目录，运行时选择自定义模板选项。

### Q: 如何清理临时文件？
A: 可以手动删除 `temp/` 和 `output/` 目录下的文件。

## 技术支持

如遇到问题，请检查：
1. Python版本 >= 3.8
2. 所有依赖已正确安装
3. 目录权限正常
