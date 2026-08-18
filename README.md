# zj公式转化工具（LaTeX → Word 公式一键转换）

把 Word 文档中**以纯文本形式存在的 LaTeX 公式**，一键编译成 **Word 原生可编辑公式对象（OMML）**，全程无需打开 Word、无需任何插件。

> 适用场景：论文 / 竞赛说明书里公式显示为 `\frac{...}{...}`、`E=mc^2` 之类的原始代码，想要它们变成 Word 里可编辑的"真公式"。

---

## 原理

```
LaTeX 文本 ──▶ latex2mathml ──▶ MathML ──▶ 微软官方 MML2OMML.XSL ──▶ OMML 公式对象
（纯文本）   （Python 库）                    （XSLT 转换）                （注入 docx）
```

转换直接改写 docx 的 XML（`word/document.xml` 等），**不依赖 Word 运行**，速度快（数十处公式约 2 秒）。

## 使用

### 方式一：直接运行成品 exe（推荐）

下载 `dist/公式一键转换.exe`，任选其一：

1. **双击** → 弹出文件选择框 → 选择 .docx；
2. **拖放**：把 .docx 直接拖到 exe 图标上；
3. **命令行**：`公式一键转换.exe 文档路径.docx`

输出：`原名_公式版.docx`（与源文件同目录）。完成弹窗会显示成功转换处数；解析失败处按原文保留并在弹窗中列出。

另有 `dist/公式一键转换_诊断版.exe`（控制台版）：报错信息直接显示在黑窗口中，用于排查"没反应"类问题。

### 方式二：源码运行

```bash
pip install -r requirements.txt
python fetch_xsl.py        # 从本机 Office 安装目录复制 MML2OMML.XSL（微软组件，不随仓库分发）
python latex2docx.py 文档路径.docx
```

## 构建 exe（PyInstaller）

```bash
pip install -r requirements.txt
python fetch_xsl.py
build.bat
```

- 产物在 `dist/`：`公式一键转换.exe`（windowed 主版）+ `公式一键转换_诊断版.exe`（console 版）。
- **注意事项**：`latex2mathml` 运行需要 `unimathsymbols.txt` 数据文件，必须用 `--add-data` 一并打包，否则打包后运行时报 FileNotFoundError。
- 环境提示：若在 WorkBuddy 等沙箱化终端里构建报 `os.remove` 被拦截（safe-delete），可用
  `PYTHONPATH=<venv site-packages> <venv python> -S -m PyInstaller ...` 跳过 site 钩子恢复原生删除。

## 公式识别策略（保守，避免误伤正文）

- 含 **LaTeX 反斜杠** / **上下标 `_ ^`** / **强数学运算符**（× · ÷ ± ≤ ≥ ≈ √ ∑ ∫ …）的片段 → 转公式；
- 纯中文比较句（如 `CO₂ > 1000 ppm`）→ 保留为普通文本，不乱转；
- 中文下标（如 `E_{夏}`）通过括号深度感知整体转换。

## 常见问题

| 问题 | 处理 |
|---|---|
| 双击 exe 没反应 | 未签名 exe 常被 Windows SmartScreen 拦截：右键 exe → 属性 → **解除锁定**，或拦截弹窗点"仍要运行" |
| 仅支持什么格式 | 仅 `.docx`（新版），不支持旧版 `.doc` |
| 报错信息 | 程序同目录会写 `公式转换_error.log` |

## 目录结构

```
latex2docx/
├── latex2docx.py        # 主程序（转换核心 + ctypes GUI）
├── make_icon.py         # 生成 gs.ico 图标（图标为二进制，不入库，build.bat 自动生成）
├── make_test_docx.py    # 生成 LaTeX 公式测试文档
├── fetch_xsl.py         # 从本机 Office 复制 MML2OMML.XSL（微软组件，不入库）
├── requirements.txt     # Python 依赖
├── build.bat            # 一键构建两个 exe（自动生成图标 + 获取 XSL）
└── dist/                # 构建产物（exe 为二进制，不入库）
    ├── 公式一键转换.exe
    └── 公式一键转换_诊断版.exe
```

## 版权说明

- `MML2OMML.XSL` 是 **Microsoft Office** 自带组件，版权归微软所有。本仓库**不包含**该文件，请通过 `fetch_xsl.py` 从你已获授权的 Office 安装中获取。
- 其余代码可按需使用（欢迎 fork / 改进）。
