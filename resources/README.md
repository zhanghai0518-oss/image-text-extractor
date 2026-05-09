# Windows OCR 引擎说明

## PaddleOCR-json (推荐)

图片文字提取工具使用 **PaddleOCR-json** 作为 Windows 平台的 OCR 引擎，这是基于 PaddlePaddle PaddleOCR 的高性能中文 OCR 工具。

### 下载地址

从 GitHub Releases 下载：
- https://github.com/hiroi-sora/PaddleOCR-json/releases

选择最新版本的 `PaddleOCR-json_vX.X.X.7z` 文件。

### 安装步骤

1. 下载 `PaddleOCR-json_vX.X.X.7z` 文件
2. 使用 7-Zip 或其他解压软件解压
3. 将解压后的文件夹重命名为 `PaddleOCR-json`
4. 将 `PaddleOCR-json` 文件夹放到图片文字提取工具的 `resources` 目录下

### 目录结构

```
图片文字提取工具/
├── resources/
│   ├── PaddleOCR-json/
│   │   ├── PaddleOCR-json.exe    <-- OCR 主程序
│   │   ├── models/               <-- 语言模型（中文、英文等）
│   │   │   ├── ch_PP-OCRv3_det_infer/
│   │   │   ├── ch_PP-OCRv3_rec_infer/
│   │   │   └── ...
│   │   └── ...
│   └── win-ocr.exe               <-- 备用 OCR（识别效果较差）
└── ...
```

### 语言支持

默认包含以下语言：
- 简体中文 (ch)
- 繁体中文 (chinese_cht)
- 英文 (en)
- 日文 (japan)
- 韩文 (korean)

### 性能特点

- **准确率高**：基于 PaddleOCR v3/v4，中文识别效果极佳
- **速度快**：C++ 原生实现，支持 mkldnn 加速
- **离线运行**：无需联网，完全本地识别
- **内存占用**：约 1.5-2GB

## 备用 OCR (win-ocr.exe)

如果未安装 PaddleOCR-json，应用会回退到内置的 win-ocr.exe（基于 Windows 内置 OCR）。

**缺点**：
- 识别准确率较低
- 可能丢失部分文字
- 中文识别效果差

**强烈建议安装 PaddleOCR-json 以获得最佳体验！**
