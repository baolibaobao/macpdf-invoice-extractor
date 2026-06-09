# Mac 虚拟机测试指南

## 1. 下载构建产物

代码推送到 GitHub `main` 分支后，进入仓库页面：

1. 打开 `Actions`
2. 选择 `Build macOS App`
3. 打开最新一次成功运行
4. 在 `Artifacts` 下载 `DZPDFExtractor-macOS-BigSur-Intel`

下载后会得到以下文件之一：

- `DZPDFExtractor-macOS-BigSur-Intel.tar.gz`

## 2. 解压应用

推荐使用终端解压 `tar.gz`：

```bash
cd ~/Downloads
tar -xzf DZPDFExtractor-macOS-BigSur-Intel.tar.gz
```

解压后会得到：

```text
DZPDFExtractor.app
```

## 3. 解除 Gatekeeper 隔离

从 GitHub Actions 下载的应用没有签名，macOS 可能提示“已损坏”或阻止打开。

在终端执行：

```bash
xattr -cr ~/Downloads/DZPDFExtractor.app
```

如果应用放在其他路径，把命令里的路径换成实际路径。

## 4. 启动测试

```bash
open ~/Downloads/DZPDFExtractor.app
```

也可以在 Finder 中双击打开。

## 5. 测试内容

建议在纯净 Intel Mac 或 Big Sur 虚拟机中测试：

- 能否正常启动
- 能否选择发票文件夹
- 能否选择保存文件夹
- 能否提取多张 PDF
- UI 处理时是否卡顿
- 日志是否正常显示中文
- Excel 是否成功生成
- `打开` 按钮是否能打开保存目录

## 6. 常见问题

如果提示应用损坏：

```bash
xattr -cr /路径/到/DZPDFExtractor.app
```

如果仍无法打开，可尝试：

```bash
spctl --assess --verbose /路径/到/DZPDFExtractor.app
```

如果应用打开后闪退，请从终端启动主程序，复制报错内容：

```bash
~/Downloads/DZPDFExtractor.app/Contents/MacOS/DZPDFExtractor
```

本项目当前没有做 Apple Developer ID 签名和公证，所以 Gatekeeper 提示属于正常现象。
