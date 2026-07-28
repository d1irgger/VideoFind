<p align="center">
  <img src="https://img.icons8.com/fluency/96/000000/video.png" alt="VideoFind Logo" width="80" height="80"/>
  <h1 align="center">🎬 VideoFind</h1>
  <p align="center">✨ 智能视频检索 · 一键直达 · 极速体验</p>
</p>

<p align="center">
  <a href="#-功能特性">功能</a> 
  <a href="#-快速开始">使用</a> 
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/Platform-Windows-blue?style=for-the-badge&logo=windows" alt="Platform">
  <img src="https://img.shields.io/badge/License-MIT-green?style=for-the-badge" alt="License">
  <img src="https://img.shields.io/badge/UI-iOS%20Style-007aff?style=for-the-badge" alt="UI">
</p>
---
## 📌 简介

**VideoFind** 是一款专为视频爱好者设计的桌面工具，帮你**快速检索**本地视频文件，**智能提取**文件属性中的网址链接，并提供 **一键播放**、**搜索过滤**、**导出链接** 等实用功能。界面采用 iOS 风格设计，简洁高效，让你在庞大的视频库中轻松定位所需内容。

---

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| 📁 **批量扫描** | 支持递归扫描多种视频格式（MP4, MOV, MKV, AVI, FLV, WEBM, WMV, MPG, MPEG） |
| 🔗 **智能链接提取** | 需要提前将链接输入到文件详情标题栏内 |
| ▶️ **本地播放** | 一键用系统默认播放器或浏览器打开视频（可切换）浏览器可使用VSR |
| 🌐 **网页跳转** | 直接打开提取的链接，若无则执行google搜索 |
| 🔍 **实时搜索** | 按文件名即时过滤，不区分大小写，单字即可匹配 |
| 📤 **导出链接** | 将文件名与对应链接导出为 `log.txt`，更便捷视频管理 |

---

## 🚀 快速开始
python -m PyInstaller --onefile --windowed --name "VideoFind" --hidden-import win32com --hidden-import win32com.client video_finder.py

### 环境要求
- Windows 10 以上操作系统（标题读取依赖 Windows Shell API）
- Python 3.8 及以上

