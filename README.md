# 科技商学院教师教学工作量填报系统

面向西京学院科技商学院教师的轻量化教学工作量填报工具，支持：

- 教师基本信息与教学工作量明细填报
- 工作量自动合计及必填项校验
- 按审核人顺序整理并导出 Excel 报表
- 将当前填报进度暂存为 JSON，并在后续导入恢复
- A4 横向 Excel 输出，宽度固定为一页、纵向按内容自动分页

## 本地运行

建议使用 Python 3.11。

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

浏览器打开 `http://localhost:8501`。

## 部署到 Streamlit Community Cloud

1. 将本项目推送到 GitHub。
2. 登录 [Streamlit Community Cloud](https://share.streamlit.io/) 并连接 GitHub。
3. 点击 **Create app**，选择对应的 GitHub 仓库。
4. Branch 选择 `main`，Main file path 填写 `app.py`。
5. 在 Advanced settings 中选择 Python 3.11。
6. 点击 **Deploy**。

本系统当前不依赖数据库和服务器端密钥。教师填写的数据仅保存在当前浏览器会话中，除非用户主动下载 Excel 或 JSON 暂存文件。

## 项目结构

```text
.
├── app.py                 # Streamlit 应用入口
├── excel_export.py        # 云端兼容的 Excel 导出模块
├── requirements.txt       # Python 依赖
├── assets/                # 学院 Logo 与网站图标
├── .streamlit/config.toml # Streamlit 主题配置
└── 教学工作量类别与计算标准清单.md
```
