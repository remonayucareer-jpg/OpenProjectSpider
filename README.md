# OpenProject AI运营工单导出

这是一个用于抓取 OpenProject AI运营工单并导出 Excel 的 Streamlit 应用。

## 本地运行

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Streamlit Cloud 部署

1. 将本目录推送到 GitHub。
2. 在 Streamlit Cloud 新建应用，入口文件选择 `streamlit_app.py`。
3. 在应用的 Secrets 中配置 OpenProject API Key：

```toml
OPENPROJECT_API_KEY = "你的 OpenProject API Key"
```

也可以直接配置完整鉴权头：

```toml
OPENPROJECT_AUTH_HEADER = "Basic xxxxx"
```

## 导出字段

工单类别、工单编号、机器人编号、酒店ID、所属集团、酒店名称、非酒店名称、具体需求、需求数量、问题方、AI应用场景、需求/问题、具体需求/问题、工单创建时间、提交人、工单开始处理时间、工单解决时间、工单状态
