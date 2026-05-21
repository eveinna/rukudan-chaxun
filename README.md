# 入库单查询系统 - Web版

基于 Streamlit 开发的入库单查询系统，支持多用户登录、数据筛选、统计和导出功能。

## 功能特点

- 🔐 用户登录验证
- 📅 日期范围筛选（支持快捷选项）
- 🔍 多条件筛选（供应商、采购员、商品、发货地）
- 📊 实时统计（总金额、总数量、总重量）
- 📥 数据导出（Excel/CSV）
- ☁️ 免费部署到 Streamlit Cloud

## 快速开始

### 本地运行

1. 安装依赖：
```bash
pip install -r requirements.txt
```

2. 运行应用：
```bash
streamlit run app.py
```

3. 浏览器访问：http://localhost:8501

### 默认账号

- 管理员：admin / admin123
- 用户1：user1 / 123456
- 用户2：user2 / 123456

## 部署到 Streamlit Cloud（免费）

### 步骤1：准备 GitHub 仓库

1. 在 GitHub 创建新仓库（如：`inventory-query-web`）
2. 将本文件夹内容上传到仓库：
   - `app.py`
   - `requirements.txt`
   - `README.md`
   - 数据文件 `入库单数据.xlsx`（如有）

### 步骤2：部署到 Streamlit Cloud

1. 访问 [share.streamlit.io](https://share.streamlit.io)
2. 使用 GitHub 账号登录
3. 点击 "New app"
4. 选择仓库和分支
5. 主文件路径填 `app.py`
6. 点击 "Deploy"

### 步骤3：配置（可选）

如果需要修改用户账号密码：
1. 在 Streamlit Cloud 进入 App 设置
2. 添加 Secrets：
```toml
[users]
admin = "your_password"
user1 = "password1"
```

然后修改 `app.py` 中的用户验证逻辑，从 secrets 读取。

## 数据文件

将 Excel 数据文件命名为 `入库单数据.xlsx`，放在与 `app.py` 同目录下。

支持的列名：
- 日期、供应商、商品名称、数量、重量、单价、箱费、金额、采购员、发货地

## 技术栈

- **前端框架**：Streamlit
- **数据处理**：Pandas
- **Excel 支持**：openpyxl

## 许可证

MIT License
