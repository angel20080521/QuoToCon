# QuoToCon

将 Excel 格式的**报价单**自动转换为填写好的**商务合同**。

上传报价单（`.xlsx`）和合同模板（`.docx`），系统自动提取报价单字段并将其填入合同模板的 `{{占位符}}` 中，生成可直接下载的合同文件。

---

## 功能特性

- 自动从 Excel 报价单两列表格中提取字段（全角字段名，支持带或不带全角冒号）
- 支持单元格内 `字段名：字段值` 全角冒号格式（作为补充）
- 合同模板中的 `{{字段名}}` 占位符被自动替换为对应值
- 处理 Word 跨 Run 分割占位符的边缘情况
- 未匹配的占位符原样保留，方便排查
- 生成文件名自动从"报价编号 / 项目名称 / 客户名称"等字段推导
- 单页 Web 界面，支持拖拽上传

---

## 快速开始（本地开发）

### 1. 克隆项目

```bash
git clone https://github.com/<your-org>/QuoToCon.git
cd QuoToCon
```

### 2. 创建并激活虚拟环境

```bash
python3 -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 配置环境变量

```bash
cp .env.example .env
# 编辑 .env，填入 SECRET_KEY
```

### 5. 启动开发服务器

```bash
python app.py
```

访问 [http://localhost:5000](http://localhost:5000)。

---

## 使用方法

### 准备报价单

报价单（`.xlsx`）应使用两列表格格式，A列为全角字段名（可带全角冒号），B列为对应值，例如：

| A列（字段名） | B列（字段值） |
|---|---|
| `报价编号：` | `QT-2024-001` |
| `客户名称：` | `某某有限公司` |
| `项目名称：` | `办公设备采购` |
| `合计金额：` | `46,000 元` |

也支持单元格内 `字段名：字段值` 全角冒号格式。

### 准备合同模板

合同模板（`.docx`）中在需要填入数据的位置插入 `{{字段名}}` 占位符，例如：

```
甲方（买方）：{{客户名称}}
合同编号：{{报价编号}}
合同总价：{{合计金额}}
```

字段名必须与报价单中提取到的字段名**完全一致**（包括大小写和空格）。

### 生成合同

1. 打开 [http://localhost:5000](http://localhost:5000)
2. 上传报价单文件（① 报价单文件）
3. 上传合同模板文件（② 合同模板文件）
4. 点击 **生成合同** 按钮
5. 检查已提取的字段数据，确认无误后点击 **下载合同文件**

---

## 生产部署（Ubuntu + PM2 + Nginx）

### 1. 服务器环境准备

```bash
# 安装 Python 3.10+（Ubuntu 22.04 已内置）
sudo apt update && sudo apt install -y python3 python3-venv python3-pip

# 安装 Node.js 和 PM2
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs
sudo npm install -g pm2
```

### 2. 部署代码

```bash
git clone https://github.com/<your-org>/QuoToCon.git /home/ubuntu/QuoToCon
cd /home/ubuntu/QuoToCon

python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 3. 配置环境变量

```bash
export SECRET_KEY="$(python3 -c 'import secrets; print(secrets.token_hex(32))')"
```

> 建议将 `SECRET_KEY` 写入服务器的 shell profile（`~/.bashrc` 或 `/etc/environment`）以便持久化。

### 4. 用 PM2 启动

```bash
pm2 start ecosystem.config.js --env production
pm2 save
pm2 startup   # 按输出提示执行一条 sudo 命令以设置开机自启
```

常用 PM2 命令：

```bash
pm2 list                  # 查看进程状态
pm2 logs QuoToCon         # 查看实时日志
pm2 restart QuoToCon      # 重启服务
pm2 stop QuoToCon         # 停止服务
```

### 5. 配置 Nginx 反向代理（推荐）

```nginx
server {
    listen 80;
    server_name your-domain.com;

    client_max_body_size 35M;

    location / {
        proxy_pass         http://127.0.0.1:5000;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_set_header   X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_read_timeout 120s;
    }
}
```

```bash
sudo nginx -t && sudo systemctl reload nginx
```

### 6. HTTPS（使用 Let's Encrypt）

```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d your-domain.com
```

### 7. 定期清理输出文件

```bash
# 添加到 crontab（每天凌晨 3 点清理超过 24 小时的生成文件）
0 3 * * * find /home/ubuntu/QuoToCon/outputs -mtime +1 -delete
```

---

## 目录结构

```
QuoToCon/
├── app.py                  # Flask 应用（路由与文件处理逻辑）
├── ecosystem.config.js     # PM2 生产部署配置
├── requirements.txt        # Python 依赖
├── .env.example            # 环境变量示例
├── utils/
│   ├── __init__.py
│   └── docx_processor.py   # 报价单数据提取 & 模板填充核心逻辑
├── templates/
│   ├── index.html          # 上传页面
│   └── result.html         # 结果页面
├── static/
│   ├── css/style.css       # 样式表
│   └── js/main.js          # 前端交互
├── tests/
│   ├── test_docx_processor.py  # 核心逻辑单元测试
│   └── test_app.py             # Flask 路由测试
├── uploads/                # 运行时（.gitignore 忽略）
├── outputs/                # 运行时（.gitignore 忽略）
└── logs/                   # 运行时（.gitignore 忽略）
```

---

## 环境变量

| 变量名       | 必填 | 说明                                        | 示例值                     |
|------------|------|---------------------------------------------|--------------------------|
| `SECRET_KEY` | 是   | Flask session 签名密钥，生产环境必须设置      | `abc123...`（随机长字符串）  |
| `PORT`       | 否   | 监听端口，默认 `5000`                         | `8080`                   |

---

## 运行测试

```bash
source venv/bin/activate
pip install pytest
pytest tests/ -v
```

---

## 常见问题

**Q：占位符没有被替换怎么办？**  
A：检查报价单中的字段名与模板中 `{{字段名}}` 是否完全一致（大小写、空格、全半角）。结果页面的"已提取的字段数据"表格可以帮助确认实际提取到的字段名。

**Q：Word 把占位符拆成多个 Run，导致替换失败？**  
A：系统已通过合并段落内所有 Run 文本的方式处理此问题。若仍有遗漏，可尝试在 Word 中重新输入占位符（不要复制粘贴）。

**Q：生成的合同文件名乱码？**  
A：现代浏览器均支持 UTF-8 编码的文件名，如遇问题请更新浏览器。

**Q：上传文件大小限制是多少？**  
A：默认 32 MB。如需调整，修改 `app.py` 中的 `MAX_CONTENT_LENGTH`，并同步更新 Nginx 的 `client_max_body_size`。
