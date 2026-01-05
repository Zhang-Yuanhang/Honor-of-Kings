# ========================================
# 云服务器部署指南
# ========================================

## 推荐云服务商

1. **阿里云** - 学生优惠 ¥9.9/月
2. **腾讯云** - 轻量服务器 ¥45/月起
3. **华为云** - 新用户优惠

## 部署步骤

### 1. 购买服务器
- 选择 Ubuntu 20.04/22.04 系统
- 配置：1核2G内存足够
- 开放端口：22(SSH), 80(HTTP), 5000(Flask)

### 2. 连接服务器
```bash
ssh root@你的服务器IP
```

### 3. 安装环境
```bash
# 更新系统
apt update && apt upgrade -y

# 安装 Python 和 pip
apt install python3 python3-pip python3-venv -y

# 创建项目目录
mkdir -p /opt/hok-analyzer
cd /opt/hok-analyzer
```

### 4. 上传代码
在本地执行：
```bash
# 使用 scp 上传文件
scp match_analyzer.py app.py requirements.txt root@你的服务器IP:/opt/hok-analyzer/
scp -r templates root@你的服务器IP:/opt/hok-analyzer/
scp "内战计分表 - 2026.xlsx" root@你的服务器IP:/opt/hok-analyzer/
```

### 5. 服务器端配置
```bash
cd /opt/hok-analyzer

# 创建虚拟环境
python3 -m venv venv
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
pip install gunicorn

# 修改 app.py 中的数据文件路径
# DATA_FILE = '/opt/hok-analyzer/内战计分表 - 2026.xlsx'
```

### 6. 使用 Gunicorn 运行（生产环境）
```bash
# 前台运行测试
gunicorn -w 2 -b 0.0.0.0:5000 app:app

# 后台运行
nohup gunicorn -w 2 -b 0.0.0.0:5000 app:app > app.log 2>&1 &
```

### 7. 配置 Nginx（可选，用于域名访问）
```bash
apt install nginx -y

# 创建配置文件
cat > /etc/nginx/sites-available/hok-analyzer << 'EOF'
server {
    listen 80;
    server_name 你的域名或IP;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
EOF

# 启用配置
ln -s /etc/nginx/sites-available/hok-analyzer /etc/nginx/sites-enabled/
nginx -t
systemctl restart nginx
```

### 8. 配置 Systemd 服务（开机自启）
```bash
cat > /etc/systemd/system/hok-analyzer.service << 'EOF'
[Unit]
Description=HOK Analyzer Web Service
After=network.target

[Service]
User=root
WorkingDirectory=/opt/hok-analyzer
Environment="PATH=/opt/hok-analyzer/venv/bin"
ExecStart=/opt/hok-analyzer/venv/bin/gunicorn -w 2 -b 127.0.0.1:5000 app:app
Restart=always

[Install]
WantedBy=multi-user.target
EOF

# 启动服务
systemctl daemon-reload
systemctl enable hok-analyzer
systemctl start hok-analyzer

# 查看状态
systemctl status hok-analyzer
```

## 访问地址

部署完成后，访问：
- http://你的服务器IP:5000 （直接访问）
- http://你的服务器IP （如果配置了Nginx）
- http://你的域名 （如果绑定了域名）

## 更新数据

上传新的Excel文件后，刷新缓存：
```bash
curl http://localhost:5000/api/refresh
```

或者重启服务：
```bash
systemctl restart hok-analyzer
```
