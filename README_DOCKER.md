# 物品ID映射工具 - Docker 部署

本项目提供一个基于 Flask 的网页工具，可将 Excel 中形如 `物品=ID$数量&…` 或 `220-221-…$1$80` 的文本按映射表（A→B）替换为“名称$数量…”，支持 `&`、`|` 分段与 `-` 组合 ID。已提供 Docker 化部署与持久化存储。

## 1. 快速开始（docker compose）

```bash
docker compose up -d
# 首次会构建镜像，完成后访问 http://<服务器IP或域名>:8000
```

- 数据目录：`./data` 会挂载为容器内 `/app/data`
- 上传文件保存到 `/app/data/uploads`
- 输出文件保存到 `/app/data/outputs`
- 管理面板：`/admin`

## 2. 直接使用 Docker 运行

```bash
docker build -t item-map-web .
docker run -d --name item-map \
  -p 8000:8000 \
  -e FLASK_SECRET_KEY=change-me \
  -e ITEMMAP_DATA_DIR=/app/data \
  -v $(pwd)/data:/app/data \
  item-map-web
```

## 3. 环境变量

- `FLASK_SECRET_KEY`: Flask 会话密钥，务必修改为强随机字符串
- `ITEMMAP_DATA_DIR`: 持久化数据目录（容器内），默认 `/app/data`

## 4. 目录说明

- `webtool/app.py`: Flask 应用，上传/转换/下载逻辑，`/admin` 列表管理
- `webtool/templates/index.html`: 上传与参数表单
- `webtool/templates/admin.html`: 管理面板
- `wsgi.py`: WSGI 入口，Gunicorn 从此启动
- `requirements.txt`: 运行依赖
- `Dockerfile`: 构建镜像
- `docker-compose.yml`: 一键运行

## 5. Nginx 反向代理（可选）

```
server {
  listen 80;
  server_name your-domain.com;

  location / {
    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
    proxy_pass http://127.0.0.1:8000;
  }
}
```

配合 Certbot 配置 HTTPS。

## 6. 常见问题

- `.xls` 读写：使用 `xlrd==1.2.0` 读取，输出统一为 `.xlsx`
- 未命中 ID：可在 `/admin` 下载输出检查；如需“未命中ID报表”，可提出需求后续增强
- 前缀与分隔：支持是否保留“物品=”，支持 `&` 与 `|` 分段，`-` 组内逐个映射

## 7. 下一步增强（建议）

- 未命中 ID 统计导出（CSV）
- 登录鉴权与多用户
- 批处理与任务历史（元数据）
- 自定义分隔符与高级规则配置
