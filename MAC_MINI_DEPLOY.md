# Mac mini 自托管部署

这个项目可以直接跑在你的 Mac mini Docker 里。服务默认监听 `7860` 端口。

## 1. 准备环境变量

复制模板：

```bash
cp .env.example .env
```

编辑 `.env`，填入 Supabase 配置：

```text
SUPABASE_URL=你的 Supabase URL
SUPABASE_SERVICE_ROLE_KEY=你的 service role key
SECRET_KEY=随便生成一串很长的随机字符串
ADMIN_EMAILS=2818242447@qq.com
```

`ADMIN_EMAILS` 可以写多个管理员邮箱，用英文逗号分隔。

## 2. 启动服务

```bash
docker compose up -d --build
```

本机访问：

```text
http://localhost:7860
```

局域网访问：

```text
http://你的Mac mini局域网IP:7860
```

查看日志：

```bash
docker compose logs -f
```

停止服务：

```bash
docker compose down
```

更新代码后重启：

```bash
git pull
docker compose up -d --build
```

## 3. 让外网访问

### 方案 A：路由器端口映射

如果你的宽带有公网 IP，可以在路由器里把外网端口转发到 Mac mini：

```text
外网端口 7860 -> Mac mini 局域网 IP:7860
```

然后别人访问：

```text
http://你的公网IP:7860
```

如果有域名，可以把域名解析到公网 IP。

### 方案 B：内网穿透

如果没有公网 IP，用内网穿透更省事。可以选：

- cpolar
- natapp
- Sakura Frp
- 花生壳

穿透目标填：

```text
127.0.0.1:7860
```

穿透平台会给你一个公网 URL，别人打开那个 URL 就能访问。

## 4. 注意

- Mac mini 要保持开机，Docker Desktop 要保持运行。
- 如果用免费内网穿透，公网地址可能会变。
- 不要把 `.env` 发给别人，里面有数据库密钥。
