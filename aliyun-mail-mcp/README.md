# 阿里邮箱 MCP 服务器

让 Claude Code 通过 IMAP/SMTP 读取、搜索和发送阿里邮箱邮件。

提供 4 个工具：`list_folders`、`search_emails`、`read_email`、`send_email`。

## 1. 生成三方客户端安全密码
阿里邮箱网页版 → 设置 → 账户与安全 → 账户安全 → 三方客户端登录安全管理 → **生成新密码**。
用途可以填 `Claude`。密码只显示一次，请先复制保存。

## 2. 确认 IMAP/SMTP 已开启
设置 → 邮箱设置 → POP3/SMTP/IMAP，勾选 IMAP 和 SMTP 服务并保存。

## 3. 安装（在你自己的电脑上）
```bash
git clone <本仓库地址>
cd <仓库目录>/aliyun-mail-mcp
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt     # Windows: .venv\Scripts\pip
```

## 4. 注册到 Claude Code
```bash
claude mcp add aliyun-mail --scope user \
  -e ALIMAIL_USER=you@yourdomain.com \
  -e ALIMAIL_PASSWORD='第1步生成的安全密码' \
  -- "$PWD/.venv/bin/python" "$PWD/server.py"
```
- 服务器地址按邮箱域名自动选择：`@aliyun.com` 用 `imap.aliyun.com` / `smtp.aliyun.com`，其他域名用企业版 `imap.qiye.aliyun.com` / `smtp.qiye.aliyun.com`。
- 如需指定，可加 `-e ALIMAIL_IMAP_HOST=... -e ALIMAIL_SMTP_HOST=...`。

## 5. 验证
运行 `claude`，输入 `/mcp`，看到 `aliyun-mail` 显示 connected 即可。然后可以直接说：
- “列出我阿里邮箱最近 10 封未读邮件”
- “搜索主题包含‘合同’的邮件，总结第一封”
- “帮我回复这封邮件，说周五前给答复”（发送前 Claude 会请你确认）

## 常见问题
| 现象 | 处理 |
|---|---|
| `LOGIN failed` / 认证失败 | 用的是登录密码而不是安全密码；或 IMAP 未开启 |
| 连接超时 | 检查网络或防火墙是否允许 993 / 465 端口 |
| 企业邮箱仍然登录失败 | 请管理员确认允许三方客户端使用 IMAP/SMTP |

安全密码只保存在本机的 Claude Code 配置里，不要写进代码或提交到 Git。如果泄露，在第 1 步的列表里删掉它，再生成一个新的。
