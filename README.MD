# Pulse VPN 自动登录脚本

一个用于自动化Pulse VPN登录过程的Python脚本，支持关闭现有进程、禁用代理、自动输入账号密码完成登录。

## 功能特性

- ✅ 自动关闭Pulse VPN和Clash for Windows
- ✅ 禁用Windows系统代理设置
- ✅ 自动启动Pulse VPN客户端
- ✅ 智能识别登录界面并自动填写账号密码
- ✅ 完整的错误处理和日志记录
- ✅ 安全加密存储用户凭据

## 环境要求

- Windows 10/11 操作系统
- Python 3.7+ 
- Pulse VPN 客户端已安装

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 首次运行
```bash
python pulse_vpn_auto_login.py
```

### 2. 输入凭据
首次运行时会提示输入：
- VPN账号（用户名）
- VPN密码

凭据将安全加密存储在本地，后续运行无需重复输入。

### 3. 自动执行
脚本将自动：
1. 关闭正在运行的Pulse VPN和Clash
2. 禁用系统代理设置
3. 启动Pulse VPN客户端
4. 等待界面加载
5. 点击"连接"按钮
6. 输入账号密码并登录

## 配置文件

配置文件位于：`config/settings.json`
- 可设置Pulse VPN安装路径
- 可调整等待超时时间
- 可启用/禁用日志记录

## 日志文件

运行日志保存在：`logs/pulse_vpn_login.log`
包含详细的执行过程和错误信息，便于问题排查。

## 常见问题

### Q: 脚本无法识别Pulse VPN窗口？
A: 确保Pulse VPN已正确安装，并检查配置文件中的安装路径。

### Q: 登录失败怎么办？
A: 检查日志文件，确认账号密码是否正确，网络是否畅通。

### Q: 如何更新账号密码？
A: 删除配置文件`config/credentials.json`，重新运行脚本即可重新输入。

## 注意事项

- 运行脚本前请保存所有工作，避免意外关闭应用程序
- 确保Pulse VPN客户端版本兼容性
- 建议以管理员权限运行脚本

## 技术支持

如有问题，请查看日志文件或提交Issue。

## 免责声明

本脚本仅供学习和合法使用，请确保你有使用VPN的合法权限。