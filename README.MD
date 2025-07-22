# Pulse VPN 自动登录脚本

一个用于自动化Pulse VPN登录过程的Python脚本，支持关闭现有进程、禁用代理、自动输入账号密码完成登录。

## 功能特性

- ✅ 自动关闭Pulse VPN和Clash for Windows进程
- ✅ 禁用Windows系统代理设置
- ✅ 自动启动Pulse VPN客户端
- ✅ 使用固定屏幕坐标精确点击输入框
- ✅ 智能输入法切换（自动切换到英文输入）
- ✅ 完整的错误处理和日志记录
- ✅ 安全加密存储用户凭据
- ✅ 退出时自动重启Clash for Windows

## 环境要求

- Windows 10/11 操作系统
- Python 3.7+ 
- Pulse VPN 客户端已安装
- Clash for Windows（可选）

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
1. 关闭正在运行的Pulse VPN和Clash进程
2. 禁用系统代理设置
3. 启动Pulse VPN客户端
4. 等待界面加载完成
5. 使用"f o c"快捷键触发连接
6. 识别登录界面并自动填写账号密码
7. 完成后自动重启Clash for Windows

## 配置文件

配置文件位于：`config/settings.json`

### 主要配置项：

```json
{
  "pulse_vpn_path": "Pulse VPN安装路径",
  "clash_path": "Clash for Windows安装路径",
  "wait_timeout": 30,
  "login_positions": {
    "userid_x": 0.2962,
    "userid_y": 0.3517,
    "password_x": 0.2962,
    "password_y": 0.3947
  }
}
```

- `pulse_vpn_path`: Pulse VPN的安装路径
- `clash_path`: Clash for Windows的安装路径
- `login_positions`: 登录界面输入框的相对屏幕坐标
- `wait_timeout`: 等待超时时间（秒）

## 特色功能

### 智能输入法切换
- 自动在输入用户名和密码前切换到英文输入法
- 使用简单的Shift键切换，兼容性更好

### 精确坐标定位
- 使用相对屏幕坐标定位输入框
- 适应不同分辨率的屏幕
- 可在配置文件中自定义坐标

### 进程管理
- 智能识别并关闭相关进程
- 退出时自动重启必要的应用程序

## 日志文件

运行日志保存在：`logs/pulse_vpn_login.log`
包含详细的执行过程和错误信息，便于问题排查。

## 常见问题

### Q: 脚本无法识别Pulse VPN窗口？
A: 确保Pulse VPN已正确安装，并检查配置文件中的安装路径。

### Q: 登录界面坐标不准确？
A: 修改`config/settings.json`中的`login_positions`参数，调整输入框坐标。

### Q: 如何更新账号密码？
A: 删除配置文件`config/credentials.enc`，重新运行脚本即可重新输入。

### Q: Clash for Windows无法自动启动？
A: 检查`config/settings.json`中的`clash_path`路径是否正确。

## 注意事项

- 运行脚本前请保存所有工作，避免意外关闭应用程序
- 确保Pulse VPN客户端版本兼容性
- 建议以管理员权限运行脚本
- 首次运行可能需要调整登录界面坐标

## 技术支持

如有问题，请查看日志文件或提交Issue。

## 免责声明

本脚本仅供学习和合法使用，请确保你有使用VPN的合法权限。