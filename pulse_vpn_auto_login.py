#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pulse VPN 自动登录脚本
自动完成Pulse VPN的登录流程：
1. 关闭现有Pulse VPN和Clash进程
2. 禁用系统代理
3. 启动Pulse VPN
4. 自动输入账号密码登录
"""

import os
import sys
import time
import json
import logging
import subprocess
import psutil
import pyautogui
import pygetwindow as gw
import winreg
from cryptography.fernet import Fernet
import getpass
import keyboard

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/pulse_vpn_login.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class PulseVPNAutoLogin:
    def __init__(self):
        self.config_dir = "config"
        self.settings_file = os.path.join(self.config_dir, "settings.json")
        self.credentials_file = os.path.join(self.config_dir, "credentials.enc")
        self.settings = self.load_settings()
        self.cipher_suite = None
        
    def load_settings(self):
        """加载配置文件"""
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
            
        if not os.path.exists(self.settings_file):
            default_settings = {
                "pulse_vpn_path": "C:\\Program Files (x86)\\Common Files\\Pulse Secure\\JamUI\\Pulse.exe",
                "clash_path": "C:\\Users\\Admin\\Softwares\\Clash\\clash\\Clash for Windows.exe",
                "clash_processes": ["clash.exe", "Clash for Windows.exe"],
                "pulse_processes": ["Pulse.exe", "JamUI.exe"],
                "wait_timeout": 30,
                "retry_count": 3,
                "window_titles": {
                    "main": "Pulse Secure",
                    "login": "登录",
                    "connect": "连接"
                },
                "login_positions": {
                    "userid_x": 0.2962,
                    "userid_y": 0.3517,
                    "password_x": 0.2962,
                    "password_y": 0.3947
                }
            }
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(default_settings, f, ensure_ascii=False, indent=4)
            return default_settings
        else:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    
    def close_existing_processes(self):
        """关闭现有Pulse VPN和Clash进程"""
        logger.info("正在关闭现有进程...")
        
        # 关闭Clash相关进程
        for process_name in self.settings["clash_processes"]:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if process_name.lower() in proc.info['name'].lower():
                        logger.info(f"正在关闭进程: {proc.info['name']} (PID: {proc.info['pid']})")
                        subprocess.run(['taskkill', '/F', '/PID', str(proc.info['pid'])], 
                                     capture_output=True)
                        time.sleep(1)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        
        # 关闭Pulse VPN相关进程
        for process_name in self.settings["pulse_processes"]:
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if process_name.lower() in proc.info['name'].lower():
                        logger.info(f"正在关闭进程: {proc.info['name']} (PID: {proc.info['pid']})")
                        subprocess.run(['taskkill', '/F', '/PID', str(proc.info['pid'])], 
                                     capture_output=True)
                        time.sleep(1)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        
        time.sleep(1)
        logger.info("进程清理完成")
    
    def disable_proxy(self):
        """禁用Windows系统代理"""
        logger.info("正在禁用系统代理...")
        try:
            # 修改Internet Settings注册表
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                               r'Software\Microsoft\Windows\CurrentVersion\Internet Settings',
                               0, winreg.KEY_SET_VALUE)
            
            # 禁用代理
            winreg.SetValueEx(key, 'ProxyEnable', 0, winreg.REG_DWORD, 0)
            winreg.SetValueEx(key, 'ProxyServer', 0, winreg.REG_SZ, '')
            winreg.CloseKey(key)
            
            # 刷新系统代理设置
            os.system('ipconfig /flushdns')
            
            logger.info("系统代理已禁用")
        except Exception as e:
            logger.error(f"禁用代理失败: {e}")
    
    def get_credentials(self):
        """获取或输入用户凭据"""
        if os.path.exists(self.credentials_file):
            return self.load_credentials()
        else:
            print("首次运行，请输入VPN凭据：")
            username = input("用户名: ").strip()
            password = getpass.getpass("密码: ").strip()
            
            # 加密并保存凭据
            self.save_credentials(username, password)
            return username, password
    
    def save_credentials(self, username, password):
        """加密保存用户凭据"""
        key = Fernet.generate_key()
        self.cipher_suite = Fernet(key)
        
        credentials = {"username": username, "password": password}
        encrypted_data = self.cipher_suite.encrypt(
            json.dumps(credentials).encode()
        )
        
        with open(self.credentials_file, 'wb') as f:
            f.write(key + b'\n' + encrypted_data)
    
    def load_credentials(self):
        """加载并解密用户凭据"""
        try:
            with open(self.credentials_file, 'rb') as f:
                key = f.readline().strip()
                encrypted_data = f.read()
            
            self.cipher_suite = Fernet(key)
            decrypted_data = self.cipher_suite.decrypt(encrypted_data)
            credentials = json.loads(decrypted_data.decode())
            
            return credentials["username"], credentials["password"]
        except Exception as e:
            logger.error(f"加载凭据失败: {e}")
            # 删除损坏的凭据文件，重新输入
            if os.path.exists(self.credentials_file):
                os.remove(self.credentials_file)
            return self.get_credentials()
    
    def launch_pulse_vpn(self):
        """启动Pulse VPN客户端"""
        logger.info("正在启动Pulse VPN...")
        
        if os.path.exists(self.settings["pulse_vpn_path"]):
            try:
                subprocess.Popen([self.settings["pulse_vpn_path"]])
                logger.info("Pulse VPN已启动")
                time.sleep(1)
                return True
            except Exception as e:
                logger.error(f"启动Pulse VPN失败: {e}")
                return False
        else:
            logger.error(f"Pulse VPN路径不存在: {self.settings['pulse_vpn_path']}")
            return False
    
    def trigger_connection(self, timeout=30):
        """触发VPN连接"""
        logger.info("正在触发VPN连接...")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # 查找Pulse窗口
                all_windows = gw.getAllWindows()
                pulse_windows = []
                
                for window in all_windows:
                    if window.title and ("pulse" in window.title.lower() or 
                                       "vpn" in window.title.lower() or
                                       "secure" in window.title.lower()):
                        pulse_windows.append(window)
                
                if pulse_windows:
                    # 选择第一个可见窗口
                    target_window = None
                    for window in pulse_windows:
                        if window.visible and window.width > 0 and window.height > 0:
                            target_window = window
                            break
                    
                    if not target_window:
                        target_window = pulse_windows[0]
                    
                    logger.info(f"激活窗口: {target_window.title}")
                    target_window.activate()
                    time.sleep(1)
                    
                    # 切换到英文输入法
                    pyautogui.hotkey('shift', 'alt')  # 中英文切换
                    time.sleep(0.5)
                    
                    # 使用f o c快捷键触发连接
                    pyautogui.press('f')
                    time.sleep(0.2)
                    pyautogui.press('o')
                    time.sleep(0.2)
                    pyautogui.press('c')
                    
                    logger.info("已发送f o c快捷键，等待登录界面...")
                    return True
                    
            except Exception as e:
                logger.debug(f"触发连接时出错: {e}")
            
            time.sleep(1)
        
        logger.warning("无法触发VPN连接")
        return False
    
    def find_login_window(self):
        """查找登录窗口"""
        # 首先尝试获取当前激活窗口
        try:
            active_window = gw.getActiveWindow()
            if active_window and active_window.visible and active_window.title:
                title_lower = active_window.title.lower()
                # 检查激活窗口是否为登录相关窗口
                if (any(keyword in title_lower for keyword in ['login', '登录', 'sign', 'auth', 'credential']) or
                    (active_window.width < 800 and active_window.height < 600 and 
                     active_window.width > 200 and active_window.height > 150)):
                    logger.info(f"使用当前激活窗口作为登录窗口: {active_window.title}")
                    return active_window
        except Exception as e:
            logger.debug(f"获取激活窗口失败: {e}")
        
        # 如果激活窗口不合适，查找所有可能的登录窗口
        all_windows = gw.getAllWindows()
        login_candidates = []
        
        for window in all_windows:
            if not window.title or not window.visible:
                continue
                
            title_lower = window.title.lower()
            # 查找包含登录相关关键词的窗口
            if any(keyword in title_lower for keyword in ['login', '登录', 'sign', 'auth', 'credential']):
                login_candidates.append(window)
                continue
                
            # 查找较小的弹窗（登录窗口通常比主窗口小）
            if window.width < 800 and window.height < 600 and window.width > 200 and window.height > 150:
                login_candidates.append(window)
        
        if not login_candidates:
            return None
        
        # 优先选择标题明确包含登录的窗口
        for window in login_candidates:
            title_lower = window.title.lower()
            if any(keyword in title_lower for keyword in ['login', '登录', 'sign in']):
                return window
        
        # 选择最前端的窗口（模拟最顶层）
        # 由于pygetwindow没有直接的Z-order信息，我们选择最小的窗口作为近似
        return min(login_candidates, key=lambda w: w.width * w.height)
    
    def find_input_field_by_label(self, window, label_text):
        """通过标签文本查找输入框位置"""
        try:
            # 获取窗口截图区域
            screenshot = pyautogui.screenshot(region=(window.left, window.top, window.width, window.height))
            
            # 尝试查找标签文本
            for text in label_text:
                try:
                    label_location = pyautogui.locateOnScreen(text, region=(window.left, window.top, window.width, window.height), confidence=0.7)
                    if label_location:
                        # 假设输入框在标签右侧或下方
                        label_center = pyautogui.center(label_location)
                        # 尝试右侧
                        input_x = label_center.x + 100
                        input_y = label_center.y
                        if input_x < window.left + window.width:
                            return (input_x, input_y)
                        # 尝试下方
                        input_x = label_center.x
                        input_y = label_center.y + 30
                        if input_y < window.top + window.height:
                            return (input_x, input_y)
                except:
                    continue
        except Exception as e:
            logger.debug(f"图像识别失败: {e}")
        
        return None
    
    def find_input_fields_by_position(self, window):
        """通过位置推测输入框"""
        # 假设用户名框在窗口上半部分，密码框在下半部分
        username_x = window.left + window.width // 2
        username_y = window.top + window.height // 3
        
        password_x = window.left + window.width // 2
        password_y = window.top + window.height * 2 // 3
        
        return (username_x, username_y), (password_x, password_y)
    
    def get_login_input_positions(self):
        """获取登录界面输入框的固定屏幕坐标"""
        # 获取屏幕尺寸
        screen_width, screen_height = pyautogui.size()
        
        # 从配置文件读取相对位置
        login_pos = self.settings.get("login_positions", {})
        userid_rel_x = login_pos.get("userid_x", 0.3)
        userid_rel_y = login_pos.get("userid_y", 0.2)
        password_rel_x = login_pos.get("password_x", 0.3)
        password_rel_y = login_pos.get("password_y", 0.22)
        
        # 计算实际像素坐标
        userid_x = int(screen_width * userid_rel_x)
        userid_y = int(screen_height * userid_rel_y)
        password_x = int(screen_width * password_rel_x)
        password_y = int(screen_height * password_rel_y)
        
        logger.info(f"屏幕尺寸: {screen_width}x{screen_height}")
        logger.info(f"UserID坐标: ({userid_x}, {userid_y}) - 相对位置({userid_rel_x}, {userid_rel_y})")
        logger.info(f"密码坐标: ({password_x}, {password_y}) - 相对位置({password_rel_x}, {password_rel_y})")
        
        return (userid_x, userid_y), (password_x, password_y)
    
    def input_credentials(self, username, password):
        """输入账号密码"""
        logger.info("正在输入账号密码...")
        
        # 等待登录界面出现
        time.sleep(1)
        
        try:
            # 获取输入框的固定屏幕坐标
            userid_pos, password_pos = self.get_login_input_positions()
            
            # 点击UserID框
            logger.info(f"点击UserID输入框坐标: ({userid_pos[0]}, {userid_pos[1]})")
            pyautogui.click(userid_pos[0], userid_pos[1])
            time.sleep(0.5)
            
            # 切换到英文输入法
            logger.info("切换到英文输入法")
            pyautogui.press('shift')  # 输入法切换
            time.sleep(0.5)
            
            # 输入用户名
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('delete')
            pyautogui.write(username)
            time.sleep(0.5)
            
            # 再次确保英文输入法
            pyautogui.press('shift')
            time.sleep(0.2)
            
            # 点击密码框并输入密码
            logger.info(f"点击密码输入框坐标: ({password_pos[0]}, {password_pos[1]})")
            pyautogui.click(password_pos[0], password_pos[1])
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('delete')
            pyautogui.write(password)
            time.sleep(0.5)
            
            # 按回车登录
            pyautogui.press('enter')
            logger.info("账号密码已输入，正在登录...")
            return True
            
        except Exception as e:
            logger.error(f"输入凭据时出错: {e}")
            return False
    
    def restart_clash(self):
        """重新启动Clash for Windows"""
        logger.info("正在重新启动Clash for Windows...")
        
        # 优先使用配置文件中的路径
        clash_paths = []
        if "clash_path" in self.settings and self.settings["clash_path"]:
            clash_paths.append(self.settings["clash_path"])
        
        # 添加常见的Clash安装路径作为备选
        clash_paths.extend([
            os.path.expanduser("~\\Softwares\\Clash\\clash\\Clash for Windows.exe"),  # 动态用户路径
            os.path.expanduser("~\\AppData\\Local\\Clash for Windows\\Clash for Windows.exe"),
            "C:\\Program Files\\Clash for Windows\\Clash for Windows.exe",
            "C:\\Program Files (x86)\\Clash for Windows\\Clash for Windows.exe",
            os.path.expanduser("~\\Desktop\\Clash for Windows.exe"),
        ])
        
        # 尝试从注册表或快捷方式找到Clash路径
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\Applications\Clash for Windows.exe\shell\open\command")
            clash_cmd, _ = winreg.QueryValueEx(key, "")
            if clash_cmd:
                # 提取exe路径（去除参数）
                clash_path = clash_cmd.split('"')[1] if '"' in clash_cmd else clash_cmd.split()[0]
                clash_paths.insert(0, clash_path)
            winreg.CloseKey(key)
        except:
            pass
        
        # 尝试启动Clash
        for path in clash_paths:
            if os.path.exists(path):
                try:
                    subprocess.Popen([path], shell=True)
                    logger.info(f"Clash for Windows已启动: {path}")
                    return True
                except Exception as e:
                    logger.debug(f"启动Clash失败 {path}: {e}")
                    continue
        
        logger.warning("未找到Clash for Windows，请手动启动")
        return False
    
    def run(self):
        """运行完整流程"""
        try:
            logger.info("开始Pulse VPN自动登录流程")
            
            # 1. 关闭现有进程
            self.close_existing_processes()
            
            # 2. 禁用代理
            self.disable_proxy()
            
            # 3. 获取凭据
            username, password = self.get_credentials()
            
            # 4. 启动Pulse VPN
            if not self.launch_pulse_vpn():
                logger.error("无法启动Pulse VPN")
                return False
            
            # 5. 触发VPN连接
            time.sleep(1)
            if not self.trigger_connection():
                logger.error("无法触发VPN连接")
                return False
            
            # 6. 输入账号密码
            time.sleep(1)
            if not self.input_credentials(username, password):
                logger.error("无法输入账号密码")
                return False
            
            logger.info("Pulse VPN自动登录完成")
            return True
            
        except KeyboardInterrupt:
            logger.info("用户中断操作")
            return False
        except Exception as e:
            logger.error(f"运行过程中出错: {e}")
            return False

def main():
    """主函数"""
    if not os.path.exists("logs"):
        os.makedirs("logs")
    
    print("=== Pulse VPN 自动登录脚本 ===")
    print("按 Ctrl+C 可随时中断")
    
    vpn_automation = PulseVPNAutoLogin()
    
    try:
        if vpn_automation.run():
            print("\n✅ 登录流程已完成！")
            print("脚本将在5秒后自动退出...")
            time.sleep(5)
        else:
            print("\n❌ 登录流程失败！")
            print("请查看日志文件获取详细信息")
            input("按回车键退出...")
    finally:
        # 退出时重新启动Clash for Windows
        print("\n正在重新启动Clash for Windows...")
        vpn_automation.restart_clash()
        time.sleep(2)

if __name__ == "__main__":
    main()