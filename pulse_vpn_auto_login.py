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
                "clash_processes": ["clash.exe", "Clash for Windows.exe"],
                "pulse_processes": ["Pulse.exe", "JamUI.exe"],
                "wait_timeout": 30,
                "retry_count": 3,
                "window_titles": {
                    "main": "Pulse Secure",
                    "login": "登录",
                    "connect": "连接"
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
        
        time.sleep(3)
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
                time.sleep(5)
                return True
            except Exception as e:
                logger.error(f"启动Pulse VPN失败: {e}")
                return False
        else:
            logger.error(f"Pulse VPN路径不存在: {self.settings['pulse_vpn_path']}")
            return False
    
    def find_and_click_button(self, button_text, timeout=30):
        """查找并点击按钮"""
        logger.info(f"正在查找按钮: {button_text}")
        
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # 查找窗口
                windows = gw.getWindowsWithTitle(self.settings["window_titles"]["main"])
                if windows:
                    window = windows[0]
                    window.activate()
                    time.sleep(1)
                    
                    # 使用键盘快捷键找到按钮
                    pyautogui.hotkey('alt', 'c')  # 尝试连接快捷键
                    time.sleep(2)
                    return True
                    
            except Exception as e:
                logger.debug(f"查找按钮时出错: {e}")
            
            time.sleep(1)
        
        logger.warning(f"未找到按钮: {button_text}")
        return False
    
    def input_credentials(self, username, password):
        """输入账号密码"""
        logger.info("正在输入账号密码...")
        
        start_time = time.time()
        while time.time() - start_time < self.settings["wait_timeout"]:
            try:
                # 查找登录窗口
                login_windows = gw.getWindowsWithTitle(self.settings["window_titles"]["login"])
                if login_windows:
                    login_window = login_windows[0]
                    login_window.activate()
                    time.sleep(1)
                    
                    # 输入用户名
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('delete')
                    pyautogui.write(username)
                    time.sleep(0.5)
                    
                    # 按Tab切换到密码框
                    pyautogui.press('tab')
                    time.sleep(0.5)
                    
                    # 输入密码
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('delete')
                    pyautogui.write(password)
                    time.sleep(0.5)
                    
                    # 按回车登录
                    pyautogui.press('enter')
                    logger.info("账号密码已输入，正在登录...")
                    return True
                    
            except Exception as e:
                logger.debug(f"输入凭据时出错: {e}")
            
            time.sleep(1)
        
        logger.error("未找到登录窗口")
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
            
            # 5. 点击连接按钮
            time.sleep(5)
            if not self.find_and_click_button("连接"):
                logger.error("无法找到连接按钮")
                return False
            
            # 6. 输入账号密码
            time.sleep(3)
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
    
    if vpn_automation.run():
        print("\n✅ 登录流程已完成！")
        print("脚本将在5秒后自动退出...")
        time.sleep(5)
    else:
        print("\n❌ 登录流程失败！")
        print("请查看日志文件获取详细信息")
        input("按回车键退出...")

if __name__ == "__main__":
    main()