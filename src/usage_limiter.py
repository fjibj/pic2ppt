"""
使用限制管理器 - 增强防篡改版
1. 单实例检测
2. 每日3次转换限制
3. 加密存储 + 校验和防篡改
"""

import os
import sys
import json
import tempfile
import hashlib
import base64
from pathlib import Path
from datetime import datetime


class UsageLimiter:
    """使用限制管理器"""
    
    MAX_DAILY = 3
    
    def __init__(self):
        self.app_name = "pic2ppt"
        self.lock_file = Path(tempfile.gettempdir()) / f"{self.app_name}.lock"
        # 存储到多个位置，互相校验
        self.data_files = [
            Path.home() / ".config" / self.app_name / "data.bin",
            Path(tempfile.gettempdir()) / f".{self.app_name}_data",
            Path.home() / f".{self.app_name}_cfg",
        ]
        self._lock_handle = None
        self._key = self._generate_key()
    
    def _generate_key(self):
        """基于机器信息生成密钥"""
        # 收集机器特征
        try:
            import uuid
            mac = uuid.getnode()
            machine = str(mac)
        except:
            machine = "unknown"
        
        # 添加一些系统信息
        key_data = f"{machine}:{os.getenv('USERNAME', 'user')}:pic2ppt_secret_v1"
        return hashlib.sha256(key_data.encode()).digest()[:32]
    
    def _encrypt(self, data):
        """简单异或加密"""
        text = json.dumps(data).encode('utf-8')
        encrypted = bytearray()
        for i, b in enumerate(text):
            encrypted.append(b ^ self._key[i % len(self._key)])
        return base64.b64encode(bytes(encrypted)).decode('ascii')
    
    def _decrypt(self, encrypted_text):
        """解密"""
        try:
            data = base64.b64decode(encrypted_text)
            decrypted = bytearray()
            for i, b in enumerate(data):
                decrypted.append(b ^ self._key[i % len(self._key)])
            return json.loads(bytes(decrypted).decode('utf-8'))
        except:
            return None
    
    def _compute_hash(self, data):
        """计算数据哈希"""
        text = f"{data.get('last_date', '')}:{data.get('daily_count', 0)}:{data.get('total_count', 0)}:{self._key.hex()}"
        return hashlib.sha256(text.encode()).hexdigest()[:16]
    
    def check_single_instance(self):
        """检查单实例"""
        try:
            import msvcrt
            self._lock_handle = open(self.lock_file, 'w')
            msvcrt.locking(self._lock_handle.fileno(), msvcrt.LK_NBLCK, 1)
            self._lock_handle.write(str(os.getpid()))
            self._lock_handle.flush()
            return True, None
        except:
            return False, "程序已经在运行中，请勿重复打开"
    
    def _save_to_all_locations(self, data):
        """保存到所有位置"""
        # 添加校验和
        data['check'] = self._compute_hash(data)
        encrypted = self._encrypt(data)
        
        for path in self.data_files:
            try:
                path.parent.mkdir(parents=True, exist_ok=True)
                with open(path, 'w', encoding='ascii') as f:
                    f.write(encrypted)
            except:
                pass
    
    def _load_from_locations(self):
        """从所有位置加载，互相校验"""
        valid_data = None
        
        for path in self.data_files:
            try:
                if not path.exists():
                    continue
                with open(path, 'r', encoding='ascii') as f:
                    encrypted = f.read()
                
                data = self._decrypt(encrypted)
                if not data:
                    continue
                
                # 验证校验和
                stored_hash = data.pop('check', '')
                computed_hash = self._compute_hash(data)
                
                if stored_hash == computed_hash:
                    valid_data = data
                    break
            except:
                continue
        
        return valid_data
    
    def check_daily_limit(self):
        """检查每日限制"""
        try:
            data = self._load_from_locations() or self._create_initial_data()
            today = datetime.now().date()
            last_date_str = data.get('last_date')
            
            # 检查日期倒退
            if last_date_str:
                try:
                    last_date = datetime.strptime(last_date_str, '%Y-%m-%d').date()
                    if today < last_date:
                        return False, "检测到异常日期设置，请检查系统时间"
                except:
                    pass
            
            if last_date_str != str(today):
                return True, 0
            
            daily_count = data.get('daily_count', 0)
            if daily_count >= self.MAX_DAILY:
                return False, f"今日转换次数已达上限（{self.MAX_DAILY}次），请明天再试"
            
            return True, daily_count
            
        except Exception as e:
            return True, 0
    
    def record_conversion(self):
        """记录转换"""
        try:
            data = self._load_from_locations() or self._create_initial_data()
            today = str(datetime.now().date())
            
            if data.get('last_date') != today:
                data['daily_count'] = 0
            
            data['daily_count'] = data.get('daily_count', 0) + 1
            data['total_count'] = data.get('total_count', 0) + 1
            data['last_date'] = today
            
            self._save_to_all_locations(data)
            return True
        except:
            return False
    
    def _create_initial_data(self):
        """创建初始数据"""
        today = str(datetime.now().date())
        data = {
            'last_date': today,
            'daily_count': 0,
            'total_count': 0,
            'first_use': datetime.now().isoformat(),
        }
        self._save_to_all_locations(data)
        return data
    
    def get_remaining(self):
        """获取剩余次数"""
        ok, count = self.check_daily_limit()
        if not ok:
            return 0
        return self.MAX_DAILY - count
    
    def release_lock(self):
        """释放锁"""
        try:
            if self._lock_handle:
                import msvcrt
                msvcrt.locking(self._lock_handle.fileno(), msvcrt.LK_UNLCK, 1)
                self._lock_handle.close()
        except:
            pass
        try:
            self.lock_file.unlink(missing_ok=True)
        except:
            pass
