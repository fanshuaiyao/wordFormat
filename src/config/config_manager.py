import json
from pathlib import Path
from typing import Dict, Any

class ConfigManager:
    def __init__(self):
        self.config_file = Path.home() / '.w0rdF0rmat' / 'config.json'
        self.config = self._load_config()

    def _load_config(self):
        """加载配置文件，先加载内置默认配置，再用用户配置覆盖"""
        # 加载内置默认配置
        default_config = {}
        default_config_path = Path(__file__).parent / "config.json"
        try:
            if default_config_path.exists():
                with open(default_config_path, 'r', encoding='utf-8') as f:
                    default_config = json.load(f)
        except Exception as e:
            print(f"加载默认配置失败: {str(e)}")

        # 加载用户配置并覆盖默认值
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                # 深度合并：用户配置覆盖默认配置
                for key, value in user_config.items():
                    if isinstance(value, dict) and isinstance(default_config.get(key), dict):
                        default_config[key].update(value)
                    else:
                        default_config[key] = value
        except Exception as e:
            print(f"加载用户配置文件失败: {str(e)}")

        return default_config

    def _save_config(self):
        """保存配置到文件"""
        try:
            # 确保配置目录存在
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            
            # 保存配置
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
                
        except Exception as e:
            print(f"保存配置文件失败: {str(e)}")

    def get(self, key, default=None):
        """获取配置值"""
        return self.config.get(key, default)

    def set(self, key, value):
        """设置配置值"""
        self.config[key] = value
        self._save_config()

    def get_format_presets(self):
        """获取格式预设"""
        return self.get('format_presets', {})

    def save_format_preset(self, name, preset):
        """保存格式预设"""
        presets = self.get_format_presets()
        presets[name] = preset
        self.set('format_presets', presets)

    def delete_format_preset(self, name):
        """删除格式预设"""
        presets = self.get_format_presets()
        if name in presets:
            del presets[name]
            self.set('format_presets', presets)
    
    def save_user_template(self, template: Dict[str, Any], project_path: str) -> str:
        """
        保存用户的格式要求为JSON文件
        Args:
            template: 格式要求
            project_path: 项目路径
        Returns:
            保存的文件路径
        """
        template_path = Path(project_path) / "format_template.json"
        try:
            # 确保目录存在
            template_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(template_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, indent=2, ensure_ascii=False)
            
            # 更新配置
            self.config.setdefault("formatting", {})["user_template_path"] = str(template_path)
            self._save_config()
            
            return str(template_path)
        except Exception as e:
            print(f"保存用户模板失败: {str(e)}")
            return None
    
    def is_ai_enabled(self) -> bool:
        """检查是否启用AI功能"""
        return self.config.get("ai_assistant", {}).get("enabled", False)

    def get_ai_model(self) -> str:
        """获取AI模型名称"""
        return self.config.get("ai_assistant", {}).get("model", "gpt-3.5-turbo")

    def get_template_path(self) -> str:
        """获取当前使用的模板路径"""
        formatting = self.config.get("formatting", {})
        if not formatting.get("use_default_template", True):
            user_path = formatting.get("user_template_path")
            if user_path:
                return user_path

        default_path = Path(__file__).parent.parent / "core" / "presets" / "default.json"
        return str(default_path)