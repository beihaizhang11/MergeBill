"""
配置管理器 - 负责预设配置的存储和管理
"""
import json
import os
from pathlib import Path


class ConfigManager:
    def __init__(self, config_file="config.json"):
        """初始化配置管理器"""
        self.config_file = config_file
        self.config = self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"加载配置文件失败: {e}")
                return self.create_default_config()
        else:
            return self.create_default_config()
    
    def create_default_config(self):
        """创建默认配置"""
        return {
            "presets": {
                "默认预设": {
                    "name": "默认预设",
                    "description": "默认的账单配置",
                    "settlement_search_column": "D",
                    "settlement_search_keyword": "折后总计",
                    "mappings": [
                        {
                            "name": "日期",
                            "cell": "A1",
                            "description": "账单日期"
                        },
                        {
                            "name": "金额",
                            "cell": "B2",
                            "description": "账单金额"
                        },
                        {
                            "name": "备注",
                            "cell": "C3",
                            "description": "备注信息"
                        }
                    ]
                }
            }
        }
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False
    
    def get_preset_names(self):
        """获取所有预设名称列表"""
        return list(self.config.get("presets", {}).keys())
    
    def get_preset(self, preset_name):
        """获取指定预设配置"""
        return self.config.get("presets", {}).get(preset_name)
    
    def add_preset(self, preset_name, description="", mappings=None, 
                   settlement_search_column="D", settlement_search_keyword="折后总计"):
        """添加新预设"""
        if mappings is None:
            mappings = []
        
        if "presets" not in self.config:
            self.config["presets"] = {}
        
        self.config["presets"][preset_name] = {
            "name": preset_name,
            "description": description,
            "settlement_search_column": settlement_search_column,
            "settlement_search_keyword": settlement_search_keyword,
            "mappings": mappings
        }
        return self.save_config()
    
    def update_preset(self, preset_name, description=None, mappings=None,
                     settlement_search_column=None, settlement_search_keyword=None):
        """更新预设配置"""
        if preset_name not in self.config.get("presets", {}):
            return False
        
        preset = self.config["presets"][preset_name]
        
        if description is not None:
            preset["description"] = description
        
        if mappings is not None:
            preset["mappings"] = mappings
        
        if settlement_search_column is not None:
            preset["settlement_search_column"] = settlement_search_column
        
        if settlement_search_keyword is not None:
            preset["settlement_search_keyword"] = settlement_search_keyword
        
        return self.save_config()
    
    def delete_preset(self, preset_name):
        """删除预设"""
        if preset_name in self.config.get("presets", {}):
            del self.config["presets"][preset_name]
            return self.save_config()
        return False
    
    def rename_preset(self, old_name, new_name):
        """重命名预设"""
        if old_name in self.config.get("presets", {}):
            preset = self.config["presets"][old_name]
            preset["name"] = new_name
            self.config["presets"][new_name] = preset
            del self.config["presets"][old_name]
            return self.save_config()
        return False
    
    def duplicate_preset(self, preset_name, new_name):
        """复制预设"""
        if preset_name in self.config.get("presets", {}):
            preset = self.config["presets"][preset_name].copy()
            preset["name"] = new_name
            # 深拷贝mappings
            preset["mappings"] = [m.copy() for m in preset["mappings"]]
            # 确保有默认的结算配置
            if "settlement_search_column" not in preset:
                preset["settlement_search_column"] = "D"
            if "settlement_search_keyword" not in preset:
                preset["settlement_search_keyword"] = "折后总计"
            self.config["presets"][new_name] = preset
            return self.save_config()
        return False
    
    def validate_cell_reference(self, cell_ref):
        """验证单元格引用格式"""
        import re
        # 匹配如 A1, B2, AA10, Z999 等格式
        pattern = r'^[A-Z]+\d+$'
        return bool(re.match(pattern, cell_ref.upper()))

