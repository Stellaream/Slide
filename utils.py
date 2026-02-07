# utils.py
import os
import json
import random
from config import DEBUG_DIR, ASSETS_DIR, BACKGROUND_DIR 

def save_debug_file(filename, data, is_json=True):
    """保存调试日志"""
    path = os.path.join(DEBUG_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        if is_json:
            json.dump(data, f, indent=4, ensure_ascii=False)
        else:
            f.write(str(data))

def get_random_background():
    """从 assets/background 文件夹中随机选择一张图片"""
    valid_extensions = ('.png', '.jpg', '.jpeg', '.bmp')
    
    # 检查文件夹是否存在
    if not os.path.exists(BACKGROUND_DIR):
        print(f"⚠️ 提示：未找到背景目录 {BACKGROUND_DIR}")
        return None
        
    files = [f for f in os.listdir(BACKGROUND_DIR) if f.lower().endswith(valid_extensions)]
    
    if not files:
        print(f"⚠️ 提示：{BACKGROUND_DIR} 文件夹为空，将生成无背景 PPT。")
        return None
        
    return os.path.join(BACKGROUND_DIR, random.choice(files))

def extract_elements_robust(slide_data):
    """提取 elements 数组，兼容各种 LLM 返回格式"""
    if not slide_data: return []
    if isinstance(slide_data, list): return slide_data
    if isinstance(slide_data, dict):
        if 'elements' in slide_data: return slide_data['elements']
        if 'slides' in slide_data and len(slide_data['slides']) > 0:
            return slide_data['slides'][0].get('elements', [])
        # 尝试遍历查找第一个列表
        for v in slide_data.values():
            if isinstance(v, list) and len(v) > 0: return v
    return []