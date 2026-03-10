# utils.py
import os
import json
import math
import random
from config import DEBUG_DIR, ASSETS_DIR, BACKGROUND_DIR 
import itertools

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

def calculate_overlap(elements: list) -> float:
    """
    计算布局的重叠度 (Overlap)
    计算公式：所有两两相交元素的重叠面积之和 / 页面所有元素的总面积
    返回值范围：[0, 1+)，0 表示完美（完全无重叠）
    """
    if len(elements) < 2:
        return 0.0

    total_area = 0.0
    overlap_area = 0.0

    # 1. 计算所有元素的总面积
    for el in elements:
        w = el['pos']['w']
        h = el['pos']['h']
        total_area += w * h

    if total_area == 0:
        return 0.0

    # 2. 计算所有元素对的交集面积
    # itertools.combinations 生成所有不重复的两两组合
    for el1, el2 in itertools.combinations(elements, 2):
        p1, p2 = el1['pos'], el2['pos']
        
        # 元素1的边界
        l1, t1 = p1['x'], p1['y']
        r1, b1 = l1 + p1['w'], t1 + p1['h']
        
        # 元素2的边界
        l2, t2 = p2['x'], p2['y']
        r2, b2 = l2 + p2['w'], t2 + p2['h']
        
        # 计算交集矩形的边界
        inter_left = max(l1, l2)
        inter_top = max(t1, t2)
        inter_right = min(r1, r2)
        inter_bottom = min(b1, b2)
        
        # 如果存在重叠 (right > left 且 bottom > top)
        if inter_right > inter_left and inter_bottom > inter_top:
            inter_area = (inter_right - inter_left) * (inter_bottom - inter_top)
            overlap_area += inter_area

    # 返回重叠率
    return overlap_area / total_area


def calculate_alignment(elements: list) -> float:
    """
    计算布局的对齐度误差 (Alignment)
    计算公式：平均每个元素与最近的一个其他元素在（左、右、顶、底、水平中心、垂直中心）对齐线上的最小距离
    返回值：0 表示存在完美的几何对齐，值越小对齐越好
    """
    if len(elements) < 2:
        return 0.0

    total_min_dist = 0.0

    def get_align_lines(pos):
        """获取一个元素的6条对齐参考线：[垂直参考线(X轴), 水平参考线(Y轴)]"""
        left, top = pos['x'], pos['y']
        right, bottom = left + pos['w'], top + pos['h']
        center_x = left + pos['w'] / 2.0
        center_y = top + pos['h'] / 2.0
        return [left, center_x, right], [top, center_y, bottom]

    for i, el1 in enumerate(elements):
        lines_x1, lines_y1 = get_align_lines(el1['pos'])
        
        # 初始化该元素的最小对齐误差为无穷大
        min_dist_for_el = float('inf')
        
        for j, el2 in enumerate(elements):
            if i == j:
                continue
                
            lines_x2, lines_y2 = get_align_lines(el2['pos'])
            
            # 计算X轴（垂直对齐线）上的最小差值
            # 比较 el1 的左、中、右 与 el2 的左、中、右
            min_x_dist = min(abs(lx1 - lx2) for lx1 in lines_x1 for lx2 in lines_x2)
            
            # 计算Y轴（水平对齐线）上的最小差值
            # 比较 el1 的上、中、下 与 el2 的上、中、下
            min_y_dist = min(abs(ly1 - ly2) for ly1 in lines_y1 for ly2 in lines_y2)
            
            # 找到X轴和Y轴中最接近的一条对齐线（任何维度的对齐都算对齐）
            dist = min(min_x_dist, min_y_dist)
            
            if dist < min_dist_for_el:
                min_dist_for_el = dist
                
        total_min_dist += min_dist_for_el

    # 返回所有元素的平均最小对齐距离
    return total_min_dist / len(elements)

def calculate_layout_score(overlap_error: float, align_error: float) -> float:
    """
    基于重叠度和对齐度误差，计算版面美观度综合得分 (0~100分)
    
    :param overlap_error: calculate_overlap() 的返回值 (重叠面积占比)
    :param align_error: calculate_alignment() 的返回值 (对齐偏差均值)
    :return: 综合总分 (float, 0-100)
    """
    # 1. 计算单项得分 (使用指数衰减将误差映射为 0-100 的得分)
    # k=5.0: 对重叠惩罚较重； k=3.0: 对未对齐惩罚相对平缓
    score_overlap = 100 * math.exp(-5.0 * overlap_error)
    score_align = 100 * math.exp(-3.0 * align_error)
    
    # 2. 加权计算总分 (重叠作为致命错误占 60%，对齐作为美观瑕疵占 40%)
    total_score = (score_overlap * 0.6) + (score_align * 0.4)
    
    # 保留两位小数返回
    return round(total_score, 2)