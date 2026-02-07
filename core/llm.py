# core/llm.py
import json
from openai import OpenAI
from config import API_KEY, BASE_URL, MODEL_NAME
from utils import save_debug_file

# 初始化客户端
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

def get_ppt_outline(chunks):
    """第一阶段：生成大纲"""
    print("正在规划 PPT 逻辑大纲 (15-20页)...")

    # 造带有显式 ID 标记的上下文
    context_text = ""
    for item in chunks:
        context_text += f"[片段ID: {item['chunk_id']}]\n{item['content']}\n\n"
    
    prompt = f"""
    请阅读以下文档，将其转化为一份 15-20 页的答辩 PPT 大纲。
    要求涵盖：封面、项目背景、行业痛点、核心技术方案(多页详细拆解)、创新点、性能对比、商业模式、团队介绍、发展规划、总结问答。
    封面页严禁包含正文描述，只需包含：项目题目、汇报人信息、指导老师、学校、日期。
    合理控制每页内容，避免过度拥挤。
    封面页内容不宜过多，正文页需突出核心论点和数据支撑。
    对于每页 PPT，列出该页内容主要参考的原文[片段ID]（整数列表）。
    
    返回 JSON 格式如下：
    {{
        "outline": [
        {{"index": 1, "title": "标题", "focus": "本页核心论点及需要展示的数据/图表建议", "ref_chunks": [1, 3, 5]}},
        ]
    }}
    文档内容：{context_text[:15000]}
    """
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": "你是一个只输出 JSON 的策划专家。"},
                {"role": "user", "content": prompt}
            ],
            response_format={ "type": "json_object" } 
        )
        res_data = json.loads(response.choices[0].message.content)
        return res_data["outline"]
    except Exception as e:
        print(f"❌ 生成大纲失败: {e}")
        return []

def generate_single_slide(slide_info, md_content, user_asset_hints=None):
    """第二阶段：生成单页布局"""
    idx = slide_info['index']
    print(f"正在设计第 {idx} 页: {slide_info['title']}")

    assets_hint_text = "无可用用户图片信息"
    if user_asset_hints:
        hint_lines = []
        for item in user_asset_hints:
            hint_lines.append(
                f"[{item.get('asset_id')}] tags={item.get('tags', [])}, ratio={item.get('aspect_ratio')}, size={item.get('width')}x{item.get('height')}"
            )
        assets_hint_text = "\n".join(hint_lines)
    
    prompt = f"""
    你是一名 PPT 专家，针对主题 '{slide_info['title']}'，根据参考内容设计一页高水平科创大赛答辩 PPT 布局。
    核心要点：{slide_info['focus']}
    参考原文：{md_content[:8000]}

    用户图片库信息（用于提前规划图框比例和位置）：
    {assets_hint_text}
    
    Action Protocol 规范
    1. 使用 16x9 栅格系统 (pos: x, y, w, h)，必须确保：
        - (当前元素的 x + w) 永远小于等于 16。
        - (当前元素的 y + h) 永远小于等于 9。
        - 元素之间必须保留合适单位的“视觉呼吸感”间距。
        - 严禁重叠。
    2. 类型仅限: "text" (文字容器), "image" (图片留白框)。
    3. style 包含: bold (加粗), align (center/left), color (十六进制), bg_color (背景色/默认transparent), border (布尔值)。
    3.1 图片框比例建议：若用户图片库存在明显的横图(>1.3)或竖图(<0.8)，请优先采用对应宽高比的 image 框，避免极端拉伸裁切。
    4. 样式一致： 一个元素(element)内部只能有一种 color。禁止在同一个 content 中混合多种颜色。
    5. 视觉对齐：标题通常占据 x:1, y:1, w:10, h:1.5，正文与图片应水平对齐（y相同）或垂直分布。
    6. 标题标签化：对于大标题和小标题，配色体现科创标签感（比如浅蓝色系）。
    7. 内容卡片化：对于正文块，建议设置 bg_color: "#FFFFFF", border: true，使其呈现为白色卡片感。
    8. 所有元素平级放在 elements 数组中。
    9. 文字容器中内容不要使用 markdown 语法，只需纯文本。
    10. 不要重叠（即一个对象的 x+w 和 y+h 不能超过另一个对象的 x,y）。
    
    排版建议示例 (Few-shot)
    - 左右分栏示例: [
        {{"type": "text", "pos": {{"x": 1, "y": 1.5, "w": 10, "h": 1}}, "content": "标题", "style": {{"bold": true}}}},
        {{"type": "text", "pos": {{"x": 1, "y": 3, "w": 5, "h": 7}}, "content": "要点描述...", "style": {{}}}},
        {{"type": "image", "pos": {{"x": 6.5, "y": 3, "w": 4.5, "h": 7}}, "content": "此处应展示一张算法架构逻辑图"}}
        ]
    
    视觉重心优化准则
    1. 利用顶部空间，不要产生“内容下沉”感。
    2. 黄金起跑线：
        - 主标题建议放在 y: 0.5 到 y: 1 之间。
        - 正文内容建议从 y: 2 或 y: 2.5 开始。
    3. 留白平衡：底部的 y=8 到 y=9 区域应留出更多空白（用于放置页码或作为视觉呼吸区），除非内容极多。
    4. 纵向紧凑：在 16x9 系统下，纵向高度非常宝贵，请尽量压缩组件的 h (高度)，确保核心内容处于视觉上半区。
    
    直接输出 JSON 对象，必须包含 "elements" 键。
    """
    
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[{"role": "user", "content": prompt}]
        )
        raw_content = response.choices[0].message.content
        clean_json = raw_content.replace("```json", "").replace("```", "").strip()
        slide_data = json.loads(clean_json)
        return slide_data
    except Exception as e:
        print(f"⚠️ 第 {idx} 页生成出错: {e}")
        return None
