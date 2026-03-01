# core/llm.py
import json
import re
from openai import OpenAI
from config import API_KEY, BASE_URL, MODEL_NAME

# 初始化客户端 (阿里云 DashScope)
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

def clean_json_response(content):
    """
    清洗模型返回的内容，提取 JSON 部分。
    即使模型开启了 JSON 模式，有时也会包裹 ```json ... ```，需要去除。
    """
    try:
        # 尝试直接解析
        return json.loads(content)
    except json.JSONDecodeError:
        # 如果失败，尝试去除 markdown 标记
        content = content.replace("```json", "").replace("```", "").strip()
        # 有时候模型会在 JSON 后跟一些废话，尝试提取第一个 {} 闭合区间（简单版）
        # 这里使用正则提取最外层的 JSON 对象
        match = re.search(r'(\{.*\})', content, re.DOTALL)
        if match:
            return json.loads(match.group(1))
        # 再次尝试直接解析清洗后的文本
        return json.loads(content)

def get_ppt_outline(chunks):
    """第一阶段：生成大纲"""
    print(f"[{MODEL_NAME}] 正在规划 PPT 逻辑大纲 (15-20页)...")

    # 构造带有显式 ID 标记的上下文
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
    
    【输出要求】
    必须返回符合 JSON 语法的纯文本，不要包含任何 markdown 格式标记（如 ```json）。
    格式如下：
    {{
        "outline": [
            {{"index": 1, "title": "标题", "focus": "本页核心论点及需要展示的数据/图表建议", "ref_chunks": [1, 3, 5]}},
            ...
        ]
    }}
    文档内容：{context_text[:15000]}
    """
    
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": "你是一个专业的 PPT 策划专家，请直接输出 JSON 数据。"},
                {"role": "user", "content": prompt}
            ],
            # 只有支持 json_object 的模型才加这个参数，Qwen-plus 通常支持，但为了兼容性，依靠 prompt 约束也可
            response_format={"type": "json_object"}, 
            temperature=0.2, # 降低随机性，提高 JSON 结构稳定性
        )
        
        content = response.choices[0].message.content
        res_data = clean_json_response(content)
        return res_data.get("outline", [])
        
    except Exception as e:
        print(f"❌ 生成大纲失败: {e}")
        return []

def generate_single_slide(slide_info, md_content, user_asset_hints=None):
    """第二阶段：生成单页布局"""
    idx = slide_info['index']
    print(f"[{MODEL_NAME}] 正在设计第 {idx} 页: {slide_info['title']}")

    assets_hint_text = "无可用用户图片信息"
    if user_asset_hints:
        hint_lines = []
        for item in user_asset_hints:
            hint_lines.append(
                f"[{item.get('asset_id')}] tags={item.get('tags', [])}, ratio(width/height)={item.get('aspect_ratio')}"
            )
        assets_hint_text = "\n".join(hint_lines)
    
    prompt = f"""
    你是一名 PPT 设计专家。请根据以下要求设计第 {idx} 页 PPT "{slide_info['title']}" 的布局 JSON。
    
    核心要点：{slide_info['focus']}
    参考原文：{md_content[:8000]}
    用户图片库：{assets_hint_text}

    【设计目标】根据核心要点提炼出本页的关键信息，合理利用用户图片库中的素材（如果有合适的），并设计一个清晰、专业、美观的布局方案。
    
    【排版规范 (16x9 网格)】
    1. 画布大小 16x9 (x=0~16, y=0~9)。严禁元素重叠。
    2. 允许的元素类型 (type)：
    - "title": 页面大标题，通常位于顶部 (y=0~1.5)，高度约 1-1.5。
    - "card": 复合卡片，包含小标题和正文。**必须包含 "subtitle" 和 "content" 两个字段**。适合展示分点论述。
    - "text": 普通文本框。适合纯段落或简单列表。
    - "image": 图片框。
    3. 样式 (style) 要求：
    - 对于 "text" 类型，必须在 style 中指定 "font_size" 和 "align" ("left", "center")。
    - 对于 "card" 类型，建议背景设为 "#FFFFFF"。
    4. 图片策略：优先使用用户图片库中的 asset_id；若无合适图片，type="image" 且 content 留空作为占位符。
    
    【输出格式示例】
    {{
    "elements": [
        {{
        "type": "title", 
        "pos": {{"x": 1, "y": 0.5, "w": 14, "h": 1.2}}, 
        "content": "核心技术架构"
        }},
        {{
        "type": "text", 
        "pos": {{"x": 1, "y": 2, "w": 14, "h": 0.8}}, 
        "content": "本架构采用分层设计，确保高可用性与扩展性。",
        "style": {{"font_size": 16, "align": "center", "bold": false}}
        }},
        {{
        "type": "card", 
        "pos": {{"x": 1, "y": 3, "w": 4, "h": 4}}, 
        "subtitle": "感知层",
        "content": "集成多模态传感器\n实时数据采集", 
        "style": {{"bg_color": "#FFFFFF"}}
        }},
        {{
        "type": "image", 
        "pos": {{"x": 6, "y": 3, "w": 9, "h": 4}}, 
        "content": "img_asset_123" 
        }}
    ]
    }}
    """
    
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": "你是一个 PPT 布局算法，只输出 JSON。"},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.2, 
        )
        
        content = response.choices[0].message.content
        slide_data = clean_json_response(content)
        
        # 简单校验数据完整性
        if "elements" not in slide_data:
            print(f"⚠️ 第 {idx} 页生成数据缺失 elements 字段")
            return None
            
        return slide_data

    except Exception as e:
        print(f"⚠️ 第 {idx} 页生成出错: {e}")
        return None