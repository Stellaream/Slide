import os
import json
import random
from openai import OpenAI
from config import API_KEY, BASE_URL, MODEL_NAME
from utils import save_debug_file

# 初始化 LLM
client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

class GlobalImageMatcher:
    def __init__(self, user_assets):
        """
        user_assets: [{'path': '...', 'tags': ['描述', '关键词']}, ...]
        """
        self.img_map = {}
        # 建立资源映射表 I1, I2...
        if user_assets:
            for idx, asset in enumerate(user_assets):
                img_id = f"I{idx+1}"
                desc_str = ", ".join(asset.get('tags', []))
                self.img_map[img_id] = {
                    "path": asset['path'],
                    "desc": desc_str
                }

    def run_matching(self, slides_data):
        """
        :param slides_data: LLM 生成的 PPT 页面数据
        :return: { (page_idx, element_idx): image_path }
        """
        slot_map = {}
        prompt_slots_text = []
        
        # 初始化全局计数器，从0开始
        global_slot_idx = 0 
        
        # 1. 提取所有图片坑位
        for p_idx, slide in enumerate(slides_data):
            elements = slide.get('elements', [])
            for e_idx, el in enumerate(elements):
                if el.get('type') == 'image':
                    # 生成简化的 Slot ID (例如 S0, S1, S2...)
                    slot_id = f"S{global_slot_idx}"
                    
                    # 仅提取 content，去除关键词和主题拼接
                    content = el.get('content', '无描述')
                    
                    # 记录映射关系，以便后续根据 S0 找到具体的页码和元素索引
                    slot_map[slot_id] = {"page": p_idx, "el_idx": e_idx, "desc": content}
                    
                    # 按照要求的格式拼接文本: [S0]: 描述内容
                    prompt_slots_text.append(f"[{slot_id}]: {content}")
                    
                    # 计数器自增
                    global_slot_idx += 1

        # 如果没有坑位或没有用户图，直接返回空
        if not slot_map or not self.img_map:
            return {}

        # 2. 构造资源文本
        prompt_images_text = []
        for img_id, info in self.img_map.items():
            prompt_images_text.append(f"[{img_id}] {info['desc']}")

        debug_context = {
            "slots_sent_to_llm": prompt_slots_text,
            "images_sent_to_llm": prompt_images_text
        }
        save_debug_file("match_context.json", debug_context)

        # 3. 调用 LLM
        match_result = self._call_llm_arbitrator(prompt_slots_text, prompt_images_text)
        
        save_debug_file("match_result.json", match_result)

        # 4. 解析结果
        final_mapping = {}
        
        for s_id, i_id in match_result.items():
            # 过滤无效值
            if not i_id or not isinstance(i_id, str):
                continue
                
            # 校验 ID 是否存在
            if s_id in slot_map and i_id in self.img_map:
                slot_info = slot_map[s_id]
                img_path = self.img_map[i_id]['path']
                
                # 记录结果
                key = (slot_info['page'], slot_info['el_idx'])
                final_mapping[key] = img_path
                
        return final_mapping

    def _call_llm_arbitrator(self, slots_text, images_text):
        # 同步更新 Prompt 中的示例，使其符合 S0, S1 的格式
        prompt = f"""
        你是一个图片资源分配系统的核心逻辑模块。
        任务：将【资源列表】中的图片 ID 分配给【需求列表】中的坑位 ID。

        【需求列表 (Slots)】
        {chr(10).join(slots_text)}

        【资源列表 (Images)】
        {chr(10).join(images_text)}

        【分配规则 (重要)】
        1. 语义匹配：根据描述的相关性进行匹配（如“团队”匹配“合照”）。
        2. 宁缺毋滥：如果某张图完全不相关，**不要强行分配**，直接忽略该 Slot ID。
        3. 格式严格：返回 JSON 对象。Key 是 Slot ID，Value 是 Image ID。
        4. 禁止幻觉：绝对不要返回资源列表中不存在的 Image ID。

        【输出示例】
        {{
            "S0": "I1", 
            "S3": "I2"
        }}
        (注意：如果没有匹配项，返回空 JSON {{}})
        """

        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                temperature=0.2 
            )
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"❌ 匹配过程 LLM 调用失败: {e}")
            return {}

class StockManager:
    def __init__(self, assets_dir="assets"):
        self.pool = []
        if assets_dir and os.path.exists(assets_dir):
            valid_exts = {'.jpg', '.jpeg', '.png', '.bmp'}
            for root, dirs, files in os.walk(assets_dir):
                for f in files:
                    if os.path.splitext(f)[1].lower() in valid_exts:
                        self.pool.append(os.path.join(root, f))
        
        self.queue = list(self.pool)
        random.shuffle(self.queue)
        
        if self.pool:
            # 静默加载，不打印日志防止刷屏
            pass

    def pick_next(self):
        """
        获取一张图片，保证尽量不重复
        """
        if not self.pool:
            return None
            
        if not self.queue:
            self.queue = list(self.pool)
            random.shuffle(self.queue)
            
        return self.queue.pop()