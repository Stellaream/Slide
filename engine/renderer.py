import os
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

# ==========================================
# 1. 配色主题配置 (Academic Blue)
# ==========================================
THEME = {
    "primary":    "#005691",  # 主色：深蓝
    "accent":     "#1E88E5",  # 辅色：亮蓝
    "bg_card":    "#FFFFFF",  # 卡片背景
    "border":     "#D1E1EF",  # 边框色
    "text_main":  "#333333",  # 正文
    "text_sub":   "#5F6368",  # 辅助文字
    "line":       "#E0E0E0"   # 分隔线
}

class ProRenderer:
    def __init__(self, prs):
        self.prs = prs
        # 16:9 宽屏设置
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        
        # 栅格系统
        self.cols = 16 
        self.rows = 9
        self.grid_w = self.prs.slide_width / self.cols
        self.grid_h = self.prs.slide_height / self.rows

    def _hex_to_rgb(self, hex_str):
        if not hex_str or str(hex_str).lower() in ['transparent', 'none', '']: 
            return None
        try:
            hex_str = str(hex_str).lstrip('#')
            if len(hex_str) == 3: 
                hex_str = ''.join([c*2 for c in hex_str])
            return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))
        except: 
            return RGBColor(0, 0, 0)

    def _set_font_face(self, run, font_name='Microsoft YaHei'):
        run.font.name = font_name
        rPr = run._r.get_or_add_rPr()
        ea = OxmlElement('a:ea')
        ea.set('typeface', font_name)
        rPr.append(ea)

    def _apply_text_frame_style(self, shape, content, align='left', font_size=18, color='#333333', bold=False, line_spacing=1.2, margin=5, auto_size=True):
        """ 
        style应用函数 
        :param auto_size: True=开启自动缩放(font_size作为上限), False=强制固定字号(font_size作为绝对值)
        """
        tf = shape.text_frame
        tf.clear()
        
        # 设置内边距
        tf.margin_left = Pt(margin)
        tf.margin_right = Pt(margin)
        tf.margin_top = Pt(margin)
        tf.margin_bottom = Pt(margin)
        
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE 
        tf.word_wrap = True 
        
        # 【核心修改】根据 auto_size 参数决定策略
        if auto_size:
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        else:
            tf.auto_size = MSO_AUTO_SIZE.NONE # 关闭自动调整，严格执行字号

        align_map = {
            'left': PP_ALIGN.LEFT,
            'center': PP_ALIGN.CENTER,
            'right': PP_ALIGN.RIGHT,
            'justify': PP_ALIGN.JUSTIFY
        }
        
        content_str = str(content).strip()
        lines = content_str.split('\n') if content_str else [""]
        rgb_color = self._hex_to_rgb(color)

        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align_map.get(align, PP_ALIGN.LEFT)
            p.line_spacing = line_spacing
            p.space_after = Pt(6) if (len(lines) > 1 and i < len(lines)-1) else Pt(0)

            run = p.add_run()
            run.text = line.strip()
            run.font.size = Pt(font_size) # 若 auto_size=True，这只是最大值；若 False，这是固定值
            run.font.bold = bold
            if rgb_color:
                run.font.color.rgb = rgb_color
            
            self._set_font_face(run, 'Microsoft YaHei')

    def render_element(self, slide, el, force_image_path=None):
        pos = el.get('pos', {})
        gx = max(0, min(float(pos.get('x', 0)), 16))
        gy = max(0, min(float(pos.get('y', 0)), 9))
        gw = max(0.5, min(float(pos.get('w', 4)), 16 - gx))
        gh = max(0.5, min(float(pos.get('h', 2)), 9 - gy))
        
        l = int(gx * self.grid_w)
        t = int(gy * self.grid_h)
        w = int(gw * self.grid_w)
        h = int(gh * self.grid_h)
        
        el_type = str(el.get('type', 'text')).lower().strip()
        content = el.get('content', '')
        subtitle = el.get('subtitle', '')
        style = el.get('style', {})

        # ====================================================
        # 1. Title (大标题 - 保持醒目风格)
        # ====================================================
        if el_type == 'title':
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
            shape.adjustments[0] = 0.2
            
            shape.fill.solid()
            shape.fill.fore_color.rgb = self._hex_to_rgb("#FFFFFF")
            
            shape.line.color.rgb = self._hex_to_rgb(THEME['primary'])
            shape.line.width = Pt(2.0)

            # 标题通常文字较少，建议默认开启 AutoSize 防止溢出，但也允许手动覆盖（这里暂定默认Auto）
            self._apply_text_frame_style(
                shape, content, 
                align='center', 
                font_size=40, 
                color=THEME['primary'], 
                bold=True,
                auto_size=True 
            )

        # ====================================================
        # 2. Text (普通文本 - 智能判定 Explicit vs Auto)
        # ====================================================
        elif el_type == 'text':
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
            
            # 样式处理
            bg_color = style.get('bg_color', 'transparent')
            if bg_color != 'transparent':
                shape.fill.solid()
                shape.fill.fore_color.rgb = self._hex_to_rgb(bg_color)
            else:
                shape.fill.background()

            if style.get('border'):
                shape.line.color.rgb = self._hex_to_rgb(THEME['border'])
                shape.line.width = Pt(1)
            else:
                shape.line.fill.background()

            json_align = style.get('align', 'left')
            is_bold = style.get('bold', False)
            text_color = style.get('color', THEME['primary'] if is_bold else THEME['text_main'])
            
            # 【核心逻辑】检查 font_size 是否设置
            specified_fs = style.get('font_size')
            
            if specified_fs:
                # 显式设置了字号 -> 关闭自动缩放，使用指定值
                use_auto_size = False
                fs_val = specified_fs
            else:
                # 未设置字号 -> 开启自动缩放，使用默认最大值
                use_auto_size = True
                fs_val = 24 # 默认最大字号

            self._apply_text_frame_style(
                shape, content, 
                align=json_align, 
                font_size=fs_val,  
                color=text_color, 
                bold=is_bold,
                auto_size=use_auto_size # 传入判定结果
            )

        # ====================================================
        # 3. Card (经典分割线风格)
        # ====================================================
        elif el_type == 'card':
            # 背景容器
            bg_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
            bg_shape.adjustments[0] = 0.04
            bg_shape.fill.solid()
            bg_shape.fill.fore_color.rgb = self._hex_to_rgb(THEME['bg_card'])
            bg_shape.line.color.rgb = self._hex_to_rgb(THEME['border'])
            bg_shape.line.width = Pt(1.5)

            # 布局计算
            header_h = max(int(Inches(0.5)), min(int(h * 0.25), int(Inches(0.9))))
            body_h = h - header_h

            # Header
            header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, header_h)
            header_box.fill.background() 
            header_box.line.fill.background() 
            
            self._apply_text_frame_style(
                header_box, subtitle, 
                align='center', font_size=22, color=THEME['primary'], bold=True, margin=3,
                auto_size=True # 卡片标题建议自动适配
            )

            # Line
            line_y = t + header_h
            line_shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, l + int(w * 0.1), line_y, l + int(w * 0.9), line_y)
            line_shape.line.color.rgb = self._hex_to_rgb(THEME['line'])
            line_shape.line.width = Pt(1)

            # Body
            body_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, line_y, w, body_h)
            body_box.fill.background()
            body_box.line.fill.background()
            
            # 卡片内容通常由 AutoSize 兜底，防止溢出
            self._apply_text_frame_style(
                body_box, content, 
                align='left', font_size=16, color=THEME['text_main'], bold=False, margin=8,
                auto_size=True
            )

        # ====================================================
        # 4. Image
        # ====================================================
        elif el_type == 'image':
            img_path = force_image_path
            
            if img_path and os.path.exists(img_path):
                try:
                    pic = slide.shapes.add_picture(img_path, l, t)
                    self._crop_image_to_fit(pic, w, h)
                    pic.line.color.rgb = self._hex_to_rgb(THEME['border'])
                    pic.line.width = Pt(1)
                    return 
                except: pass

            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
            shape.adjustments[0] = 0.02
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(248, 248, 248)
            shape.line.color.rgb = self._hex_to_rgb(THEME['border'])
            shape.line.dash_style = 4
            shape.line.width = Pt(1.5)
            
            self._apply_text_frame_style(
                shape, f"🖼️\n{content}", 
                align="center", font_size=14, color=THEME['text_sub'], bold=False,
                auto_size=True
            )

    def _crop_image_to_fit(self, pic, target_w, target_h):
        img_ratio = pic.width / pic.height
        target_ratio = target_w / target_h
        
        pic.crop_left = 0
        pic.crop_right = 0
        pic.crop_top = 0
        pic.crop_bottom = 0

        if img_ratio > target_ratio:
            new_width = target_h * img_ratio
            crop_amount = (new_width - target_w) / new_width
            pic.crop_left = crop_amount / 2
            pic.crop_right = crop_amount / 2
            pic.width = target_w
            pic.height = target_h
        else:
            new_height = target_w / img_ratio
            crop_amount = (new_height - target_h) / new_height
            pic.crop_top = crop_amount / 2
            pic.crop_bottom = crop_amount / 2
            pic.width = target_w
            pic.height = target_h

    def add_background_to_all_slides(self, image_path):
        if not image_path or not os.path.exists(image_path): return
        for slide in self.prs.slides:
            try:
                pic = slide.shapes.add_picture(image_path, 0, 0, width=self.prs.slide_width, height=self.prs.slide_height)
                try:
                    blip = pic._element.blipFill.blip
                    lum = OxmlElement('a:lum')
                    lum.set('bright', '40000')    
                    lum.set('contrast', '-40000') 
                    blip.append(lum)
                except: pass
                spTree = slide.shapes._spTree
                element = pic._element
                spTree.remove(element)
                spTree.insert(2, element)
            except: pass

# ==========================================
# 测试代码
# ==========================================
if __name__ == "__main__":
    
    test_slide_json = {
        "index": 1,
        "title": "核心技术方案概述页",
        "elements": [
            # 1. 标题 (默认 AutoSize)
            {
                "type": "title",
                "pos": {"x": 1, "y": 0.5, "w": 14, "h": 1},
                "content": "核心技术方案概述",
                "style": {"bg_color": "transparent"}
            },
            
            # 2. Text (显式设置 font_size，测试固定字号)
            {
                "type": "text",
                "pos": { "x": 1, "y": 1.6, "w": 6, "h": 0.5 },
                "content": "",
                "style": { 
                    "bold": False, 
                    "bg_color": "#EFEFEF", 
                    "align": "left",
                    "font_size": 12 # <--- 显式设置
                }
            },
            
            
            # 4. Cards (经典风格：Line + Subtitle)
            {
                "type": "card",
                "subtitle": "材料创新",
                "pos": { "x": 0.5, "y": 2.5, "w": 2.8, "h": 3.5 },
                "content": "液态金属基导电薄膜\n低迟滞聚氨酯介质\n仿水膜 - 鱼网结构",
                "style": { "bg_color": "#FFFFFF" }
            },
            {
                "type": "card",
                "subtitle": "硬件架构",
                "pos": { "x": 3.5, "y": 2.5, "w": 2.8, "h": 3.5 },
                "content": "光纤-IMU 混合架构\n硬件级漂移抑制\n目标精度 0.1mm",
                "style": { "bg_color": "#FFFFFF" }
            },
            {
                "type": "card",
                "subtitle": "核心算法",
                "pos": { "x": 6.5, "y": 2.5, "w": 2.8, "h": 3.5 },
                "content": "分层手势识别模型\n多滑动窗口自适应分割\n分割精度>97%",
                "style": { "bg_color": "#FFFFFF" }
            },
            {
                "type": "card",
                "subtitle": "系统集成",
                "pos": { "x": 9.5, "y": 2.5, "w": 2.8, "h": 3.5 },
                "content": "多模态闭环交互\n压力梯度与温控反馈\n标准化协议对接",
                "style": { "bg_color": "#FFFFFF" }
            },
            {
                "type": "card",
                "subtitle": "验证测试",
                "pos": { "x": 12.5, "y": 2.5, "w": 2.8, "h": 3.5 },
                "content": "国产设备无缝对接\n场景压力测试\n10+ 主流设备兼容",
                "style": { "bg_color": "#FFFFFF" }
            },
            
            {
                "type": "image",
                "pos": {"x": 0.5, "y": 6.5, "w": 7, "h": 2},
                "content": "架构图占位",
                "style": {"bg_color": "#F0F0F0", "border": "dashed"}
            },
            {
                "type": "card",
                "subtitle": "关键性能指标",
                "pos": {"x": 8, "y": 6.5, "w": 7.5, "h": 2},
                "content": "● 端到端延迟：<50ms\n● 空间分辨率：0.1mm 级\n● 抗干扰能力：复杂环境鲁棒性\n● 自主可控：全栈自研技术链",
                "style": {"bg_color": "#FFFFFF"}
            }
        ]
    }

    print("🚀 开始 PPT 渲染测试 (混合字号模式)...")
    prs = Presentation()
    blank_layout = prs.slide_layouts[6] 
    slide = prs.slides.add_slide(blank_layout)
    
    renderer = ProRenderer(prs)
    
    for el in test_slide_json["elements"]:
        renderer.render_element(slide, el)

    output_filename = "mixed_fontsize_logic.pptx"
    
    try:
        prs.save(output_filename)
        print(f"✅ 测试完成！文件已保存为: {os.path.abspath(output_filename)}")
    except PermissionError:
        print(f"❌ 保存失败：请先关闭 {output_filename} 文件。")