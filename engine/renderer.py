import os
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement

class ProRenderer:
    def __init__(self, prs, image_manager=None):
        self.prs = prs
        # image_manager 在新逻辑中其实已经不太需要了，但保留着防止报错
        self.image_manager = image_manager
        
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.cols = 16 
        self.rows = 9
        self.grid_w = self.prs.slide_width / self.cols
        self.grid_h = self.prs.slide_height / self.rows

    def _hex_to_rgb(self, hex_str):
        if not hex_str or str(hex_str).lower() == 'transparent': return None
        try:
            hex_str = str(hex_str).lstrip('#')
            if len(hex_str) == 3: hex_str = ''.join([c*2 for c in hex_str])
            return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))
        except: return RGBColor(0, 0, 0)

    def _crop_image_to_fit(self, pic, target_w, target_h):
        """ 核心算法: Object-Fit Cover """
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

    def _apply_style_text(self, shape, content, s, box_w_grid, box_h_grid, force_color=None):
        tf = shape.text_frame
        tf.clear()
        content_str = str(content).strip()
        lines = content_str.split('\n') if content_str else [""]
        
        is_title_like = s.get('bold', False) or (box_h_grid < 2.0)
        base_size = 32 if is_title_like else 24

        target_color = force_color if force_color else s.get('color', '#333333')
        target_rgb = self._hex_to_rgb(target_color)

        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = str(line).strip()
            run.font.size = Pt(base_size) 
            run.font.bold = s.get('bold', False)
            run.font.name = '微软雅黑'
            if target_rgb:
                run.font.color.rgb = target_rgb
            try:
                rPr = run._r.get_or_add_rPr()
                ea = OxmlElement('a:ea')
                ea.set('typeface', '微软雅黑')
                rPr.append(ea)
            except: pass 
            
            align = str(s.get('align', 'left')).lower()
            if 'center' in align: p.alignment = PP_ALIGN.CENTER
            elif 'right' in align: p.alignment = PP_ALIGN.RIGHT
            else: p.alignment = PP_ALIGN.LEFT

        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_top = Pt(2)
        tf.margin_bottom = Pt(2)
        tf.margin_left = Pt(2)
        tf.margin_right = Pt(2)
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    def render_element(self, slide, el, force_image_path=None):
        pos = el.get('pos', {})
        if isinstance(pos, list) and len(pos) >= 4:
            pos = {'x': pos[0], 'y': pos[1], 'w': pos[2], 'h': pos[3]}
        if not isinstance(pos, dict): return

        visual_lift = -0.1 
        gx = max(0, min(float(pos.get('x', 0)), 15.5))
        gy = max(0, min(float(pos.get('y', 0)) + visual_lift, 8.5))
        gw = max(0.5, min(float(pos.get('w', 4)), 16 - gx))
        gh = max(0.5, min(float(pos.get('h', 2)), 9 - gy))
        
        l, t = gx * self.grid_w, gy * self.grid_h
        w, h = gw * self.grid_w, gh * self.grid_h
        
        emu_l, emu_t = int(l), int(t)
        emu_w, emu_h = int(w), int(h)
        
        el_type = str(el.get('type', 'text')).lower()
        content = el.get('content', '')
        style = el.get('style', {})

        if 'text' in el_type:
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
            bg_color = style.get('bg_color', 'transparent')
            if bg_color == 'transparent' and style.get('bold') and gh < 1.2:
                bg_color = "#E3F2FD" 
            
            bg_rgb = self._hex_to_rgb(bg_color)
            if bg_rgb:
                shape.fill.solid()
                shape.fill.fore_color.rgb = bg_rgb
            else:
                shape.fill.background() 
            
            if style.get('border', False):
                shape.line.color.rgb = self._hex_to_rgb('#A0C4E3')
                shape.line.width = Pt(1)
            else:
                shape.line.fill.background()
            
            self._apply_style_text(shape, content, style, gw, gh)

        elif 'image' in el_type:
            # 直接使用外部传入的 force_image_path (包含用户匹配图 或 商务兜底图)
            img_path = force_image_path

            if img_path and os.path.exists(img_path):
                try:
                    pic = slide.shapes.add_picture(img_path, emu_l, emu_t)
                    self._crop_image_to_fit(pic, emu_w, emu_h)
                    pic.line.color.rgb = self._hex_to_rgb("#CCCCCC")
                    pic.line.width = Pt(0.5)
                    return 
                except Exception as e:
                    print(f"⚠️ 图片插入异常: {e}")

            # 兜底占位符
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(245, 245, 245)
            shape.line.color.rgb = RGBColor(200, 200, 200)
            shape.line.dash_style = 2 
            
            keywords = el.get('keywords', [])
            kw_str = "/".join(keywords[:2]) if keywords else "无"
            placeholder_text = f"🖼️ {content}\n(关键词: {kw_str})"
            
            self._apply_style_text(
                shape, 
                placeholder_text, 
                {"align": "center", "color": "#999999", "bold": False}, 
                gw, gh
            )

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