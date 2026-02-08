import mammoth
import re

def clean_text(text):
    """
    深度清洗 Markdown 文本，并强制段落分行。
    (保持不变：去噪、去空格、双换行重组)
    """
    # 1. 去除图片占位符
    text = re.sub(r'\[image:.*?\]', '', text)

    # 2. 修复转义字符
    text = text.replace(r'\.', '.') 
    text = text.replace(r'\_', '_') 

    # 3. 去除 Markdown 格式噪音
    text = re.sub(r'_{2,}', '', text)
    text = re.sub(r'\*{2,}', '', text)

    # 4. 去除中文间空格
    text = re.sub(r'(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])', '', text)

    # 5. 规范化空白
    text = text.replace('\u3000', ' ')
    text = re.sub(r'[ \t\f\v]+', ' ', text)
    
    return text

def chunk_text_with_overlap(text, chunk_size=800, overlap=100):
    """
    带重叠的滑动窗口分块，保持段落间距。
    (保持不变：处理超长段落、段落累积、重叠回溯)
    """
    if not text:
        return []

    paragraphs = text.split('\n\n')
    chunks = []
    current_chunk = []
    current_length = 0

    for para in paragraphs:
        para_len = len(para)
        
        # 1. 处理超长段落
        if para_len > chunk_size:
            if current_chunk:
                chunks.append("\n\n".join(current_chunk))
                # 计算重叠
                joined_prev = "\n\n".join(current_chunk)
                backtrack_len = min(len(joined_prev), overlap)
                current_chunk = [joined_prev[-backtrack_len:]] if backtrack_len > 0 else []
                current_length = sum(len(c) for c in current_chunk)

            # 强制切分
            for i in range(0, para_len, chunk_size - overlap):
                chunks.append(para[i : i + chunk_size])
            
            current_chunk = []
            current_length = 0
            continue

        # 2. 正常累积
        if current_length + para_len + 2 > chunk_size:
            chunks.append("\n\n".join(current_chunk))
            
            # 计算重叠
            overlap_buffer = []
            overlap_len = 0
            for p in reversed(current_chunk):
                overlap_buffer.insert(0, p)
                overlap_len += len(p)
                if overlap_len >= overlap:
                    break
            
            current_chunk = overlap_buffer
            current_length = overlap_len
        
        current_chunk.append(para)
        current_length += para_len + 2 

    if current_chunk:
        chunks.append("\n\n".join(current_chunk))

    return chunks

def docx_to_markdown(file_path):
    """
    对外接口：读取 -> 清洗 -> 分块 -> 格式化为精简 JSON
    """
    def ignore_image(image): return []
    
    try:
        with open(file_path, "rb") as docx_file:
            # 1. 转换
            result = mammoth.convert_to_markdown(
                docx_file, 
                convert_image=mammoth.images.img_element(ignore_image)
            )
            raw_text = result.value
            
            # 2. 清洗
            cleaned_text = clean_text(raw_text)
            
            # 3. 分块
            raw_chunks = chunk_text_with_overlap(cleaned_text, chunk_size=800, overlap=100)
            
            # 4. 格式化 JSON 数据 (只保留 id 和 content)
            structured_data = []
            for index, chunk_content in enumerate(raw_chunks):
                structured_data.append({
                    "chunk_id": index + 1,
                    "content": chunk_content
                })
            
            return structured_data

    except Exception as e:
        print(f"❌ 解析错误: {file_path} - {e}")
        return []
    
def collect_ref_chunks(slide_info, chunks, max_len=8000):
    """
    根据 slide_info["ref_chunks"] 抽取对应 chunk 原文
    """
    ref_ids = slide_info.get("ref_chunks", [])
    contents = []

    for cid in ref_ids:
        for ch in chunks:
            if ch["chunk_id"] == cid:
                contents.append(f"{ch['content']}")
                break

    merged = "\n\n".join(contents)
    return merged[:max_len]
