import os
import json
import queue
import threading
from flask import Flask, request, send_file, jsonify, Response, stream_with_context
from flask_cors import CORS
from werkzeug.utils import secure_filename

try:
    from PIL import Image
except Exception:
    Image = None

# 导入核心逻辑
from core.pipeline import run_pipeline
from config import OUTPUT_DIR

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def get_image_meta(image_path):
    """读取图片宽高与宽高比；环境不支持时返回空值。"""
    if not Image:
        return {"width": None, "height": None, "aspect_ratio": None}

    try:
        with Image.open(image_path) as img:
            width, height = img.size
            ratio = round(width / height, 4) if height else None
            return {
                "width": width,
                "height": height,
                "aspect_ratio": ratio
            }
    except Exception:
        return {"width": None, "height": None, "aspect_ratio": None}

# --- 1. 下载接口 ---
@app.route('/api/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        # 安全处理文件名
        filename = secure_filename(filename)
        file_path = os.path.join(OUTPUT_DIR, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({"error": "File not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- 2. 生成接口 (流式响应) ---
@app.route('/api/generate', methods=['POST'])
def generate_ppt_stream():
    # 基础校验
    if 'file' not in request.files:
        return jsonify({"error": "No docx file uploaded"}), 400
    
    docx_file = request.files['file']
    if docx_file.filename == '':
        return jsonify({"error": "No filename"}), 400

    filename = secure_filename(docx_file.filename)
    save_path = os.path.join(UPLOAD_FOLDER, filename)
    docx_file.save(save_path)
    print(f"📥 [后端] 收到文档: {filename}")

    # --- 处理图片与描述 ---
    user_assets = []
    
    # 获取文件列表
    uploaded_images = request.files.getlist('images') 
    # 获取描述列表 (注意：这里对应前端 FormData 的 key: 'image_descriptions')
    uploaded_descs = request.form.getlist('image_descriptions')
    
    if uploaded_images and len(uploaded_images) > 0:
        print(f"📥 [后端] 收到 {len(uploaded_images)} 张素材图片")
        
        # 配对处理
        for img, desc_str in zip(uploaded_images, uploaded_descs):
            if img.filename == '': continue
            
            # 保存图片
            safe_img_name = secure_filename(img.filename)
            img_save_path = os.path.join(UPLOAD_FOLDER, safe_img_name)
            img.save(img_save_path)
            img_meta = get_image_meta(img_save_path)
            
            # 处理描述/标签
            # 1. 保留原始描述用于 LLM 语义理解
            # 2. 同时也拆分成单词用于简单的兜底匹配
            # 例如: "核心团队在会议室开会" -> ["核心团队在会议室开会", "核心", "团队", "会议室"]
            raw_tags = desc_str.replace(',', ' ').replace('，', ' ').split()
            tags = [t.strip() for t in raw_tags if t.strip()]
            
            # 将完整描述也加进去，增加 LLM 匹配的上下文
            if desc_str.strip() and desc_str not in tags:
                tags.insert(0, desc_str.strip())
            
            user_assets.append({
                "path": img_save_path,
                "tags": tags,
                "width": img_meta["width"],
                "height": img_meta["height"],
                "aspect_ratio": img_meta["aspect_ratio"]
            })
            # print(f"   + 素材已入库: {safe_img_name}")

    # --- 后台任务 ---
    msg_queue = queue.Queue()

    def pipeline_callback(msg, log_type="info"):
        msg_queue.put({"type": log_type, "msg": msg})

    def background_worker():
        try:
            # 运行 Pipeline
            final_path = run_pipeline(
                save_path, 
                log_callback=pipeline_callback, 
                user_assets=user_assets
            )
            
            if final_path and os.path.exists(final_path):
                final_name = os.path.basename(final_path)
                msg_queue.put({"type": "done", "msg": "Done", "filename": final_name})
            else:
                msg_queue.put({"type": "error", "msg": "生成失败，未返回路径"})
                
        except Exception as e:
            error_msg = str(e)
            print(f"❌ 后台任务出错: {error_msg}")
            msg_queue.put({"type": "error", "msg": f"服务器内部错误: {error_msg}"})
            import traceback
            traceback.print_exc()
        finally:
            msg_queue.put(None)

    # 启动线程
    thread = threading.Thread(target=background_worker)
    thread.daemon = True
    thread.start()

    # 流式响应
    def event_stream():
        while True:
            data = msg_queue.get()
            if data is None:
                break
            yield json.dumps(data, ensure_ascii=False) + "\n"

    return Response(stream_with_context(event_stream()), mimetype='application/x-ndjson')

if __name__ == '__main__':
    print("🚀 后端服务已启动: http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)
