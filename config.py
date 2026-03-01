import os

API_KEY = "sk-3c0fa0eddd5d4a47a340258641cd99fe"
BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
MODEL_NAME = "qwen3.5-plus" 

if not API_KEY:
    raise ValueError("未检测到环境变量 DEEPSEEK_API_KEY，请先设置你的 API Key")

try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()

DEBUG_DIR = os.path.join(BASE_DIR, "debug_logs")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
STOCK_DIR = os.path.join(ASSETS_DIR, "stock")
BACKGROUND_DIR = os.path.join(ASSETS_DIR, "background")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

for path in [DEBUG_DIR, ASSETS_DIR, STOCK_DIR, BACKGROUND_DIR, OUTPUT_DIR]:
    os.makedirs(path, exist_ok=True)