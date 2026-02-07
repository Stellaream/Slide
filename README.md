# AI PPT Generator - 开发者指南

这是一个基于 **大语言模型 (DeepSeek/OpenAI)** 与 **Python-pptx** 构建的智能 PPT 生成系统。它能将 Word 文档自动转化为排版精美、图文并茂的演示文稿。

## 核心特性

1.  **全链路自动化**：Docx 解析 -> 结构化拆分 -> AI 大纲规划 -> 页面内容生成 -> PPT 渲染。
2.  **智能语义配图**：
    *   用户上传图片并输入描述（如“团队合照”）。
    *   系统利用 **LLM 上帝视角**，理解 PPT 每一页的内容，自动将最匹配的图片“指派”给对应的页面。
3.  **三级资源兜底**：
    *   **Level 1**: 用户上传图片 (语义精准匹配)。
    *   **Level 2**: 本地素材库 (随机不重复填充，保证不留白)。
    *   **Level 3**: 灰色占位符 (最后的防线)。
4.  **Bento Grid 排版**：生成的 PPT 采用现代化的卡片式布局，拒绝大段文字堆砌。
5.  **版式自适应修复**：利用 Windows COM 接口，模拟人工操作触发 PPT 的自动排版引擎，修复文字溢出问题。

---

## 📂 项目结构说明

```text
Project_Root/
├── app.py                  # [入口] Flask 后端，处理前端请求与流式响应
├── index.html              # [入口] 前端交互界面
├── config.py               # [配置] API Key, 路径常量
├── utils.py                # [工具] 通用工具函数
│
├── assets/                 # [资源] 本地素材库
│   ├── background/         # PPT 背景图 (.jpg/.png)
│   └── stock/              # 商务/科技类兜底插图
│
├── core/                   # [核心] 业务逻辑层
│   ├── pipeline.py         # 串联解析、生成、渲染全流程
│   ├── llm.py              # 负责与 DeepSeek/OpenAI 交互
│   ├── content.py          # 解析 Word 文档
│   └── ...
│
└── engine/                 # [引擎] 底层执行层
    ├── renderer.py         # 画笔：基于 python-pptx 绘制幻灯片
    ├── image_manager.py    # 调度员：负责图片的语义匹配与分发
    └── size.py             # 维修工：调用 Windows PPT 接口修复版式
```

---

## 快速开始

### 1. 环境准备
*   **OS**: 推荐 Windows (为了支持版式修复)，Linux/Mac 仅支持基础生成。
*   **Python**: 3.8+
*   **Office**: 必须安装 Microsoft PowerPoint。

### 2. 安装依赖
```bash
pip install flask flask-cors python-pptx openai mammoth pywin32
```

### 3. 配置 API
打开 `config.py`，填入你的大模型 Key：
```python
API_KEY = "sk-xxxxxxxx"
BASE_URL = "https://api.deepseek.com" # 或 OpenAI 地址
```

### 4. 准备素材
在项目根目录创建 `assets` 文件夹：
*   把 **背景图** 放入 `assets/background/`。
*   把 **通用插图** (如会议、握手、服务器等) 放入 `assets/stock/`。

### 5. 运行
启动后端服务：
```bash
python app.py
```
然后直接在浏览器打开 `index.html` 即可使用。

---