"""
图片转结构化MD文档工具
功能：批量将指定文件夹中的图片通过AI分析转换为结构化Markdown文档
"""
import os
import time
import re
import warnings
from PIL import Image
import json

warnings.filterwarnings('ignore')

try:
    import dashscope
    DASHSCOPE_SDK_AVAILABLE = True
except ImportError:
    DASHSCOPE_SDK_AVAILABLE = False
    import requests

try:
    from openai import OpenAI
    OPENAI_SDK_AVAILABLE = True
except ImportError:
    OPENAI_SDK_AVAILABLE = False

SUPPORTED_FORMATS = ['.png', '.jpg', '.jpeg', '.bmp', '.gif']

SECTION_KEYWORDS = [
    "需求背景", "需求说明", "新增说明", "功能说明", "页面说明", "功能描述",
    "需求描述", "产品说明", "业务说明", "规则说明", "处理办法", "显示规则",
    "后台", "前端", "安全提示", "变更时间",
    "活动说明", "活动时间", "活动描述", "奖励说明", "规则", "参与方式", "奖励", "活动规则",
    "功能", "模块", "更多", "PS"
]

BAILIANT_API_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
BAILIANT_API_KEY = ""  # TODO: 在这里配置你的 API Key
BAILIANT_MODEL = "qwen3-vl-flash"

API_TIMEOUT = 120
MAX_RETRIES = 2
RETRY_DELAY = 5


def sanitize_filename(name):
    name = name.replace("\n", " ").replace("\r", " ")
    name = " ".join(name.split())
    return name.replace("/", "-").replace("\\", "-").replace(":", "-").replace("*", "-") \
                .replace("?", "-").replace("\"", "-").replace("<", "-").replace(">", "-") \
                .replace("|", "-").strip()

def get_image_files(folder_path):
    image_files = []
    for file in os.listdir(folder_path):
        ext = os.path.splitext(file)[1].lower()
        if ext in SUPPORTED_FORMATS:
            image_files.append(os.path.join(folder_path, file))
    return image_files

_interrupt_flag = False

def setup_signal_handler():
    import signal
    import sys
    
    def signal_handler(sig, frame):
        global _interrupt_flag
        _interrupt_flag = True
        print("\n\n检测到中断信号，正在停止...")
        sys.exit(0)
    
    signal.signal(signal.SIGINT, signal_handler)

def check_interrupt():
    global _interrupt_flag
    return _interrupt_flag

def get_api_key():
    api_key = BAILIANT_API_KEY
    
    if not api_key:
        print("错误：未配置 API key")
        print("请在代码中设置 BAILIANT_API_KEY 的值")
        return None
    
    return api_key

def analyze_image_with_ai(image_path, timeout=None, max_retries=None):
    global _interrupt_flag
    
    if timeout is None:
        timeout = API_TIMEOUT
    if max_retries is None:
        max_retries = MAX_RETRIES
    
    retry_count = 0
    
    while retry_count <= max_retries:
        try:
            if check_interrupt():
                print("操作已中断")
                return ""
            
            print(f"正在分析图片：{os.path.basename(image_path)} (超时：{timeout}秒)")
            
            api_key = get_api_key()
            if not api_key:
                raise Exception("未配置 API key")
            
            model = BAILIANT_MODEL
            
            prompt = """你是专业的需求文档提取助手。你的任务是**忠实、完整**地将图片内容转换为结构化文档。

【核心原则】
1. 不编造：不添加任何图片中不存在的内容、文字或数据
2. 不遗漏：完整提取所有可见的文字、标注、编号、符号
3. 保留原文：严格按照原图的表述方式，不改写、不翻译
4. 保留结构：还原原图的排版、层级、顺序、分组关系

【输出格式】
输出JSON对象，字段根据图片内容自适应：
- 文档标题
- 内容区块（按原图顺序/层级组织）
- 标注信息（序号、箭头标记等）
- 页面元素（按钮、输入框等）

请直接输出JSON，不要有其他内容。"""
            
            print(f"正在调用 AI API (第 {retry_count + 1}/{max_retries + 1} 次尝试)...")
            print(f"按 Ctrl+C 可中断当前操作")
            
            if OPENAI_SDK_AVAILABLE:
                client = OpenAI(
                    api_key=api_key,
                    base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
                    timeout=timeout
                )
                
                import base64
                with open(image_path, "rb") as f:
                    image_data = f.read()
                base64_image = base64.b64encode(image_data).decode('utf-8')
                
                try:
                    response = client.chat.completions.create(
                        model=model,
                        messages=[
                            {
                                "role": "user",
                                "content": [
                                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{base64_image}"}},
                                    {"type": "text", "text": prompt}
                                ]
                            }
                        ],
                        temperature=0.3,
                        max_tokens=4096
                    )
                    analysis_result = response.choices[0].message.content
                except Exception as api_error:
                    if "timeout" in str(api_error).lower() or "timed out" in str(api_error).lower():
                        raise TimeoutError(f"API 调用超时（超过{timeout}秒）")
                    elif "connection" in str(api_error).lower() or "network" in str(api_error).lower():
                        raise ConnectionError(f"网络连接错误：{str(api_error)}")
                    else:
                        raise api_error
            else:
                raise Exception("请安装 openai SDK: pip install openai")
            
            if not analysis_result:
                raise Exception("API 返回空结果")
            
            print("分析完成")
            return analysis_result
            
        except TimeoutError as e:
            print(f"\n警告：{e}")
            retry_count += 1
            if retry_count <= max_retries:
                print(f"将在 {RETRY_DELAY} 秒后重试... (按 Ctrl+C 中断)")
                time.sleep(RETRY_DELAY)
            else:
                print("已达到最大重试次数，跳过此图片")
                return ""
        
        except ConnectionError as e:
            print(f"\n警告：{e}")
            retry_count += 1
            if retry_count <= max_retries:
                print(f"将在 {RETRY_DELAY} 秒后重试... (按 Ctrl+C 中断)")
                time.sleep(RETRY_DELAY)
            else:
                print("已达到最大重试次数，跳过此图片")
                return ""
        
        except KeyboardInterrupt:
            print("\n\n用户中断操作")
            return ""
        
        except Exception as e:
            print(f"分析图片时出错：{e}")
            retry_count += 1
            if retry_count <= max_retries:
                print(f"将在 {RETRY_DELAY} 秒后重试... (按 Ctrl+C 中断)")
                time.sleep(RETRY_DELAY)
            else:
                print("已达到最大重试次数，跳过此图片")
                return ""
    
    return ""

def structure_content(raw_content):
    if not raw_content:
        return {"未分类": raw_content}
    
    try:
        import json
        import re
        
        json_match = re.search(r'\{[\s\S]*\}', raw_content)
        if json_match:
            json_str = json_match.group()
            data = json.loads(json_str)
            
            field_mapping = {
                "doc_title": "文档标题", "文档标题": "文档标题",
                "background": "需求背景", "需求背景": "需求背景",
                "modules": "功能模块", "功能模块": "功能模块",
                "global_info": "全局说明", "全局说明": "全局说明",
                "module_name": "模块名称", "模块名称": "模块名称",
                "description": "模块描述", "模块描述": "模块描述",
                "changes": "变更点", "变更点": "变更点",
                "page_structure": "页面结构", "页面结构": "页面结构",
                "page_name": "页面名称", "页面名称": "页面名称",
                "page_description": "页面描述", "页面描述": "页面描述",
                "ui_components": "UI 组件", "UI 组件": "UI 组件",
                "interactions": "交互说明", "交互说明": "交互说明",
                "related_docs": "相关文档", "相关文档": "相关文档",
                "component_type": "组件类型", "组件类型": "组件类型",
                "component_name": "组件名称", "组件名称": "组件名称",
                "properties": "属性", "属性": "属性",
                "rules": "规则", "规则": "规则",
                "notes": "备注", "备注": "备注",
                "business_rules": "业务规则", "业务规则": "业务规则",
                "attention": "注意事项", "注意事项": "注意事项"
            }
            
            def standardize_keys(obj):
                if isinstance(obj, dict):
                    result = {}
                    for k, v in obj.items():
                        new_key = field_mapping.get(k, k)
                        result[new_key] = standardize_keys(v)
                    return result
                elif isinstance(obj, list):
                    return [standardize_keys(item) for item in obj]
                else:
                    return obj
            
            return standardize_keys(data)
            
    except (json.JSONDecodeError, AttributeError) as e:
        print(f"JSON 解析失败：{e}")
    
    skip_patterns = [
        r"^\d+%$", r"^画布\(.*\)$", r"^画布（.*）$", r"^\d+$",
        r"^暂无批注$", r"^\d{2}-.+", r"^（废纸）$",
        r"^展开全部$", r"^展开$", r"^[\u4e00-\u9fa5]+V\d+\.\d+\.\d+$",
    ]
    
    lines = raw_content.strip().split('\n')
    result = {}
    current_section = "内容"
    current_lines = []
    
    for line in lines:
        line = line.strip()
        if not line:
            if current_lines:
                current_lines.append("")
            continue
        
        if any(re.match(pattern, line) for pattern in skip_patterns):
            continue
        
        is_section, keyword, remaining = parse_section(line)
        
        if is_section:
            if current_lines:
                filtered_lines = [l for l in current_lines if l.strip()]
                if filtered_lines:
                    result[current_section] = '\n'.join(filtered_lines)
                current_lines = []
            
            current_section = keyword
            if remaining:
                current_lines.append(remaining)
        else:
            current_lines.append(line)
    
    if current_lines:
        filtered_lines = [l for l in current_lines if l.strip()]
        if filtered_lines:
            result[current_section] = '\n'.join(filtered_lines)
    
    result = {k: v.strip() for k, v in result.items() if v.strip()}
    if "内容" in result:
        content_val = result["内容"].strip()
        if len(content_val) < 5 or re.match(r"^[\u4e00-\u9fa5]+V\d+", content_val):
            del result["内容"]
    
    return result

def parse_section(line):
    for keyword in SECTION_KEYWORDS:
        if line == keyword:
            return True, keyword, ""
        if line.startswith(keyword + "：") or line.startswith(keyword + ":"):
            remaining = line[len(keyword):].strip()
            if remaining.startswith("：") or remaining.startswith(":"):
                remaining = remaining[1:].strip()
            return True, keyword, remaining
        if line.startswith("-" + keyword) or line.startswith("—" + keyword):
            remaining = line[len(keyword)+1:].strip()
            return True, keyword, remaining
    
    return False, "", ""

def generate_structured_md(title, structured_data, screenshot_name):
    md_lines = [f"# {title}\n"]
    
    field_mapping = {
        "doc_title": "文档标题", "文档标题": "文档标题",
        "background": "需求背景", "需求背景": "需求背景",
        "modules": "功能模块", "功能模块": "功能模块",
        "global_info": "全局说明", "全局说明": "全局说明",
        "module_name": "模块名称", "模块名称": "模块名称",
        "description": "模块描述", "模块描述": "模块描述",
        "changes": "变更点", "变更点": "变更点",
        "page_structure": "页面结构", "页面结构": "页面结构",
        "page_name": "页面名称", "页面名称": "页面名称",
        "page_description": "页面描述", "页面描述": "页面描述",
        "ui_components": "UI 组件", "UI 组件": "UI 组件",
        "interactions": "交互说明", "交互说明": "交互说明",
        "related_docs": "相关文档", "相关文档": "相关文档",
        "component_type": "组件类型", "组件类型": "组件类型",
        "component_name": "组件名称", "组件名称": "组件名称",
        "properties": "属性", "属性": "属性",
        "rules": "规则", "规则": "规则",
        "notes": "备注", "备注": "备注",
        "business_rules": "业务规则", "业务规则": "业务规则",
        "attention": "注意事项", "注意事项": "注意事项"
    }
    
    def translate_field(field_name):
        return field_mapping.get(field_name, field_name.replace("_", " "))
    
    def format_value(key, value, indent_level=0):
        lines = []
        indent = "  " * indent_level
        
        if isinstance(value, str):
            if value.strip():
                lines.append(f"{indent}- **{translate_field(key)}**: {value}\n")
        
        elif isinstance(value, list):
            if not value:
                return lines
            
            if isinstance(value[0], dict):
                for idx, item in enumerate(value, 1):
                    item_title = item.get('组件名称') or item.get('component_name') or \
                                item.get('模块名称') or item.get('module_name') or \
                                item.get('页面名称') or item.get('page_name') or \
                                item.get('文档名') or item.get('文档名称') or f"项{idx}"
                    
                    lines.append(f"\n{indent}### {item_title}\n")
                    
                    for k, v in item.items():
                        if k in ['组件名称', 'component_name', '模块名称', 'module_name', 
                                '页面名称', 'page_name', '文档名', '文档名称']:
                            continue
                        
                        if isinstance(v, str):
                            lines.append(f"{indent}- **{translate_field(k)}**: {v}\n")
                        elif isinstance(v, list):
                            if v and isinstance(v[0], dict):
                                for sub_item in v:
                                    sub_title = sub_item.get('element') or sub_item.get('name') or str(sub_item)
                                    lines.append(f"{indent}- **{sub_title}**\n")
                                    for sk, sv in sub_item.items():
                                        if sk not in ['element', 'name']:
                                            lines.append(f"{indent}  - {translate_field(sk)}: {sv}\n")
                            else:
                                items_str = ', '.join(str(x) for x in v)
                                lines.append(f"{indent}- **{translate_field(k)}**: {items_str}\n")
                        elif isinstance(v, dict):
                            lines.append(f"{indent}- **{translate_field(k)}**:\n")
                            for nk, nv in v.items():
                                lines.append(f"{indent}  - {translate_field(nk)}: {nv}\n")
                        elif isinstance(v, (int, float, bool)):
                            lines.append(f"{indent}- **{translate_field(k)}**: {v}\n")
            else:
                items_str = '  \n'.join(f"{indent}- {item}" for item in value)
                lines.append(f"{items_str}\n")
        
        elif isinstance(value, dict):
            for k, v in value.items():
                if isinstance(v, str):
                    lines.append(f"{indent}- **{translate_field(k)}**: {v}\n")
                elif isinstance(v, list):
                    if v and isinstance(v[0], dict):
                        for sub_item in v:
                            sub_title = sub_item.get('组件名称') or sub_item.get('element') or \
                                       sub_item.get('name') or str(sub_item)
                            lines.append(f"\n{indent}#### {sub_title}\n")
                            for sk, sv in sub_item.items():
                                if sk not in ['组件名称', 'element', 'name']:
                                    if isinstance(sv, str):
                                        lines.append(f"{indent}- {translate_field(sk)}: {sv}\n")
                                    elif isinstance(sv, (int, float, bool)):
                                        lines.append(f"{indent}- {translate_field(sk)}: {sv}\n")
                                    elif isinstance(sv, list):
                                        items_str = ', '.join(str(x) for x in sv)
                                        lines.append(f"{indent}- {translate_field(sk)}: {items_str}\n")
                    else:
                        items_str = ', '.join(str(x) for x in v)
                        lines.append(f"{indent}- **{translate_field(k)}**: {items_str}\n")
                elif isinstance(v, dict):
                    lines.append(f"{indent}- **{translate_field(k)}**:\n")
                    for nk, nv in v.items():
                        if isinstance(nv, str):
                            lines.append(f"{indent}  - {translate_field(nk)}: {nv}\n")
                        elif isinstance(nv, (int, float, bool)):
                            lines.append(f"{indent}  - {translate_field(nk)}: {nv}\n")
                elif v:
                    lines.append(f"{indent}- **{translate_field(k)}**: {v}\n")
        
        elif isinstance(value, (int, float, bool)):
            lines.append(f"{indent}- **{translate_field(key)}**: {value}\n")
        
        return lines
    
    for key, value in structured_data.items():
        if not value:
            continue
        
        section_title = translate_field(key)
        
        md_lines.append(f"\n## {section_title}\n")
        
        formatted_lines = format_value(key, value, indent_level=0)
        md_lines.extend(formatted_lines)
    
    md_lines.extend([f"\n---\n", f"\n## 截图\n", f"\n![{title}]({screenshot_name})\n"])
    
    return '\n'.join(md_lines)

def format_content_lines(content):
    formatted_lines = []
    content_lines = content.split('\n')
    
    for line in content_lines:
        line = line.strip()
        if not line:
            formatted_lines.append("")
            continue
        
        is_list_item = (
            re.match(r'^\d+[、.。]', line) or
            re.match(r'^[-+*]\s+', line) or
            re.match(r'^[a-zA-Z][.、]\s*', line) or
            re.match(r'^[•●○■□◆◇►▶•◦▪▸▹›»]', line)
        )
        
        if is_list_item:
            formatted_lines.append(f"- {line}")
        else:
            formatted_lines.append(line)
    
    return formatted_lines

def image_to_md_batch(input_folder, output_folder, debug_mode=False):
    setup_signal_handler()
    
    if not os.path.exists(input_folder):
        print(f"错误：输入文件夹 {input_folder} 不存在")
        return
    
    os.makedirs(output_folder, exist_ok=True)
    
    image_files = get_image_files(input_folder)
    if not image_files:
        print(f"错误：文件夹 {input_folder} 中没有图片文件")
        return
    
    print(f"找到 {len(image_files)} 张图片，开始处理...")
    print(f"提示：按 Ctrl+C 可随时中断处理过程\n")
    
    success_count = 0
    failed_count = 0
    skip_count = 0
    processed_files = []
    
    for i, image_path in enumerate(image_files, 1):
        try:
            if check_interrupt():
                print("\n操作已中断，停止处理")
                break
            
            print(f"\n{'='*60}")
            print(f"处理第 {i}/{len(image_files)} 张图片：{os.path.basename(image_path)}")
            print(f"{'='*60}")
            
            image_name = os.path.basename(image_path)
            title = os.path.splitext(image_name)[0]
            md_filename = f"{sanitize_filename(title)}.md"
            md_path = os.path.join(output_folder, md_filename)
            
            if os.path.exists(md_path) and not debug_mode:
                print(f"跳过：{md_filename} (文档已存在)")
                skip_count += 1
                processed_files.append(image_path)
                continue
            
            analysis_result = analyze_image_with_ai(image_path)
            
            if not analysis_result:
                print(f"跳过图片：{os.path.basename(image_path)} (分析失败)")
                failed_count += 1
                continue
            
            structured_data = structure_content(analysis_result)
            
            image_name = os.path.basename(image_path)
            title = os.path.splitext(image_name)[0]
            
            screenshot_name = os.path.join("..", "images", image_name).replace('\\', '/')
            md_content = generate_structured_md(title, structured_data, screenshot_name)
            
            md_filename = f"{sanitize_filename(title)}.md"
            md_path = os.path.join(output_folder, md_filename)
            
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_content)
            
            print(f"✓ 已生成：{md_filename}")
            success_count += 1
            processed_files.append(image_path)
            
        except KeyboardInterrupt:
            print("\n\n用户中断操作")
            break
        
        except Exception as e:
            print(f"处理图片 {os.path.basename(image_path)} 时出错：{e}")
            failed_count += 1
            continue
    
    if processed_files:
        generate_index_file(output_folder, processed_files)
    
    print(f"\n{'='*60}")
    print(f"处理完成！")
    print(f"总计：{len(image_files)} 张")
    print(f"成功：{success_count} 张")
    print(f"跳过：{skip_count} 张")
    print(f"失败：{failed_count} 张")
    print(f"{'='*60}")

def generate_index_file(output_folder, image_files):
    import time
    
    index_content = "# 图片分析索引\n\n"
    index_content += f"**生成时间**: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n"
    index_content += "---\n\n"
    index_content += "## 目录\n\n"
    
    for image_path in image_files:
        image_name = os.path.basename(image_path)
        title = os.path.splitext(image_name)[0]
        md_filename = f"{sanitize_filename(title)}.md"
        index_content += f"- [{title}]({md_filename})\n"
    
    index_path = os.path.join(output_folder, "index.md")
    with open(index_path, "w", encoding="utf-8") as f:
        f.write(index_content)
    
    print(f"已生成索引文件：index.md")

def process_images(input_folder, output_folder, debug_mode=False):
    print("=" * 60)
    print("图片转结构化MD文档工具")
    print("=" * 60)
    
    image_to_md_batch(input_folder, output_folder, debug_mode)

if __name__ == "__main__":
    import sys
    
    input_folder = 'images'
    output_folder = 'ai-md'
    debug_mode = False
    
    if len(sys.argv) >= 2:
        project_root = sys.argv[1]
        
        if len(sys.argv) >= 3 and sys.argv[2].lower() == 'debug':
            debug_mode = True
            print("⚠️  警告：Debug 模式将覆盖已有文件")
        
        input_folder = os.path.join(project_root, 'images')
        output_folder = os.path.join(project_root, 'ai-md')
        
        print(f"项目根目录：{project_root}")
        print(f"输入目录：{input_folder}")
        print(f"输出目录：{output_folder}")
    
    process_images(input_folder, output_folder, debug_mode)
