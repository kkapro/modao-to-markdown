"""
墨刀页面采集工具 v2
功能：采集墨刀原型页面，生成结构化 Markdown 文档和截图
"""
from playwright.sync_api import sync_playwright
from PIL import Image
import os
import time
import re
from multiprocessing import Pool, cpu_count

# ==================== 常量定义 ====================

# 文件命名过滤关键词
CANVAS_SKIP_KEYWORDS = [
    "总览", "演示", "标注", "登录", "免费使用", "页面", "图层", "批注", "评论",
    "需求列表", "欢迎来到墨刀", "立即登录", "分享", "废纸篓", "废纸", "废弃"
]

PAGE_SKIP_KEYWORDS = ["废纸", "废弃"]

# 内容结构化关键词
SECTION_KEYWORDS = [
    "需求背景", "需求说明", "新增说明", "功能说明", "页面说明", "功能描述",
    "需求描述", "产品说明", "业务说明", "规则说明", "处理办法", "显示规则",
    "后台", "前端","安全提示","活动说明", "活动时间", "更多", "PS"
]

# 截图配置
VIEWPORT_WIDTH = 1920
VIEWPORT_HEIGHT = 1080
HEADER_OFFSET = 60
SCROLL_STEP = 550
RIGHT_SIDEBAR_WIDTH = 280
MIN_ZOOM_PERCENT = 60

# ==================== 工具函数 ====================

def sanitize_filename(name):
    """清理文件名，移除特殊字符"""
    name = name.replace("\n", " ").replace("\r", " ")
    name = " ".join(name.split())
    return name.replace("/", "-").replace("\\", "-").replace(":", "-").replace("*", "-") \
                .replace("?", "-").replace("\"", "-").replace("<", "-").replace(">", "-") \
                .replace("|", "-").strip()

def should_skip_item(text, skip_keywords):
    """检查是否应该跳过该项目"""
    if not text:
        return True
    for keyword in skip_keywords:
        if keyword in text:
            return True
    return False

# ==================== 内容处理函数 ====================

def structure_content(raw_content):
    """将原始内容转换为结构化数据"""
    if not raw_content:
        return {"未分类": raw_content}
    
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
        
        # 跳过不需要的内容
        if any(re.match(pattern, line) for pattern in skip_patterns):
            continue
        
        # 检查是否是新段落的开始
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
    
    # 保存最后一段内容
    if current_lines:
        filtered_lines = [l for l in current_lines if l.strip()]
        if filtered_lines:
            result[current_section] = '\n'.join(filtered_lines)
    
    # 清理空内容和无效的"内容"段落
    result = {k: v.strip() for k, v in result.items() if v.strip()}
    if "内容" in result:
        content_val = result["内容"].strip()
        if len(content_val) < 5 or re.match(r"^[\u4e00-\u9fa5]+V\d+", content_val):
            del result["内容"]
    
    return result

def parse_section(line):
    """解析行是否为 section 开始"""
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
    """生成结构化的 Markdown 文档"""
    md_lines = [f"# {title}\n"]
    
    # 按优先级输出各个部分
    priority_sections = ['需求背景', '需求说明', '页面说明', '功能说明', '业务说明']
    
    for section in priority_sections:
        if section in structured_data:
            content = structured_data[section].strip()
            md_lines.extend([f"\n## {section}\n", f"{content}\n"])
    
    # 输出其他部分
    for section, content in structured_data.items():
        if section in priority_sections or not content.strip():
            continue
        
        md_lines.append(f"\n## {section}\n")
        formatted_lines = format_content_lines(content)
        md_lines.extend(formatted_lines)
    
    # 添加截图部分
    md_lines.extend([f"\n---\n", f"\n## 截图\n", f"\n![{title}]({screenshot_name})\n"])
    
    return '\n'.join(md_lines)

def format_content_lines(content):
    """格式化内容行，智能识别列表"""
    formatted_lines = []
    content_lines = content.split('\n')
    
    for line in content_lines:
        line = line.strip()
        if not line:
            formatted_lines.append("")
            continue
        
        # 智能识别列表项
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

# ==================== 页面操作函数 ====================

def find_canvas_list(page):
    """查找左侧画布列表"""
    canvas_items = []
    
    try:
        lis = page.query_selector_all("li")
        for li in lis:
            try:
                text = li.inner_text().strip()
                if not text:
                    continue
                    
                first_line = text.split('\n')[0].strip()
                if not first_line or len(first_line) >= 50:
                    continue
                
                # 过滤无效项
                if first_line.isdigit() or should_skip_item(first_line, CANVAS_SKIP_KEYWORDS):
                    continue
                
                if first_line not in canvas_items:
                    canvas_items.append(first_line)
            except:
                pass
    except Exception as e:
        print(f"查找画布列表失败：{e}")
    
    return canvas_items

def click_canvas(page, canvas_name):
    """点击左侧画布"""
    try:
        lis = page.query_selector_all("li")
        for li in lis:
            try:
                text = li.inner_text().strip()
                first_line = text.split('\n')[0].strip() if text else ""
                if first_line == canvas_name or canvas_name in first_line:
                    li.click()
                    print(f"已点击画布：{canvas_name}")
                    try:
                        page.wait_for_load_state('networkidle', timeout=5000)
                    except:
                        time.sleep(1)
                    return True
            except:
                pass
    except Exception as e:
        print(f"点击画布失败：{e}")
    return False

def get_page_list_in_canvas(page):
    """获取当前画布下的页面列表"""
    page_list = []
    
    try:
        elements = page.query_selector_all(".canvas-sortable-list > ul > li")
        for el in elements:
            try:
                text = el.inner_text().strip()
                data_cid = el.get_attribute("data-cid")
                if text:
                    page_name = text.replace('\n', '-').strip()[:50]
                    if not should_skip_item(page_name, PAGE_SKIP_KEYWORDS):
                        page_list.append({"name": page_name, "data_cid": data_cid})
            except:
                pass
    except Exception as e:
        print(f"获取页面列表失败：{e}")
    
    return page_list

def click_page_in_canvas(page, page_index):
    """点击画布下的页面"""
    try:
        elements = page.query_selector_all(".canvas-sortable-list > ul > li")
        if elements and len(elements) > page_index:
            elements[page_index].click()
            try:
                page.wait_for_load_state('networkidle', timeout=3000)
            except:
                time.sleep(0.5)
            page.mouse.move(0, 0)
            time.sleep(0.3)
            print(f"已点击第 {page_index + 1} 个页面")
            return True
    except Exception as e:
        print(f"点击页面失败：{e}")
    return False

def adjust_zoom_if_needed(page):
    """检查并调整缩放比例，如果小于 60% 则调整到 60%"""
    try:
        body = page.query_selector("body")
        if not body:
            return
            
        text = body.inner_text()
        match = re.search(r'(\d+)%', text)
        if not match:
            return
            
        current_percent = int(match.group(1))
        print(f"当前缩放比例：{current_percent}%")
        
        if current_percent < MIN_ZOOM_PERCENT:
            while current_percent < MIN_ZOOM_PERCENT:
                page.keyboard.press("Control+=")
                time.sleep(0.5)
                body = page.query_selector("body")
                if not body:
                    break
                    
                text = body.inner_text()
                match = re.search(r'(\d+)%', text)
                if match:
                    new_percent = int(match.group(1))
                    if new_percent == current_percent:
                        break
                    current_percent = new_percent
                    print(f"调整后缩放比例：{current_percent}%")
            time.sleep(1)
    except Exception as e:
        print(f"调整缩放失败：{e}")

def click_to_blur_focus(page):
    """点击页面空白处移开焦点"""
    try:
        page.click(".tree-node.rResCanvas", offset={"x": 10, "y": 10})
        time.sleep(0.2)
    except:
        pass

# ==================== 截图处理函数 ====================

def get_zoom_transform(page):
    """获取 .zoom-area 的 transform 信息"""
    try:
        return page.evaluate("""
            () => {
                const zoom = document.querySelector('.zoom-area');
                if (!zoom) return null;
                const style = window.getComputedStyle(zoom);
                const transform = style.transform;
                const matrix = new DOMMatrix(transform);
                const scale = style.transform.split('(')[1]?.split(',')[0] || '1';
                return {
                    translateX: matrix.m41,
                    translateY: matrix.m42,
                    scale: parseFloat(scale)
                };
            }
        """)
    except:
        return None

def crop_canvas_region(img, box, viewport_height=VIEWPORT_HEIGHT, header_offset=HEADER_OFFSET):
    """裁剪画布区域，返回裁剪后的图像"""
    left = int(box['x'])
    top = int(box['y'])
    width = int(box['width'])
    height = int(box['height'])
    
    # 边界检查
    if left < 0:
        left = 0
    if left + width > img.width:
        width = img.width - left
    
    # # 计算裁剪区域 - 优化顶部裁剪，减少灰色区域
    # # 如果画布位置在 header_offset 之上，从画布顶部开始裁，但去掉顶部 20px 的灰色边框
    # if top < header_offset:
    #     # 画布在顶部之上，从画布顶部 +20px 开始裁（去掉灰色边框）
    #     actual_crop_top = max(0, top + 20)
    # else:
        # 画布在正常位置，从 header_offset 开始裁
    actual_crop_top = header_offset
    
    # 计算实际内容的底部位置
    content_bottom = top + height
    actual_crop_bottom = min(viewport_height, content_bottom)
    
    # 确保裁剪区域有效
    actual_crop_top = max(0, actual_crop_top)
    actual_crop_bottom = max(actual_crop_top + 1, actual_crop_bottom)
    
    # 使用完整宽度，不去掉右侧内容
    full_width = width
    
    if (actual_crop_bottom - actual_crop_top) <= 0:
        return None
    
    return img.crop((left, actual_crop_top, left + full_width, actual_crop_bottom))

def stitch_images(images, positions, current_scale):
    """拼接多张截图 - 使用改进的特征匹配算法和平滑融合
    
    核心思路：
    1. 使用图像特征点匹配（类似 SIFT/ORB 的简化版）
    2. 自动计算最佳重叠区域
    3. 平滑融合，消除接缝
    
    Args:
        images: 截图列表
        positions: 位置信息列表
        current_scale: 当前缩放比例
    
    Returns:
        拼接后的完整图像
    """
    try:
        if not images:
            print("  错误：图像列表为空")
            return None
        
        # 检查图像列表是否有效
        valid_images = []
        for img in images:
            if img and hasattr(img, 'width') and hasattr(img, 'height'):
                valid_images.append(img)
            else:
                print("  警告：跳过无效图像")
        
        if not valid_images:
            print("  错误：没有有效的图像")
            return None
        
        if len(valid_images) == 1:
            return valid_images[0]
        
        # 检查所有图像的宽度是否一致
        result_width = valid_images[0].width
        for i, img in enumerate(valid_images[1:]):
            if img.width != result_width:
                print(f"  警告：图像 {i+1} 宽度不一致，使用第一张图像的宽度")
                # 调整图像宽度
                if img.width > result_width:
                    img = img.crop((0, 0, result_width, img.height))
                else:
                    # 创建新图像并填充空白
                    new_img = Image.new('RGB', (result_width, img.height), color='white')
                    new_img.paste(img, (0, 0))
                    img = new_img
                valid_images[i+1] = img
        
        total_height = 0
        paste_positions = [0]  # 第一张图从 0 开始
        overlap_heights = []  # 存储重叠高度
        
        for i in range(1, len(valid_images)):
            prev_img = valid_images[i - 1]
            curr_img = valid_images[i]
            
            try:
                # 计算最佳重叠区域
                overlap_height = find_optimal_overlap(prev_img, curr_img)
                overlap_heights.append(overlap_height)
                
                print(f"  图像 {i-1} 和 {i} 的重叠高度：{overlap_height}px (前图高度：{prev_img.height}, 当前图高度：{curr_img.height})")
                
                # 改进：即使重叠为 0，也直接拼接，确保内容不丢失
                if overlap_height > 0:
                    # 有重叠，减去重叠部分
                    total_height += prev_img.height - overlap_height
                    print(f"    → 使用重叠拼接：{prev_img.height} - {overlap_height} = {prev_img.height - overlap_height}")
                else:
                    # 无重叠，直接拼接（可能有轻微间隙，但保证内容完整）
                    total_height += prev_img.height
                    print(f"    → 直接拼接：{prev_img.height}")
                
                paste_positions.append(total_height)
            except Exception as e:
                print(f"  警告：计算重叠区域时出错：{e}")
                # 出错时直接拼接
                total_height += prev_img.height
                overlap_heights.append(0)
                paste_positions.append(total_height)
        
        # 加上最后一张图的高度
        total_height += valid_images[-1].height
        
        # 检查总高度是否合理
        if total_height <= 0:
            print("  错误：总高度计算错误")
            return None
        
        print(f"拼接结果：总高度={total_height}px, 宽度={result_width}px, 图像数量={len(valid_images)}")
        
        try:
            # 创建结果图像
            result = Image.new('RGB', (result_width, total_height), color='white')
            
            # 粘贴所有图像并应用融合
            for i, img in enumerate(valid_images):
                if i == 0:
                    # 第一张图像直接粘贴
                    result.paste(img, (0, paste_positions[i]))
                else:
                    # 计算当前图像的位置
                    current_pos = paste_positions[i]
                    
                    # 获取重叠高度
                    overlap_height = overlap_heights[i-1]
                    
                    if overlap_height > 0:
                        try:
                            # 有重叠，应用融合
                            prev_img = valid_images[i-1]
                            prev_pos = paste_positions[i-1]
                            
                            # 计算融合区域的起始和结束位置
                            blend_start = prev_pos + prev_img.height - overlap_height
                            blend_end = current_pos + overlap_height
                            
                            # 先粘贴当前图像
                            result.paste(img, (0, current_pos))
                            
                            # 然后在重叠区域应用融合
                            apply_blend(result, prev_img, img, prev_pos, current_pos, overlap_height)
                        except Exception as e:
                            print(f"  警告：应用融合时出错：{e}")
                            # 出错时直接粘贴
                            result.paste(img, (0, current_pos))
                    else:
                        # 无重叠，直接粘贴
                        result.paste(img, (0, current_pos))
            
            return result
        except Exception as e:
            print(f"  错误：创建或处理结果图像时出错：{e}")
            # 尝试返回第一张图像作为 fallback
            return valid_images[0]
    except Exception as e:
        print(f"  错误：拼接图像时出错：{e}")
        import traceback
        traceback.print_exc()
        # 尝试返回第一张有效图像作为 fallback
        if images and hasattr(images[0], 'width') and hasattr(images[0], 'height'):
            return images[0]
        return None

def apply_blend(result, prev_img, curr_img, prev_pos, curr_pos, overlap_height):
    """在重叠区域应用平滑融合
    
    Args:
        result: 结果图像
        prev_img: 前一张图像
        curr_img: 当前图像
        prev_pos: 前一张图像的位置
        curr_pos: 当前图像的位置
        overlap_height: 重叠高度
    """
    if overlap_height <= 0:
        return
    
    width = prev_img.width
    
    # 计算融合区域的范围
    blend_start = prev_pos + prev_img.height - overlap_height
    blend_end = blend_start + overlap_height
    
    # 遍历重叠区域的每一行
    for y in range(overlap_height):
        # 计算当前行在结果图像中的位置
        result_y = blend_start + y
        
        # 计算融合权重（从 1.0 到 0.0 渐变）
        weight = 1.0 - (y / overlap_height)
        
        # 遍历每一列
        for x in range(width):
            # 获取前一张图像对应位置的像素
            prev_pixel = prev_img.getpixel((x, prev_img.height - overlap_height + y))
            
            # 获取当前图像对应位置的像素
            curr_pixel = curr_img.getpixel((x, y))
            
            # 计算融合后的像素值
            blended_pixel = (
                int(prev_pixel[0] * weight + curr_pixel[0] * (1 - weight)),
                int(prev_pixel[1] * weight + curr_pixel[1] * (1 - weight)),
                int(prev_pixel[2] * weight + curr_pixel[2] * (1 - weight))
            )
            
            # 设置融合后的像素
            result.putpixel((x, result_y), blended_pixel)

def find_optimal_overlap(prev_img, curr_img, max_overlap_ratio=0.5, min_overlap_ratio=0.05):
    """查找两张图像之间的最佳重叠区域
    
    使用改进的特征匹配算法，结合多区域分析和边缘检测
    
    Args:
        prev_img: 前一张图像
        curr_img: 当前图像
        max_overlap_ratio: 最大重叠比例（50%）
        min_overlap_ratio: 最小重叠比例（5%）
    
    Returns:
        最佳重叠高度（像素）
    """
    prev_height = prev_img.height
    curr_height = curr_img.height
    width = prev_img.width
    
    # 限制搜索范围
    max_overlap = int(min(prev_height, curr_height) * max_overlap_ratio)
    min_overlap = int(prev_height * min_overlap_ratio)
    
    # 确保最小搜索范围
    max_overlap = max(max_overlap, 50)
    min_overlap = max(min_overlap, 20)  # 确保最小重叠至少为20像素
    
    # 从 prev 图像底部和 curr 图像顶部提取候选重叠区域
    best_overlap = min_overlap
    best_score = -1
    best_debug_info = ""
    
    # 采样步长，平衡精度和性能
    step = 1  # 更细的步长，提高精度
    
    # 多区域分析：分别分析中心区域和边缘区域
    regions = [
        (0.2, 0.8),  # 中心区域
        (0.1, 0.3),  # 左侧区域
        (0.7, 0.9)   # 右侧区域
    ]
    
    # 准备并行处理的任务
    overlap_range = range(min_overlap, max_overlap + 1, step)
    
    # 并行处理阈值
    PARALLEL_THRESHOLD = 100
    
    if len(overlap_range) > PARALLEL_THRESHOLD:
        # 使用并行处理
        print(f"  使用并行处理计算重叠区域（{len(overlap_range)}个候选值）")
        
        # 准备任务参数
        tasks = []
        for overlap in overlap_range:
            # prev 图像底部 overlap 像素
            prev_bottom = prev_img.crop((0, prev_height - overlap, width, prev_height))
            # curr 图像顶部 overlap 像素
            curr_top = curr_img.crop((0, 0, width, overlap))
            tasks.append((prev_bottom, curr_top, regions))
        
        # 使用进程池
        num_processes = min(cpu_count(), 4)  # 限制最大进程数
        with Pool(processes=num_processes) as pool:
            results = pool.starmap(calculate_overlap_score, tasks)
        
        # 处理结果
        for i, (overlap, score) in enumerate(zip(overlap_range, results)):
            if score > best_score:
                best_score = score
                best_overlap = overlap
                best_debug_info = f"overlap={overlap}, score={score:.3f}"
    else:
        # 串行处理
        for overlap in overlap_range:
            # prev 图像底部 overlap 像素
            prev_bottom = prev_img.crop((0, prev_height - overlap, width, prev_height))
            
            # curr 图像顶部 overlap 像素
            curr_top = curr_img.crop((0, 0, width, overlap))
            
            # 计算多区域相似度
            score = calculate_overlap_score(prev_bottom, curr_top, regions)
            
            # 记录最佳匹配
            if score > best_score:
                best_score = score
                best_overlap = overlap
                best_debug_info = f"overlap={overlap}, score={score:.3f}"
    
    # 改进的阈值判断
    if best_score > 0.9:
        # 高置信度，直接使用
        print(f"  找到高置信重叠：{best_debug_info}")
        return best_overlap
    elif best_score > 0.8:
        # 中等置信度，保守使用
        print(f"  中等重叠：{best_debug_info}（保守使用）")
        return int(best_overlap * 0.98)  # 稍微减少一点
    elif best_score > 0.7:
        # 低置信度，但可能有效
        print(f"  低置信重叠：{best_debug_info}（谨慎使用）")
        return int(best_overlap * 0.95)
    elif best_score > 0.3:
        # 低置信度，但可能有效
        print(f"  低置信重叠：{best_debug_info}（谨慎使用）")
        return int(best_overlap * 0.9)
    else:
        # 相似度很低，返回 0（不重叠）
        print(f"  无有效重叠：{best_debug_info}（直接拼接）")
        return 0

def calculate_overlap_score(prev_bottom, curr_top, regions):
    """计算重叠区域的相似度得分
    
    Args:
        prev_bottom: 前一张图像的底部区域
        curr_top: 当前图像的顶部区域
        regions: 分析区域列表
    
    Returns:
        相似度得分
    """
    width, height = prev_bottom.size
    total_score = 0
    region_weights = [0.6, 0.2, 0.2]  # 中心区域权重更高
    
    for i, (start, end) in enumerate(regions):
        region_width = int(width * (end - start))
        region_x = int(width * start)
        
        prev_region = prev_bottom.crop((region_x, 0, region_x + region_width, height))
        curr_region = curr_top.crop((region_x, 0, region_x + region_width, height))
        
        region_score = calculate_image_similarity(prev_region, curr_region)
        total_score += region_score * region_weights[i]
    
    return total_score

def calculate_image_similarity(img1, img2):
    """计算两张图像的相似度（增强版 SSIM）
    
    Args:
        img1: 图像 1
        img2: 图像 2（尺寸应与 img1 相同）
    
    Returns:
        相似度分数（0-1 之间）
    """
    if img1.size != img2.size:
        return 0.0
    
    width, height = img1.size
    
    # 尝试使用 numpy 进行快速计算
    try:
        import numpy as np
        
        arr1 = np.array(img1).astype(float)
        arr2 = np.array(img2).astype(float)
        
        # 方法 1：均方误差 (MSE)
        mse = np.mean((arr1 - arr2) ** 2)
        if mse == 0:
            return 1.0
        
        # 转换为相似度
        max_pixel = 255.0
        mse_similarity = 1.0 / (1.0 + mse / (max_pixel ** 2))
        
        # 方法 2：结构相似性（改进版 SSIM）
        try:
            # 分通道计算 SSIM
            ssim_channels = []
            for c in range(3):  # RGB 三个通道
                mean1 = np.mean(arr1[:,:,c])
                mean2 = np.mean(arr2[:,:,c])
                std1 = np.std(arr1[:,:,c])
                std2 = np.std(arr2[:,:,c])
                
                # 协方差
                covariance = np.mean((arr1[:,:,c] - mean1) * (arr2[:,:,c] - mean2))
                
                # SSIM 公式
                C1 = (0.01 * 255) ** 2
                C2 = (0.03 * 255) ** 2
                
                ssim_c = ((2 * mean1 * mean2 + C1) * (2 * covariance + C2)) / \
                        ((mean1 ** 2 + mean2 ** 2 + C1) * (std1 ** 2 + std2 ** 2 + C2))
                ssim_channels.append(ssim_c)
            
            # 取三个通道的平均值
            ssim = np.mean(ssim_channels)
            
            # 防止 NaN 或负值
            if np.isnan(ssim) or np.isinf(ssim):
                ssim = 0.5  # 降级为中等分数
            
        except Exception as e:
            # SSIM 计算失败，使用 MSE 结果
            ssim = mse_similarity
        
        # 方法 3：边缘相似度
        try:
            from scipy import ndimage
            
            # 简单的边缘检测
            edge1 = np.sqrt(
                ndimage.sobel(arr1[:,:,0], axis=0)**2 + 
                ndimage.sobel(arr1[:,:,0], axis=1)**2
            )
            edge2 = np.sqrt(
                ndimage.sobel(arr2[:,:,0], axis=0)**2 + 
                ndimage.sobel(arr2[:,:,0], axis=1)**2
            )
            
            edge_mse = np.mean((edge1 - edge2) ** 2)
            edge_similarity = 1.0 / (1.0 + edge_mse / (max_pixel ** 2))
        except ImportError:
            edge_similarity = (mse_similarity + ssim) / 2
        
        # 综合得分：SSIM 权重最高，边缘相似度次之，MSE 最后
        final_score = 0.5 * ssim + 0.3 * edge_similarity + 0.2 * mse_similarity
        
        return final_score
        
    except ImportError:
        # 如果没有 numpy，使用改进的像素比较
        pixels1 = list(img1.getdata())
        pixels2 = list(img2.getdata())
        
        if len(pixels1) != len(pixels2):
            return 0.0
        
        # 改进的像素比较，考虑颜色差异和亮度差异
        total_diff = 0
        total_pixels = len(pixels1)
        
        for p1, p2 in zip(pixels1, pixels2):
            # 计算 RGB 差异
            r_diff = abs(p1[0] - p2[0])
            g_diff = abs(p1[1] - p2[1])
            b_diff = abs(p1[2] - p2[2])
            
            # 计算亮度差异
            brightness1 = (p1[0] + p1[1] + p1[2]) / 3
            brightness2 = (p2[0] + p2[1] + p2[2]) / 3
            brightness_diff = abs(brightness1 - brightness2)
            
            # 综合差异
            total_diff += (r_diff + g_diff + b_diff) + brightness_diff * 0.5
        
        # 计算相似度
        max_possible_diff = total_pixels * (255 * 3 + 255 * 0.5)
        if max_possible_diff == 0:
            return 1.0
        
        similarity = 1.0 - (total_diff / max_possible_diff)
        return max(0.0, min(1.0, similarity))

def take_screenshot(page, screenshot_path, data_cid):
    """截图 - 通过 data-cid 匹配元素，滚动拼接完整页面"""
    click_to_blur_focus(page)
    time.sleep(0.3)
    
    # 获取画布元素
    canvas_el = None
    if data_cid:
        try:
            canvas_el = page.query_selector(f'.tree-node.rResCanvas[data-cid="{data_cid}"]')
        except:
            pass
    
    if not canvas_el:
        canvas_el = page.query_selector(".tree-node.rResCanvas")
    
    if not canvas_el:
        print("未找到 .tree-node.rResCanvas 元素")
        return False
    
    try:
        # 记录初始位置和缩放
        initial_transform = get_zoom_transform(page)
        if not initial_transform:
            print("无法获取 .zoom-area 初始位置")
            canvas_el.screenshot(path=screenshot_path)
            return True
        
        recorded_x = initial_transform.get('translateX', 0)
        before_y = initial_transform['translateY']
        recorded_scale = initial_transform.get('scale', 1)
        print(f"记录初始位置：translateX={recorded_x}, translateY={before_y}, scale={recorded_scale}")
        
        # 调整缩放
        adjust_zoom_if_needed(page)
        
        # 获取当前缩放比例
        current_scale = recorded_scale
        body = page.query_selector("body")
        match = re.search(r'(\d+)%', body.inner_text())
        if match:
            current_percent = int(match.group(1))
            current_scale = current_percent / 100.0
            print(f"当前缩放：{current_percent}%, scale={current_scale}")
        
        # 更新位置
        new_transform = get_zoom_transform(page)
        if new_transform:
            recorded_y = new_transform['translateY']
            recorded_x = new_transform['translateX']
            print(f"调整缩放后位置：translateX={recorded_x}, translateY={recorded_y}")

        if recorded_scale<0.6:
            # 重置位置
            page.evaluate(f"""
                () => {{
                    const zoom = document.querySelector('.zoom-area');
                    if (zoom) {{
                        zoom.style.transform = 'translate({recorded_x}px, {before_y}px) scale({current_scale})';
                    }}
                }}
            """)
            time.sleep(0.5)
        
        # 获取画布信息
        box = canvas_el.bounding_box()
        if not box:
            print("无法获取画布元素位置")
            canvas_el.screenshot(path=screenshot_path)
            return True
            
        el_style = canvas_el.get_attribute("style")
        el_height = int(re.search(r'height:\s*(\d+)px', el_style).group(1))
        content_height = int(canvas_el.evaluate("el => el.scrollHeight"))
        
        # 初始化截图列表
        images = []
        positions = []
        temp_screenshot_path = screenshot_path.replace('.png', '_temp.png')
        
        # 判断是否需要滚动截图
        visible_content_height = (VIEWPORT_HEIGHT - HEADER_OFFSET) * current_scale
        content_actual_height = el_height * current_scale
        
        # 增加容差判断：如果超出不多（在 100px 以内），也不需要滚动，直接截图即可
        height_diff = content_actual_height - visible_content_height
        needs_scrolling = height_diff > 100  # 只有当内容超出可视区域 100px 以上才需要滚动
        
        print(f"可视高度：{visible_content_height:.2f}px, 内容高度：{content_actual_height:.2f}px")
        print(f"高度差：{height_diff:.2f}px")
        print(f"是否需要滚动：{needs_scrolling} (超出 {height_diff:.2f}px > 100px)")
        
        # 拍摄初始位置截图
        print("拍摄第 1 张图（初始位置）")
        page.screenshot(path=temp_screenshot_path, full_page=False)
        time.sleep(0.2)
        
        cropped = capture_and_crop(temp_screenshot_path, page, canvas_el, recorded_y, 
                                   images, positions, current_scale)
        if not cropped:
            return False
        
        # 如果不需要滚动，直接保存
        if not needs_scrolling:
            images[0].save(screenshot_path)
            print(f"截图已保存（无需滚动）")
            
            # 恢复初始位置
            page.evaluate(f"""
                () => {{
                    const zoom = document.querySelector('.zoom-area');
                    if (zoom) {{
                        zoom.style.transform = 'translate({recorded_x}px, {recorded_y}px) scale({current_scale})';
                    }}
                }}
            """)
            time.sleep(0.2)
            
            # 清理临时文件
            try:
                if os.path.exists(temp_screenshot_path):
                    os.remove(temp_screenshot_path)
            except:
                pass
            
            return True
        
        # 滚动拍摄后续截图
        current_y = recorded_y
        scroll_count = 0
        
        # 内存优化：限制同时在内存中的图像数量
        MAX_IMAGES_IN_MEMORY = 5
        
        while -(current_y - recorded_y) < el_height * current_scale:
            current_y -= SCROLL_STEP
            
            # 执行滚动
            page.evaluate(f"""
                () => {{
                    const zoom = document.querySelector('.zoom-area');
                    if (zoom) {{
                        zoom.style.transform = 'translate({recorded_x}px, {current_y}px) scale({current_scale})';
                    }}
                }}
            """)
            time.sleep(0.4)
            
            # 截图
            page.screenshot(path=temp_screenshot_path, full_page=False)
            cropped = capture_and_crop(temp_screenshot_path, page, canvas_el, current_y, 
                                       images, positions, current_scale)
            
            if not cropped:
                break
            
            # 内存优化：当图像数量超过阈值时，进行中间拼接
            if len(images) >= MAX_IMAGES_IN_MEMORY:
                print(f"  内存优化：图像数量达到 {MAX_IMAGES_IN_MEMORY}，执行中间拼接")
                # 拼接前半部分
                intermediate_result = stitch_images(images[:-1], positions[:-1], current_scale)
                if intermediate_result:
                    # 保留最后一张图像和新图像进行拼接
                    images = [intermediate_result, images[-1]]
                    positions = [0, intermediate_result.height]
                    print(f"  中间拼接完成，图像数量减少到 {len(images)}")
            
            # 检查是否还有剩余内容
            scroll_offset = (current_y - recorded_y) / current_scale
            remaining_content = content_height - scroll_offset - visible_content_height / current_scale
            
            if remaining_content <= 0:
                print(f"    内容已全部显示，停止滚动")
                break
            
            scroll_count += 1
        
        # 清理临时文件
        try:
            if os.path.exists(temp_screenshot_path):
                os.remove(temp_screenshot_path)
        except:
            pass
        
        # 拼接图像
        if len(images) > 1:
            result = stitch_images(images, positions, current_scale)
            if result:
                result.save(screenshot_path)
                print(f"截图已保存 (滚动拼接，总高度：{result.height}px, 宽度：{result.width}px)")
            else:
                print("图像拼接失败")
                return False
        else:
            images[0].save(screenshot_path)
            print(f"截图已保存")
        
        # 恢复初始位置
        page.evaluate(f"""
            () => {{
                const zoom = document.querySelector('.zoom-area');
                if (zoom) {{
                    zoom.style.transform = 'translate({recorded_x}px, {recorded_y}px) scale({current_scale})';
                }}
            }}
        """)
        time.sleep(0.2)
        
        # 清理内存
        del images
        
        return True
        
    except Exception as e:
        print(f"截图失败：{e}")
        import traceback
        traceback.print_exc()
        return False

def capture_and_crop(temp_path, page, canvas_el, y_position, images, positions, current_scale):
    """截图并裁剪，添加到 images 列表"""
    try:
        img = Image.open(temp_path)
    except Exception as e:
        print(f"读取截图失败：{e}")
        return False
    
    box = canvas_el.bounding_box()
    if not box:
        print("无法获取画布位置")
        return False
    
    cropped = crop_canvas_region(img, box)
    
    # 检查裁剪后的图像是否有效
    if cropped and cropped.width >= 10 and cropped.height >= 10:
        images.append(cropped)
        positions.append({
            'y_position': y_position,
            'viewport_top': HEADER_OFFSET,
            'viewport_bottom': VIEWPORT_HEIGHT,
            'width': cropped.width
        })
        return True
    elif cropped:
        # 图像太小，可能是空白区域
        print(f"  警告：裁剪图像过小 ({cropped.width}x{cropped.height})，可能是空白区域")
        return False
    else:
        print(f"  裁剪失败")
        return False

def get_page_content(page):
    """获取页面内容"""
    body = page.query_selector("body")
    if not body:
        return ""
    
    content = body.inner_text()
    lines = content.split('\n')
    cleaned_lines = []
    
    ui_tags = [
        "总览", "演示", "标注", "登录", "免费使用", "页面", "图层", "批注", "评论", 
        "需求列表", "欢迎来到墨刀", "立即登录",
        "头像", "用户昵称", "昵称", "私聊", "聊天", "消息", "发送", "取消",
        "设置", "保存", "删除", "修改", "添加", "关闭", "确定", "返回",
        "时间", "日期", "状态", "类型", "名称", "简介", "介绍", "描述",
        "操作", "编辑", "查看", "详情", "更多", "全部", "选择", "筛选",
        "关注", "粉丝", "礼物", "排行榜", "榜单", "等级", "经验",
        "认证", "资质", "公会", "家族", "会长", "成员",
        "开", "关", "是", "否", "已读", "未读", "在线", "离线",
        "扩列", "派对", "一键", "开始", "结束", "参与", "报名",
        "广场", "互动", "专区", "热门", "推荐", "最新", "关注",
        "加入", "离开", "申请", "同意", "拒绝", "通过",
        "用户", "信息", "资料", "内容", "分享", "复制",
        "刷新", "上传", "下载", "播放", "暂停", "停止",
        "上一页", "下一页", "上一步", "下一步", "完成",
        "个人", "我的", "首页", "发现", "消息", "通知",
        "公会", "家族", "房间", "厅", "派对", "麦位",
        "马明伦"
    ]
    
    for line in lines:
        line = line.strip()
        if not line:
            cleaned_lines.append("")
            continue
        if line.isdigit():
            continue
        if len(line) <= 2 and line in ui_tags:
            continue
        if len(line) <= 4:
            is_ui_only = True
            for tag in ui_tags[:20]:
                if tag in line:
                    is_ui_only = False
                    break
            if is_ui_only and not re.search(r'\d', line):
                continue
        if any(line == p or line.startswith(p) for p in ui_tags):
            if len(line) < 4:
                continue
        cleaned_lines.append(line)
    
    # 移除连续空行
    content_lines = []
    prev_empty = False
    for line in cleaned_lines:
        if not line:
            if not prev_empty:
                content_lines.append(line)
            prev_empty = True
        else:
            content_lines.append(line)
            prev_empty = False
    
    return '\n'.join(content_lines)

# ==================== 主流程函数 ====================

def save_modao_page_v2(url, base_output_dir=r".\modao-export", canvas_index=None, debug=False):
    """采集墨刀页面并生成 Markdown 文档
    
    Args:
        url: 墨刀页面 URL
        base_output_dir: 基础输出目录
        canvas_index: 指定采集第几个画布（从 1 开始）
        debug: 是否为 debug 模式，True 时覆盖已有文件，不创建新目录
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            viewport={"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT},
            ignore_https_errors=True
        )
        page = context.new_page()
        
        # 打开页面
        print(f"正在打开页面：{url}")
        page.goto(url)
        time.sleep(3)
        print("页面主资源加载完成")
        
        page.set_viewport_size({"width": VIEWPORT_WIDTH, "height": VIEWPORT_HEIGHT})
        
        # 等待页面渲染
        try:
            page.wait_for_load_state('networkidle', timeout=10000)
            print("页面网络请求完成")
        except:
            pass
        
        try:
            page.wait_for_selector('.canvas-sortable-list', timeout=5000)
            print("画布列表已加载")
        except:
            pass
        
        # 关闭登录窗口
        try:
            signup_btn = page.query_selector(".signup-btn")
            if signup_btn:
                close_btn = page.query_selector("#fixed-area > div > svg")
                if close_btn:
                    close_btn.click()
                    print("已关闭登录窗口")
                    time.sleep(0.5)
        except:
            pass
        
        # 获取项目标题
        project_title = "墨刀项目"
        try:
            title = page.title()
            if title:
                project_title = sanitize_filename(title.split('-')[0].strip()[:50])
        except:
            pass
        
        if not project_title:
            project_title = "未命名项目"
        
        print(f"项目标题：{project_title}")
        
        # 创建输出目录结构
        base_output_dir = os.path.join(base_output_dir, project_title)
        images_dir = os.path.join(base_output_dir, "images")
        md_dir = os.path.join(base_output_dir, "md")
        
        # Debug 模式下不检查目录是否存在，直接覆盖
        if not debug:
            # 检查并处理目录已存在的情况
            if os.path.exists(base_output_dir):
                counter = 1
                while True:
                    new_base = f"{base_output_dir}_{counter}"
                    if not os.path.exists(new_base):
                        base_output_dir = new_base
                        images_dir = os.path.join(base_output_dir, "images")
                        md_dir = os.path.join(base_output_dir, "md")
                        break
                    counter += 1
        
        os.makedirs(images_dir, exist_ok=True)
        os.makedirs(md_dir, exist_ok=True)
        
        if debug:
            print(f"Debug 模式：覆盖已有文件")
        print(f"输出目录：{base_output_dir}")
        print(f"图片目录：{images_dir}")
        print(f"MD 目录：{md_dir}")
        
        # 采集画布和页面
        all_pages_content = []
        page_counter = 1
        
        canvas_list = find_canvas_list(page)
        print(f"画布数量：{len(canvas_list)}")
        print(f"画布列表：{canvas_list}")
        
        if canvas_index is not None:
            if 1 <= canvas_index <= len(canvas_list):
                canvas_list = [canvas_list[canvas_index - 1]]
                print(f"只采集第 {canvas_index} 个画布：{canvas_list[0]}")
            else:
                print(f"画布索引 {canvas_index} 超出范围，将采集所有画布")
        
        if not canvas_list:
            canvas_list = ["默认画布"]
        
        # 遍历画布
        for canvas_idx, canvas_name in enumerate(canvas_list):
            print(f"\n{'='*50}")
            print(f"处理画布 {canvas_idx + 1}: {canvas_name}")
            print(f"{'='*50}")
            
            click_canvas(page, canvas_name)
            time.sleep(0.5)
            
            page_list = get_page_list_in_canvas(page)
            print(f"画布 '{canvas_name}' 下有 {len(page_list)} 个页面")
            for p in page_list:
                print(f"  - 页面：{p['name']}, data-cid: {p['data_cid']}")
            
            if not page_list:
                page_list = [{"name": "页面 1", "data_cid": None}]
            
            # 遍历页面
            for page_idx, page_info in enumerate(page_list):
                page_name = page_info["name"]
                data_cid = page_info["data_cid"]
                
                print(f"\n--- 画布：{canvas_name} - 页面 {page_idx + 1}: {page_name} ---")
                
                click_page_in_canvas(page, page_idx)
                time.sleep(0.3)
                
                # 截图保存到 images 目录
                screenshot_name = f"page_{page_counter}_{sanitize_filename(canvas_name)}_{sanitize_filename(page_name)}.png"
                screenshot_path = os.path.join(images_dir, screenshot_name)
                
                if not take_screenshot(page, screenshot_path, data_cid):
                    print("截图失败，跳过")
                    continue
                
                content = get_page_content(page)
                
                all_pages_content.append({
                    "index": page_counter,
                    "canvas": canvas_name,
                    "page": page_name,
                    "content": content
                })
                
                page_counter += 1
        
        print(f"\n共处理 {len(all_pages_content)} 个页面")
        
        # 生成 Markdown 文件到 md 目录
        for item in all_pages_content:
            title = f"{item['canvas']} - {item['page']}"
            md_file = f"page_{item['index']}_{sanitize_filename(item['canvas'])}_{sanitize_filename(item['page'])}.md"
            md_path = os.path.join(md_dir, md_file)
            
            # 图片相对路径（从 md 目录到 images 目录）
            screenshot_name = f"../images/page_{item['index']}_{sanitize_filename(item['canvas'])}_{sanitize_filename(item['page'])}.png"
            
            structured = structure_content(item['content'])
            md_content = generate_structured_md(title, structured, screenshot_name)
            
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_content)
            print(f"已创建：{md_file}")
        
        # 生成索引文件到 md 目录
        index_content = f"# {project_title}\n\n"
        index_content += f"**项目 URL**: {url}\n\n"
        index_content += "---\n\n"
        index_content += "## 目录\n\n"
        
        for item in all_pages_content:
            title = f"{item['canvas']} - {item['page']}"
            md_file = f"page_{item['index']}_{sanitize_filename(item['canvas'])}_{sanitize_filename(item['page'])}.md"
            index_content += f"- [{title}]({md_file})\n"
        
        index_path = os.path.join(md_dir, "index.md")
        with open(index_path, "w", encoding="utf-8") as f:
            f.write(index_content)
        print(f"已创建：index.md")
        
        browser.close()
        print(f"\n完成！保存在：{base_output_dir}")
        return base_output_dir

# ==================== 程序入口 ====================

if __name__ == "__main__":
    import sys
    
    url = None
    canvas_index = None
    debug_mode = False
    
    if len(sys.argv) > 1:
        # 第一个参数是 URL
        url = sys.argv[1].strip()
    
    if len(sys.argv) > 2:
        # 检查最后一个参数是否为 debug
        last_arg = sys.argv[-1].strip().lower()
        if last_arg == "debug":
            debug_mode = True
            
            # 如果有 3 个或更多参数，尝试解析倒数第二个参数为 canvas_index
            if len(sys.argv) >= 3:
                try:
                    potential_index = sys.argv[-2].strip().lower()
                    # 确保倒数第二个参数不是 "debug"
                    if potential_index != "debug":
                        canvas_index = int(potential_index)
                except:
                    pass
    
    if not url:
        url = input("请输入墨刀页面 URL: ").strip()
    
    if not url:
        print("错误：URL 不能为空")
        exit(1)
    
    if debug_mode:
        print("=== Debug 模式：将覆盖已有文件 ===")
    
    if canvas_index is not None:
        print(f"将只采集第 {canvas_index} 个画布")
    else:
        print("将采集所有画布")
    
    output_dir = save_modao_page_v2(url, canvas_index=canvas_index, debug=debug_mode)
