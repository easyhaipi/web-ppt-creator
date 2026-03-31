#!/usr/bin/env python3
"""
网页PPT生成器 - Web PPT Generator
根据配置生成可交互的网页版PPT (HTML格式)

使用示例:
    python generate_ppt.py --config config.json --output my-ppt.html
    python generate_ppt.py --config config.json --output my-ppt.html --reference-image ./style-ref.png
    python generate_ppt.py --title "我的PPT" --theme 科技蓝 --animation 滑动进入 --output presentation.html
"""

import json
import argparse
import os
import base64
from pathlib import Path
from datetime import datetime


# ==================== 主题颜色配置 ====================
THEMES = {
    "科技蓝": {
        "primary": "#0066FF",
        "secondary": "#00AAFF",
        "accent": "#00FFFF",
        "gradient": "linear-gradient(135deg, #0066FF 0%, #00AAFF 100%)",
        "background": "linear-gradient(135deg, #0a0a20 0%, #1a1a3a 100%)",
        "text": "#ffffff",
        "text_secondary": "#aaccff",
        "card_bg": "rgba(0, 102, 255, 0.1)",
        "border": "rgba(0, 170, 255, 0.3)"
    },
    "商务黑": {
        "primary": "#2c2c2c",
        "secondary": "#4a4a4a",
        "accent": "#888888",
        "gradient": "linear-gradient(135deg, #2c2c2c 0%, #4a4a4a 100%)",
        "background": "linear-gradient(135deg, #0d0d0d 0%, #1a1a1a 100%)",
        "text": "#ffffff",
        "text_secondary": "#cccccc",
        "card_bg": "rgba(255, 255, 255, 0.05)",
        "border": "rgba(255, 255, 255, 0.1)"
    },
    "清新绿": {
        "primary": "#00C853",
        "secondary": "#69F0AE",
        "accent": "#B2FF59",
        "gradient": "linear-gradient(135deg, #00C853 0%, #69F0AE 100%)",
        "background": "linear-gradient(135deg, #0a1f0a 0%, #1a2f1a 100%)",
        "text": "#ffffff",
        "text_secondary": "#aaffcc",
        "card_bg": "rgba(0, 200, 83, 0.1)",
        "border": "rgba(105, 240, 174, 0.3)"
    },
    "渐变紫": {
        "primary": "#8B5CF6",
        "secondary": "#D946EF",
        "accent": "#F472B6",
        "gradient": "linear-gradient(135deg, #8B5CF6 0%, #D946EF 100%)",
        "background": "linear-gradient(135deg, #1a0a2e 0%, #2d1b4e 50%, #1a0a2e 100%)",
        "text": "#ffffff",
        "text_secondary": "#e0ccff",
        "card_bg": "rgba(139, 92, 246, 0.1)",
        "border": "rgba(217, 70, 239, 0.3)"
    },
    "渐变橙": {
        "primary": "#FF6B35",
        "secondary": "#FFB347",
        "accent": "#FFD93D",
        "gradient": "linear-gradient(135deg, #FF6B35 0%, #FFB347 100%)",
        "background": "linear-gradient(135deg, #1a0f0a 0%, #2f1a0a 100%)",
        "text": "#ffffff",
        "text_secondary": "#ffe0cc",
        "card_bg": "rgba(255, 107, 53, 0.1)",
        "border": "rgba(255, 179, 71, 0.3)"
    },
    "毛玻璃": {
        "primary": "rgba(255, 255, 255, 0.25)",
        "secondary": "rgba(255, 255, 255, 0.15)",
        "accent": "rgba(255, 255, 255, 0.9)",
        "gradient": "linear-gradient(135deg, rgba(255,255,255,0.25) 0%, rgba(255,255,255,0.15) 100%)",
        "background": "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
        "text": "#ffffff",
        "text_secondary": "rgba(255,255,255,0.8)",
        "card_bg": "rgba(255, 255, 255, 0.1)",
        "border": "rgba(255, 255, 255, 0.2)",
        "glass": True
    },
    "中国红": {
        "primary": "#DC2626",
        "secondary": "#EF4444",
        "accent": "#F87171",
        "gradient": "linear-gradient(135deg, #DC2626 0%, #EF4444 100%)",
        "background": "linear-gradient(135deg, #1a0505 0%, #2d0a0a 100%)",
        "text": "#ffffff",
        "text_secondary": "#fecaca",
        "card_bg": "rgba(220, 38, 38, 0.1)",
        "border": "rgba(239, 68, 68, 0.3)"
    },
    "深海蓝": {
        "primary": "#0EA5E9",
        "secondary": "#06B6D4",
        "accent": "#22D3EE",
        "gradient": "linear-gradient(135deg, #0EA5E9 0%, #06B6D4 50%, #22D3EE 100%)",
        "background": "linear-gradient(180deg, #0c4a6e 0%, #075985 50%, #0369a1 100%)",
        "text": "#ffffff",
        "text_secondary": "#bae6fd",
        "card_bg": "rgba(14, 165, 233, 0.15)",
        "border": "rgba(6, 182, 212, 0.3)"
    },
    "极简白": {
        "primary": "#1a1a1a",
        "secondary": "#4a4a4a",
        "accent": "#0066FF",
        "gradient": "linear-gradient(135deg, #f5f5f5 0%, #ffffff 100%)",
        "background": "linear-gradient(135deg, #f8f9fa 0%, #ffffff 100%)",
        "text": "#1a1a1a",
        "text_secondary": "#4a4a4a",
        "card_bg": "rgba(0, 0, 0, 0.03)",
        "border": "rgba(0, 0, 0, 0.1)"
    },
    "赛博朋克": {
        "primary": "#FF00FF",
        "secondary": "#00FFFF",
        "accent": "#FFFF00",
        "gradient": "linear-gradient(135deg, #FF00FF 0%, #00FFFF 100%)",
        "background": "linear-gradient(135deg, #0a0014 0%, #1a0033 50%, #0a0014 100%)",
        "text": "#ffffff",
        "text_secondary": "#ff99ff",
        "card_bg": "rgba(255, 0, 255, 0.1)",
        "border": "rgba(0, 255, 255, 0.3)"
    }
}

# ==================== 动画配置 ====================
ANIMATIONS = {
    "无动画": "none",
    "淡入淡出": "fade",
    "滑动进入": "slide",
    "缩放效果": "zoom",
    "组合动画": "combo",
    "3D翻转": "flip",
    "立方体": "cube",
    "覆盖": "cover",
    "揭开": "uncover"
}

# ==================== 模板配置 ====================
TEMPLATES = {
    "空白模板": "blank",
    "标题页": "title",
    "两栏布局": "two-column",
    "三栏布局": "three-column",
    "图片展示": "gallery",
    "代码演示": "code",
    "时间轴": "timeline",
    "卡片网格": "card-grid",
    "对比布局": "comparison",
    "数据展示": "data",
    "引用页": "quote",
    "结束页": "end",
    "自定义描述": "custom"
}


def get_animation_css(animation_type):
    """获取动画CSS"""
    animations = {
        "fade": """
    .slide { transition: opacity 0.6s ease; }
    .slide.from-left { transform: translateX(-30px); opacity: 0; }
    .slide.from-right { transform: translateX(30px); opacity: 0; }
    .slide.active { transform: translateX(0); opacity: 1; }
        """,
        "slide": """
    .slide { transition: transform 0.5s ease, opacity 0.5s ease; }
    .slide.from-left { transform: translateX(-100%); opacity: 0; }
    .slide.from-right { transform: translateX(100%); opacity: 0; }
    .slide.active { transform: translateX(0); opacity: 1; }
        """,
        "zoom": """
    .slide { transition: transform 0.5s ease, opacity 0.5s ease; }
    .slide.from-left { transform: scale(0.8); opacity: 0; }
    .slide.from-right { transform: scale(1.2); opacity: 0; }
    .slide.active { transform: scale(1); opacity: 1; }
        """,
        "combo": """
    .slide { transition: all 0.6s cubic-bezier(0.4, 0, 0.2, 1); }
    .slide.from-left { transform: translateX(-50px) scale(0.95); opacity: 0; }
    .slide.from-right { transform: translateX(50px) scale(0.95); opacity: 0; }
    .slide.active { transform: translateX(0) scale(1); opacity: 1; }
        """,
        "flip": """
    .slide { transition: transform 0.7s ease, opacity 0.7s ease; transform-style: preserve-3d; }
    .slide.from-left { transform: rotateY(-90deg); opacity: 0; }
    .slide.from-right { transform: rotateY(90deg); opacity: 0; }
    .slide.active { transform: rotateY(0); opacity: 1; }
    .slides-container { perspective: 1000px; }
        """,
        "cube": """
    .slide { transition: transform 0.8s ease, opacity 0.8s ease; transform-style: preserve-3d; }
    .slide.from-left { transform: translateX(-50%) rotateY(-90deg); opacity: 0; }
    .slide.from-right { transform: translateX(50%) rotateY(90deg); opacity: 0; }
    .slide.active { transform: translateX(0) rotateY(0); opacity: 1; }
    .slides-container { perspective: 1200px; }
        """,
        "cover": """
    .slide { transition: transform 0.6s ease, opacity 0.6s ease; z-index: 1; }
    .slide.from-left { transform: translateY(-100%); opacity: 0; z-index: 2; }
    .slide.from-right { transform: translateY(100%); opacity: 0; z-index: 0; }
    .slide.active { transform: translateY(0); opacity: 1; z-index: 1; }
        """,
        "uncover": """
    .slide { transition: transform 0.6s ease, opacity 0.6s ease; }
    .slide.from-left { transform: translateX(0); opacity: 0; }
    .slide.from-right { transform: translateX(0); opacity: 1; }
    .slide.active { transform: translateX(0); opacity: 1; }
    .slide.prev { transform: translateX(-30%) scale(0.9); opacity: 0.5; }
        """,
        "none": """
    .slide { transition: none; }
    .slide.from-left { opacity: 0; }
    .slide.from-right { opacity: 0; }
    .slide.active { opacity: 1; }
        """
    }
    return animations.get(animation_type, animations["combo"])


def get_glass_css(theme):
    """获取毛玻璃效果CSS"""
    if theme.get("glass"):
        return """
    .slide { backdrop-filter: blur(10px); background: rgba(255,255,255,0.1); }
    .card { background: rgba(255,255,255,0.15); backdrop-filter: blur(5px); }
        """
    return ""


def generate_css(theme_name, animation_type, reference_image=None):
    """生成CSS样式"""
    theme = THEMES.get(theme_name, THEMES["渐变紫"])

    # 如果有参考图片，添加自定义样式支持
    ref_style = ""
    if reference_image:
        ref_style = f"""
    /* 参考图片样式 - 可基于此调整 */
    .reference-overlay {{
        position: fixed;
        top: 10px;
        left: 10px;
        width: 100px;
        height: 60px;
        opacity: 0.3;
        border-radius: 4px;
        z-index: 1000;
        pointer-events: none;
    }}
        """

    css = f"""
    * {{
        margin: 0;
        padding: 0;
        box-sizing: border-box;
    }}

    body {{
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
        background: {theme['background']};
        color: {theme['text']};
        overflow: hidden;
        height: 100vh;
    }}

    .slides-container {{
        position: relative;
        width: 100%;
        height: 100vh;
    }}

    .slide {{
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        padding: 60px 80px;
        opacity: 0;
        visibility: hidden;
        transition: opacity 0.6s ease, visibility 0.6s ease;
    }}

    .slide.active {{
        opacity: 1;
        visibility: visible;
    }}

    /* 动画效果 */
    {get_animation_css(animation_type)}

    /* 毛玻璃效果 */
    {get_glass_css(theme)}

    /* 标题样式 */
    .slide h1 {{
        font-size: 3.5rem;
        font-weight: 700;
        margin-bottom: 1.5rem;
        background: {theme['gradient']};
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }}

    .slide h2 {{
        font-size: 2.5rem;
        font-weight: 600;
        margin-bottom: 1rem;
        color: {theme['text']};
    }}

    .slide h3 {{
        font-size: 1.8rem;
        font-weight: 500;
        margin-bottom: 0.8rem;
        color: {theme['text_secondary']};
    }}

    .slide p {{
        font-size: 1.3rem;
        line-height: 1.8;
        color: {theme['text_secondary']};
        max-width: 800px;
        text-align: center;
    }}

    .slide ul {{
        font-size: 1.2rem;
        line-height: 2;
        color: {theme['text_secondary']};
        list-style: none;
        text-align: left;
    }}

    .slide li {{
        padding: 0.5rem 0;
        padding-left: 2rem;
        position: relative;
    }}

    .slide li::before {{
        content: "▸";
        position: absolute;
        left: 0;
        color: {theme['primary']};
    }}

    /* 模板样式 */
    .template-title {{
        text-align: center;
    }}

    .template-two-column {{
        flex-direction: row;
        gap: 60px;
        justify-content: center;
        align-items: flex-start;
    }}

    .template-two-column .column {{
        flex: 1;
        max-width: 500px;
    }}

    .template-three-column {{
        flex-direction: row;
        gap: 40px;
        justify-content: center;
        align-items: flex-start;
    }}

    .template-three-column .column {{
        flex: 1;
        max-width: 350px;
    }}

    .template-gallery {{
        flex-direction: row;
        flex-wrap: wrap;
        gap: 30px;
        align-content: center;
        justify-content: center;
    }}

    .template-code {{
        font-family: 'Monaco', 'Menlo', monospace;
        text-align: left;
        padding: 40px 60px;
    }}

    .template-code pre {{
        background: rgba(0,0,0,0.3);
        padding: 30px;
        border-radius: 10px;
        overflow-x: auto;
        font-size: 1rem;
        line-height: 1.6;
    }}

    .template-timeline {{
        align-items: flex-start;
        padding-left: 100px;
    }}

    .timeline-item {{
        position: relative;
        padding-left: 40px;
        margin-bottom: 30px;
        border-left: 3px solid {theme['primary']};
    }}

    .timeline-item::before {{
        content: "";
        position: absolute;
        left: -9px;
        top: 5px;
        width: 15px;
        height: 15px;
        border-radius: 50%;
        background: {theme['primary']};
    }}

    .template-card-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        gap: 30px;
        width: 100%;
        max-width: 1200px;
    }}

    .template-comparison {{
        flex-direction: row;
        gap: 0;
    }}

    .comparison-side {{
        flex: 1;
        padding: 40px;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}

    .comparison-side.left {{
        background: rgba(0,0,0,0.2);
        border-right: 2px solid {theme['border']};
    }}

    .comparison-side.right {{
        background: rgba(0,0,0,0.1);
    }}

    .template-data {{
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 40px;
        width: 100%;
        max-width: 1000px;
    }}

    .data-item {{
        text-align: center;
        padding: 30px;
    }}

    .data-number {{
        font-size: 4rem;
        font-weight: 700;
        background: {theme['gradient']};
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }}

    .template-quote {{
        text-align: center;
        max-width: 900px;
    }}

    .quote-text {{
        font-size: 2rem;
        font-style: italic;
        line-height: 1.6;
        margin-bottom: 30px;
        color: {theme['text']};
    }}

    .quote-author {{
        font-size: 1.2rem;
        color: {theme['text_secondary']};
    }}

    .template-end {{
        text-align: center;
    }}

    .end-title {{
        font-size: 4rem;
        font-weight: 700;
        margin-bottom: 20px;
        background: {theme['gradient']};
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }}

    /* 卡片样式 */
    .card {{
        background: {theme['card_bg']};
        border: 1px solid {theme['border']};
        border-radius: 16px;
        padding: 30px;
        margin: 15px;
        min-width: 280px;
        text-align: center;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }}

    .card:hover {{
        transform: translateY(-5px);
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
    }}

    .card h4 {{
        font-size: 1.4rem;
        margin-bottom: 10px;
        color: {theme['text']};
    }}

    .card p {{
        font-size: 1rem;
        color: {theme['text_secondary']};
    }}

    /* 进度指示器 */
    .progress-bar {{
        position: fixed;
        bottom: 0;
        left: 0;
        height: 4px;
        background: {theme['gradient']};
        transition: width 0.3s ease;
        z-index: 100;
    }}

    .page-indicator {{
        position: fixed;
        bottom: 20px;
        right: 30px;
        font-size: 1rem;
        color: {theme['text_secondary']};
        background: rgba(0,0,0,0.3);
        padding: 8px 16px;
        border-radius: 20px;
        z-index: 100;
    }}

    /* 导航按钮 */
    .nav-button {{
        position: fixed;
        top: 50%;
        transform: translateY(-50%);
        width: 50px;
        height: 50px;
        background: {theme['primary']};
        border: none;
        border-radius: 50%;
        color: {theme['text']};
        font-size: 1.5rem;
        cursor: pointer;
        opacity: 0.5;
        transition: opacity 0.3s, transform 0.3s;
        z-index: 100;
    }}

    .nav-button:hover {{
        opacity: 1;
        transform: translateY(-50%) scale(1.1);
    }}

    .nav-prev {{ left: 20px; }}
    .nav-next {{ right: 20px; }}

    /* 全屏按钮 */
    .fullscreen-btn {{
        position: fixed;
        top: 20px;
        right: 30px;
        padding: 10px 20px;
        background: {theme['primary']};
        border: none;
        border-radius: 8px;
        color: {theme['text']};
        cursor: pointer;
        font-size: 0.9rem;
        z-index: 100;
        transition: opacity 0.3s;
    }}

    .fullscreen-btn:hover {{
        opacity: 0.8;
    }}

    /* 缩略图导航 */
    .thumbnail-nav {{
        position: fixed;
        bottom: 60px;
        left: 50%;
        transform: translateX(-50%);
        display: flex;
        gap: 10px;
        z-index: 100;
        background: rgba(0,0,0,0.5);
        padding: 10px 20px;
        border-radius: 30px;
    }}

    .thumbnail {{
        width: 40px;
        height: 30px;
        background: rgba(255,255,255,0.2);
        border-radius: 4px;
        cursor: pointer;
        transition: all 0.3s;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.7rem;
        color: {theme['text_secondary']};
    }}

    .thumbnail:hover, .thumbnail.active {{
        background: {theme['primary']};
        color: {theme['text']};
    }}

    /* 激光笔效果 */
    .laser-pointer {{
        position: fixed;
        width: 20px;
        height: 20px;
        border-radius: 50%;
        background: radial-gradient(circle, rgba(255,0,0,0.8) 0%, rgba(255,0,0,0) 70%);
        pointer-events: none;
        z-index: 9999;
        display: none;
    }}

    .laser-pointer.active {{
        display: block;
    }}

    /* 触摸滑动提示 */
    .swipe-hint {{
        position: fixed;
        bottom: 100px;
        left: 50%;
        transform: translateX(-50%);
        color: {theme['text_secondary']};
        font-size: 0.9rem;
        opacity: 0.5;
        animation: pulse 2s infinite;
    }}

    @keyframes pulse {{
        0%, 100% {{ opacity: 0.5; }}
        50% {{ opacity: 1; }}
    }}

    {ref_style}
    """
    return css


def generate_js(interaction_config):
    """生成JavaScript交互代码"""
    keyboard_nav = "keyboard" in interaction_config or "keyboard-nav" in interaction_config
    mouse_click = "click" in interaction_config or "mouse-click" in interaction_config
    progress = "progress" in interaction_config or "progress-bar" in interaction_config
    fullscreen = "fullscreen" in interaction_config
    thumbnail = "thumbnail" in interaction_config or "thumbnail-nav" in interaction_config
    laser = "laser" in interaction_config or "laser-pointer" in interaction_config
    touch = "touch" in interaction_config or "swipe" in interaction_config

    # 基础JS
    js_parts = ["""
    const slides = document.querySelectorAll('.slide');
    let currentSlide = 0;
    const totalSlides = slides.length;

    function showSlide(index) {
        slides.forEach((slide, i) => {
            slide.classList.remove('active', 'from-left', 'from-right', 'prev');
            if (i < index) {
                slide.classList.add('from-left');
            } else if (i > index) {
                slide.classList.add('from-right');
            }
        });
        slides[index].classList.add('active');

        // 更新进度条
        const progressBar = document.querySelector('.progress-bar');
        if (progressBar) {
            const progress = (index + 1) / totalSlides * 100;
            progressBar.style.width = progress + '%';
        }

        // 更新页码
        const pageIndicator = document.querySelector('.page-indicator');
        if (pageIndicator) {
            pageIndicator.textContent = (index + 1) + ' / ' + totalSlides;
        }

        // 更新缩略图
        const thumbnails = document.querySelectorAll('.thumbnail');
        thumbnails.forEach((thumb, i) => {
            thumb.classList.toggle('active', i === index);
        });
    }

    function nextSlide() {
        if (currentSlide < totalSlides - 1) {
            currentSlide++;
            showSlide(currentSlide);
        }
    }

    function prevSlide() {
        if (currentSlide > 0) {
            currentSlide--;
            showSlide(currentSlide);
        }
    }

    function goToSlide(index) {
        if (index >= 0 && index < totalSlides) {
            currentSlide = index;
            showSlide(currentSlide);
        }
    }
"""]

    # 键盘导航
    if keyboard_nav:
        js_parts.append("""
    // 键盘导航
    document.addEventListener('keydown', (e) => {
        if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') {
            e.preventDefault();
            nextSlide();
        } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
            e.preventDefault();
            prevSlide();
        } else if (e.key === 'Home') {
            e.preventDefault();
            goToSlide(0);
        } else if (e.key === 'End') {
            e.preventDefault();
            goToSlide(totalSlides - 1);
        } else if (e.key === 'f' || e.key === 'F') {
            e.preventDefault();
            toggleFullscreen();
        } else if (e.key === 'Escape') {
            if (document.fullscreenElement) {
                document.exitFullscreen();
            }
        }
    });
""")

    # 鼠标点击
    if mouse_click:
        js_parts.append("""
    // 鼠标点击导航
    document.addEventListener('click', (e) => {
        if (e.target.classList.contains('nav-button') ||
            e.target.classList.contains('thumbnail') ||
            e.target.classList.contains('fullscreen-btn')) return;

        const x = e.clientX;
        const width = window.innerWidth;
        const threshold = width * 0.25;

        if (x > width - threshold) {
            nextSlide();
        } else if (x < threshold) {
            prevSlide();
        }
    });

    // 导航按钮事件
    const prevBtn = document.querySelector('.nav-prev');
    const nextBtn = document.querySelector('.nav-next');
    if (prevBtn) prevBtn.addEventListener('click', prevSlide);
    if (nextBtn) nextBtn.addEventListener('click', nextSlide);
""")

    # 缩略图导航
    if thumbnail:
        js_parts.append("""
    // 缩略图导航
    const thumbnails = document.querySelectorAll('.thumbnail');
    thumbnails.forEach((thumb, index) => {
        thumb.addEventListener('click', () => goToSlide(index));
    });
""")

    # 全屏模式
    if fullscreen:
        js_parts.append("""
    // 全屏模式
    function toggleFullscreen() {
        if (document.fullscreenElement) {
            document.exitFullscreen();
        } else {
            document.documentElement.requestFullscreen().catch(err => {
                console.log('全屏模式不支持:', err);
            });
        }
    }
    const fullscreenBtn = document.querySelector('.fullscreen-btn');
    if (fullscreenBtn) fullscreenBtn.addEventListener('click', toggleFullscreen);
""")

    # 激光笔效果
    if laser:
        js_parts.append("""
    // 激光笔效果
    const laser = document.querySelector('.laser-pointer');
    let laserActive = false;

    document.addEventListener('keydown', (e) => {
        if (e.key === 'l' || e.key === 'L') {
            laserActive = !laserActive;
            if (laser) laser.classList.toggle('active', laserActive);
        }
    });

    document.addEventListener('mousemove', (e) => {
        if (laserActive && laser) {
            laser.style.left = (e.clientX - 10) + 'px';
            laser.style.top = (e.clientY - 10) + 'px';
        }
    });
""")

    # 触摸滑动
    if touch:
        js_parts.append("""
    // 触摸滑动支持
    let touchStartX = 0;
    let touchEndX = 0;

    document.addEventListener('touchstart', (e) => {
        touchStartX = e.changedTouches[0].screenX;
    }, { passive: true });

    document.addEventListener('touchend', (e) => {
        touchEndX = e.changedTouches[0].screenX;
        handleSwipe();
    }, { passive: true });

    function handleSwipe() {
        const swipeThreshold = 50;
        const diff = touchStartX - touchEndX;

        if (Math.abs(diff) > swipeThreshold) {
            if (diff > 0) {
                nextSlide();
            } else {
                prevSlide();
            }
        }
    }
""")

    # 初始化
    js_parts.append("""
    // 初始化
    showSlide(0);
    console.log('🎨 网页PPT已加载');
    console.log('📖 快捷键: ← → 翻页 | F 全屏 | Home/End 首尾页 | L 激光笔');
""")

    return "\n".join(js_parts)


def generate_slide_content(slide_data, theme):
    """生成单个幻灯片内容"""
    slide_type = slide_data.get("type", "content")
    title = slide_data.get("title", "")
    subtitle = slide_data.get("subtitle", "")
    content = slide_data.get("content", [])
    template = slide_data.get("template", "")

    # 根据类型生成不同内容
    if slide_type == "title":
        return f"""
    <div class="slide template-title active">
        <h1>{title}</h1>
        <p>{subtitle}</p>
    </div>
"""

    elif slide_type == "content":
        content_html = ""
        if isinstance(content, list):
            content_html = "<ul>" + "".join([f"<li>{item}</li>" for item in content]) + "</ul>"
        else:
            content_html = f"<p>{content}</p>"

        return f"""
    <div class="slide">
        <h2>{title}</h2>
        {content_html}
    </div>
"""

    elif slide_type == "comparison":
        left = slide_data.get("left", {})
        right = slide_data.get("right", {})
        left_items = "<ul>" + "".join([f"<li>{item}</li>" for item in left.get("items", [])]) + "</ul>"
        right_items = "<ul>" + "".join([f"<li>{item}</li>" for item in right.get("items", [])]) + "</ul>"

        return f"""
    <div class="slide template-comparison">
        <div class="comparison-side left">
            <h2>{left.get("title", "")}</h2>
            {left_items}
        </div>
        <div class="comparison-side right">
            <h2>{right.get("title", "")}</h2>
            {right_items}
        </div>
    </div>
"""

    elif slide_type == "timeline":
        items_html = ""
        for item in content:
            items_html += f"""
        <div class="timeline-item">
            <h3>{item.get("time", "")}</h3>
            <p>{item.get("event", "")}</p>
        </div>"""

        return f"""
    <div class="slide template-timeline">
        <h2>{title}</h2>
        {items_html}
    </div>
"""

    elif slide_type == "card-grid":
        cards_html = ""
        for card in content:
            cards_html += f"""
        <div class="card">
            <h4>{card.get("title", "")}</h4>
            <p>{card.get("desc", "")}</p>
        </div>"""

        return f"""
    <div class="slide">
        <h2>{title}</h2>
        <div class="template-card-grid">
            {cards_html}
        </div>
    </div>
"""

    elif slide_type == "data":
        items_html = ""
        for item in content:
            items_html += f"""
        <div class="data-item">
            <div class="data-number">{item.get("number", "")}</div>
            <p>{item.get("label", "")}</p>
        </div>"""

        return f"""
    <div class="slide">
        <h2>{title}</h2>
        <div class="template-data">
            {items_html}
        </div>
    </div>
"""

    elif slide_type == "quote":
        return f"""
    <div class="slide template-quote">
        <div class="quote-text">"{slide_data.get("quote", "")}"</div>
        <div class="quote-author">—— {slide_data.get("author", "")}</div>
    </div>
"""

    elif slide_type == "end":
        return f"""
    <div class="slide template-end">
        <div class="end-title">{title}</div>
        <p>{subtitle}</p>
    </div>
"""

    elif slide_type == "demo":
        steps_html = "<ul>" + "".join([f"<li>{step}</li>" for step in slide_data.get("steps", [])]) + "</ul>"
        return f"""
    <div class="slide">
        <h2>{title}</h2>
        {steps_html}
    </div>
"""

    elif slide_type == "code-showcase":
        dialogue_html = ""
        for line in slide_data.get("dialogue", []):
            dialogue_html += f"<p>{line}</p>"
        return f"""
    <div class="slide template-code">
        <h2>{title}</h2>
        <div class="code-content">
            {dialogue_html}
        </div>
    </div>
"""

    elif slide_type == "action-plan":
        week_html = "<ul>" + "".join([f"<li>{day}</li>" for day in slide_data.get("week", [])]) + "</ul>"
        return f"""
    <div class="slide">
        <h2>{title}</h2>
        {week_html}
    </div>
"""

    else:
        # 默认内容页
        content_html = ""
        if isinstance(content, list):
            content_html = "<ul>" + "".join([f"<li>{item}</li>" for item in content]) + "</ul>"
        else:
            content_html = f"<p>{content}</p>"

        return f"""
    <div class="slide">
        <h2>{title}</h2>
        {content_html}
    </div>
"""


def generate_thumbnails_html(slide_count):
    """生成缩略图导航HTML"""
    thumbnails = ""
    for i in range(slide_count):
        thumbnails += f'<div class="thumbnail" data-index="{i}">{i+1}</div>'
    return f'<div class="thumbnail-nav">{thumbnails}</div>'


def generate_html(ppt_config, reference_image=None):
    """生成完整的HTML文件"""
    theme = ppt_config.get("theme", "渐变紫")
    animation = ppt_config.get("animation", "组合动画")
    interactions = ppt_config.get("interactions", ppt_config.get("features", ["keyboard-nav", "progress-bar", "fullscreen"]))
    slides = ppt_config.get("slides", [])
    title = ppt_config.get("title", "网页PPT")

    # 获取主题配置
    theme_config = THEMES.get(theme, THEMES["渐变紫"])

    # 生成幻灯片HTML
    slides_html = ""
    for i, slide in enumerate(slides):
        slides_html += generate_slide_content(slide, theme_config)

    # 生成缩略图导航
    thumbnail_nav = ""
    if "thumbnail" in interactions or "thumbnail-nav" in interactions:
        thumbnail_nav = generate_thumbnails_html(len(slides))

    # 生成激光笔元素
    laser_pointer = ""
    if "laser" in interactions or "laser-pointer" in interactions:
        laser_pointer = '<div class="laser-pointer"></div>'

    # 生成完整HTML
    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
{generate_css(theme, animation, reference_image)}
    </style>
</head>
<body>
    <div class="slides-container">
{slides_html}
    </div>

    {'<div class="progress-bar"></div>' if 'progress' in interactions or 'progress-bar' in interactions else ''}
    {'<div class="page-indicator">1 / ' + str(len(slides)) + '</div>' if 'progress' in interactions or 'progress-bar' in interactions else ''}
    {'<button class="nav-button nav-prev">&#10094;</button>' if 'click' in interactions or 'mouse-click' in interactions else ''}
    {'<button class="nav-button nav-next">&#10095;</button>' if 'click' in interactions or 'mouse-click' in interactions else ''}
    {'<button class="fullscreen-btn">全屏 (F)</button>' if 'fullscreen' in interactions else ''}
    {thumbnail_nav}
    {laser_pointer}
    {'<div class="swipe-hint">← 左右滑动或点击边缘翻页 →</div>' if 'touch' in interactions or 'swipe' in interactions else ''}

    <script>
{generate_js(interactions)}
    </script>
</body>
</html>"""

    return html


def image_to_base64(image_path):
    """将图片转换为base64编码"""
    try:
        with open(image_path, 'rb') as f:
            return base64.b64encode(f.read()).decode('utf-8')
    except Exception as e:
        print(f"⚠️  无法读取参考图片: {e}")
        return None


def load_config(config_path):
    """加载配置文件"""
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_html(html, output_path):
    """保存HTML文件"""
    # 确保输出路径以.html结尾
    if not output_path.endswith('.html'):
        output_path += '.html'

    os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else '.', exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    return output_path


def main():
    parser = argparse.ArgumentParser(
        description="网页PPT生成器 - 生成可交互的HTML格式PPT",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 使用配置文件生成
  python generate_ppt.py --config config.json --output my-ppt.html

  # 使用配置文件并传入参考图片
  python generate_ppt.py --config config.json --output my-ppt.html --reference-image ./style.png

  # 快速生成（使用命令行参数）
  python generate_ppt.py --title "我的PPT" --theme 科技蓝 --animation 滑动进入 --output presentation.html

  # 查看所有可用选项
  python generate_ppt.py --list-themes
  python generate_ppt.py --list-animations
  python generate_ppt.py --list-templates

可用主题: 科技蓝、商务黑、清新绿、渐变紫、渐变橙、毛玻璃、中国红、深海蓝、极简白、赛博朋克
可用动画: 无动画、淡入淡出、滑动进入、缩放效果、组合动画、3D翻转、立方体、覆盖、揭开
        """
    )

    parser.add_argument("--config", "-c", help="配置文件路径 (JSON格式)")
    parser.add_argument("--output", "-o", default="ppt-output/index.html", help="输出HTML文件路径")
    parser.add_argument("--title", "-t", help="PPT标题")
    parser.add_argument("--theme", default="渐变紫", help="主题风格")
    parser.add_argument("--animation", default="组合动画", help="动画效果")
    parser.add_argument("--reference-image", "-i", help="参考图片路径（用于样式参考）")
    parser.add_argument("--list-themes", action="store_true", help="列出所有可用主题")
    parser.add_argument("--list-animations", action="store_true", help="列出所有可用动画")
    parser.add_argument("--list-templates", action="store_true", help="列出所有可用模板")

    args = parser.parse_args()

    # 列出可用选项
    if args.list_themes:
        print("\n🎨 可用主题:")
        for theme in THEMES.keys():
            print(f"  - {theme}")
        return

    if args.list_animations:
        print("\n✨ 可用动画:")
        for anim in ANIMATIONS.keys():
            print(f"  - {anim}")
        return

    if args.list_templates:
        print("\n📐 可用模板:")
        for template in TEMPLATES.keys():
            print(f"  - {template}")
        return

    # 处理参考图片
    reference_image = None
    if args.reference_image and os.path.exists(args.reference_image):
        reference_image = image_to_base64(args.reference_image)
        print(f"📷 已加载参考图片: {args.reference_image}")

    # 加载配置
    if args.config and os.path.exists(args.config):
        config = load_config(args.config)
        print(f"📄 已加载配置文件: {args.config}")
    else:
        # 使用命令行参数创建默认配置
        config = {
            "title": args.title or "我的网页PPT",
            "theme": args.theme,
            "animation": args.animation,
            "interactions": ["keyboard-nav", "mouse-click", "progress-bar", "fullscreen", "thumbnail-nav", "swipe"],
            "slides": [
                {"type": "title", "title": args.title or "我的网页PPT", "subtitle": "点击开始演示"},
                {"type": "content", "title": "使用说明", "content": ["使用方向键 ← → 切换页面", "点击屏幕左右边缘翻页", "按 F 键进入全屏", "按 Home/End 跳转到首尾页"]}
            ]
        }

    # 生成HTML
    html = generate_html(config, reference_image)

    # 保存（确保是.html格式）
    output_path = save_html(html, args.output)

    print(f"\n✅ PPT已成功生成: {output_path}")
    print(f"\n📖 使用说明:")
    print("  • 双击文件即可在浏览器中打开")
    print("  • 方向键 ← → 或点击边缘翻页")
    print("  • 按 F 键进入全屏模式")
    print("  • 按 Home/End 跳转到首尾页")
    print("  • 支持触摸滑动（移动设备）")


if __name__ == "__main__":
    main()
