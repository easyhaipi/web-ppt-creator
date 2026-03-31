#!/usr/bin/env python3
"""
交互式PPT创建器 - Interactive PPT Creator
通过对话式问答一步步引导用户创建PPT

使用方式:
    python interactive_creator.py
"""

import json
import os
from datetime import datetime


class InteractivePPTCreator:
    """交互式PPT创建器"""

    def __init__(self):
        self.ppt_data = {
            "title": "",
            "subtitle": "",
            "slides": [],
            "theme": "渐变紫",
            "animation": "组合动画",
            "interactions": ["keyboard-nav", "mouse-click", "progress-bar", "fullscreen"]
        }
        self.current_step = 0
        self.total_slides = 0

    def print_header(self, title):
        """打印步骤标题"""
        print("\n" + "=" * 60)
        print(f"  {title}")
        print("=" * 60)

    def print_progress(self, step, total, description):
        """打印进度"""
        progress = int((step / total) * 100)
        bar_length = 30
        filled = int(bar_length * step / total)
        bar = "█" * filled + "░" * (bar_length - filled)
        print(f"\n[{bar}] {progress}% - {description}")

    def ask_question(self, question, options=None, allow_multiple=False):
        """
        提问并获取用户回答
        
        Args:
            question: 问题文本
            options: 选项列表 [(value, description), ...]
            allow_multiple: 是否允许多选
        """
        print(f"\n❓ {question}")
        
        if options:
            print("\n选项：")
            for i, (value, desc) in enumerate(options, 1):
                print(f"  {i}. {value}")
                if desc:
                    print(f"     {desc}")
        
        if allow_multiple:
            print("\n💡 提示：可以输入多个选项，用逗号分隔（如：1,3,5）")
        
        print("\n你的回答：", end=" ")
        answer = input().strip()
        
        if not answer:
            return None
            
        if options:
            try:
                if allow_multiple:
                    # 多选模式
                    indices = [int(x.strip()) - 1 for x in answer.split(",")]
                    selected = []
                    for idx in indices:
                        if 0 <= idx < len(options):
                            selected.append(options[idx][0])
                    return selected if selected else None
                else:
                    # 单选模式
                    idx = int(answer) - 1
                    if 0 <= idx < len(options):
                        return options[idx][0]
                    return answer
            except ValueError:
                return answer
        
        return answer

    def confirm(self, message):
        """确认对话框"""
        print(f"\n{message}")
        print("  1. ✅ 确认继续")
        print("  2. 🔄 重新输入")
        print("  3. ⏭️  跳过此步骤")
        print("\n你的选择：", end=" ")
        choice = input().strip()
        return choice

    def step_1_basic_info(self):
        """步骤1: 基本信息"""
        self.print_header("步骤 1/8: PPT基本信息")
        self.print_progress(1, 8, "设置标题和主题")
        
        print("\n📝 让我们开始创建你的PPT！")
        print("首先，我需要了解一些基本信息。\n")
        
        # PPT标题
        while True:
            title = self.ask_question("请给你的PPT起一个标题：")
            if title:
                self.ppt_data["title"] = title
                break
            print("⚠️  标题不能为空，请重新输入。")
        
        # 副标题（可选）
        subtitle = self.ask_question("请添加副标题（可选，直接回车跳过）：")
        if subtitle:
            self.ppt_data["subtitle"] = subtitle
        
        # 用途/场景
        purpose_options = [
            ("工作汇报", "向上级汇报工作进展、成果"),
            ("产品发布", "新产品/功能发布、路演"),
            ("培训教学", "内部培训、知识分享"),
            ("项目提案", "项目立项、方案汇报"),
            ("年终总结", "年度工作总结、回顾"),
            ("技术分享", "技术方案、架构讲解"),
            ("个人展示", "简历、作品集、自我介绍"),
            ("其他", "自定义场景")
        ]
        purpose = self.ask_question("这个PPT主要用于什么场景？", purpose_options)
        self.ppt_data["purpose"] = purpose
        
        print(f"\n✅ 已设置：")
        print(f"   标题：{self.ppt_data['title']}")
        if self.ppt_data['subtitle']:
            print(f"   副标题：{self.ppt_data['subtitle']}")
        print(f"   场景：{purpose}")

    def step_2_outline_discussion(self):
        """步骤2: 大纲讨论"""
        self.print_header("步骤 2/8: PPT大纲规划")
        self.print_progress(2, 8, "设计页面结构")
        
        print("\n📋 现在我们来规划PPT的整体结构。")
        print("根据你的场景，我推荐以下几种大纲结构：\n")
        
        # 根据用途推荐大纲
        outline_templates = {
            "工作汇报": [
                ("封面", "标题+副标题"),
                ("工作概述", "本期工作重点"),
                ("主要成果", "关键数据和成果"),
                ("问题与挑战", "遇到的困难"),
                ("下一步计划", "后续工作安排"),
                ("结束页", "感谢/联系方式")
            ],
            "产品发布": [
                ("封面", "产品名称+口号"),
                ("市场痛点", "解决的问题"),
                ("产品介绍", "核心功能和亮点"),
                ("产品演示", "功能展示"),
                ("竞争优势", "与竞品对比"),
                ("发展规划", "路线图"),
                ("结束页", "行动号召")
            ],
            "培训教学": [
                ("封面", "课程标题"),
                ("课程目标", "学习收获"),
                ("知识讲解", "核心概念"),
                ("案例分析", "实际案例"),
                ("实践练习", "动手操作"),
                ("总结回顾", "要点梳理"),
                ("结束页", "Q&A/资源")
            ],
            "项目提案": [
                ("封面", "项目名称"),
                ("项目背景", "为什么做"),
                ("项目目标", "要达成什么"),
                ("解决方案", "怎么做"),
                ("项目计划", "时间安排"),
                ("资源需求", "需要什么支持"),
                ("预期收益", "价值回报"),
                ("结束页", "呼吁支持")
            ],
            "年终总结": [
                ("封面", "年度主题"),
                ("年度回顾", "整体概况"),
                ("主要成就", "亮点工作"),
                ("数据展示", "关键指标"),
                ("经验总结", "收获与反思"),
                ("明年规划", "新年目标"),
                ("结束页", "感谢")
            ],
            "技术分享": [
                ("封面", "技术主题"),
                ("背景介绍", "技术背景"),
                ("核心原理", "技术详解"),
                ("实现方案", "架构/代码"),
                ("实践案例", "应用场景"),
                ("总结展望", "未来方向"),
                ("结束页", "Q&A")
            ],
            "个人展示": [
                ("封面", "姓名+定位"),
                ("关于我", "个人简介"),
                ("专业技能", "核心能力"),
                ("项目经验", "代表作品"),
                ("成果展示", "数据/奖项"),
                ("联系方式", "找到我"),
                ("结束页", "感谢")
            ],
            "其他": [
                ("封面", "标题"),
                ("内容页1", "第一部分"),
                ("内容页2", "第二部分"),
                ("内容页3", "第三部分"),
                ("结束页", "结尾")
            ]
        }
        
        recommended = outline_templates.get(self.ppt_data.get("purpose", "其他"), outline_templates["其他"])
        
        print("推荐大纲结构：")
        for i, (page_type, desc) in enumerate(recommended, 1):
            print(f"  {i}. {page_type} - {desc}")
        
        print("\n💡 你可以：")
        print("  1. 使用推荐大纲")
        print("  2. 自定义修改大纲")
        print("  3. 完全自定义创建")
        
        choice = self.ask_question("你的选择？", [
            ("使用推荐", "直接使用上述大纲"),
            ("修改推荐", "在推荐基础上调整"),
            ("完全自定义", "从头开始创建")
        ])
        
        if choice == "使用推荐":
            self.ppt_data["outline"] = recommended
        elif choice == "修改推荐":
            print("\n📝 请告诉我需要如何调整：")
            print("  - 添加页面：告诉我要在哪一页后添加什么")
            print("  - 删除页面：告诉我要删除哪一页")
            print("  - 修改页面：告诉我要修改哪一页的内容")
            print("  - 确认完成：输入'完成'结束修改")
            
            custom_outline = list(recommended)
            while True:
                print(f"\n当前大纲：")
                for i, (page_type, desc) in enumerate(custom_outline, 1):
                    print(f"  {i}. {page_type} - {desc}")
                
                modify = input("\n调整（或输入'完成'）：").strip()
                if modify.lower() in ["完成", "done", "ok"]:
                    break
                # 这里可以添加更复杂的修改逻辑
                print("💡 已记录你的修改需求")
            
            self.ppt_data["outline"] = custom_outline
        else:
            # 完全自定义
            print("\n📝 让我们一步步创建你的自定义大纲。")
            custom_outline = []
            
            # 封面
            custom_outline.append(("封面", "标题页"))
            
            # 内容页
            page_num = 1
            while True:
                page_title = input(f"\n第{page_num}页内容页标题（或输入'结束'完成）：").strip()
                if page_title.lower() in ["结束", "done", "finish"]:
                    break
                if page_title:
                    custom_outline.append((page_title, "内容页"))
                    page_num += 1
            
            # 结束页
            custom_outline.append(("结束页", "结尾"))
            self.ppt_data["outline"] = custom_outline
        
        self.total_slides = len(self.ppt_data["outline"])
        print(f"\n✅ 大纲已确定，共 {self.total_slides} 页")

    def step_3_page_details(self):
        """步骤3: 逐页内容细化"""
        self.print_header("步骤 3/8: 逐页内容细化")
        self.print_progress(3, 8, "填写每页内容")
        
        print("\n📝 现在让我们详细设计每一页的内容。")
        print("我会逐页询问，你可以告诉我具体的内容。\n")
        
        slides = []
        
        for i, (page_type, desc) in enumerate(self.ppt_data["outline"], 1):
            print(f"\n{'─' * 50}")
            print(f"📄 第 {i}/{self.total_slides} 页: {page_type}")
            print(f"{'─' * 50}")
            
            slide = {
                "type": "content",
                "title": "",
                "content": []
            }
            
            # 根据页面类型选择模板
            if page_type == "封面" or i == 1:
                slide["type"] = "title"
                slide["title"] = self.ppt_data["title"]
                slide["subtitle"] = self.ppt_data["subtitle"] or "点击开始演示"
                print(f"✅ 封面页已自动设置")
                
            elif page_type == "结束页" or i == self.total_slides:
                slide["type"] = "end"
                print("\n结束页内容：")
                end_title = input("  大标题（如：谢谢观看/Questions?）：").strip() or "谢谢观看"
                end_subtitle = input("  副标题（可选）：").strip()
                slide["title"] = end_title
                if end_subtitle:
                    slide["subtitle"] = end_subtitle
                
            else:
                # 内容页
                print(f"\n这页是：{desc}")
                
                # 选择页面模板
                template_options = [
                    ("内容列表", "要点列表形式，适合多个并列内容"),
                    ("卡片网格", "卡片式展示，适合特性/优势"),
                    ("对比布局", "左右对比，适合前后/优劣对比"),
                    ("时间轴", "时间线形式，适合发展历程"),
                    ("数据展示", "大数字展示，适合统计数据"),
                    ("引用页", "引用名言/金句"),
                    ("图文混排", "文字+图片描述")
                ]
                
                template = self.ask_question("选择页面布局：", template_options)
                
                # 页面标题
                title = input(f"  页面标题：").strip()
                if title:
                    slide["title"] = title
                else:
                    slide["title"] = page_type
                
                # 根据模板收集内容
                if template == "内容列表":
                    slide["type"] = "content"
                    print("\n请输入要点内容（每行一个，输入空行结束）：")
                    points = []
                    while True:
                        point = input(f"  要点 {len(points)+1}：").strip()
                        if not point:
                            break
                        points.append(point)
                    slide["content"] = points if points else ["内容待补充"]
                    
                elif template == "卡片网格":
                    slide["type"] = "card-grid"
                    print("\n请输入卡片内容（每个卡片包含标题和描述）：")
                    cards = []
                    while True:
                        card_title = input(f"  卡片 {len(cards)+1} 标题（输入空行结束）：").strip()
                        if not card_title:
                            break
                        card_desc = input(f"    描述：").strip()
                        cards.append({"title": card_title, "desc": card_desc or "详细说明"})
                    slide["content"] = cards if cards else [{"title": "卡片1", "desc": "描述"}]
                    
                elif template == "对比布局":
                    slide["type"] = "comparison"
                    left_title = input("  左侧标题：").strip() or "方案A"
                    right_title = input("  右侧标题：").strip() or "方案B"
                    
                    print(f"\n  {left_title} 的要点：")
                    left_items = []
                    while True:
                        item = input(f"    要点 {len(left_items)+1}（空行结束）：").strip()
                        if not item:
                            break
                        left_items.append(item)
                    
                    print(f"\n  {right_title} 的要点：")
                    right_items = []
                    while True:
                        item = input(f"    要点 {len(right_items)+1}（空行结束）：").strip()
                        if not item:
                            break
                        right_items.append(item)
                    
                    slide["left"] = {"title": left_title, "items": left_items or ["待补充"]}
                    slide["right"] = {"title": right_title, "items": right_items or ["待补充"]}
                    
                elif template == "时间轴":
                    slide["type"] = "timeline"
                    print("\n请输入时间节点（时间 + 事件）：")
                    events = []
                    while True:
                        time = input(f"  节点 {len(events)+1} 时间（如：2024年Q1，空行结束）：").strip()
                        if not time:
                            break
                        event = input(f"    事件：").strip()
                        events.append({"time": time, "event": event or "重要事件"})
                    slide["content"] = events if events else [{"time": "时间", "event": "事件"}]
                    
                elif template == "数据展示":
                    slide["type"] = "data"
                    print("\n请输入数据（数字 + 标签）：")
                    data_items = []
                    while True:
                        number = input(f"  数据 {len(data_items)+1} 数字（如：99%，空行结束）：").strip()
                        if not number:
                            break
                        label = input(f"    标签：").strip()
                        data_items.append({"number": number, "label": label or "指标"})
                    slide["content"] = data_items if data_items else [{"number": "0", "label": "待补充"}]
                    
                elif template == "引用页":
                    slide["type"] = "quote"
                    quote = input("  引用内容：").strip()
                    author = input("  作者/来源：").strip()
                    slide["quote"] = quote or "引用内容"
                    slide["author"] = author or "未知"
                    
                else:  # 图文混排
                    slide["type"] = "content"
                    print("\n请输入要点内容：")
                    points = []
                    while True:
                        point = input(f"  要点 {len(points)+1}（空行结束）：").strip()
                        if not point:
                            break
                        points.append(point)
                    slide["content"] = points if points else ["内容待补充"]
            
            slides.append(slide)
            print(f"✅ 第 {i} 页已保存")
        
        self.ppt_data["slides"] = slides
        print(f"\n✅ 所有页面内容已收集完成！")

    def step_4_theme_selection(self):
        """步骤4: 主题风格选择"""
        self.print_header("步骤 4/8: 选择主题风格")
        self.print_progress(4, 8, "选择视觉风格")
        
        print("\n🎨 选择一个适合你PPT的视觉主题：\n")
        
        theme_options = [
            ("科技蓝", "蓝色渐变，现代科技感，适合技术/产品主题"),
            ("商务黑", "深灰黑色，专业商务，适合正式汇报"),
            ("清新绿", "绿色系，自然清新，适合环保/健康主题"),
            ("渐变紫", "紫粉渐变，梦幻优雅，适合创意/设计主题"),
            ("渐变橙", "橙黄渐变，活力温暖，适合营销/推广主题"),
            ("毛玻璃", "半透明效果，时尚现代，通用型"),
            ("中国红", "红色系，喜庆大气，适合节日/庆典"),
            ("深海蓝", "深蓝渐变，深邃沉稳，适合金融/企业"),
            ("极简白", "白色主调，简洁干净，适合设计/艺术"),
            ("赛博朋克", "霓虹色，高对比，适合游戏/科幻/创意")
        ]
        
        # 根据用途推荐
        purpose = self.ppt_data.get("purpose", "")
        recommendations = {
            "工作汇报": ["商务黑", "深海蓝", "极简白"],
            "产品发布": ["科技蓝", "渐变紫", "赛博朋克"],
            "培训教学": ["清新绿", "渐变橙", "极简白"],
            "项目提案": ["商务黑", "深海蓝", "科技蓝"],
            "年终总结": ["中国红", "渐变紫", "商务黑"],
            "技术分享": ["科技蓝", "赛博朋克", "渐变紫"],
            "个人展示": ["渐变紫", "赛博朋克", "极简白"]
        }
        
        if purpose in recommendations:
            print(f"💡 根据「{purpose}」场景，推荐主题：")
            for theme in recommendations[purpose]:
                print(f"   • {theme}")
            print()
        
        theme = self.ask_question("选择主题：", theme_options)
        self.ppt_data["theme"] = theme
        print(f"\n✅ 已选择主题：{theme}")

    def step_5_animation_selection(self):
        """步骤5: 动画效果选择"""
        self.print_header("步骤 5/8: 选择动画效果")
        self.print_progress(5, 8, "选择切换动画")
        
        print("\n✨ 选择幻灯片切换时的动画效果：\n")
        
        animation_options = [
            ("淡入淡出", "透明度渐变，优雅简洁，最常用"),
            ("滑动进入", "左右滑动，流程感强"),
            ("缩放效果", "大小变化，强调重点"),
            ("组合动画", "滑动+缩放，现代感强，推荐"),
            ("3D翻转", "Y轴旋转，创意展示"),
            ("立方体", "3D立方体旋转，科技感"),
            ("覆盖", "上下覆盖，层次感强"),
            ("揭开", "像翻页一样，正式传统"),
            ("无动画", "静态切换，快速演示")
        ]
        
        animation = self.ask_question("选择动画：", animation_options)
        self.ppt_data["animation"] = animation
        print(f"\n✅ 已选择动画：{animation}")

    def step_6_interaction_selection(self):
        """步骤6: 交互功能选择"""
        self.print_header("步骤 6/8: 选择交互功能")
        self.print_progress(6, 8, "配置交互方式")
        
        print("\n🎮 选择你需要的交互功能（可多选）：\n")
        
        interaction_options = [
            ("keyboard-nav", "键盘导航 - 方向键翻页（推荐）"),
            ("mouse-click", "鼠标点击 - 点击边缘翻页（推荐）"),
            ("touch-swipe", "触摸滑动 - 支持手机/平板滑动"),
            ("progress-bar", "进度指示 - 显示页码和进度条（推荐）"),
            ("thumbnail-nav", "缩略图导航 - 底部快速跳转"),
            ("laser-pointer", "激光笔效果 - 按L键开启/关闭"),
            ("fullscreen", "全屏模式 - 按F键切换（推荐）")
        ]
        
        interactions = self.ask_question("选择交互功能（输入数字，多选用逗号分隔）：", 
                                        interaction_options, 
                                        allow_multiple=True)
        
        if interactions:
            self.ppt_data["interactions"] = interactions
        else:
            self.ppt_data["interactions"] = ["keyboard-nav", "mouse-click", "progress-bar", "fullscreen"]
        
        print(f"\n✅ 已选择 {len(self.ppt_data['interactions'])} 项交互功能")

    def step_7_preview_confirm(self):
        """步骤7: 预览确认"""
        self.print_header("步骤 7/8: 预览确认")
        self.print_progress(7, 8, "确认配置信息")
        
        print("\n📋 PPT配置预览：")
        print("=" * 50)
        print(f"📝 标题：{self.ppt_data['title']}")
        if self.ppt_data.get('subtitle'):
            print(f"📝 副标题：{self.ppt_data['subtitle']}")
        print(f"🎯 场景：{self.ppt_data.get('purpose', '未设置')}")
        print(f"📄 页数：{len(self.ppt_data['slides'])} 页")
        print(f"🎨 主题：{self.ppt_data['theme']}")
        print(f"✨ 动画：{self.ppt_data['animation']}")
        print(f"🎮 交互：{', '.join(self.ppt_data['interactions'])}")
        print("=" * 50)
        
        print("\n📑 页面大纲：")
        for i, slide in enumerate(self.ppt_data['slides'], 1):
            slide_type = slide.get('type', 'content')
            title = slide.get('title', '未命名')
            print(f"  {i}. [{slide_type}] {title}")
        
        print("\n💡 确认信息无误吗？")
        choice = self.confirm("是否继续生成PPT？")
        
        if choice == "2":
            print("\n🔄 请告诉我需要修改哪里：")
            print("  1. 修改基本信息（标题/副标题）")
            print("  2. 修改页面内容")
            print("  3. 修改主题/动画")
            print("  4. 全部重新来")
            
            modify_choice = input("\n选择：").strip()
            if modify_choice == "1":
                self.step_1_basic_info()
            elif modify_choice == "2":
                self.step_3_page_details()
            elif modify_choice == "3":
                self.step_4_theme_selection()
                self.step_5_animation_selection()
            elif modify_choice == "4":
                return False
            return self.step_7_preview_confirm()
        
        elif choice == "3":
            print("⏭️  跳过确认，继续生成...")
        
        return True

    def step_8_generate(self):
        """步骤8: 生成PPT"""
        self.print_header("步骤 8/8: 生成PPT")
        self.print_progress(8, 8, "生成HTML文件")
        
        print("\n🚀 正在生成你的PPT...")
        
        # 生成文件名
        safe_title = "".join(c for c in self.ppt_data['title'] if c.isalnum() or c in (' ', '-', '_')).strip()
        safe_title = safe_title.replace(' ', '_') or 'presentation'
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{safe_title}_{timestamp}.html"
        output_path = os.path.join('ppt-output', filename)
        
        # 确保输出目录存在
        os.makedirs('ppt-output', exist_ok=True)
        
        # 保存配置
        config_path = os.path.join('ppt-output', f"{safe_title}_{timestamp}_config.json")
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(self.ppt_data, f, ensure_ascii=False, indent=2)
        
        print(f"\n✅ 配置已保存：{config_path}")
        
        # 调用generate_ppt.py生成HTML
        import subprocess
        try:
            result = subprocess.run(
                ['python', 'scripts/generate_ppt.py', 
                 '--config', config_path, 
                 '--output', output_path],
                capture_output=True,
                text=True,
                cwd=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            )
            
            if result.returncode == 0:
                print(f"\n🎉 PPT生成成功！")
                print(f"\n📁 文件位置：{output_path}")
                print(f"\n📖 使用说明：")
                print("  1. 双击HTML文件在浏览器中打开")
                print("  2. 方向键 ← → 翻页")
                print("  3. 点击屏幕左右边缘翻页")
                print("  4. 按 F 键进入全屏")
                print("  5. 按 Home/End 跳转到首尾页")
                
                # 询问是否预览
                print("\n🌐 是否立即启动预览服务器？")
                print("  1. 是，立即预览")
                print("  2. 否，稍后手动打开")
                
                preview = input("\n选择：").strip()
                if preview == "1":
                    print("\n🚀 启动预览服务器...")
                    subprocess.Popen(
                        ['python', 'scripts/preview.py', 
                         '--html', output_path, 
                         '--port', '8080', 
                         '--open'],
                        cwd=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                    )
            else:
                print(f"\n❌ 生成失败：")
                print(result.stderr)
                
        except Exception as e:
            print(f"\n❌ 生成出错：{e}")
            print("\n💡 你可以手动运行：")
            print(f"  python scripts/generate_ppt.py --config {config_path} --output {output_path}")

    def run(self):
        """运行完整的交互式创建流程"""
        print("\n" + "=" * 60)
        print("  🎨 网页PPT创建器 - 交互式向导")
        print("=" * 60)
        print("\n欢迎使用！我将通过一系列问题帮你创建精美的网页PPT。")
        print("整个过程大约需要 5-10 分钟。\n")
        
        input("按回车键开始...")
        
        try:
            self.step_1_basic_info()
            self.step_2_outline_discussion()
            self.step_3_page_details()
            self.step_4_theme_selection()
            self.step_5_animation_selection()
            self.step_6_interaction_selection()
            
            if self.step_7_preview_confirm():
                self.step_8_generate()
            else:
                # 重新开始
                self.run()
                
        except KeyboardInterrupt:
            print("\n\n👋 已取消创建。你的进度已自动保存。")
            # 保存当前进度
            try:
                os.makedirs('ppt-output', exist_ok=True)
                save_path = 'ppt-output/interrupted_draft.json'
                with open(save_path, 'w', encoding='utf-8') as f:
                    json.dump(self.ppt_data, f, ensure_ascii=False, indent=2)
                print(f"💾 草稿已保存：{save_path}")
            except:
                pass


def main():
    """主函数"""
    creator = InteractivePPTCreator()
    creator.run()


if __name__ == "__main__":
    main()
