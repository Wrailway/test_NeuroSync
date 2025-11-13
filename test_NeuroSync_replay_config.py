import random
import time
import os
import warnings
from pywinauto import Application, mouse
from pywinauto.findwindows import find_window, ElementNotFoundError
from pywinauto.timings import wait_until_passes

# 过滤窗口聚焦警告
warnings.filterwarnings("ignore", category=RuntimeWarning, message="The window has not been focused due to COMError")

# 基础配置
UI_TIMIEOUT = 3
TEST_DURATION = 60#12 * 3600  # 12小时
CYCLE_INTERVAL = 10  # 循环间隔（秒）

# 功能开关配置（控制是否执行）
RUN_CONFIG = {
    "dropdown_config": True,    # 下拉框配置
    "nav_buttons": True,        # 导航按钮
    "drag_progress_bar": True,  # 进度条拖拽
    "tag_marking": True,        # 标签标记
    "channel_selection": True   # 通道选择（默认初始化后关闭）
}

# 业务参数配置（核心：下拉框保留auto_id，新增序号定位开关）
CONFIG = {
    # 核心路径配置（移至最前，方便修改）
    "APP_PATH": r"D:\Program Files\NeuroSync\Replay3\NeuroSync.Client.Replay.exe",
    "FILE_DIR": r"D:\edfx\V25OB3000test20250924191457\V25OB3000test20250924191457",
    
    # 通道选择配置（紧随路径配置）
    "CHANNEL_CONFIG": {
        "target_channels": [2, 3],
        "btn_cl_auto_id": "btn_cl",
        "btn_confirm_auto_id": "btn_confirm"
    },

    # 下拉框配置（优先用auto_id，无则开启use_index）
    "DROPDOWNS": {
        "DAO_LIAN": {
            "auto_id": "cb_daolian",  # 导联有auto_id，优先使用
            "use_index": False,       # 不启用序号定位
            "fixed_index": 0,         # 固定选择第1项
            "option_name": "导联"
        },
        "SWEEP_SPEED": {
            "auto_id": "cb_zouzhisudu",  # 走纸速度有auto_id
            "use_index": False,
            "range": range(0, 22),
            "is_random": True,
            "option_name": "走纸速度"
        },
        "SENSITIVITY": {
            "auto_id": "cb_lingmindu",   # 灵敏度有auto_id
            "use_index": False,
            "range": range(0, 16),
            "is_random": True,
            "option_name": "灵敏度"
        },
        "HIGH_PASS": {
            "auto_id": "cb_lvboxiaxian",  # 高通滤波有auto_id
            "use_index": False,
            "range": range(0, 11),
            "is_random": True,
            "option_name": "高通滤波"
        },
        "LOW_PASS": {
            "auto_id": "cb_lvboshangxian", # 低通滤波有auto_id
            "use_index": False,
            "range": range(0, 6),
            "is_random": True,
            "option_name": "低通滤波"
        },
        "PLAYBACK_SPEED": {
            "auto_id": "cb_bofangbeishu",  # 播放倍速有auto_id
            "use_index": False,
            "range": range(0, 8),
            "is_random": True,
            "option_name": "播放倍速"
        },
        # 示例：无auto_id的下拉框（启用序号定位）
        # "CUSTOM_COMBO": {
        #     "auto_id": "",
        #     "use_index": True,    # 启用序号定位
        #     "combo_index": 0,     # 第1个ComboBox
        #     "range": range(0, 5),
        #     "is_random": True,
        #     "option_name": "自定义下拉框"
        # }
    },

    # 其他配置
    "TAG_LIST": {
        'Eyes Open', 'Eyes Closed', 'Seizure', 'Deep Breath', 'Background', 'Awake',
        'Eyes Closed PPR', 'Eyes Shut PPR', 'Eyes Open PCR', 'Electrical Seizure', 'End',
        'Identify Event', 'Seizure*', 'Drug-induced Sleep', 'Mechanical Ventilation', 'Facial Twitching'
    },
    "TAG_CONFIG": {
        "count_range": (2, 10),
        "max_down_retries": 5,
        "rollback_after_fail": True
    },
    "NAV_BUTTONS": [
        {"title_re": "Previous Second", "name": "上一秒", "click_count": 5, "interval": 0.3},
        {"title_re": "Next Second", "name": "下一秒", "click_count": 5, "interval": 0.3},
        {"title_re": "Previous Page", "name": "上一页", "click_count": 3, "interval": 0.3},
        {"title_re": "Next Page", "name": "下一页", "click_count": 3, "interval": 0.3}
    ],
    "PROGRESS_BAR": {
        "auto_id": "slider_play_jd",
        "drag_cycles": 10,
        "min_percent": 1,
        "max_percent": 99
    }
}
# 统计变量
STATS = {
    "total_cycles": 0,
    "success_cycles": 0,
    "fail_cycles": 0,
    "module_stats": {
        "dropdown_config": {"success": 0, "fail": 0},
        "nav_buttons": {"success": 0, "fail": 0},
        "drag_progress_bar": {"success": 0, "fail": 0},
        "tag_marking": {"success": 0, "fail": 0},
        "channel_selection": {"success": 0, "fail": 0}
    },
    "error_log": []
}

def safe_set_focus(window, max_retries=3, delay=0.5):
    """安全聚焦窗口，带重试机制"""
    for i in range(max_retries):
        try:
            window.set_focus()
            return True
        except Exception:
            if i == max_retries - 1:
                print(f"警告：窗口聚焦失败（已重试{max_retries}次）")
            time.sleep(delay)
    return False

def select_dropdown_option(main_window, config):
    """下拉框选择函数（优先auto_id，无则用序号）"""
    try:
        option_name = config["option_name"]
        # 1. 定位下拉框（优先用auto_id，无则用序号）
        if not config["use_index"] and config["auto_id"]:
            # 原有逻辑：用auto_id定位（不修改）
            dropdown = main_window.child_window(
                auto_id=config["auto_id"],
                control_type="ComboBox"
            )
        else:
            # 新增逻辑：用序号定位（仅对无auto_id的控件启用）
            if not config.get("combo_index"):
                raise Exception(f"{option_name}未配置序号（combo_index）")
            dropdown = main_window.child_window(
                control_type="ComboBox",
                found_index=config["combo_index"]
            )

        dropdown.wait("visible", timeout=UI_TIMIEOUT)
        dropdown.wait("enabled", timeout=UI_TIMIEOUT)
        safe_set_focus(dropdown)
        dropdown.click_input()
        time.sleep(1.5)

        # 2. 获取选项列表
        items = dropdown.descendants(control_type="ListItem")
        if not items:
            raise Exception(f"未找到任何{option_name}选项")
        max_valid_index = len(items) - 1

        # 3. 确定目标索引
        if config.get("is_random", False):
            valid_indices = [idx for idx in config["range"] if 0 <= idx <= max_valid_index]
            if not valid_indices:
                raise Exception(f"{option_name}无有效索引（最大{max_valid_index}）")
            target_index = random.choice(valid_indices)
        else:
            target_index = config["fixed_index"]
            if not (0 <= target_index <= max_valid_index):
                raise Exception(f"{option_name}索引{target_index}超出范围（最大{max_valid_index}）")

        # 4. 检查当前选中项
        current_selected = next((item for item in items if item.is_selected()), None)
        current_index = items.index(current_selected) if current_selected else 0

        if current_index == target_index:
            print(f"当前已选中{option_name}第{target_index+1}项，无需切换")
            dropdown.type_keys("{ESC}")
            return

        # 5. 切换选项
        dropdown.type_keys("{HOME}")  # 跳至第一项
        time.sleep(0.3)
        dropdown.type_keys(f"{{DOWN {target_index}}}")
        time.sleep(0.3)
        target_item = items[target_index]
        safe_set_focus(target_item)
        target_item.click_input()
        dropdown.type_keys("{ESC}")
        print(f"{option_name}已选择第{target_index+1}项")

    except Exception as e:
        raise Exception(f"{option_name}选择失败：{str(e)}")

def execute_dropdown_config(main_window):
    """执行下拉框配置（保留auto_id逻辑，兼容序号定位）"""
    if not RUN_CONFIG["dropdown_config"]:
        print("\n【下拉框配置】已关闭，跳过该模块")
        return True
    try:
        print("\n===== 开始下拉框参数配置 =====")
        # 遍历所有下拉框配置（优先auto_id）
        for dropdown_key in CONFIG["DROPDOWNS"]:
            select_dropdown_option(main_window, CONFIG["DROPDOWNS"][dropdown_key])
        print("===== 下拉框参数配置完成 =====")
        STATS["module_stats"]["dropdown_config"]["success"] += 1
        return True
    except Exception as e:
        error_msg = f"下拉框配置失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["dropdown_config"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

# 以下函数逻辑不变（标签标记、导航按钮、通道选择、进度条拖拽等）
def find_and_click_tag(main_window, target_title):
    try:
        tag_btn = main_window.child_window(
            title_re=target_title,
            control_type="Text",
            found_index=0
        )
        tag_btn.wait("visible", timeout=UI_TIMIEOUT)
        tag_btn.click_input(coords=(5, 5))
        return True
    except:
        return False

def execute_tag_marking(main_window):
    if not RUN_CONFIG["tag_marking"]:
        print("\n【标签标记】已关闭，跳过该模块")
        return True
    try:
        print("\n===== 开始标签标记 =====")
        tag_count = random.randint(*CONFIG["TAG_CONFIG"]["count_range"])
        print(f"本次循环计划标记 {tag_count} 个标签")

        for tag_idx in range(1, tag_count + 1):
            target_tag = random.choice(list(CONFIG["TAG_LIST"]))
            print(f"\n第{tag_idx}/{tag_count}个标签：尝试标记「{target_tag}」")
            down_count = 0

            if find_and_click_tag(main_window, target_tag):
                print(f"成功标记标签「{target_tag}」")
                time.sleep(0.8)
                continue

            down_button = main_window.child_window(auto_id="DownButton", control_type="Button")
            down_button.wait("visible", timeout=UI_TIMIEOUT)
            found = False
            for down_count in range(1, CONFIG["TAG_CONFIG"]["max_down_retries"] + 1):
                down_button.click_input()
                time.sleep(0.6)
                if find_and_click_tag(main_window, target_tag):
                    print(f"第{down_count}次翻页后，成功标记标签「{target_tag}」")
                    found = True
                    time.sleep(0.8)
                    break

            if not found:
                print(f"翻页{CONFIG['TAG_CONFIG']['max_down_retries']}次仍未找到标签「{target_tag}」，跳过")
                if CONFIG["TAG_CONFIG"]["rollback_after_fail"] and down_count > 0:
                    down_button.type_keys(f"{{UP {down_count}}}")
                    time.sleep(0.3)

        print("===== 标签标记完成 =====")
        STATS["module_stats"]["tag_marking"]["success"] += 1
        return True
    except Exception as e:
        error_msg = f"标签标记失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["tag_marking"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

def click_button_multiple_times(main_window, title_re, button_name, click_count=1, interval=0.5):
    try:
        btn = main_window.child_window(title_re=title_re, control_type="Button")
        btn.wait("enabled", timeout=UI_TIMIEOUT)
        safe_set_focus(btn)

        for i in range(click_count):
            btn.click_input()
            if i < click_count - 1:
                time.sleep(interval)
        print(f"{button_name}共{click_count}次点击成功")
        return True
    except Exception as e:
        raise Exception(f"{button_name}失败：{str(e)}")

def execute_nav_buttons(main_window):
    if not RUN_CONFIG["nav_buttons"]:
        print("\n【导航按钮】已关闭，跳过该模块")
        return True
    try:
        print("\n===== 开始导航按钮操作 =====")
        for btn_config in CONFIG["NAV_BUTTONS"]:
            click_button_multiple_times(
                main_window=main_window,
                title_re=btn_config["title_re"],
                button_name=btn_config["name"],
                click_count=btn_config["click_count"],
                interval=btn_config["interval"]
            )
        print("===== 导航按钮操作完成 =====")
        STATS["module_stats"]["nav_buttons"]["success"] += 1
        return True
    except Exception as e:
        error_msg = f"导航按钮操作失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["nav_buttons"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

def select_specific_channels(main_window):
    target_channels = CONFIG["CHANNEL_CONFIG"]["target_channels"]
    for num in target_channels:
        success = False
        for retry in range(3):  # 3次重试
            try:
                channel = main_window.child_window(
                    title_re=rf"^\s*{num}\s*$",
                    control_type="CheckBox"
                )
                channel.wait("visible", timeout=UI_TIMIEOUT * 2)
                channel.wait("enabled", timeout=UI_TIMIEOUT * 2)
                
                safe_set_focus(main_window)
                time.sleep(0.2)
                safe_set_focus(channel)
                
                if channel.get_toggle_state() == 0:
                    channel.click_input()
                    time.sleep(0.3)
                    if channel.get_toggle_state() == 1:
                        print(f"已选中通道 {num}")
                        success = True
                        break
                else:
                    print(f"通道 {num} 已选中")
                    success = True
                    break
            except Exception as e:
                if retry < 2:
                    print(f"通道 {num} 第{retry+1}次操作失败，重试...")
                    time.sleep(0.5)
                else:
                    raise Exception(f"通道 {num} 操作失败：{str(e)}")
        if not success:
            raise Exception(f"通道 {num} 未选中")

def execute_channel_selection(main_window):
    if not RUN_CONFIG["channel_selection"]:
        print("\n【通道选择】已关闭，跳过该模块")
        return True
    try:
        print("\n===== 开始通道选择 =====")
        # 选择第一个匹配的通道列表按钮（关键修改）
        btn_cl = main_window.child_window(
            auto_id=CONFIG["CHANNEL_CONFIG"]["btn_cl_auto_id"],
            control_type="Button",
            found_index=0  # 强制选择第一个
        )
        btn_cl.wait("enabled", timeout=UI_TIMIEOUT)
        btn_cl.click_input()
        time.sleep(3)
        print("已点击第一个通道列表展开按钮")

        # 后续逻辑不变...
        cbx_all = main_window.child_window(title='All', control_type="CheckBox")
        cbx_all.wait("enabled", timeout=UI_TIMIEOUT)
        if cbx_all.get_toggle_state() == 1:
            cbx_all.click_input()
            time.sleep(0.5)
            print("已取消全选通道")
        else:
            print("通道已处于未全选状态")

        select_specific_channels(main_window)

        btn_confirm = main_window.child_window(
            auto_id=CONFIG["CHANNEL_CONFIG"]["btn_confirm_auto_id"],
            control_type="Button",
            found_index=0  # 确认按钮也建议加索引，确保唯一性
        )
        btn_confirm.wait("enabled", timeout=UI_TIMIEOUT)
        btn_confirm.click_input()
        time.sleep(3)
        print(f"已确认选中指定通道：{CONFIG['CHANNEL_CONFIG']['target_channels']}")
        print("===== 通道选择完成 =====")
        STATS["module_stats"]["channel_selection"]["success"] += 1
        return True
    except Exception as e:
        error_msg = f"通道选择失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["channel_selection"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

def drag_progress_in_cycles(main_window):
    if not RUN_CONFIG["drag_progress_bar"]:
        print("\n【进度条拖拽】已关闭，跳过该模块")
        return True
    try:
        print("\n===== 开始进度条拖拽 =====")
        progress_config = CONFIG["PROGRESS_BAR"]
        
        # 1. 校验进度条控件是否存在
        try:
            progress_bar = main_window.child_window(
                auto_id=progress_config["auto_id"],
                control_type="Slider"
            )
            # 等待控件可见且启用
            progress_bar.wait("visible", timeout=UI_TIMIEOUT * 2)
            progress_bar.wait("enabled", timeout=UI_TIMIEOUT * 2)
        except Exception as e:
            raise Exception(f"进度条控件不存在或未就绪：{str(e)}")

        # 2. 校验滑块（Thumb）控件是否存在
        try:
            thumb = progress_bar.child_window(control_type="Thumb")
            thumb.wait("visible", timeout=UI_TIMIEOUT)
            thumb.wait("enabled", timeout=UI_TIMIEOUT)
        except Exception as e:
            raise Exception(f"进度条滑块（Thumb）不存在：{str(e)}")

        # 3. 安全获取坐标信息（增加非空校验）
        progress_rect = progress_bar.rectangle()
        if not progress_rect:
            raise Exception("无法获取进度条坐标信息（rectangle为空）")
        
        thumb_rect = thumb.rectangle()
        if not thumb_rect:
            raise Exception("无法获取滑块坐标信息（rectangle为空）")

        # 4. 计算有效拖拽长度（增加非负校验）
        valid_length = progress_rect.width() - thumb_rect.width()
        if valid_length <= 0:
            raise Exception(f"进度条有效长度异常（{valid_length}），无法拖拽")

        current_percent = 0
        target_x_prev = None

        # 5. 循环拖拽（修复循环内变量作用域问题）
        for i in range(progress_config["drag_cycles"]):
            # 计算目标百分比
            if i == 0:
                target_percent = random.randint(1, 30)
            else:
                if i % 5 == 0:
                    target_percent = int(current_percent * 0.4)
                else:
                    target_percent = int(current_percent * 1.5)
                # 限制在有效范围
                target_percent = max(progress_config["min_percent"], 
                                    min(target_percent, progress_config["max_percent"]))

            # 计算目标X坐标（确保为整数）
            target_x = progress_rect.left + int(valid_length * (target_percent / 100))
            # 限制在进度条范围内
            target_x = max(progress_rect.left, 
                         min(target_x, progress_rect.right - thumb_rect.width()))
            target_y = progress_rect.top + (progress_rect.height() // 2)

            # 确定拖拽起点
            if i == 0:
                start_x = thumb_rect.left + (thumb_rect.width() // 2)
                start_y = thumb_rect.top + (thumb_rect.height() // 2)
            else:
                start_x = target_x_prev
                start_y = target_y

            # 模拟人工拖拽
            mouse.move(coords=(start_x, start_y))
            time.sleep(0.4)
            mouse.press(button="left", coords=(start_x, start_y))
            time.sleep(0.3)

            # 分步移动（避免一次性拖拽）
            step_count = 3
            step_x = (target_x - start_x) // step_count
            step_y = (target_y - start_y) // step_count
            for step in range(1, step_count + 1):
                current_step_x = start_x + step_x * step
                current_step_y = start_y + step_y * step
                mouse.move(coords=(current_step_x, current_step_y))
                time.sleep(0.15)

            mouse.release(button="left", coords=(target_x, target_y))
            print(f"第{i+1}次拖拽完成，位置：{target_percent}%")

            # 更新状态变量
            current_percent = target_percent
            target_x_prev = target_x
            time.sleep(1.5)

        print("===== 进度条拖拽完成 =====")
        STATS["module_stats"]["drag_progress_bar"]["success"] += 1
        return True

    except Exception as e:
        error_msg = f"进度条拖拽失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["drag_progress_bar"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

def init_application():
    app = None
    try:
        print("===== 开始初始化应用 =====")
        
        # 校验文件
        bdf_files = [f for f in os.listdir(CONFIG['FILE_DIR']) if f.lower().endswith(".bdf")]
        if not bdf_files:
            raise Exception(f"目录 {CONFIG['FILE_DIR']} 中未找到.bdf文件")
        file_path = os.path.join(CONFIG['FILE_DIR'], bdf_files[0])
        print(f"找到目标文件：{file_path}")

        # 启动应用
        app = Application(backend="uia").start(CONFIG['APP_PATH'])
        main_window = app.window(title_re="NeuroSync Replay.*")
        if not main_window.wait("visible", timeout=UI_TIMIEOUT*10):
            raise TimeoutError("主窗口30秒内未可见")
        safe_set_focus(main_window)
        print("应用启动成功")

        # 加载文件
        file_btn = main_window.child_window(title="File", control_type="Button")
        file_btn.wait("enabled", timeout=UI_TIMIEOUT)
        file_btn.click_input()
        print("File按钮点击成功")

        # 处理文件对话框
        def get_dialog_handle():
            return find_window(title="打开", class_name="#32770")
        dialog_hwnd = wait_until_passes(UI_TIMIEOUT*2, 0.5, get_dialog_handle)
        file_dialog = app.window(handle=dialog_hwnd)
        file_dialog.wait("visible", timeout=UI_TIMIEOUT)
        safe_set_focus(file_dialog)

        file_edit = file_dialog.child_window(class_name="Edit")
        safe_set_focus(file_edit)
        file_edit.type_keys("{BACKSPACE}"*200)
        file_edit.type_keys(file_path, with_spaces=True, pause=0.01)
        time.sleep(0.5)
        open_btn = file_dialog.child_window(title="打开(O)", control_type="Button")
        open_btn.click_input()
        print("文件加载中...")
        time.sleep(5)

        # 启动播放
        play_btn = main_window.child_window(title_re="Play", control_type="Button")
        play_btn.wait("enabled", timeout=UI_TIMIEOUT)
        play_btn.click_input()
        time.sleep(1)
        print("播放功能启动成功")

        # 关闭Video Playback窗口
        video_window = main_window.child_window(control_type="Window", title='Video Playback')
        if video_window.exists():
            safe_set_focus(video_window)
            close_btn = video_window.child_window(control_type="Button", title='关闭')
            close_btn.click_input()
            print("Video Playback窗口已关闭")

        # 0.2s线切换
        toggle_0_2s_line = main_window.child_window(title_re="0.2s Line", control_type="CheckBox")
        toggle_0_2s_line.wait("enabled", timeout=UI_TIMIEOUT)
        toggle_0_2s_line.click_input()
        time.sleep(2)
        print("0.2s line 点击成功")

        # # 初始化通道选择
        # if RUN_CONFIG["channel_selection"]:
        #     execute_channel_selection(main_window)
        #     RUN_CONFIG["channel_selection"] = False  # 初始化后关闭，避免重复执行

        # print("===== 应用初始化完成 =====")
        return app, main_window

    except Exception as e:
        print(f"应用初始化失败：{str(e)}")
        if app and app.is_process_running():
            app.kill()
        raise

def run_cycle_operations(main_window):
    try:
        STATS['total_cycles'] += 1
        print(f"\n===== 开始第 {STATS['total_cycles']} 次循环 =====")
        
        module_results = [
            execute_dropdown_config(main_window),
            execute_channel_selection(main_window),
            execute_nav_buttons(main_window),
            drag_progress_in_cycles(main_window),
            execute_tag_marking(main_window),
            execute_channel_selection(main_window)
        ]

        if all(module_results):
            STATS['success_cycles'] += 1
            print(f"===== 第 {STATS['total_cycles']} 次循环成功 =====")
        else:
            STATS['fail_cycles'] += 1
            print(f"===== 第 {STATS['total_cycles']} 次循环部分模块失败 =====")
        
        # 每次循环结束后输出累计统计
        print(f"当前累计：总循环 {STATS['total_cycles']} 次，成功 {STATS['success_cycles']} 次，失败 {STATS['fail_cycles']} 次")
        return True

    except Exception as e:
        error_msg = f"第 {STATS['total_cycles']} 次循环异常：{str(e)}"
        print(error_msg)
        STATS['fail_cycles'] += 1
        STATS['error_log'].append(error_msg)
        # 异常时也输出累计统计
        print(f"当前累计：总循环 {STATS['total_cycles']} 次，成功 {STATS['success_cycles']} 次，失败 {STATS['fail_cycles']} 次")
        return False

def main():
    app = None
    main_window = None
    start_time = time.time()

    try:
        app, main_window = init_application()

        while (time.time() - start_time) < TEST_DURATION:
            run_cycle_operations(main_window)

            elapsed_time = time.time() - start_time
            remaining_time = TEST_DURATION - elapsed_time
            print(f"\n已运行：{elapsed_time / 3600:.2f}小时，剩余：{remaining_time / 3600:.2f}小时")

            if remaining_time > CYCLE_INTERVAL:
                time.sleep(CYCLE_INTERVAL)
            else:
                # 只在剩余时间为正数时睡眠，否则直接结束
                if remaining_time > 0:
                    time.sleep(remaining_time)
                break  # 无论剩余时间是否为正，都退出循环

        print(f"\n===== 测试完成！=====")
        print(f"总循环次数：{STATS['total_cycles']}")
        print(f"成功次数：{STATS['success_cycles']}，失败次数：{STATS['fail_cycles']}")
        print(f"总成功率：{STATS['success_cycles']/STATS['total_cycles']*100:.2f}%" if STATS['total_cycles'] else "无数据")
        
        print("\n===== 模块执行统计 =====")
        for module, stats in STATS["module_stats"].items():
            total = stats["success"] + stats["fail"]
            rate = f"{stats['success']/total*100:.2f}%" if total else "未执行"
            print(f"{module}：成功{stats['success']}次，失败{stats['fail']}次，成功率{rate}")

        if STATS['error_log']:
            print(f"\n错误日志（{len(STATS['error_log'])}条）：")
            for idx, err in enumerate(STATS['error_log'], 1):
                print(f"{idx}. {err}")

    except Exception as e:
        print(f"\n测试异常终止：{str(e)}")
    finally:
        print(f"\n===== 退出应用 =====")
        if main_window and main_window.exists():
            main_window.close()
            time.sleep(2)
        if app and app.is_process_running():
            app.kill()
        print("测试结束")

if __name__ == "__main__":
    main()