import random
import time
import os
import ctypes
import warnings
from datetime import datetime
from pywinauto import Application, mouse
from pywinauto.findwindows import find_window
from pywinauto.timings import wait_until_passes
from pywinauto.win32functions import GetSystemMetrics
from pywinauto import win32defines

# 过滤窗口聚焦警告
warnings.filterwarnings("ignore", category=RuntimeWarning, message="The window has not been focused due to COMError")

# 基础配置
UI_TIMIEOUT = 3
TEST_DURATION = 12 * 3600  # 12小时
CYCLE_INTERVAL = 10  # 循环间隔（秒）

# 功能开关配置（控制是否执行）
RUN_CONFIG = {
    "dropdown_config": True,    # 下拉框配置
    "nav_buttons": True,        # 导航按钮
    "drag_progress_bar": True,  # 进度条拖拽
    "tag_marking": True,        # 标签标记
    "channel_selection": True,  # 通道选择
    "move_video_window": True   # 视频窗口拖动   
}

# 业务参数配置
CONFIG = {
    # 核心路径配置
    "APP_PATH": r"D:\Program Files\NeuroSync\Replay3\NeuroSync.Client.Replay.exe",
    "FILE_DIR": r"D:\edfx\V25OB3000test20250924191457\V25OB3000test20250924191457",
    
    # 通道选择配置
    "CHANNEL_CONFIG": {
        "target_channels": [2, 3],
        "btn_cl_auto_id": "btn_cl",
        "btn_confirm_auto_id": "btn_confirm"
    },

    # 下拉框配置（优先用auto_id）
    "DROPDOWNS": {
        "DAO_LIAN": {
            "auto_id": "cb_daolian",
            "use_index": False,
            "fixed_index": 0,
            "option_name": "导联"
        },
        "SWEEP_SPEED": {
            "auto_id": "cb_zouzhisudu",
            "use_index": False,
            "range": range(0, 22),
            "is_random": True,
            "option_name": "走纸速度"
        },
        "SENSITIVITY": {
            "auto_id": "cb_lingmindu",
            "use_index": False,
            "range": range(0, 16),
            "is_random": True,
            "option_name": "灵敏度"
        },
        "HIGH_PASS": {
            "auto_id": "cb_lvboxiaxian",
            "use_index": False,
            "range": range(0, 11),
            "is_random": True,
            "option_name": "高通滤波"
        },
        "LOW_PASS": {
            "auto_id": "cb_lvboshangxian",
            "use_index": False,
            "range": range(0, 6),
            "is_random": True,
            "option_name": "低通滤波"
        },
        "PLAYBACK_SPEED": {
            "auto_id": "cb_bofangbeishu",
            "use_index": False,
            "range": range(0, 8),
            "is_random": True,
            "option_name": "播放倍速"
        }
    },

    # 其他配置
    "TAG_LIST": {
        'Eyes Open', 'Eyes Closed', 'Seizure', 'Deep Breath', 'Background', 'Awake',
        'Eyes Closed PPR', 'Eyes Shut PPR', 'Eyes Open PCR', 'Electrical Seizure', 'End',
        'Identify Event', 'Seizure*', 'Drug-induced Sleep', 'Mechanical Ventilation', 'Facial Twitching'
    },
    "TAG_CONFIG": {
        "count_range": (2, 5),
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
        "channel_selection": {"success": 0, "fail": 0},
        "move_video_window": {"success": 0, "fail": 0}
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

# 新增全局变量：记录播放倍速选择时的索引（初始值-1）
PLAYBACK_SPEED_INDEX = -1
def select_dropdown_option(main_window, config):
    """下拉框选择函数（优先auto_id）"""
    global PLAYBACK_SPEED_INDEX  # 引用全局变量
    try:
        option_name = config["option_name"]
        # 定位下拉框
        if not config["use_index"] and config["auto_id"]:
            dropdown = main_window.child_window(
                auto_id=config["auto_id"],
                control_type="ComboBox"
            )
        else:
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

        # 获取选项列表
        items = dropdown.descendants(control_type="ListItem")
        if not items:
            raise Exception(f"未找到任何{option_name}选项")
        max_valid_index = len(items) - 1

        # 确定目标索引
        if config.get("is_random", False):
            valid_indices = [idx for idx in config["range"] if 0 <= idx <= max_valid_index]
            if not valid_indices:
                raise Exception(f"{option_name}无有效索引（最大{max_valid_index}）")
            target_index = random.choice(valid_indices)
        else:
            target_index = config["fixed_index"]
            if not (0 <= target_index <= max_valid_index):
                raise Exception(f"{option_name}索引{target_index}超出范围（最大{max_valid_index}）")

        # 检查当前选中项
        current_selected = next((item for item in items if item.is_selected()), None)
        current_index = items.index(current_selected) if current_selected else 0

        if current_index == target_index:
            print(f"当前已选中{option_name}第{target_index+1}项，无需切换")
            dropdown.type_keys("{ESC}")
            # ===== 在这里添加（未切换选项时也需要打印）=====
            if config["option_name"] == "播放倍速":
                PLAYBACK_SPEED_INDEX = target_index
                print(f"选中的播放倍速选项：（索引{target_index+1}）")
            return

        # 切换选项
        dropdown.type_keys("{HOME}")
        time.sleep(0.3)
        dropdown.type_keys(f"{{DOWN {target_index}}}")
        time.sleep(0.3)
        target_item = items[target_index]
        safe_set_focus(target_item)
        target_item.click_input()
        dropdown.type_keys("{ESC}")
        print(f"{option_name}已选择第{target_index+1}项")
        # ===== 在这里添加（未切换选项时也需要打印）=====
        if config["option_name"] == "播放倍速":
            # selected_text = items[target_index].window_text()
            PLAYBACK_SPEED_INDEX = target_index  # 保存当前选择的索引
            print(f"选中的播放倍速选项：（索引{target_index+1}）")

    except Exception as e:
        raise Exception(f"{option_name}选择失败：{str(e)}")

def execute_dropdown_config(main_window):
    """执行下拉框配置"""
    if not RUN_CONFIG["dropdown_config"]:
        print("\n【下拉框配置】已关闭，跳过该模块")
        return True
    try:
        print("\n===== 开始下拉框参数配置 =====")
        for dropdown_key in CONFIG["DROPDOWNS"]:
            select_dropdown_option(main_window, CONFIG["DROPDOWNS"][dropdown_key])
        print("===== 下拉框参数配置完成 =====")
        STATS["module_stats"]["dropdown_config"]["success"] += 1
        return True
    except Exception as e:
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 下拉框配置失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["dropdown_config"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

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
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 标签标记失败：{str(e)}"
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
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 导航按钮操作失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["nav_buttons"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

def select_specific_channels(main_window):
    target_channels = CONFIG["CHANNEL_CONFIG"]["target_channels"]
    for num in target_channels:
        success = False
        for retry in range(3):
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
        btn_cl = main_window.child_window(
            auto_id=CONFIG["CHANNEL_CONFIG"]["btn_cl_auto_id"],
            control_type="Button",
            found_index=0
        )
        btn_cl.wait("enabled", timeout=UI_TIMIEOUT)
        btn_cl.click_input()
        time.sleep(3)
        print("已点击第一个通道列表展开按钮")

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
            found_index=0
        )
        btn_confirm.wait("enabled", timeout=UI_TIMIEOUT)
        btn_confirm.click_input()
        time.sleep(3)
        print(f"已确认选中指定通道：{CONFIG['CHANNEL_CONFIG']['target_channels']}")
        print("===== 通道选择完成 =====")
        STATS["module_stats"]["channel_selection"]["success"] += 1
        return True
    except Exception as e:
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 通道选择失败：{str(e)}"
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
        
        # 校验进度条控件
        try:
            progress_bar = main_window.child_window(
                auto_id=progress_config["auto_id"],
                control_type="Slider"
            )
            progress_bar.wait("visible", timeout=UI_TIMIEOUT * 2)
            progress_bar.wait("enabled", timeout=UI_TIMIEOUT * 2)
        except Exception as e:
            raise Exception(f"进度条控件不存在或未就绪：{str(e)}")

        # 校验滑块控件
        try:
            thumb = progress_bar.child_window(control_type="Thumb")
            thumb.wait("visible", timeout=UI_TIMIEOUT)
            thumb.wait("enabled", timeout=UI_TIMIEOUT)
        except Exception as e:
            raise Exception(f"进度条滑块（Thumb）不存在：{str(e)}")

        # 获取坐标信息
        progress_rect = progress_bar.rectangle()
        if not progress_rect:
            raise Exception("无法获取进度条坐标信息（rectangle为空）")
        
        thumb_rect = thumb.rectangle()
        if not thumb_rect:
            raise Exception("无法获取滑块坐标信息（rectangle为空）")

        # 计算有效拖拽长度
        valid_length = progress_rect.width() - thumb_rect.width()
        if valid_length <= 0:
            raise Exception(f"进度条有效长度异常（{valid_length}），无法拖拽")

        current_percent = 0
        target_x_prev = None

        # 循环拖拽
        for i in range(progress_config["drag_cycles"]):
            # 计算目标百分比
            if i == 0:
                target_percent = random.randint(1, 30)
            else:
                if i % 3 == 0:
                    target_percent = int(current_percent * 0.5)
                else:
                    target_percent = int(current_percent * 1.75)
                target_percent = max(progress_config["min_percent"], 
                                    min(target_percent, progress_config["max_percent"]))

            # 计算目标X坐标
            target_x = progress_rect.left + int(valid_length * (target_percent / 100))
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

            # 模拟拖拽
            mouse.move(coords=(start_x, start_y))
            time.sleep(0.4)
            mouse.press(button="left", coords=(start_x, start_y))
            time.sleep(0.3)

            # 分步移动
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

            current_percent = target_percent
            target_x_prev = target_x
            time.sleep(1.5)

        print("===== 进度条拖拽完成 =====")
        STATS["module_stats"]["drag_progress_bar"]["success"] += 1
        return True

    except Exception as e:
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 进度条拖拽失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["drag_progress_bar"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

# 系统API：直接移动窗口
def move_window(hwnd, x, y):
    if ctypes.sizeof(ctypes.c_void_p) == 8:
        get_window_long = ctypes.windll.user32.GetWindowLongPtrW
    else:
        get_window_long = ctypes.windll.user32.GetWindowLongW
    style = get_window_long(hwnd, win32defines.GWL_STYLE)
    
    if style & win32defines.WS_THICKFRAME:
        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
        client_rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.GetClientRect(hwnd, ctypes.byref(client_rect))
        border_x = (rect.right - rect.left) - client_rect.right
        border_y = (rect.bottom - rect.top) - client_rect.bottom
        x -= border_x // 2
        y -= border_y // 3
    ctypes.windll.user32.SetWindowPos(
        hwnd, None, x, y, 0, 0,
        win32defines.SWP_NOSIZE | win32defines.SWP_NOZORDER
    )

def move_and_close_video_window(main_window, window_title='Video Playback', close_after_move=False):
    if "move_video_window" not in STATS["module_stats"]:
        STATS["module_stats"]["move_video_window"] = {"success": 0, "fail": 0}
    
    if not RUN_CONFIG.get("move_video_window", True):
        print("\n【视频窗口移动】已关闭，跳过该模块")
        return True
    try:
        print("\n===== 开始视频窗口移动 =====")
        # 校验窗口是否存在
        if not main_window.child_window(control_type="Window", title=window_title).exists(timeout=5):
            raise Exception(f"视频窗口'{window_title}'未出现，无法移动")
        
        video_window = main_window.child_window(control_type="Window", title=window_title)
        
        # 获取窗口句柄（带重试）
        max_retries = 3
        hwnd = None
        for retry in range(max_retries):
            hwnd = video_window.handle
            if hwnd:
                break
            if retry < max_retries - 1:
                time.sleep(0.5)
        if not hwnd:
            raise Exception("无法获取窗口句柄，无法移动（已重试3次）")

        # 获取窗口信息
        window_rect = video_window.rectangle()
        window_width, window_height = window_rect.width(), window_rect.height()

        # 计算目标位置
        screen_width = GetSystemMetrics(0)
        screen_height = GetSystemMetrics(1)
        target_x = screen_width - window_width - 10
        target_y = int((screen_height - window_height) / 2)
        print(f"目标位置：({target_x}, {target_y})")

        # 移动窗口
        move_window(hwnd, target_x, target_y)
        time.sleep(0.5)

        # 验证移动结果
        final_rect = video_window.rectangle()
        if abs(final_rect.left - target_x) > 10:
            raise Exception(f"窗口移动失败（目标X：{target_x}，实际X：{final_rect.left}）")

        print(f"'{window_title}'窗口已移动到目标位置")
        time.sleep(1.5)

        # 关闭窗口
        if close_after_move:
            try:
                video_window.child_window(control_type="Button", title='关闭').click_input()
            except:
                ctypes.windll.user32.SendMessageW(hwnd, win32defines.WM_CLOSE, 0, 0)

        print("===== 视频窗口操作完成 =====")
        STATS["module_stats"]["move_video_window"]["success"] += 1
        return True

    except Exception as e:
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 视频窗口操作失败：{str(e)}"
        print(f"===== {error_msg} =====")
        STATS["module_stats"]["move_video_window"]["fail"] += 1
        STATS["error_log"].append(error_msg)
        return False

def show_video(main_window):
    """显示视频窗口（带状态判断和重试）"""
    try:
        for retry in range(3):
            try:
                btn_show_video = main_window.child_window(
                    auto_id="cb_shipin",
                    control_type="Button"
                )
                btn_show_video.wait("enabled", timeout=UI_TIMIEOUT)
                
                # 检查按钮状态（选中表示已显示）
                print(f'btn_show_video.get_toggle_state()={btn_show_video.get_toggle_state()}')
                if btn_show_video.get_toggle_state() == 1:
                    print("视频已显示，无需操作")
                    return True
                
                btn_show_video.click_input()
                print("视频显示按钮点击成功")
                return True
            except Exception as e:
                if retry < 2:
                    print(f"视频显示按钮第{retry+1}次点击失败，重试...")
                    time.sleep(0.5)
                else:
                    raise Exception(f"视频显示按钮点击失败：{str(e)}")
    except Exception as e:
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 视频显示操作异常：{str(e)}"
        print(error_msg)
        STATS["error_log"].append(error_msg)
        return False

def get_playback_speed_index():
    """直接返回全局变量中记录的索引（无需再读取控件）"""
    global PLAYBACK_SPEED_INDEX
    if PLAYBACK_SPEED_INDEX != -1:
        print(f"当前播放倍速索引（记录值）：{PLAYBACK_SPEED_INDEX+1}")
        return PLAYBACK_SPEED_INDEX
    else:
        print("警告：尚未选择播放倍速，索引记录为空")
        return -1

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
        
        # 动态等待播放按钮可用
        wait_until_passes(
            timeout=60,
            retry_interval=1,
            func=lambda: main_window.child_window(title_re="Play", control_type="Button").is_enabled()
        )

        # 启动播放
        play_btn = main_window.child_window(title_re="Play", control_type="Button")
        play_btn.click_input()
        time.sleep(1)
        print("播放功能启动成功")

        # 初始化视频窗口
        show_video(main_window)
        time.sleep(1)
        move_and_close_video_window(main_window)

        # 0.2s线切换
        toggle_0_2s_line = main_window.child_window(title_re="0.2s Line", control_type="CheckBox")
        toggle_0_2s_line.wait("enabled", timeout=UI_TIMIEOUT)
        toggle_0_2s_line.click_input()
        time.sleep(2)
        print("0.2s line 点击成功")

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
            execute_tag_marking(main_window)
        ]
        
        # 根据播放倍速控制视频显示
        playback_speed_idx = get_playback_speed_index()
        if playback_speed_idx != -1 and playback_speed_idx <= 4:  # 倍速≤5倍
            if not main_window.child_window(control_type="Window", title="Video Playback").exists(timeout=1):
                show_video(main_window)
                time.sleep(1)
                move_and_close_video_window(main_window)
            else:
                move_and_close_video_window(main_window)
        else:
            print("播放倍速>5倍，不显示视频")

        if all(module_results):
            STATS['success_cycles'] += 1
            print(f"===== 第 {STATS['total_cycles']} 次循环成功 =====")
        else:
            STATS['fail_cycles'] += 1
            print(f"===== 第 {STATS['total_cycles']} 次循环部分模块失败 =====")
        
        # 输出累计统计
        print(f"当前累计：总循环 {STATS['total_cycles']} 次，成功 {STATS['success_cycles']} 次，失败 {STATS['fail_cycles']} 次")
        return True

    except Exception as e:
        error_msg = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 第 {STATS['total_cycles']} 次循环异常：{str(e)}"
        print(error_msg)
        STATS['fail_cycles'] += 1
        STATS['error_log'].append(error_msg)
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
                if remaining_time > 0:
                    time.sleep(remaining_time)
                break

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