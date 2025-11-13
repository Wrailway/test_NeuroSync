import random
import time
import os
import warnings
from pywinauto import Application, mouse
from pywinauto.findwindows import find_window, ElementNotFoundError
from pywinauto.timings import wait_until_passes

# 过滤窗口聚焦警告（避免冗余输出）
warnings.filterwarnings("ignore", category=RuntimeWarning, message="The window has not been focused due to COMError")

UI_TIMIEOUT = 3  # 延长超时时间，提升稳定性
TEST_DURATION = 12 * 3600  # 测试总时长：10小时（单位：秒）
CYCLE_INTERVAL = 10  # 每次循环（核心操作）间隔时间（单位：秒），可调整

# 配置参数（集中管理，便于修改）
CONFIG = {
    "APP_PATH": r"D:\Program Files\NeuroSync\Replay3\NeuroSync.Client.Replay.exe",
    "FILE_DIR": r"D:\edfx\V25OB3000test20250924191457\V25OB3000test20250924191457",
    "DAO_LIAN_INDEX": 0,  # 固定导联索引
    "SWEEP_SPEED_RANGE": range(0, 22),  # 走纸速度范围（循环时随机）
    "SENSITIVITY_RANGE": range(0, 16),  # 灵敏度范围（循环时随机）
    "HIGH_PASS_RANGE": range(0, 11),    # 高通滤波范围（循环时随机）
    "LOW_PASS_RANGE": range(0, 6),      # 低通滤波范围（循环时随机）
    "PLAYBACK_SPEED_RANGE": range(0, 8),# 播放倍速范围（循环时随机）
    "MAX_DOWN_RETRIES": 5,
    "TAG_LIST": {
        'Eyes Open', 'Eyes Closed', 'Seizure', 'Deep Breath', 'Background', 'Awake',
        'Eyes Closed PPR', 'Eyes Shut PPR', 'Eyes Open PCR', 'Electrical Seizure', 'End',
        'Identify Event', 'Seizure*', 'Drug-induced Sleep', 'Mechanical Ventilation', 'Facial Twitching'
    },
    "NAV_BUTTONS": [  # 导航按钮配置（循环时执行）
        {"title_re": "Previous Second", "name": "上一秒", "click_count": 5, "interval": 0.3},
        {"title_re": "Next Second", "name": "下一秒", "click_count": 5, "interval": 0.3},
        {"title_re": "Previous Page", "name": "上一页", "click_count": 3, "interval": 0.3},
        {"title_re": "Next Page", "name": "下一页", "click_count": 3, "interval": 0.3}
    ],
    "TARGET_CHANNELS": [2, 3],  # 固定选中的通道
    "PROGRESS_BAR_AUTO_ID": "slider_play_jd",  # 进度条auto_id（需替换为实际值）
    "DRAG_CYCLES": 10  # 每次循环中拖拽进度条的次数
}

# 统计变量
STATS = {
    "total_cycles": 0,
    "success_cycles": 0,
    "fail_cycles": 0,
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

def select_dropdown_option(
    main_window, 
    combo_box_auto_id, 
    target_range, 
    option_name,
    is_random=True
):
    """下拉框选择函数（支持随机选择或固定索引）"""
    try:
        dropdown = main_window.child_window(auto_id=combo_box_auto_id, control_type="ComboBox")
        dropdown.wait("visible", timeout=UI_TIMIEOUT)
        dropdown.wait("enabled", timeout=UI_TIMIEOUT)
        safe_set_focus(dropdown)
        dropdown.click_input()
        time.sleep(1.5)

        items = dropdown.descendants(control_type="ListItem")
        if not items:
            raise Exception(f"未找到任何{option_name}选项")
        max_valid_index = len(items) - 1

        # 处理目标索引
        if is_random and isinstance(target_range, (range, list, tuple)):
            valid_indices = [idx for idx in target_range if 0 <= idx <= max_valid_index]
            if not valid_indices:
                raise Exception(f"{option_name}范围中无有效索引（最大{max_valid_index}）")
            target_index = random.choice(valid_indices)
        else:
            target_index = target_range if isinstance(target_range, int) else target_range[0]
            if not (0 <= target_index <= max_valid_index):
                raise Exception(f"{option_name}索引{target_index}超出范围（最大{max_valid_index}）")

        # 检查当前选中项
        current_selected = next((item for item in items if item.is_selected()), None)
        current_index = items.index(current_selected) if current_selected else 0

        if current_index == target_index:
            print(f"当前已选中{option_name}第{target_index+1}项，无需切换")
            dropdown.type_keys("{ESC}")
            return

        # 切换选项
        dropdown.type_keys("{UP 20}")
        time.sleep(0.3)
        dropdown.type_keys(f"{{DOWN {target_index}}}")
        time.sleep(0.3)
        target_item = items[target_index]
        safe_set_focus(target_item)
        time.sleep(0.3)
        target_item.click_input()
        dropdown.type_keys("{ESC}")
        print(f"{option_name}已选择第{target_index+1}项")

    except Exception as e:
        raise Exception(f"{option_name}选择失败：{str(e)}")

def find_and_click_tag(main_window, target_title):
    """通用标签查找点击函数"""
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

def click_button_multiple_times(
    main_window, 
    title_re, 
    button_name, 
    click_count=1, 
    interval=0.5, 
    timeout=2
):
    """通用按钮多次点击函数"""
    try:
        btn = main_window.child_window(title_re=title_re, control_type="Button")
        btn.wait("enabled", timeout=timeout)
        safe_set_focus(btn)

        for i in range(click_count):
            btn.click_input()
            print(f"{button_name}第{i+1}/{click_count}次点击完成")
            if i < click_count - 1:
                time.sleep(interval)

        time.sleep(0.5)
        print(f"{button_name}共{click_count}次点击成功")
        return True

    except ElementNotFoundError:
        raise Exception(f"【{button_name}失败】未找到该按钮")
    except TimeoutError:
        raise Exception(f"【{button_name}失败】{timeout}秒内未启用")
    except Exception as e:
        raise Exception(f"【{button_name}失败】{str(e)}")

def select_specific_channels(main_window, target_numbers, max_retries=2):
    """选中指定数字名称的通道（带重试和容错）"""
    for num in target_numbers:
        success = False
        for retry in range(max_retries + 1):
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
                
                current_state = channel.get_toggle_state()
                if current_state == 0:
                    channel.click_input()
                    time.sleep(0.3)
                    if channel.get_toggle_state() == 1:
                        print(f"已选中通道 {num}")
                        success = True
                        break
                    else:
                        raise Exception("点击后状态未改变")
                else:
                    print(f"通道 {num} 已处于选中状态")
                    success = True
                    break
                
            except Exception as e:
                if retry < max_retries:
                    print(f"通道 {num} 第{retry+1}次操作失败（{str(e)}），重试...")
                    time.sleep(0.5)
                else:
                    raise Exception(f"通道 {num} 多次操作失败：{str(e)}")
        
        if not success:
            raise Exception(f"通道 {num} 未成功选中，终止流程")

def drag_progress_in_cycles(main_window, progress_bar_auto_id, cycles=5):
    """分周期拖拽进度条（模拟人工操作）"""
    try:
        progress_bar = main_window.child_window(
            auto_id=progress_bar_auto_id,
            control_type="Slider"  # 或 "ProgressBar"，根据实际类型修改
        )
        progress_bar.wait("visible", timeout=UI_TIMIEOUT * 2)
        progress_bar.wait("enabled", timeout=UI_TIMIEOUT * 2)
        thumb = progress_bar.child_window(control_type="Thumb")
        thumb.wait("visible", timeout=UI_TIMIEOUT)
        thumb.wait("enabled", timeout=UI_TIMIEOUT)
        print(f"\n开始分{cycles}次拖拽进度条")

        progress_rect = progress_bar.rectangle()
        thumb_rect = thumb.rectangle()
        valid_length = progress_rect.width() - thumb_rect.width()
        current_percent = 0
        target_x_prev = None

        for i in range(cycles):
            if i == 0:
                target_percent = random.randint(1, 30)  # 首次0-7%随机
            else:
                if i % 5 ==0:
                    target_percent =current_percent*0.4
                    
                else:
                    target_percent = current_percent*1.5
                    if target_percent > 100:
                        target_percent = 100

            target_x = progress_rect.left + int(valid_length * (target_percent / 100))
            target_x = max(progress_rect.left, min(target_x, progress_rect.right - thumb_rect.width()))
            target_y = progress_rect.top + (progress_rect.height() // 2)

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

            # 分步拖动
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

    except Exception as e:
        raise Exception(f"周期性拖拽失败：{str(e)}")

def init_application():
    """初始化应用：启动+加载文件+固定配置（仅执行一次）"""
    app = None
    try:
        print("===== 开始初始化应用 =====")
        
        # 1. 校验文件
        bdf_files = [f for f in os.listdir(CONFIG['FILE_DIR']) if f.lower().endswith(".bdf")]
        if not bdf_files:
            raise Exception(f"目录 {CONFIG['FILE_DIR']} 中未找到.bdf文件")
        file_path = os.path.join(CONFIG['FILE_DIR'], bdf_files[0])
        print(f"找到目标文件：{file_path}")

        # 2. 启动应用
        app = Application(backend="uia").start(CONFIG['APP_PATH'])
        main_window = app.window(title_re="NeuroSync Replay.*")
        if not main_window.wait("visible", timeout=UI_TIMIEOUT*10):
            raise TimeoutError("主窗口30秒内未可见")
        safe_set_focus(main_window)
        print("应用启动成功")

        # 3. 加载文件
        file_btn = main_window.child_window(title="File", control_type="Button")
        file_btn.wait("enabled", timeout=UI_TIMIEOUT)
        file_btn.click_input()
        print("File按钮点击成功")

        # 4. 处理文件对话框
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
        time.sleep(5)  # 适配大文件加载

        # 5. 启动播放
        play_btn = main_window.child_window(title_re="Play", control_type="Button")
        play_btn.wait("enabled", timeout=UI_TIMIEOUT)
        safe_set_focus(play_btn)
        play_btn.click_input()
        time.sleep(1)
        print("播放功能启动成功")

        # 6. 关闭Video Playback窗口（固定操作，仅执行一次）
        video_window = main_window.child_window(control_type="Window", title='Video Playback')
        safe_set_focus(video_window)
        close_btn = video_window.child_window(control_type="Button", title='关闭')
        close_btn.click_input()
        print("Video Playback窗口已关闭")

        # 7. 固定配置：导联（仅执行一次）
        select_dropdown_option(
            main_window=main_window,
            combo_box_auto_id="cb_daolian",
            target_range=CONFIG['DAO_LIAN_INDEX'],
            option_name="导联",
            is_random=False
        )

        # 8. 启用0.2s线（固定操作）
        toggle_0_2s_line = main_window.child_window(title_re="0.2s Line", control_type="CheckBox")
        toggle_0_2s_line.wait("enabled", timeout=UI_TIMIEOUT)
        safe_set_focus(toggle_0_2s_line)
        toggle_0_2s_line.click_input()
        time.sleep(2)
        print("0.2s line 点击成功")

        # 9. 通道选择（固定操作，仅执行一次）
        btn_cl = main_window.child_window(auto_id="btn_cl", control_type="Button", found_index=0)
        btn_cl.wait("enabled", timeout=UI_TIMIEOUT)
        safe_set_focus(btn_cl)
        btn_cl.click_input()
        time.sleep(3)
        print("展开通道列表")

        cbx_all = main_window.child_window(title='All', control_type="CheckBox")
        cbx_all.wait("enabled", timeout=UI_TIMIEOUT)
        safe_set_focus(cbx_all)
        if cbx_all.get_toggle_state() == 1:
            cbx_all.click_input()
            time.sleep(0.5)
            print("已取消全选通道")
        else:
            print("通道已处于未全选状态")

        select_specific_channels(main_window, CONFIG['TARGET_CHANNELS'])

        btn_confirm = main_window.child_window(auto_id="btn_confirm", control_type="Button")
        btn_confirm.wait("enabled", timeout=UI_TIMIEOUT)
        btn_confirm.click_input()
        time.sleep(3)
        print(f"已确认选中指定通道：{CONFIG['TARGET_CHANNELS']}")

        print("===== 应用初始化完成，开始12小时循环测试 =====")
        return app, main_window

    except Exception as e:
        print(f"应用初始化失败：{str(e)}")
        if app and app.is_process_running():
            app.kill()
        raise

def run_cycle_operations(main_window):
    """循环执行的核心操作（反复执行）"""
    try:
        print(f"\n===== 开始第 {STATS['total_cycles'] + 1} 次循环操作 =====")
        
        # 1. 随机配置下拉框参数（每次循环随机选择）
        select_dropdown_option(main_window, "cb_zouzhisudu", CONFIG['SWEEP_SPEED_RANGE'], "走纸速度")
        select_dropdown_option(main_window, "cb_lingmindu", CONFIG['SENSITIVITY_RANGE'], "灵敏度")
        select_dropdown_option(main_window, "cb_lvboxiaxian", CONFIG['HIGH_PASS_RANGE'], "高通滤波")
        select_dropdown_option(main_window, "cb_lvboshangxian", CONFIG['LOW_PASS_RANGE'], "低通滤波")
        select_dropdown_option(main_window, "cb_bofangbeishu", CONFIG['PLAYBACK_SPEED_RANGE'], "播放倍速")

        # 2. 执行导航按钮点击
        for btn_config in CONFIG['NAV_BUTTONS']:
            click_button_multiple_times(
                main_window=main_window,
                title_re=btn_config["title_re"],
                button_name=btn_config["name"],
                click_count=btn_config["click_count"],
                interval=btn_config["interval"]
            )

        # 3. 拖拽进度条（模拟人工操作）
        drag_progress_in_cycles(
            main_window=main_window,
            progress_bar_auto_id=CONFIG['PROGRESS_BAR_AUTO_ID'],
            cycles=CONFIG['DRAG_CYCLES']
        )

        # 4. 随机打标签（修改后：一次循环打2-3个标签，可调整数量）
        tag_count = random.randint(2, 10)  # 每次循环随机打2-3个标签，可改为固定值（如3）
        print(f"\n本次循环计划标记 {tag_count} 个标签")

        for tag_idx in range(1, tag_count + 1):
            target_tag = random.choice(list(CONFIG['TAG_LIST']))
            print(f"\n第{tag_idx}/{tag_count}个标签：尝试标记「{target_tag}」")
            
            if find_and_click_tag(main_window, target_tag):
                print(f"成功标记标签「{target_tag}」")
                time.sleep(0.8)  # 标记后停顿，模拟人工操作间隔
                continue

            # 未找到则翻页查找
            print(f"未找到标签「{target_tag}」，尝试翻页查找")
            down_button = main_window.child_window(auto_id="DownButton", control_type="Button")
            down_button.wait("visible", timeout=UI_TIMIEOUT)
            found = False
            
            for down_count in range(1, CONFIG['MAX_DOWN_RETRIES'] + 1):
                down_button.click_input()
                time.sleep(0.4)
                if find_and_click_tag(main_window, target_tag):
                    print(f"第{down_count}次翻页后，成功标记标签「{target_tag}」")
                    found = True
                    time.sleep(5)  # 标记后停顿
                    break
    
            if not found:
                print(f"翻页{CONFIG['MAX_DOWN_RETRIES']}次仍未找到标签「{target_tag}」，跳过该标签")
                # 可选：翻页后回到初始位置，避免后续标签查找位置偏差
                down_button.type_keys(f"{{UP {down_count}}}")
                time.sleep(0.3)

        # 5. 循环状态更新
        STATS['total_cycles'] += 1
        STATS['success_cycles'] += 1
        print(f"===== 第 {STATS['total_cycles']} 次循环操作执行成功 =====")
        return True

    except Exception as e:
        error_msg = f"第 {STATS['total_cycles'] + 1} 次循环失败：{str(e)}"
        print(error_msg)
        STATS['total_cycles'] += 1
        STATS['fail_cycles'] += 1
        STATS['error_log'].append(error_msg)
        return False

def main():
    """主函数：初始化应用+循环执行+时间控制"""
    app = None
    main_window = None
    start_time = time.time()

    try:
        # 1. 初始化应用（仅执行一次）
        app, main_window = init_application()

        # 2. 循环执行操作，直到达到测试时长
        while (time.time() - start_time) < TEST_DURATION:
            # 执行一次循环操作
            run_cycle_operations(main_window)

            # 计算剩余时间，避免超出总时长
            elapsed_time = time.time() - start_time
            remaining_time = TEST_DURATION - elapsed_time
            print(f"\n当前已运行：{elapsed_time / 3600:.2f} 小时，剩余：{remaining_time / 3600:.2f} 小时")

            # 循环间隔（可调整，避免操作过于频繁）
            if remaining_time > CYCLE_INTERVAL:
                time.sleep(CYCLE_INTERVAL)
            else:
                # 剩余时间不足间隔，等待到总时长后退出
                time.sleep(remaining_time)
                break

        # 3. 测试结束，输出统计结果
        print(f"\n===== 12小时测试完成！=====")
        print(f"总循环次数：{STATS['total_cycles']}")
        print(f"成功次数：{STATS['success_cycles']}")
        print(f"失败次数：{STATS['fail_cycles']}")
        print(f"成功率：{STATS['success_cycles'] / STATS['total_cycles'] * 100:.2f}%" if STATS['total_cycles'] > 0 else "无循环执行")

        if STATS['error_log']:
            print(f"\n错误日志（共{len(STATS['error_log'])}条）：")
            for idx, error in enumerate(STATS['error_log'], 1):
                print(f"{idx}. {error}")

    except Exception as e:
        print(f"\n测试异常终止：{str(e)}")
    finally:
        # 4. 退出应用（仅执行一次）
        print(f"\n===== 开始退出应用 =====")
        if main_window and main_window.exists():
            main_window.close()
            time.sleep(2)
        if app and app.is_process_running():
            app.kill()
        print("应用已退出，测试结束")

if __name__ == "__main__":
    main()