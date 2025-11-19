import datetime
import random
from pywinauto import Application
from time import sleep

APP_PATH = r"D:\Program Files\NeuroSync\Recorder3\NeuroSync.Client.Recorder.exe"
COLLECTION_DURATION = 60 # 采集时长
SAMPLING_RATE = 500 # 采样率
LEAD_SOURCE_INDEX = 3 # 导联源索引
DAO_LIAN_INDEX = 3  # 固定导联索引
CLICK_INTERVAL = 3 # 标记间隔时间
SWEEP_SPEED_RANGE = range(0, 8) # 走纸速度范围
SENSITIVITY_RANGE = range(0, 16) # 灵敏度范围
HIGH_PASS_RANGE = range(0, 11) # 高通滤波范围
LOW_PASS_RANGE = range(0, 6) # 低通滤波范围
PATIENTS_NAME = "This is about NeuroSync test"
MARK_BUTTON_TITLES = [
    "过度换气","直立伸臂","图形诱发","发作*","药物睡眠","机械通气","面部抽动","四肢抽动","咳嗽","寒颤",
    "睁眼","test","闭眼","背景","发作","深呼吸","清醒","REM","睁眼PPR","闭眼PPR","合眼PPR","电发作",
    "QS","指认事件","AS"
]
TARGET_MARKS = ["发作*", "药物睡眠", "图形诱发","直立伸臂","过度换气"]

COMBOBOX_ITEM = [
    {"desc": "导联切换", "found_index": 0, "value": DAO_LIAN_INDEX},
    {"desc": "走纸速度", "found_index": 1, "value": lambda: random.choice(SWEEP_SPEED_RANGE)},
    {"desc": "灵敏度", "found_index": 2, "value": lambda: random.choice(SENSITIVITY_RANGE)},
    {"desc": "高通滤波", "found_index": 3, "value": lambda: random.choice(HIGH_PASS_RANGE)},
    {"desc": "低通滤波", "found_index": 4, "value": lambda: random.choice(LOW_PASS_RANGE)}
]

def select_combobox_item(main_window, found_index, value, desc=""):
    combo = main_window.child_window(control_type = "ComboBox", found_index = found_index)
    combo.wait('visible', timeout=5000)
    combo.wait('enabled', timeout=5000)
    combo.set_focus()
    combo.click_input()
    sleep(0.2)
    # 如果value是函数，调用它获取实际值
    v = value() if callable(value) else value
    combo.select(v)
    print(f"{desc or 'ComboBox'} 选择项: {v}")
    sleep(0.2)
def check_impedance(main_window):
    try:
        impedance_button = main_window.child_window(title="阻抗", control_type="Button")
        impedance_button.wait('visible', timeout=3000)
        impedance_button.wait('enabled', timeout=3000)
        impedance_button.click_input()
        print("打开阻抗检测")
        sleep(3)
        impedance_button.click_input()
        print("关闭阻抗检测")
    except Exception as e:
        print(f"阻抗检测操作失败：{str(e)}")
# 点击相同按钮的函数
def click_all_good_coop_radios(main_window):
    target_title = "合作良好"
    all_good_coop_radios = [
        btn for btn in main_window.descendants(
            title=target_title,
            control_type="RadioButton"
        ) if btn.is_visible() and btn.is_enabled()
    ]
    success_count = 0
    for idx, radio in enumerate(all_good_coop_radios, 1):
        try:
            # radio.set_focus()
            # sleep(0.3)
            radio.click()
            success_count += 1
        except Exception as e:
            print(f"第{idx}个'{target_title}'RadioButton点击失败:{str(e)}")
# 翻页查找标记函数
def find_mark_accross_pages(main_window, mark_title):
    mark_btn = None
    is_backword = False
    try:
        next_page_btn = main_window.child_window(
            auto_id="DownButton",
            control_type="Button",
            found_index=1 #id相同的按钮，选择第二个
        )
        prev_page_btn = main_window.child_window(
            auto_id="UpButton",
            control_type="Button",
            found_index=1
        )
    except Exception as e:
        print(f"分页按钮定位失败：{str(e)}")
        return None
    # 循环查找标记
    while True:
        try:
            # 查找当前页的目标按钮
            mark_btn = next(
                btn for btn in main_window.descendants(
                    title=mark_title,
                    control_type="Button"
                )
                if btn.is_visible() and btn.is_enabled()
            )
            print(f"找到标记：'{mark_title}'")
            return mark_btn
        except StopIteration:
            # 当前页未找到，翻页继续查找
            if not is_backword:
                # 向下一页查找
                if next_page_btn.is_enabled():
                    next_page_btn.click()
                else:
                    # 到达最后一页，开始向前翻页
                    is_backword = True
            else:
                # 向上一页查找
                if prev_page_btn.is_enabled():
                    prev_page_btn.click()
                else:
                    # 回到第一页，结束查找
                    print(f"未找到标记：'{mark_title}'")
                    break
    return None

def test_open_neurosync_app():
    try:
        # 启动应用
        app = Application(backend="uia").start(APP_PATH)
        sleep(5)

        # 获取主窗口
        main_window = app.window(title_re="NeuroSync.*")
        main_window.wait('visible', timeout=10000)
        
        # 点击查找设备
        find_device_btn = main_window.child_window(
            title="NeuroSync.Client.Recorder.Pages.Windows.MainWindowViewModel",
            control_type="Button"
        )
        find_device_btn.wait('visible', timeout=5000)
        find_device_btn.click()
        
        # 点击开始扫描
        scan_device_btn = main_window.child_window(title="开始扫描",control_type="Button")
        scan_device_btn.wait('visible', timeout=5000)
        scan_device_btn.click()
        sleep(8)
        
        # 点击连接按钮
        connect_buttons = main_window.descendants(title="连接", control_type="Button")
        connect_buttons[0].click_input()
        sleep(3)
        
        # 点击确认（确认连接）
        confirm_connect_btn = main_window.child_window(auto_id="Confirm",control_type="Button")
        confirm_connect_btn.wait('visible', timeout=5000)
        confirm_connect_btn.click_input()
        
        # 动态选择采样率
        sampling_rate_auto_id = f"Radio{SAMPLING_RATE}"
        try:
            choose_sampling_rate = main_window.child_window(
                auto_id=sampling_rate_auto_id,
                control_type="RadioButton"
            )
            choose_sampling_rate.wait('visible', timeout=5000)
            choose_sampling_rate.click_input()
            print(f"采样率设置为 {SAMPLING_RATE}Hz")
        except Exception as e:
            print(f"采样率设置失败：{str(e)}")
            return

        # 点击确认按钮（确认采样率）
        confirm_sampling_rate_btn = main_window.child_window(title="确定", control_type="Button")
        confirm_sampling_rate_btn.wait('visible', timeout=5000)
        confirm_sampling_rate_btn.click_input()
        
        # 点击确定（确定设置导联源）
        confirm_sampling_rate_btn = main_window.child_window(title="确定", control_type="Button")
        confirm_sampling_rate_btn.wait('visible', timeout=5000)
        confirm_sampling_rate_btn.click_input()
        
        # 选择导联源
        lead_source_btn = main_window.child_window(control_type="ComboBox", found_index=0)
        lead_source_btn.wait('visible', timeout=5000)
        lead_source_btn.wait('enabled', timeout=5000)
        lead_source_btn.set_focus()
        lead_source_btn.click_input()
        sleep(0.3)
        lead_source_btn.select(LEAD_SOURCE_INDEX)
        
        # 点击应用
        apply_btn = main_window.child_window(title="应用", control_type="Button")
        apply_btn.wait('visible', timeout=5000)
        if apply_btn.is_enabled():
            apply_btn.click_input()
        
        # 退出设置页面
        # 坐标信息定位退出设置按钮
        TARGET_LEFT = 1549
        TARGET_TOP = 79
        exit_setting_btn = None
        # 获取所有按钮，遍历查找目标按钮
        all_buttons = main_window.descendants(control_type="Button")
        for idx,btn in enumerate(all_buttons):
            try:
                rect = btn.rectangle()
                if rect.left == TARGET_LEFT and rect.top == TARGET_TOP:
                    exit_setting_btn = btn
                    print(f"找到目标btn:索引{idx}, left={rect.left}, top={rect.top}")
                    break
            except Exception as e:
                print(f"读取按钮索引{idx}信息失败：{str(e)}")
                return
        # 点击退出设置按钮
        if exit_setting_btn:
            try:
                if exit_setting_btn.is_enabled() and exit_setting_btn.is_visible():
                    exit_setting_btn.click_input()
                print("已退出设置页面！")
            except Exception as e:
                print(f"点击退出设置按钮失败：{str(e)}")
                
        # 导联切换、走纸速度、灵敏度、高通滤波、低通滤波设置
        for item in COMBOBOX_ITEM:
            select_combobox_item(
                main_window,
                found_index=item["found_index"],
                value=item["value"],
                desc=item["desc"]
            )
            
        # 阻抗检测
        check_impedance(main_window)
        
        # 试验标记列表
        marker_list = main_window.child_window(title="试验标记列表", control_type="Button")
        marker_list.wait('visible', timeout=3000)
        marker_list.wait('enabled', timeout=3000)
        marker_list.click_input()
        
        # 点击开始(开始采集数据)
        start_btn = main_window.child_window(title="开始",control_type="Button")
        start_btn.wait('visible', timeout=5000)
        start_btn.click_input()
        
        # 填写姓名
        name_fields = main_window.descendants(class_name="TextBox",control_type="Edit")
        name_fields[0].set_focus()
        name_fields[0].set_text(PATIENTS_NAME)
        
        # 选择出生日期
        birthday_fields = main_window.child_window(auto_id="targetElement",control_type="Button")
        birthday_fields.click()
        today = datetime.date.today()
        today_str = today.strftime("%Y年%m月%d日").replace("01日", "1日"
                                                ).replace("02日", "2日"
                                                ).replace("03日", "3日"
                                                ).replace("04日", "4日"
                                                ).replace("05日", "5日"
                                                ).replace("06日", "6日"
                                                ).replace("07日", "7日"
                                                ).replace("08日", "8日"
                                                ).replace("09日", "9日")
        choose_date_btn = main_window.child_window(title=today_str,control_type="Button")
        choose_date_btn.wait('visible', timeout=3000)
        choose_date_btn.click_input()
        
        # 点击脑电采集
        eeg_collect_btn = main_window.child_window(title="脑电采集",control_type="Button")
        eeg_collect_btn.wait('visible', timeout=5000)
        eeg_collect_btn.click()
        
        # 选择患者状态
        status_checkbox = main_window.child_window(title="清醒",control_type="CheckBox")
        status_checkbox.wait('visible', timeout=5000)
        status_checkbox.click_input()
        
        # 点击确定
        confirm_collect_btn = main_window.child_window(title="确定",control_type="Button")
        confirm_collect_btn.wait('visible', timeout=5000)
        confirm_collect_btn.click()
        
        mark_selected = {mark: False for mark in TARGET_MARKS}
        start_time = datetime.datetime.now()  # 记录开始时间
        while True:
            # 计算已消耗时间和剩余时间
            elapsed = (datetime.datetime.now() - start_time).total_seconds()
            remaining = COLLECTION_DURATION - elapsed
            if remaining <= 0:
                break  # 超时直接退出，不继续循环
            random_mark_title = random.choice(MARK_BUTTON_TITLES)
            try:
                mark_btn = next(
                    btn for btn in main_window.descendants(
                        title=random_mark_title,
                        control_type="Button"
                    )
                    if btn.is_visible() and btn.is_enabled()
                )
            except StopIteration:
                mark_btn = find_mark_accross_pages(main_window, random_mark_title)
                if not mark_btn:
                    print(f"跳过未找到的标记：'{random_mark_title}'")
                    sleep(min(remaining, CLICK_INTERVAL))
                    continue

            mark_btn.click()
            # 更新目标标记状态
            if random_mark_title in mark_selected and not mark_selected[random_mark_title]:
                mark_selected[random_mark_title] = True

            # 等待下一个点击周期
            sleep(min(remaining, CLICK_INTERVAL))
        
        # 点击结束
        end_btn = main_window.child_window(title="结束",control_type="Button")
        end_btn.wait('visible', timeout=3000)
        end_btn.click()
        collection_minutes = COLLECTION_DURATION // 60 # 采集时长（分钟）
        print(f"已采集时长{collection_minutes}分钟，结束采集")
        
        # 点击确定结束
        confirm_end_btn = main_window.child_window(title="确定",control_type="Button")
        confirm_end_btn.wait('visible', timeout=3000)
        confirm_end_btn.click()
        
        # 填写患者检查情况
        if mark_selected["药物睡眠"]:
            try:
                medication_sleep_fields = main_window.descendants(
                    class_name="TextBox",
                    control_type="Edit"
                )
                medication_sleep_fields[1].set_focus()
                medication_sleep_fields[1].set_text("123")
            except Exception as e:
                print(f"药物睡眠操作失败：{str(e)}")
        if mark_selected["直立伸臂"] or mark_selected["图形诱发"] or mark_selected["过度换气"]:
            click_all_good_coop_radios(main_window)
        if mark_selected["发作*"]:
            try:
                seizure_btn = main_window.child_window(
                    title="意识丧失",
                    control_type="RadioButton"
                )
                seizure_btn.click()
            except Exception as e:
                print(f"发作*操作失败：{str(e)}")
        
        # 点击保存
        save_btn = main_window.child_window(title="保存",control_type="Button")
        save_btn.wait('visible', timeout=3000)
        save_btn.click()
        
    except Exception as e:
        print(f"测试失败：{str(e)}")
    finally:
        sleep(1)
        # 关闭应用
        # if 'app' in locals():
        #     app.kill()
        

if __name__ == "__main__":
    test_open_neurosync_app()