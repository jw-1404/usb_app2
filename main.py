# main.py
import sys
import os
import re
import struct
import time
#import resources_rc
import threading
import queue
from collections import deque
from datetime import datetime
import pandas as pd
import openpyxl
import et_xmlfile
from dataclasses import dataclass
import numpy as np
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QFileDialog, QLabel
from PyQt5.QtCore import Qt, pyqtSignal, QThread, QObject, QTimer
from PyQt5.QtGui import QTextCursor, QTextCharFormat, QColor
from PyQt5.QtWidgets import QInputDialog, QProgressDialog
import usb.core
import usb.backend.libusb1
# 在 main.py 的导入部分添加这一行
from typing import List
# 在 main.py 开头的导入语句中添加 Optional
from typing import List, Optional
# 从其他模块导入
#from ui_usb import Ui_MainWindow
from threads import UsbReaderThread, FileWriterThread, DrainDataThread
#from utils import Entry, encode_id, decode_id, to_binary_str
from analysis import AnalysisThread
from baseline import DatParser
# main.py 顶部（import 区域）
from baseline import run_analysis   
from ui_usb import Ui_MainWindow
# 在文件开头的导入语句中添加 Dict（与已有的类型导入放在一起）
from typing import List, Optional, Dict  # 新增 Dict
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QFileDialog, QLabel, QLineEdit
# 全局后端
backend = usb.backend.libusb1.get_backend(find_library=lambda x: "libusb-1.0.dll")
# 常量
VENDOR_ID = 0x04b4
PRODUCT_ID = 0x00f1
EP_OUT = 0x01
EP_IN = 0x81


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.file_interval_minutes = 10
        self.received_count = 0
        self.log_entry_count = 0
        # === 新增：采集控制变量 ===
        self.collection_mode = "baseline"
        self.collection_duration = 0  # 秒
        self.start_time = 0
        self.is_running = False
        self.need_usb_reset = False

        # === 新增：定时器 ===
        self.trigger_timer = QTimer()
        self.trigger_timer.timeout.connect(self.on_timer_tick)

        # === 新增：USB 端点（确保已初始化）===
        self.ep_out = None  # 必须在 find_device 后赋值
        #self.run_counter = 0
        # ---------- 新增：run 前缀 & 计数器 ----------
        self.run_prefix = "run"          # 用户可在 lineEdit 中修改
        self.run_counter = 0             # 每次 Start 从 1 开始
        self.data_base_dir = os.getcwd()

        # === 添加：实时时间显示
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)  # 每1000ms（1秒）触发一次
        self.update_time()      # 立即更新一次
        self.current_data_dir = ""  # 当前采集周期的文件夹路径
        """"""
        # 初始化变量
        self.file_counter = 1 
        self.log_builder = []
        self.selected_device = None
        self.boards = 0x00
        self.commands: Dict[str, List[int]] = {}
        self.is_connected = 0
        self.recv = bytearray(16384)
        self.last_readed = 0
        self.read_thread: Optional[UsbReaderThread] = None
        self.writer_thread: Optional[FileWriterThread] = None
        self.drain_thread: Optional[DrainDataThread] = None
        self.board_ids: List[int] = []
        self.chip_ids: List[int] = []
        self.adc_ths: List[int] = []
        self.ep_out = None
        self.ep_in = None
        self.threshold_commands: List[Dict[str, List[int]]] = []
        self.is_stopping = False
        self.init_chips_per_board_combobox()
        self.chips_per_board = 0
        self.channels_per_chip = 128
        self.collection_mode = "baseline"
        # 初始化触发层数下拉框
        self.ui.Combo_trigger_value.addItems(["0 (0层)","1 (1层)", "2 (2层)", "3 (3层)", "4 (4层)", "5 (5层)", "6 (6层)"])
        self.ui.Combo_trigger_value.setCurrentIndex(0)
        self.init_hardcoded_commands()
        self.refresh_device_lists()

        # 连接信号和槽
        self.ui.combox_device_lists.currentIndexChanged.connect(self.combbox_device_lists_index_changed)
        self.ui.button_update.clicked.connect(self.on_button_update_clicked)
        self.ui.button_connect.clicked.connect(self.on_button_connect_clicked)
        self.ui.button_send.clicked.connect(self.on_button_send_clicked)
        self.ui.button_command_send.clicked.connect(self.on_button_on_command_send_clicked)
        #self.ui.button_test.clicked.connect(self.action_button_test_click)
        self.ui.button_config.clicked.connect(self.on_button_config_clicked)
        self.ui.Combo_trigger_value.currentIndexChanged.connect(self.on_trigger_value_changed)
        self.ui.comboBox123.currentIndexChanged.connect(self.on_chips_per_board_changed)
        self.ui.button_analyze.clicked.connect(self.on_analyze_dat_clicked)
        self.ui.button_send_custom_th.clicked.connect(self.on_send_custom_th_value)
        self.ui.lineEdit_run_prefix_2.textChanged.connect(self.on_run_prefix_changed)
        self.ui.pushButton_select.clicked.connect(self.select_data_dir)
        self.ui.lineEdit_run_prefix_2.textChanged.connect(self.update_run_prefix)
        self.ui.pushButton_analyze.clicked.connect(self.on_analyze_baseline)
        self.ui.pushButton_reset_usb.clicked.connect(self.on_reset_usb_clicked)
        self.ui.pushButton_send_regs.clicked.connect(self.on_send_custom_regs_clicked)


    def init_chips_per_board_combobox(self):
       
        
        self.ui.comboBox123.clear()
        
        chip_options = [
            ("6 芯片/板卡", 6),
            ("16 芯片/板卡", 16),
            ("24 芯片/板卡", 24),
        
        ]
        
        for text, value in chip_options:
            self.ui.comboBox123.addItem(text, value)
        default_index = 2  # 24芯片/板卡
        self.ui.comboBox123.setCurrentIndex(default_index)
        self.chips_per_board = self.ui.comboBox123.currentData()
        
        self.log(f"芯片数量设置为: {self.chips_per_board} 芯片/板卡", "info")

    def on_chips_per_board_changed(self, index):
        """处理芯片数量变更"""
        if index < 0:
            return
            
        new_chips_per_board = self.ui.comboBox123.currentData()
    
        if new_chips_per_board != self.chips_per_board:
            old_value = self.chips_per_board
            self.chips_per_board = new_chips_per_board
            
            self.log(f"芯片数量已从 {old_value} 更改为 {self.chips_per_board} 芯片/板卡", "info")
            
            self.clear_all_thr_data()
            self.init_hardcoded_commands()
            
            if self.adc_ths:
                self.log("警告: 芯片数量已更改，建议重新加载阈值文件", "warning")

    def check_file_format(self, file_path: str) -> bool:
        """检查文件是否包含预期的起始标记"""
        try:
            with open(file_path, 'rb') as f:
                # 读取文件前部分内容进行检查
                data = f.read(4096)  # 读取前4KB进行检查
                if len(data) < 4:
                    return False
                
                # 检查是否包含预期的起始标记
                possible_starts = [0xffab530b, 0xffab530f, 0x0fac0f53]  # 添加更多可能的起始标记
                found_markers = []
                
                for i in range(len(data) - 4):
                    value = struct.unpack_from('>I', data, i)[0]  # 大端序
                    if value in possible_starts:
                        found_markers.append((i, hex(value)))
                
                print(f"在文件中找到的标记: {found_markers}")
                return len(found_markers) > 0
        except Exception as e:
            print(f"文件格式检查失败: {e}")
            return False

    def update_layer_status(self, boards: int):
        layers = [(0x01, self.ui.label_1, "第1层"), (0x02, self.ui.label_2, "第2层"),
                  (0x04, self.ui.label_3, "第3层"), (0x08, self.ui.label_4, "第4层"),
                  (0x10, self.ui.label_5, "第5层"), (0x20, self.ui.label_6, "第6层")]
        online_count = 0
        online_boards = []
        for board_id, label, layer_name in layers:
            if boards & board_id:
                online_count += 1
                online_boards.append(board_id)
                if label:
                    label.setStyleSheet("background-color: green; color: white; font-weight: bold;")
            elif label:
                label.setStyleSheet("background-color: red; color: white; font-weight: bold;")
        self.log(f"在线层数: {online_count}，在线板卡: {', '.join(f'B-0b{board_id:08b}' for board_id in online_boards)}", "info")
        return online_boards
    

    def on_analyze_dat_clicked(self):
        dat_path, _ = QFileDialog.getOpenFileName(self, "选择.dat文件", "", "DAT Files (*.dat)")
        if not dat_path:
            return
            
        # 检查文件格式
        if not self.check_file_format(dat_path):
            self.log("警告: 文件格式检查未找到预期标记，将尝试强制解析", "warning")
            # 不直接返回，而是尝试解析
        
        output_path, _ = QFileDialog.getSaveFileName(self, "保存阈值文件", "", "Excel Files (*.xlsx)")
        if not output_path:
            return
            
        # 获取在线层信息
        online_boards = self.update_layer_status(self.boards)
        if not online_boards:
            self.log("没有检测到在线层，无法设置分层Sigma值", "error")
            self.show_warning_dialog("错误", "没有检测到在线层，请先执行Check命令")
            return
        
        # 为每一层设置Sigma值
        layer_sigma_map = {}
        layer_names = {0x01: "第1层", 0x02: "第2层", 0x04: "第3层", 
                    0x08: "第4层", 0x10: "第5层", 0x20: "第6层"}
        
        for board_id in online_boards:
            layer_name = layer_names.get(board_id, f"板卡0x{board_id:02x}")
            sigma, ok = QInputDialog.getDouble(
                self, 
                f"输入{layer_name}的Sigma系数", 
                f"{layer_name} Sigma系数:", 
                3.0, 0.1, 10.0, 1
            )
            if not ok:
                # 如果用户取消，使用默认值3.0
                sigma = 3.0
                self.log(f"{layer_name} 使用默认Sigma系数: 3.0", "info")
            layer_sigma_map[board_id] = sigma
            self.log(f"{layer_name} Sigma系数设置为: {sigma}", "info")
        chips_per = self.ui.comboBox123.currentData()
        progress = QProgressDialog("分析中...", "取消", 0, 100, self)
        progress.setWindowModality(Qt.WindowModal)
        # 修改AnalysisThread调用，传递layer_sigma_map
        self.analysis_thread = AnalysisThread(dat_path, output_path, layer_sigma_map, chips_per)
        self.analysis_thread.progress.connect(progress.setValue)
        self.analysis_thread.finished.connect(lambda msg: (progress.close(), self.log(msg, "info")))
        self.analysis_thread.error.connect(lambda err: (progress.close(), self.log(f"分析错误: {err}", "error")))
        self.analysis_thread.start()

    def send_commands_for_online_layers(self, online_boards: List[int]):
        command_sequence = [
            ("02-Reg02", [0x00ac0837] * 6),
            ("02-Reg03", [0x0080010f] * 6)
        ]
        for board_id in online_boards:
            self.log(f"正在为板卡 B-0b{board_id:08b} 发送命令序列", "info")
            for command_name, command_data in command_sequence:
                for_send = self.make_command(command_name, command_data, board_id)
                if for_send is None:
                    self.log(f"命令 {command_name} 生成失败", "error")
                    self.show_warning_dialog("错误", f"命令 {command_name} 生成失败")
                    continue
                try:
                    sended = self.send_to_usb(for_send)
                    if sended < 0:
                        return False
                    self.log(f">> 发送 {len(for_send)} 字节 (板卡 B-0b{board_id:08b}, {command_name}): {' '.join(f'{b:02x}' for b in for_send)}", "send")
                    readed = self.read_from_usb(1000)
                    if readed != 4:
                        self.log(f"命令 {command_name} 接收数据无效 (readed={readed})", "error")
                        self.show_warning_dialog("错误", f"命令 {command_name} 接收数据无效")
                        return False
                    get = self.recv[:readed]
                    hex_str = ' '.join(f'{b:02x}' for b in get)
                    self.log(f"<< 接收 {len(get):02} 字节: {hex_str}", "receive")
                    time.sleep(0.1)
                except Exception as e:
                    self.log(f"发送命令 {command_name} 时发生错误: {e}", "error")
                    self.show_warning_dialog("错误", f"发送命令 {command_name} 失败: {e}")
                    return False
            self.log(f"板卡 B-0b{board_id:08b} 命令序列发送完成", "info")
        return True

    def init_hardcoded_commands(self):
        self.clear_all_thr_data()
        self.ui.combox_commands.clear()

        command_names = [
            "Check", "02-Reg02", "02-Reg03", "03-Th_data", "05-Th_value",
            "06-Th_enable (baseline)", "06-Th_enable (Cosmic)", "Start", "Stop",
            "Select_Trigger_Layers", "08-Filter"
        ]

        self.commands = {
            "Check": [],
            "02-Reg02": [0x00ac0837] * self.chips_per_board,
            "02-Reg03": [0x0080010f] * self.chips_per_board,
            "05-Th_value": [0x07],
            "06-Th_enable (baseline)": [0x00],
            "06-Th_enable (Cosmic)": [0x01],
            "08-Filter": [0x00],
            "Start": [],
            "Stop": [],
            "Select_Trigger_Layers": [0x03]
        }

        self.board_ids = [0x01, 0x02, 0x04, 0x08, 0x10, 0x20]
        self.chip_ids = [i for i in range(self.chips_per_board)] * len(self.board_ids)
        self.adc_ths = [0x00123456] * (len(self.board_ids) * self.chips_per_board * self.channels_per_chip)
        
        for name in command_names:
            self.ui.combox_commands.addItem(name)
        self.ui.combox_commands.setCurrentIndex(0)
        if self.is_connected == 1:
            self.ui.button_command_send.setEnabled(True)

    def log(self, message: str, log_type: str = "info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"{timestamp}: {message}"
        self.log_builder.append(full_message)
        self.log_entry_count += 1
        
        if self.log_entry_count >= 500:
            self.ui.text_logger.clear()
            self.log_builder = self.log_builder[-100:]  #
            self.log_entry_count = 0
            self.log("日志已自动清理，仅保留最近记录", "info")
        
        if "芯片数量" in message or "初始化" in message:
            message = f"{message} [当前配置: {self.chips_per_board}芯片/板卡]"
        if len(self.log_builder) > 100:
            self.log_builder = self.log_builder[-100:]
        
        cursor = self.ui.text_logger.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.ui.text_logger.setTextCursor(cursor)
        
        format_obj = QTextCharFormat()
        if log_type == "send":
            format_obj.setForeground(QColor("green"))
            format_obj.setFontWeight(75)
            prefix = ">> [发送] "
        elif log_type == "receive":
            format_obj.setForeground(QColor("blue"))
            format_obj.setFontWeight(75)
            prefix = "<< [接收] "
        elif log_type == "error":
            format_obj.setForeground(QColor("red"))
            format_obj.setFontWeight(75)
            prefix = "[错误] "
        else:
            format_obj.setForeground(QColor("black"))
            prefix = ""
        
        cursor.insertText(prefix + full_message, format_obj)
        cursor.insertBlock()
        self.ui.text_logger.ensureCursorVisible()

    def on_trigger_value_changed(self, index):
        layer_values = {0: 0x00, 1: 0x01, 2: 0x02, 3: 0x03, 4: 0x04, 5: 0x05, 6: 0x06}
        selected_value = layer_values.get(index, 0x01)
        self.commands["Select_Trigger_Layers"] = [selected_value]
        self.log(f"触发层数已设置为: 0x{selected_value:02x} ({bin(selected_value)[2:].zfill(8)})", "info")

    def refresh_device_lists(self):
        self.ui.combox_device_lists.clear()
        self.log("开始扫描 USB 设备...", "info")
        devices = usb.core.find(backend=backend, find_all=True, idVendor=VENDOR_ID, idProduct=PRODUCT_ID)
        count = 0
        sstr = []
        
        if devices is None:
            self.log("未找到设备或后端加载失败。", "error")
            all_devices = usb.core.find(backend=backend, find_all=True)
            if all_devices:
                for dev in all_devices:
                    self.log(f"发现设备: VID=0x{dev.idVendor:04x}, PID=0x{dev.idProduct:04x}", "info")
            return
        for dev in devices:
            count += 1
            dev_info = f"VID: 0x{dev.idVendor:04x}, PID: 0x{dev.idProduct:04x}"
            dev_name = usb.util.get_string(dev, dev.iProduct) or "未知设备"
            self.ui.combox_device_lists.addItem(dev_info, dev)
            sstr.append(f"{dev_info} {dev_name}")
        if count > 0:
            self.ui.button_connect.setEnabled(True)
            self.log(f"找到 {count} 个设备: {' '.join(sstr)}", "info")
        else:
            self.log("未找到匹配的设备。", "error")

    def combbox_device_lists_index_changed(self, index: int):
        if index < 0:
            return
        item = self.ui.combox_device_lists.itemText(index)
        pattern = r"VID:.*PID:.*"
        if re.match(pattern, item):
            self.ui.button_connect.setEnabled(True)
            self.is_connected = 1
        else:
            self.ui.button_connect.setEnabled(False)
            self.is_connected = 0

    def on_button_update_clicked(self):
        self.refresh_device_lists()

    def on_button_connect_clicked(self):
        index = self.ui.combox_device_lists.currentIndex()
        if index < 0:
            self.log("连接失败：未选择设备。", "error")
            self.show_warning_dialog("错误", "未选择设备")
            return
        self.selected_device = self.ui.combox_device_lists.itemData(index)
        if self.selected_device:
            try:
                self.selected_device.set_configuration()
                cfg = self.selected_device.get_active_configuration()
                intf = cfg[(0, 0)]
                # 调试：打印所有端点
                for ep in intf:
                    self.log(f"端点地址: 0x{ep.bEndpointAddress:02x}, 类型: {ep.bmAttributes}", "info")
                self.ep_out = usb.util.find_descriptor(intf, bEndpointAddress=EP_OUT)
                self.ep_in = usb.util.find_descriptor(intf, bEndpointAddress=EP_IN)
                if self.ep_out and self.ep_in:
                    self.log("连接成功！", "info")
                    self.is_connected = 1
                else:
                    self.log(f"连接失败：未找到端点 EP_OUT=0x{EP_OUT:02x} 或 EP_IN=0x{EP_IN:02x}", "error")
                    self.show_warning_dialog("错误", f"未找到所需的批量端点 EP_OUT=0x{EP_OUT:02x} 或 EP_IN=0x{EP_IN:02x}")
                    self.is_connected = 0
            except usb.core.USBError as e:
                self.log(f"连接失败：USB 错误: {e}", "error")
                self.show_warning_dialog("错误", f"USB 连接错误: {e}")
                self.is_connected = 0
        else:
            self.log("连接失败：无效设备。", "error")
            self.show_warning_dialog("错误", "无效设备")
            self.is_connected = 0

    def on_button_send_clicked(self):
        if not self.selected_device or not self.ep_out or not self.ep_in:
            self.log("发送失败：设备未连接或端点未初始化", "error")
            self.show_warning_dialog("错误", "请先选择并连接有效设备")
            return
        if self.threshold_commands:
            self.log("开始发送阈值文件命令", "info")
            original_boards = self.boards
            for cmd in self.threshold_commands:
                command_name = cmd["name"]
                command_data = cmd["data"]
                if command_name in ["05-Th_value", "06-Th_enable (baseline)", "06-Th_enable (Cosmic)"]:
                    online_boards = self.update_layer_status(self.boards)
                    if not online_boards:
                        self.log(f"无在线板卡，无法发送命令 {command_name}", "error")
                        self.show_warning_dialog("错误", f"无在线板卡，请先执行 Check 命令")
                        return
                    
                    for board_id in online_boards:
                        for_send = self.make_command(command_name, command_data, board_id)
                        if for_send is None:
                            self.log(f"命令 {command_name} 生成失败", "error")
                            self.show_warning_dialog("错误", f"命令 {command_name} 生成失败")
                            continue
                        try:
                            sended = self.send_to_usb(for_send)
                            if sended < 0:
                                return
                            self.log(f">> 发送 {len(for_send)} 字节 (板卡 B-0b{board_id:08b}, {command_name}): {' '.join(f'{b:02x}' for b in for_send)}", "send")
                            readed = self.read_from_usb(1000)
                            if readed != 4:
                                self.log(f"命令 {command_name} 接收数据无效 (readed={readed})", "error")
                                self.show_warning_dialog("错误", f"命令 {command_name} 接收数据无效")
                                return
                            get = self.recv[:readed]
                            hex_str = ' '.join(f'{b:02x}' for b in get)
                            self.log(f"<< 接收 {len(get):02} 字节: {hex_str}", "receive")
                            time.sleep(0.1)
                        except Exception as e:
                            self.log(f"发送命令 {command_name} 时发生错误: {e}", "error")
                            self.show_warning_dialog("错误", f"发送命令 {command_name} 失败: {e}")
                            return
                else:
                    for_send = self.make_command(command_name, command_data)
                    if for_send is None:
                        self.log(f"命令 {command_name} 生成失败", "error")
                        self.show_warning_dialog("错误", f"命令 {command_name} 生成失败")
                        continue
                    try:
                        sended = self.send_to_usb(for_send)
                        if sended < 0:
                            return
                        self.log(f">> 发送 {len(for_send)} 字节 ({command_name}): {' '.join(f'{b:02x}' for b in for_send)}", "send")
                        readed = self.read_from_usb(1000)
                        if readed != 4:
                            self.log(f"命令 {command_name} 接收数据无效", "error")
                            self.show_warning_dialog("错误", f"命令 {command_name} 接收数据无效")
                            return
                        get = self.recv[:readed]
                        hex_str = ' '.join(f'{b:02x}' for b in get)
                        self.log(f"<< 接收 {len(get):02} 字节: {hex_str}", "receive")
                        time.sleep(0.1)
                    except Exception as e:
                        self.log(f"发送命令 {command_name} 时发生错误: {e}", "error")
                        self.show_warning_dialog("错误", f"发送命令 {command_name} 失败: {e}")
                        return
            self.log("阈值文件命令发送完成", "info")
            self.threshold_commands = []
            self.boards = original_boards
            self.update_layer_status(self.boards)
            return
        input_str = self.ui.lineedit_for_send.text().strip()
        bytes_array = self.parse_hex_string00(input_str)
        if len(bytes_array) != 4:
            self.log("输入必须为 4 字节的十六进制", "error")
            self.show_warning_dialog("错误", "输入必须为 4 字节的十六进制")
            return
        try:
            sended = self.send_to_usb(bytes_array)
            if sended < 0:
                return
            self.log(f">> 发送 {sended} 字节: {' '.join(f'{b:02x}' for b in bytes_array)}", "send")
            readed = self.read_from_usb(1000)
            if readed != 4:
                self.log(f"接收数据无效 (readed={readed})", "error")
                self.show_warning_dialog("错误", "接收数据无效")
                return
            get = self.recv[:readed]
            hex_str = ' '.join(f'{b:02x}' for b in get)
            self.log(f"<< 接收 {len(get)} 字节: {hex_str}", "receive")
            if bytes_array == struct.pack("BBBB", 0xfa, 0xff, 0x01, 0x00):
                self.boards = get[2]
                self.update_layer_status(self.boards)
        except Exception as e:
            self.log(f"发送数据时发生错误: {e}", "error")
            self.show_warning_dialog("错误", f"发送数据失败: {e}")

    def on_button_on_command_send_clicked(self):
        if not self.selected_device or not self.ep_out or not self.ep_in:
            self.log("无法发送命令：设备未连接或端点未初始化", "error")
            self.show_warning_dialog("错误", "请先连接设备")
            return
        index = self.ui.combox_commands.currentIndex()
        if index < 0:
            self.log("未选择命令", "error")
            self.show_warning_dialog("错误", "请先选择一个命令")
            return
        current_index = self.ui.combox_commands.itemText(index)
        
        try:
            if current_index == "03-Th_data":
                self.log("开始处理阈值数据", "info")
                online_boards = self.update_layer_status(self.boards)
                if not online_boards:
                    self.log("无在线板卡，无法发送阈值数据", "error")
                    self.show_warning_dialog("错误", "无在线板卡，请先执行 Check 命令")
                    return
                
                total_chips = len(online_boards) * self.chips_per_board
                self.log(f"检测到 {len(online_boards)} 层在线，共 {total_chips} 个芯片", "info")

                if not all(b in self.board_ids for b in online_boards):
                    self.log(f"在线板卡 {online_boards} 与 Excel 文件板卡 ID {self.board_ids} 不匹配", "error")
                    self.show_warning_dialog("错误", "在线板卡与 Excel 文件板卡 ID 不匹配")
                    return
                
                expected_ths_length = total_chips * self.channels_per_chip
                
                if len(self.adc_ths) < expected_ths_length:
                    self.log(f"阈值数据长度 ({len(self.adc_ths)}) 不足预期 ({expected_ths_length})", "error")
                    self.show_warning_dialog("错误", f"阈值数据与在线芯片数不匹配")
                    return
                elif len(self.adc_ths) > expected_ths_length:
                    self.log(f"警告: 阈值数据长度 ({len(self.adc_ths)}) 超过预期 ({expected_ths_length})，将使用前{expected_ths_length}个数据", "warning")
                    self.adc_ths = self.adc_ths[:expected_ths_length]
                
                idx = 0
                for board in online_boards:
                    for chip in range(self.chips_per_board):
                        start_idx = idx * self.channels_per_chip
                        subs = self.adc_ths[start_idx:start_idx + self.channels_per_chip]
                        if len(subs) != self.channels_per_chip:
                            self.log(f"板卡 B-0b{board:08b} 芯片 {chip} 阈值数据不足128个", "error")
                            self.show_warning_dialog("错误", f"阈值数据不足")
                            return
                        for_send = self.make_command_adc_thr(board, chip, subs)
                        if for_send is None:
                            self.log(f"生成阈值命令失败 (板卡 B-0b{board:08b}, 芯片 {chip})", "error")
                            self.show_warning_dialog("错误", f"生成阈值命令失败")
                            return
                        try:
                            sended = self.send_to_usb(for_send)
                            if sended < 0:
                                return
                            hex_str = ' '.join(f'{b:02x}' for b in for_send)
                            self.log(f">> 已发送: {sended} 字节 (板卡 B-0b{board:08b}, 芯片 {chip}): {hex_str}", "send")
                            readed = self.read_from_usb(300)
                            if readed != 4:
                                self.log(f"接收数据无效 (readed={readed})", "error")
                                self.show_warning_dialog("错误", f"接收数据无效")
                                return
                            get = self.recv[:readed]
                            hex_str = ' '.join(f'{b:02x}' for b in get)
                            self.log(f"<< 接收 {len(get):02} 字节: {hex_str}", "receive")
                        except Exception as e:
                            self.log(f"发送阈值数据时发生错误 (板卡 B-0b{board:08b}, 芯片 {chip}): {e}", "error")
                            self.show_warning_dialog("错误", f"发送阈值数据失败: {e}")
                            return
                        idx += 1
                self.log(f"阈值数据发送完成，共处理 {idx} 个芯片", "info")
            
            elif current_index in ["05-Th_value", "06-Th_enable (baseline)", "06-Th_enable (Cosmic)"]:
                self.log(f"原始 self.boards: 0x{self.boards:02x}", "info")
                online_boards = self.update_layer_status(self.boards)
                if not online_boards:
                    self.log("无在线板卡，无法发送命令", "error")
                    self.show_warning_dialog("错误", "无在线板卡，请先执行 Check 命令")
                    return
    
                if not all(b in self.board_ids for b in online_boards):
                    self.log(f"在线板卡 {online_boards} 与 Excel 文件板卡 ID {self.board_ids} 不匹配", "error")
                    self.show_warning_dialog("错误", "在线板卡与 Excel 文件板卡 ID 不匹配")
                    return
                
                command_data = None
                for cmd in self.threshold_commands:
                    if cmd["name"] == current_index:
                        command_data = cmd["data"]
                        break
                if command_data is None:
                    command_data = self.commands.get(current_index, [0x07 if current_index == "05-Th_value" else 0x00])
                
                for board_id in online_boards:
                    self.log(f"处理板卡 B-0b{board_id:08b}", "info")
                    for_send = self.make_command(current_index, command_data, board_id)
                    if for_send is None:
                        self.log(f"命令 {current_index} 生成失败", "error")
                        self.show_warning_dialog("错误", f"命令 {current_index} 生成失败")
                        continue
                    try:
                        sended = self.send_to_usb(for_send)
                        if sended < 0:
                            return
                        self.log(f">> 发送 {len(for_send)} 字节 (板卡 B-0b{board_id:08b}, {current_index}): {' '.join(f'{b:02x}' for b in for_send)}", "send")
                        readed = self.read_from_usb(1000)
                        if readed != 4:
                            self.log(f"命令 {current_index} 接收数据无效 (readed={readed})", "error")
                            self.show_warning_dialog("错误", f"命令 {current_index} 接收数据无效")
                            return
                        get = self.recv[:readed]
                        self.log(f"原始响应字节: {list(get)}", "info")
                        hex_str = ' '.join(f'{b:02x}' for b in get)
                        self.log(f"<< 接收 {len(get):02} 字节: {hex_str}", "receive")
                        time.sleep(0.1)
                    except Exception as e:
                        self.log(f"发送命令 {current_index} 时发生错误: {e}", "error")
                        self.show_warning_dialog("错误", f"发送命令 {current_index} 失败: {e}")
                        return
                self.log(f"命令 {current_index} 发送完成，共 {len(online_boards)} 个板卡", "info")
                if current_index in ["06-Th_enable (baseline)", "06-Th_enable (Cosmic)"]:
                    self.collection_mode = "baseline" if current_index == "06-Th_enable (baseline)" else "cosmic"
                    self.log(f"采集模式设置为: {self.collection_mode}", "info")
                self.update_layer_status(self.boards)
            
            else:
                command = self.commands.get(current_index, [])
                self.log(f"命令长度: {len(command)}", "info")
                for_send = self.make_command(current_index, command)
                if for_send is None:
                    self.log(f"命令 {current_index} 生成失败", "error")
                    self.show_warning_dialog("错误", f"命令 {current_index} 生成失败")
                    return
                try:
                    sended = self.send_to_usb(for_send)
                    if sended < 0:
                        return
                    self.log(f">> 发送 {len(for_send)} 字节: {' '.join(f'{b:02x}' for b in for_send)}", "send")
                    if current_index == "Stop":
                        self.log("停止监控", "info")
                        self.safe_stop_async_read()
                        return
                    readed = self.read_from_usb()
                    if readed != 4:
                        self.log(f"接收数据无效 (readed={readed})", "error")
                        self.show_warning_dialog("错误", f"接收数据无效")
                        return
                    get = self.recv[:readed]
                    hex_str = ' '.join(f'{b:02x}' for b in get)
                    self.log(f"<< 接收 {len(get):02} 字节: {hex_str}", "receive")
                    if current_index == "Check":
                        self.boards = get[2]
                        online_boards = self.update_layer_status(self.boards)
                        if online_boards:
                            self.log("Check 完成！检测到在线板卡: " + 
                 ', '.join(f'B-0b{board_id:08b}' for board_id in online_boards), "info")
                            self.log("提示：Reg2/Reg3 未自动发送，请手动在下拉框选择 '02-Reg02' 或 '02-Reg03' 发送", "warning")
                        else:
                            self.log("无在线板卡", "warning")
                    if current_index == "Start":
                        self.start_async_read()
                        if self.collection_mode == "cosmic":
                            self.log("Cosmic 模式：10秒后将发送 start trigger 命令", "info")
                           # QTimer.singleShot(10000, lambda: self.send_trigger_command(start=True))

                except Exception as e:
                    self.log(f"处理命令 {current_index} 时发生错误: {e}", "error")
                    self.show_warning_dialog("错误", f"处理命令 {current_index} 失败: {e}")
        except Exception as e:
            self.log(f"处理命令 {current_index} 时发生未知错误: {e}", "error")
            self.show_warning_dialog("错误", f"处理命令 {current_index} 失败: {e}")
    
    def make_command(self, command: str, data: List[int], board: Optional[int] = None) -> Optional[bytes]:
        target_board = board if board is not None else self.boards
        try:
            if command == "Check":
                return struct.pack("BBBB", 0xfa, 0xff, 0x01, 0x00)
            elif command == "02-Reg02":
                result = bytearray(4 + len(data) * 4)
                result[0] = 0xfa
                result[1] = target_board
                result[2] = 0x02
                result[3] = 0x00
                count = 0
                index = 4
                for x in data:
                    result[index] = count
                    result[index+1] = (x >> 16) & 0xff
                    result[index+2] = (x >> 8) & 0xff
                    result[index+3] = x & 0xff
                    index += 4
                    count += 1
                return bytes(result)
            elif command == "02-Reg03":
                result = bytearray(4 + len(data) * 4)
                result[0] = 0xfa
                result[1] = target_board
                result[2] = 0x03
                result[3] = 0x00
                count = 0
                index = 4
                for x in data:
                    result[index] = count
                    result[index+1] = (x >> 16) & 0xff
                    result[index+2] = (x >> 8) & 0xff
                    result[index+3] = x & 0xff
                    index += 4
                    count += 1
                return bytes(result)
            elif command == "05-Th_value": # [bugfix] name sh be Th_num for fired nr of strips
                if len(data) != 1:
                    self.show_warning_dialog("错误", "阈值命令只能包含一个值")
                    return None
                return struct.pack("BBBB", 0xfa, target_board, 0x05, data[0])
            elif command in ["06-Th_enable (baseline)", "06-Th_enable (Cosmic)"]:
                if len(data) != 1:
                    self.show_warning_dialog("错误", "阈值启用命令只能包含一个值")
                    return None
                return struct.pack("BBBB", 0xfa, target_board, 0x06, data[0])
            elif command == "Start":
                return struct.pack("BBBB", 0xfa, target_board, 0x07, 0x01)
            elif command == "Stop":
                return struct.pack("BBBB", 0xfa, target_board, 0x07, 0x00)
            elif command == "Select_Trigger_Layers":
                if len(data) != 1:
                    self.show_warning_dialog("错误", "触发层数命令只能包含一个值")
                    return None
                return struct.pack("BBBB", 0xfa, 0x00, 0x81, data[0])
            elif command == "08-Filter":
                if len(data) != 1:
                    self.show_warning_dialog("错误", "Filter命令只能包含一个值")
                    return None
                return struct.pack("BBBB", 0xfa, target_board, 0x08, data[0])
            else:
                return None
        except Exception as e:
            self.log(f"生成命令 {command} 时发生错误: {e}", "error")
            self.show_warning_dialog("错误", f"生成命令 {command} 失败: {e}")
            return None

    def make_command_adc_thr(self, board: int, chip: int, data: List[int]) -> Optional[bytes]:
        try:
            result = bytearray(4 + len(data) * 4)
            result[0] = 0xfa
            result[1] = board
            result[2] = 0x04
            result[3] = chip
            index = 4
            for channel, x in enumerate(data):
                result[index] = channel & 0xff
                result[index+1] = (x >> 16) & 0xff
                result[index+2] = (x >> 8) & 0xff
                result[index+3] = x & 0xff
                index += 4
            return bytes(result)
        except Exception as e:
            self.log(f"生成 ADC 阈值命令失败 (板卡 {board}, 芯片 {chip}): {e}", "error")
            self.show_warning_dialog("错误", f"生成 ADC 阈值命令失败: {e}")
            return None
    
    def read_from_usb(self, timeout: int = 1000) -> int:
        if not self.ep_in:
            self.log("读取失败：输入端点未初始化", "error")
            self.show_warning_dialog("错误", "USB 输入端点未初始化")
            return -1
        try:
            data = self.ep_in.read(len(self.recv), timeout)
            self.recv[:len(data)] = data
            self.last_readed = len(data)
            return self.last_readed
        except usb.core.USBError as e:
            self.log(f"USB 读取错误: {e}", "error")
            self.show_warning_dialog("错误", f"USB 读取失败: {e}")
            return -1
        except Exception as e:
            self.log(f"读取数据时发生未知错误: {e}", "error")
            self.show_warning_dialog("错误", f"读取数据失败: {e}")
            return -1

    def send_to_usb(self, bytes_array: bytes) -> int:
        if not self.ep_out:
            self.log("发送失败：输出端点未初始化", "error")
            self.show_warning_dialog("错误", "USB 输出端点未初始化，请检查设备连接")
            return -1
        try:
            return self.ep_out.write(bytes_array)
        except usb.core.USBError as e:
            self.log(f"USB 发送错误: {e}", "error")
            self.show_warning_dialog("错误", f"USB 发送失败: {e}")
            return -1
        except Exception as e:
            self.log(f"发送数据时发生未知错误: {e}", "error")
            self.show_warning_dialog("错误", f"发送数据失败: {e}")
            return -1

    def select_data_dir(self):
        """弹出文件夹选择对话框，设置数据保存根目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择数据保存目录", self.data_base_dir)
        if dir_path:
            self.data_base_dir = dir_path
            self.log(f"数据保存目录已设置为: {self.data_base_dir}", "info")

    def update_run_prefix(self, text):
        """实时更新 run 前缀"""
        self.run_prefix = text.strip() or "run"
    # 新增：处理 LineEdit 输入
    # ----------------------------------------------------------------------
    def on_run_prefix_changed(self, text: str):
        """用户在 lineEdit_run_prefix 中修改文字时触发"""
        clean = text.strip()
        if clean != self.run_prefix:
            old = self.run_prefix
            self.run_prefix = clean if clean else "run"
            #self.log(f"Run 前缀已从 '{old}' 改为 '{self.run_prefix}'", "info")

    
    def start_async_read(self):
        if self.is_running:
            self.log("采集已在运行中，忽略重复 Start", "warning")
            return

        if not self.ep_in:
            self.log("USB 输入端点未初始化", "error")
            return

        if self.collection_mode not in ("baseline", "cosmic"):
            self.log("错误：未设置采集模式！请先发送 06-Th_enable 命令", "error")
            return

        # ========== 关键：点击 Start 时生成一次时间戳文件夹 ==========
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        self.current_data_dir = os.path.join(self.data_base_dir, timestamp)
        os.makedirs(self.current_data_dir, exist_ok=True)
        self.log(f"本次采集数据将保存至: {self.current_data_dir}", "info")

        # 重置文件计数器
        self.run_counter = 0
        # ============================================================

        self.log(f"开始 {self.collection_mode.upper()} 采集，时长 {self.file_interval_minutes} 分钟", "info")

        self.collection_duration = self.file_interval_minutes * 60
        self.is_running = True
        self.stop_sent = False
 
        # 启动读取线程
        self.read_thread = UsbReaderThread(self.ep_in)
        self.read_thread.data_received.connect(self.on_data_received)
        self.read_thread.error_occurred.connect(self.on_usb_error)
        self.read_thread.finished_signal.connect(self.on_read_thread_finished)
        self.read_thread.start()

        # 创建第一个文件
        self.create_new_writer_thread()


        # 启动计时器
        self.start_time = time.time()
        self.trigger_timer.start(1000)


    def create_new_writer_thread(self):
        """创建新文件 + 新 writer（使用已存在的 self.current_data_dir）"""
        if not self.is_running or not hasattr(self, 'current_data_dir') or not self.current_data_dir:
            self.log("错误：数据目录未初始化，无法创建新文件", "error")
            return

        if not hasattr(self, 'read_thread') or not self.read_thread:
            self.log("错误：读取线程未启动，無法创建 writer", "error")
            return

        # 1. 生成文件名（不生成新目录！）
        file_name = f"{self.run_prefix}-{self.run_counter}.dat"
        file_path = os.path.join(self.current_data_dir, file_name)  # 使用当前目录

        # 2. 停止旧 writer
        self.stop_current_writer()

        # 3. 创建新 writer
        try:
            self.writer_thread = FileWriterThread(
                queue=self.read_thread.queue,
                file_path=file_path,
                collection_mode=self.collection_mode
            )
            self.writer_thread.error_occurred.connect(self.on_usb_error)
            self.writer_thread.finished_signal.connect(self.on_writer_thread_finished)
            self.writer_thread.start()

            self.log(f"新建数据文件: {file_path}", "info")
        except Exception as e:
            self.log(f"创建 writer 线程失败: {e}", "error")
            self.on_usb_error(f"创建 writer 线程失败: {e}")

        # 4. 递增计数器
        self.run_counter += 1


    def safe_stop_async_read(self):
        if self.is_stopping:
            self.log("正在停止中，请稍候...", "info")
            return
        self.is_stopping = True
        self.log("开始安全停止采集...", "info")
        QTimer.singleShot(0, self._stop_async_read_internal)

    
    def _stop_async_read_internal(self):
        try:
            stop_cmd = self.make_command("Stop", [])
            if stop_cmd and self.ep_out:
                try:
                    self.send_to_usb(stop_cmd)
                    self.log("已发送停止命令", "info")
                except Exception as e:
                    self.log(f"发送停止命令时出错: {e}", "error")
                    self.show_warning_dialog("错误", f"发送停止命令失败: {e}")
            
            time.sleep(0.5)  # 延长等待
            
            if self.read_thread:
                self.read_thread.stop()
            
            self.log("开始排空剩余数据 (延长尝试)...", "info")
            self.drain_thread = DrainDataThread(self.ep_in)
            self.drain_thread.progress_signal.connect(self.log)
            self.drain_thread.finished_signal.connect(self.on_drain_finished)
            self.drain_thread.start()
            
        except Exception as e:
            self.log(f"停止过程中发生错误: {e}", "error")
            self.show_warning_dialog("错误", f"停止异步读取失败: {e}")
        finally:
            self.is_stopping = False
    
    def on_drain_finished(self):
        self.log("数据排空完成", "info")

        # 停止 writer
        self.stop_current_writer()

        # 清理线程引用
        self.read_thread = None
        self.writer_thread = None
        self.drain_thread = None

        # 确保状态正确
        self.is_running = False
        self.is_stopping = False
        self.trigger_timer.stop()
        self.statusBar().showMessage("采集已完全停止")
    def on_read_thread_finished(self):
        self.log("读取线程已停止", "info")

    def on_writer_thread_finished(self):
        self.log("写入线程已停止", "info")

    def stop_async_read(self):
        if self.read_thread:
            self.read_thread.stop()
            self.read_thread = None
        if self.writer_thread:
            self.writer_thread.stop()
            self.writer_thread = None
        if self.drain_thread:
            self.drain_thread.stop()
            self.drain_thread = None
        self.is_stopping = False

    def on_data_received(self, data: bytes):
        self.received_count += 1
        if self.received_count % 500 == 0:  
            self.log(f"已累计推送 {self.received_count} 条数据，最新包大小: {len(data)} 字节 ", "receive")
    
    def on_usb_error(self, error: str):
        self.log(f"USB 错误: {error}", "error")
    
    def parse_hex_string00(self, hex_str: str) -> bytes:
        try:
            hex_str = hex_str.replace(" ", "")
            if len(hex_str) != 8:
                self.log("解析失败：输入必须为 4 字节的十六进制", "error")
                self.show_warning_dialog("错误", "输入必须为 4 字节的十六进制")
                return b""
            bytes_array = bytearray()
            for i in range(0, 8, 2):
                bytes_array.append(int(hex_str[i:i+2], 16))
            return bytes(bytes_array)
        except Exception as e:
            self.log(f"解析十六进制字符串失败: {e}", "error")
            self.show_warning_dialog("错误", f"解析十六进制字符串失败: {e}")
            return b""

    def clear_all_thr_data(self):
        self.board_ids.clear()
        self.chip_ids.clear()
        self.adc_ths.clear()
        self.threshold_commands.clear()
        self.log("已清空所有阈值数据", "info")

    def show_warning_dialog(self, title: str, content: str):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle(title)
        msg.setText(content)
        msg.setModal(False)
        msg.show()
    """清除日志
    def action_button_test_click(self):
        self.log_builder.clear()
        self.ui.text_logger.clear()
        self.log_entry_count = 0
    """

    def on_button_config_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择阈值配置文件", "", "Excel Files (*.xlsx)")
        if not file_path:
            self.log("未选择配置文件", "warning")
            return

        try:
            df = pd.read_excel(file_path, header=None)
            if df.empty:
                self.log("配置文件为空", "error")
                self.show_warning_dialog("错误", "配置文件为空")
                return

            total_rows, total_cols = df.shape
            self.log(f"读取文件: {total_rows} 行 × {total_cols} 列", "info")

            # === 1. 动态检测层数（通过 B-0b... 标识）===
            layer_ids = set()
            for col in range(total_cols):
                for row in range(128, total_rows):
                    cell = df.iloc[row, col]
                    if isinstance(cell, str) and cell.startswith("B-0b"):
                        try:
                            lid_bit = int(cell.split("0b")[1], 2)
                            lid = 0
                            while lid_bit > 1:
                                lid_bit >>= 1
                                lid += 1
                            layer_ids.add(lid)
                        except:
                            continue
            layer_count = len(layer_ids)
            if layer_count == 0:
                self.log("未检测到层标识 (B-0b...)", "error")
                return

            # === 2. 计算每层芯片数 ===
            chips_per_layer = total_cols // layer_count
            if total_cols % layer_count != 0:
                self.log(f"列数 {total_cols} 不能被层数 {layer_count} 整除", "error")
                return

            self.log(f"检测到 {layer_count} 层，每层 {chips_per_layer} 芯片，共 {total_cols} 列", "info")

            # === 3. 验证阈值区域：至少 128 行 ===
            if total_rows < 128:
                self.log(f"阈值区域不足 128 行，实际 {total_rows} 行", "error")
                return

            # === 4. 清空并重新设置板卡/芯片 ID ===
            self.clear_all_thr_data()
            self.board_ids = []
            self.chip_ids = []
            for lid in sorted(layer_ids):
                for chip in range(chips_per_layer):
                    self.board_ids.append(1 << lid)   # 例如 0x01, 0x02, 0x04
                    self.chip_ids.append(chip)

            # === 5. 解析阈值（前 128 行）===
            self.adc_ths = []
            for col in range(total_cols):
                for row in range(128):
                    value = df.iloc[row, col]
                    if pd.isna(value):
                        self.log(f"阈值缺失: 行{row+1} 列{col+1}", "error")
                        return
                    if isinstance(value, str) and value.startswith("0x"):
                        try:
                            threshold = int(value, 16)
                            self.adc_ths.append(threshold)
                        except ValueError:
                            self.log(f"无效十六进制: 行{row+1} 列{col+1}: {value}", "error")
                            return
                    else:
                        self.log(f"阈值格式错误: 行{row+1} 列{col+1}: {value}", "error")
                        return

            # === 6. 解析命令头（第 129 行开始）===
            self.threshold_commands = []
            for col in range(total_cols):
                col_commands = []
                current_reg = None
                for row in range(128, total_rows):
                    value = df.iloc[row, col]
                    if pd.isna(value):
                        continue

                    str_val = str(value).strip()

                    # 寄存器名
                    if str_val in ["02-Reg02", "02-Reg03", "04-Th_value", "05-Th_enable", "08-Filter", "03-Th_config"]:
                        if current_reg:
                            col_commands.append(current_reg)
                        current_reg = {"name": str_val, "data": []}
                    # B-0b... 或 C-0b... 或 ${...}
                    elif str_val.startswith("B-0b") or str_val.startswith("C-0b") or str_val.startswith("${"):
                        col_commands.append({"name": f"header_{col}_{row}", "data": [str_val]})
                    # 十六进制值
                    elif str_val.startswith("0x"):
                        try:
                            hex_val = int(str_val, 16)
                            if current_reg:
                                current_reg["data"].append(hex_val)
                        except ValueError:
                            self.log(f"无效命令值: 行{row+1} 列{col+1}: {str_val}", "error")
                            return
                if current_reg:
                    col_commands.append(current_reg)
                self.threshold_commands.extend(col_commands)
                
            self.log(f"成功加载配置文件: {file_path}\n"
                    f"  - 层数: {layer_count}\n"
                    f"  - 每层芯片: {chips_per_layer}\n"
                    f"  - 总阈值: {len(self.adc_ths)}\n"
                    f"  - 命令条数: {len(self.threshold_commands)}", "info")

        except Exception as e:
            import traceback
            self.log(f"加载配置文件失败: {e}\n{traceback.format_exc()}", "error")
            self.show_warning_dialog("错误", f"加载配置文件失败: {e}")
        

    
    def closeEvent(self, event):
        self.stop_async_read()
        if self.selected_device:
            try:
                usb.util.dispose_resources(self.selected_device)
            except Exception as e:
                self.log(f"释放USB资源失败: {e}", "error")
        event.accept()
    

    def update_time(self):#qlabel显示时间戳
            """实时更新 label_time 显示当前时间"""
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if hasattr(self.ui, 'label_time'):
                self.ui.label_time.setText(current_time)
            else:
                self.log("警告: 未找到 label_time 控件", "warning")

    def on_send_custom_th_value(self):
        if not self.selected_device or not self.ep_out or not self.ep_in:
            self.log("无法发送：设备未连接", "error")
            self.show_warning_dialog("错误", "请先连接设备")
            return
            
        # 获取在线板卡
        online_boards = self.update_layer_status(self.boards)
        if not online_boards:
            self.log("无在线板卡，无法发送 Th_value", "error")
            self.show_warning_dialog("错误", "请先执行 Check 命令")
            return

        # 动态查找 textEdit6 ~ textEdit12
        text_edits = []
        for i in range(6, 13):
            te_name = f"textEdit{i}"
            te = getattr(self.ui, te_name, None)
            if te is None:
                self.log(f"未找到 {te_name}，将使用默认值 0x07", "warning")
                text_edits.append(None)
            else:
                text_edits.append(te)

        # 解析输入值
        th_values = []
        for i, te in enumerate(text_edits):
            if te is None:
                val = 0x07
            else:
                text = te.toPlainText().strip()
                if not text:
                    val = 0x07
                    self.log(f"textEdit{i+6} 为空，使用默认值 0x07", "info")
                else:
                    try:
                        if text.lower().startswith("0x"):
                            val = int(text, 16)
                        else:
                            val = int(text)
                        if not (0 <= val <= 0xFF):
                            raise ValueError("超出 0~255 范围")
                    except:
                        self.log(f"textEdit{i+6} 输入无效: '{text}'，使用默认 0x07", "error")
                        val = 0x07
            th_values.append(val)

        if len(th_values) != 7:
            self.log(f"内部错误：th_values 长度为 {len(th_values)}，应为 7", "error")
            return

        self.log(f"准备发送自定义 Th_value: {[f'0x{v:02x}' for v in th_values]}", "info")

        # 发送给每个在线板卡
        for idx, board_id in enumerate(online_boards):
            if idx >= len(th_values):
                self.log(f"板卡 B-0b{board_id:08b} 无对应 Th_value，跳过", "warning")
                continue

            val = th_values[idx]
            for_send = self.make_command("05-Th_value", [val], board_id)
            if not for_send:
                self.log(f"生成命令失败 (板卡 B-0b{board_id:08b})", "error")
                continue

            try:
                # 显示完整命令字节
                hex_cmd = ' '.join(f'{b:02x}' for b in for_send)
                self.log(f">> 发送 Th_value 0x{val:02x} → 板卡 B-0b{board_id:08b}: {hex_cmd}", "send")

                sended = self.send_to_usb(for_send)
                if sended < 0:
                    self.log("USB 发送失败，停止后续发送", "error")
                    return

                readed = self.read_from_usb(1000)
                if readed != 4:
                    self.log(f"响应长度错误 (readed={readed})，继续下一个", "warning")
                    continue

                response = self.recv[:readed]
                resp_hex = ' '.join(f'{b:02x}' for b in response)
                self.log(f"<< 响应: {resp_hex}", "receive")
                time.sleep(0.1)

            except Exception as e:
                self.log(f"发送失败 (板卡 B-0b{board_id:08b}): {e}", "error")

        self.log("自定义 Th_value 发送完成", "info")

    def send_trigger_command(self, start: bool):
        if not self.ep_out:
            self.log("无法发送 trigger 命令：USB 未连接", "error")
            return
        cmd = 0x00 if start else 0x01
        command_bytes = bytes([0xfa, 0x00, 0x82, cmd])
        try:
            sended = self.send_to_usb(command_bytes)
            if sended > 0:
                self.log(f">> 发送 trigger {'start' if start else 'stop'}: {' '.join(f'{b:02x}' for b in command_bytes)}", "send")
        except Exception as e:
            self.log(f"发送 trigger 命令失败: {e}", "error")

    def send_start_trigger(self):
        self.send_trigger_command(start=True)

    def send_stop_trigger(self):
        self.send_trigger_command(start=False)

    def on_timer_tick(self):
        if not self.is_running:
            return

        elapsed = time.time() - self.start_time   # ← 正确计算
        if elapsed < 0:  # 防御性编程
            self.start_time = time.time()
            return

        remaining = self.collection_duration - elapsed
        mins, secs = divmod(int(remaining), 60)
        self.statusBar().showMessage(
        f"模式: {self.collection_mode.upper()} | 剩余: {mins:02d}:{secs:02d} | 文件: {self.run_counter}"
    )

        # ---------- 提前 30 秒 stop ----------
        if self.collection_mode == "cosmic" and elapsed >= self.collection_duration - 30 and not self.stop_sent:
            self.send_stop_trigger()
            self.log("Cosmic 模式：提前 30 秒发送 stop_trigger", "info")
            self.stop_sent = True

        # ---------- 时间到：只执行一次 ----------
        if elapsed >= self.collection_duration:
            self.log("时间到，准备切换文件", "info")

            # 1. 停止旧 writer
            self.stop_current_writer()

            # 2. Cosmic 重新使能
        #    if self.collection_mode == "cosmic":
         #       if not self.resend_th_enable("06-Th_enable (Cosmic)"):
          #          self.safe_stop_async_read()
           #         return

            # 3. 新文件
            self.create_new_writer_thread()


            # 5. 重置 start_time
            self.start_time = time.time()   # 正确重置
            self.stop_sent = False


    def stop_current_writer(self):
        if hasattr(self, 'writer_thread') and self.writer_thread and self.writer_thread.isRunning():
            self.writer_thread.stop()
            self.writer_thread.wait(5000)

    # def stop_collection(self):
    #     """彻底停止采集（Cosmic 结束或手动 Stop）"""
    #     self.trigger_timer.stop()
    #     if self.is_running and not self.stop_sent:
    #         self.send_stop_trigger()
            

    #     self.stop_current_writer()
    #     if hasattr(self, 'read_thread'):
    #         self.read_thread.stop()

    #     self.is_running = False
    #     self.log("采集已完全停止", "info")
    #     self.statusBar().showMessage("采集已停止")


    def on_analyze_baseline(self):
        """点击 Analyze baseline 按钮"""
        # 1. 选择文件
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 .dat 文件",
            "",
            "DAT Files (*.dat);;All Files (*)"
        )
        if not file_path:
            return  # 用户取消

        # 2. 禁用按钮，防止重复点击
        self.ui.pushButton_analyze.setEnabled(False)
        self.ui.pushButton_analyze.setText("分析中...")

        try:
            # 3. 调用 baseline.py 的主函数
            result_dir = run_analysis(file_path)

            # 4. 成功提示
            self.log(f"分析完成！结果已保存至:\n{result_dir}", "info")
            QMessageBox.information(self, "成功", f"分析完成！\n结果保存至:\n{result_dir}")

        except Exception as e:
            # 5. 错误处理
            error_msg = f"分析失败: {e}"
            self.log(error_msg, "error")
            QMessageBox.critical(self, "错误", error_msg)

        finally:
            # 6. 恢复按钮
            self.ui.pushButton_analyze.setEnabled(True)
            self.ui.pushButton_analyze.setText("Analyze baseline")


    def on_reset_usb_clicked(self):
        """用户手动点击“复位 USB”按钮"""
        if not self.is_connected:
            QMessageBox.warning(self, "警告", "请先连接 USB 设备")
            return

        reply = QMessageBox.question(
            self, "确认复位 USB",
            "此操作将：<br>"
            "• 中断所有 USB 传输<br>"
            "• 清除固件残留状态<br>"
            "• 用于软件卡退后恢复采集<br><br>"
            "<b>确定要执行吗？</b>",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # 禁用按钮，防止重复点击
        self.ui.pushButton_reset_usb.setEnabled(False)
        self.ui.pushButton_reset_usb.setText("复位中...")

        success = self.force_reset_usb_state()

        if success:
            self.log("USB 复位成功！现在可以 Start 采集", "info")
            QMessageBox.information(self, "成功", "USB 复位完成，可重新 Start 采集")
        else:
            self.log("USB 复位失败，建议重启硬件", "error")
            QMessageBox.critical(self, "失败", "复位失败，建议重启硬件")

        # 恢复按钮
        self.ui.pushButton_reset_usb.setEnabled(True)
        self.ui.pushButton_reset_usb.setText("复位 USB")
    
    def force_reset_usb_state(self) -> bool:
        if not self.selected_device:
            self.log("设备未连接", "error")
            return False

        try:
            # === 1. 停止旧读取线程 ===
            if hasattr(self, 'read_thread') and self.read_thread and self.read_thread.isRunning():
                self.read_thread.stop()
                self.read_thread.wait(1000)
                self.log("旧读取线程已停止", "info")

            # === 2. 发送 Stop 命令 ===
            self.log("1. 发送 Stop 命令", "info")
            try:
                self.send_to_usb(bytes([0xFA, 0xFF, 0x07, 0x00]))
                self.read_from_usb(200)
            except: self.log("Stop 超时（正常）", "warning")

            # === 3. 清除 HALT ===
            self.log("2. 清除端点 HALT", "info")
            for ep in (EP_IN, EP_OUT):
                try: self.selected_device.clear_halt(ep)
                except: pass

            # === 4. 完全释放旧设备 ===
            self.log("3. 完全释放旧设备", "info")
            old_dev = self.selected_device
            usb.util.dispose_resources(old_dev)
            old_dev = None
            import gc
            gc.collect()
            time.sleep(0.5)

            # === 5. 执行 reset ===
            self.log("4. 执行 USB reset", "info")
            try:
                old_dev.reset()  
            except: pass

            self.selected_device = None
            self.ep_in = None
            self.ep_out = None

            # === 6. 轮询新设备（最多 15 秒）===
            self.log("5. 轮询新设备（最多 15 秒）...", "info")
            for i in range(150):
                time.sleep(0.1)
                # 使用 libusb1 后端查找
                dev = usb.core.find(idVendor=VENDOR_ID, idProduct=PRODUCT_ID, backend=backend)
                if dev:
                    try:
                        dev.get_active_configuration()
                        self.selected_device = dev
                        self.log(f"新设备发现！耗时 {i*0.1:.1f}s", "info")
                        break
                    except:
                        continue
            else:
                raise RuntimeError("15 秒内未发现新设备")

            # === 7. 重新配置 ===
            self.log("6. 重新配置接口", "info")
            self.selected_device.set_configuration()
            cfg = self.selected_device.get_active_configuration()
            intf = cfg[(0, 0)]
            self.ep_out = usb.util.find_descriptor(intf, bEndpointAddress=EP_OUT)
            self.ep_in  = usb.util.find_descriptor(intf, bEndpointAddress=EP_IN)
            if not (self.ep_out and self.ep_in):
                raise RuntimeError("端点获取失败")

                        # === 8. 排空 ===
            self.log("7. 排空残留数据", "info")
            for _ in range(30):
                try: self.ep_in.read(16384, timeout=100)
                except: break
        
            # === 9. 自动发送 Start 命令！===
            self.log("8. 自动发送 Start 命令", "info")
            try:
                start_cmd = bytes([0xFA, 0xFF, 0x06, 0x00])
                self.send_to_usb(start_cmd)
                resp = self.read_from_usb(500)
                if resp == 4:
                    self.log("Start 命令成功，固件已启动采集", "info")
                else:
                    self.log("Start 响应异常", "warning")
            except Exception as e:
                self.log(f"Start 命令失败: {e}", "warning")

            self.log("USB 复位 + 自动 Start 完成！现在可以采集数据", "info")
            return True

        except Exception as e:
            import traceback
            self.log(f"复位失败: {e}\n{traceback.format_exc()}", "error")
            return False

    def _parse_hex_lineedit(self, le: QLineEdit, default: int) -> Optional[int]:
        """把 QLineEdit 的文本解析为 32 位整数，失败返回 None"""
        txt = le.text().strip()
        if not txt:
            return default
        try:
            if txt.lower().startswith("0x"):
                return int(txt, 16)
            else:
                return int(txt, 16)  # 直接当十六进制处理
        except ValueError:
            self.log(f"输入非法: {txt}（使用默认 0x{default:08x}）", "warning")
            return default


    def on_send_custom_regs_clicked(self):
        online_boards = self.update_layer_status(self.boards)
        if not online_boards:
            self.log("未检测到在线层，请先执行 Check", "error")
            return

        reg1_val = self._parse_hex_lineedit(self.ui.lineEdit_reg1, 0x00ac0837)
        reg2_val = self._parse_hex_lineedit(self.ui.lineEdit_reg2, 0x0080010f)
        if reg1_val is None or reg2_val is None:
            return

        chips = self.chips_per_board
        data_reg1 = [reg1_val] * chips
        data_reg2 = [reg2_val] * chips

        self.log(f"发送自定义寄存器: Reg1=0x{reg1_val:08x}, Reg2=0x{reg2_val:08x} (每层 {chips} 芯片)", "info")

        for board_id in online_boards:
            # 关键：使用原始命令名！
            cmd1 = self.make_command("02-Reg02", data_reg1, board_id)
            cmd2 = self.make_command("02-Reg03", data_reg2, board_id)

            if cmd1: self._send_one_command(cmd1, board_id, "自定义 Reg1")
            if cmd2: self._send_one_command(cmd2, board_id, "自定义 Reg2")

        self.log("自定义 Reg1+Reg2 发送完成", "info")

    def _send_one_command(self, cmd: bytes, board_id: int, name: str):
        """内部统一发送+接收+日志"""
        try:
            sended = self.send_to_usb(cmd)
            if sended < 0:
                self.log(f"{name} 发送失败 (板卡 B-0b{board_id:08b})", "error")
                return
            self.log(f">> {name} → B-0b{board_id:08b} ({sended} 字节): {' '.join(f'{b:02x}' for b in cmd)}", "send")

            readed = self.read_from_usb(1000)
            if readed != 4:
                self.log(f"{name} 响应异常 (readed={readed})", "warning")
                return
            resp = self.recv[:readed]
            self.log(f"<< {name} 响应: {' '.join(f'{b:02x}' for b in resp)}", "receive")
            time.sleep(0.1)
        except Exception as e:
            self.log(f"{name} 发送异常: {e}", "error")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet("""
    * {
        font-family: 'Segoe UI Variable Text', sans-serif;
        font-size: 14px;
        color: #000000;
        background-color: #F3F3F3;
    }
    QPushButton {
        border: 1px solid #E5E5E5;
        padding: 8px 16px;
        background-color: white;
        color: #000000;
        border-radius: 6px;
    }
    QPushButton:hover {
        background-color: #F5F5F5;
        border-color: #CCCCCC;
    }
    QPushButton:pressed {
        background-color: #E5E5E5;
    }
    QTextEdit, QLineEdit {
        border: 1px solid #E5E5E5;
        padding: 8px;
        background-color: white;
        border-radius: 6px;
        selection-background-color: #0078D4;
    }
    QTextEdit:focus, QLineEdit:focus {
        border: 1px solid #0078D4;
    }
    /* 新增 QComboBox 样式（与原有风格统一） */
    QComboBox {
        border: 1px solid #E5E5E5;
        padding: 8px;
        border-radius: 6px;
        background-color: white;
        color: #000000;
    }
    """)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
