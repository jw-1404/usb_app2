#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
VSCode 直接运行版：解析 .dat 文件 → 画 逐通道 Mean/RMS 柱状图 + 热图
横坐标 = 通道编号（每个通道一个柱子），纵坐标 = Mean / RMS
"""

# ==============================
# 1. 导入模块
# ==============================
import struct                     # 解析二进制数据（如 unpack）
import numpy as np                # 数值计算：mean, var, sqrt
import matplotlib.pyplot as plt   # 绘图核心
from typing import List, Dict, Tuple
from dataclasses import dataclass # 数据类，简洁定义结构体
from pathlib import Path          # 路径操作，跨平台
import webbrowser                 # 自动打开生成的图片

# ==============================
# 2. 配置区（只需改这里！）
# ==============================
# 输入：你的 .dat 文件路径（FX3 设备采集的原始数据）
DAT_FILE = r"D:\AAA_work\python\usbpy3.5_stop排空分开\data\output_baseline_20251104_175226.dat"
# 输出：所有图表保存到这个文件夹（会自动创建）
OUTPUT_DIR = r"D:\AAA_work\python\usbpy3.5_stop排空分开\baseline_result1\高压4转1_三层1"

# ==============================

# ==============================
def encode_id(fec_index: int, chip_id: int, channel_id: int) -> int:
    """
    将 (FEC, Chip, Channel) 编码为一个 32 位整数
    格式：FEC(8bit) << 16 | Chip(8bit) << 8 | Channel(8bit)
    例：FEC=0, Chip=5, Channel=35 → GID = 0x00050123
    """
    return (fec_index << 16) | (chip_id << 8) | channel_id

# ==============================
# 4. 数据结构：每一帧的数据
# ==============================
@dataclass
class Entry:
    trigger_id: int          # 触发 ID
    time_stamp: int          # 时间戳
    sync_id: int             # 同步 ID（头）
    sync_id_body: int        # 同步 ID（体）
    fec_index: int           # FEC 板卡编号 (0~7)
    ids: List[int]           # 所有通道的 GID 列表
    adcs: List[int]          # 对应通道的 ADC 值列表

# ==============================
# 5. 解析器类：读取 .dat 文件并解析每一帧
# ==============================
class DatParser:
    def unpack_dat(self, file_path: str) -> List[Entry]:
        """
        按 FX3 固件协议解析 .dat 文件
        协议：大端字节序，固定 Tag 标记
        返回：所有有效帧的 Entry 列表
        """
        # --- 读取整个文件到内存 ---
        try:
            with open(file_path, 'rb') as f:
                data = f.read()
        except Exception as e:
            print(f"读取文件失败: {e}")
            return []

        entries = []                    # 存储解析出的每一帧
        pos = 0                         # 当前解析位置（字节）
        total_size = len(data)
        print(f"开始解析文件，总大小: {total_size:,} 字节")

        # --- 协议常量定义 ---
        HEAD_SIZE = 12                                      # 帧头 12 字节
        BODY_START_TAG = 0xac0f                             # 数据体开始标记
        BODY_END_TAG = 0xfffeaaaa                           # 数据体结束标记
        CHIP_START_TAG = 0xfa                               # 每个 Chip 数据开始
        CHIP_CHANNEL_NUMBER = 128                           # 每个 Chip 有 128 通道
        CHIP_END0_MASK = 0x81000000                         # Chip 结束字1 掩码
        CHIP_END1_MASK = 0x82000000                         # Chip 结束字2 掩码
        CHIP_END2_MASK = 0x83000000                         # Chip 结束字3 掩码
        TAIL_START_TAG = 0xffcc0000                         # 尾部开始标记
        TAIL_SIZE = 12                                      # 尾部 12 字节

        # --- 主解析循环：滑动窗口查找帧头 ---
        while pos + HEAD_SIZE < total_size:
            # 1. 查找帧头：0xffab530b 或 0xffab530f
            head_start_tag = struct.unpack_from('>I', data, pos)[0]
            if head_start_tag not in [0xffab530b, 0xffab530f]:
                pos += 1
                continue

            # 2. 读取保留字段和 FEC 索引
            reserved = data[pos + 6]       # 保留字节，固定为 0xd4
            fec_index = data[pos + 7]       # FEC 编号 (0~7)
            sync_id = struct.unpack_from('>I', data, pos + 8)[0]  # 同步 ID

            # 校验：reserved 必须是 0xd4，fec_index < 8
            if reserved != 0xd4 or fec_index >= 0x08:
                pos += 1
                continue

            pos += HEAD_SIZE                # 跳过帧头

            # 3. 检查数据体开始标记 0xac0f
            if pos + 20 > total_size: break
            body_start_tag = struct.unpack_from('>H', data, pos)[0]
            if body_start_tag != BODY_START_TAG:
                pos += 1 - HEAD_SIZE        # 回退，重新查找
                continue

            # 4. 读取数据体头部信息
            frame_length = struct.unpack_from('>H', data, pos + 2)[0]   # 帧长
            trigger_id = struct.unpack_from('>I', data, pos + 4)[0]     # 触发 ID
            body_sync_id = struct.unpack_from('>I', data, pos + 8)[0]  # 体同步 ID
            byte_length = struct.unpack_from('>I', data, pos + 12)[0]   # 数据字节数

            # 校验：帧长与数据长度是否匹配
            if byte_length / 4 + 5 != frame_length:
                pos += 1 - HEAD_SIZE
                continue

            pos += 16                       # 跳过数据体头部

            # 5. 解析多个 Chip 数据
            ids = []
            adcs = []
            chip_count = 0
            max_chips = 24                  # 最多 24 个 Chip

            while chip_count < max_chips and pos + 8 < total_size:
                if data[pos] != CHIP_START_TAG: break
                chip_id = data[pos + 1]     # Chip ID
                pos += 8                    # 跳过 Chip 头

                # 解析 128 个通道
                for _ in range(CHIP_CHANNEL_NUMBER):
                    if pos + 4 > total_size: break
                    value = struct.unpack_from('>I', data, pos)[0]
                    channel_id = (value >> 24) & 0xFF   # 高8位
                    adc = value & 0x00FFFFFF           # 低24位
                    if channel_id > 0x80: break        # 无效通道
                    gid = encode_id(fec_index, chip_id, channel_id)
                    ids.append(gid)
                    adcs.append(adc)
                    pos += 4

                # 校验 Chip 结束字（3个）
                if pos + 12 > total_size: break
                end0 = struct.unpack_from('>I', data, pos)[0]
                end1 = struct.unpack_from('>I', data, pos + 4)[0]
                end2 = struct.unpack_from('>I', data, pos + 8)[0]
                if (end0 & 0xFF000000) != CHIP_END0_MASK or \
                   (end1 & 0xFF000000) != CHIP_END1_MASK or \
                   (end2 & 0xFF000000) != CHIP_END2_MASK:
                    break
                pos += 12
                chip_count += 1

            # 6. 检查数据体结束标记
            body_end_tag = struct.unpack_from('>I', data, pos)[0]
            if body_end_tag != BODY_END_TAG:
                pos += 1 - HEAD_SIZE - 16
                continue
            pos += 4

            # 7. 跳过填充的 0x00
            while pos < total_size and data[pos] == 0x00:
                pos += 1

            # 8. 读取尾部（时间戳）
            if pos + TAIL_SIZE > total_size: break
            time_stamp = struct.unpack_from('>I', data, pos)[0]
            tail_start_tag = struct.unpack_from('>I', data, pos + 4)[0]
            if tail_start_tag != TAIL_START_TAG:
                pos += 1 - HEAD_SIZE - 16 - 4
                continue
            pos += TAIL_SIZE

            # 9. 保存完整帧
            if adcs:
                entry = Entry(trigger_id, time_stamp, sync_id, body_sync_id, fec_index, ids, adcs)
                entries.append(entry)

        print(f"解析完成，找到 {len(entries)} 帧")
        return entries

    def ana_baseline(self, entries: List[Entry]) -> Dict[int, Tuple[float, float]]:

        """
        统计每个通道的 Mean 和 Variance
        返回：{gid: (mean, var)}
        """
        baseline_map = {}  # gid → [adc1, adc2, ...]
        for entry in entries:
            for i, gid in enumerate(entry.ids):
                baseline_map.setdefault(gid, []).append(entry.adcs[i])

        ms_map = {}
        for gid, lst in baseline_map.items():
            if lst:
                arr = np.array(lst, dtype=np.float64)
                ms_map[gid] = (np.mean(arr), np.var(arr))  # (均值, 方差)
        return ms_map

# ==============================
# 6. 绘图函数：分层柱状图 + 全局热图
# ==============================
def plot_and_save(ms_map, out_dir):
    """
    1. 每 6 个 Chip 为一层 → 画柱状图（Mean）+ 折线（RMS）
    2. 全局热图：所有 FEC-Chip-Channel 的 Mean
    """
    Path(out_dir).mkdir(parents=True, exist_ok=True)

    # --- 1. 按 (FEC, Chip) 分组 ---
    chip_groups = {}
    for gid, (m, v) in ms_map.items():
        fec = (gid >> 16) & 0xFF
        chip = (gid >> 8) & 0xFF
        ch = gid & 0xFF
        key = (fec, chip)
        if key not in chip_groups:
            chip_groups[key] = {}
        chip_groups[key][ch] = (m, np.sqrt(v))  # RMS = sqrt(var)

    sorted_keys = sorted(chip_groups.keys())
    total_chips = len(sorted_keys)
    print(f"检测到 {total_chips} 个 Chip，共 {total_chips//6} 层")

    # --- 2. 每 6 个 Chip 一层，画一张图 ---
    CHIPS_PER_LAYER = 6
    num_layers = (total_chips + CHIPS_PER_LAYER - 1) // CHIPS_PER_LAYER

    for layer in range(num_layers):
        start = layer * CHIPS_PER_LAYER
        end = min(start + CHIPS_PER_LAYER, total_chips)
        current_keys = sorted_keys[start:end]

        # 创建 2×3 子图布局
        fig, axes = plt.subplots(2, 3, figsize=(18, 10), sharex=True)
        axes = axes.ravel()

        for idx, (fec, chip) in enumerate(current_keys):
            ax = axes[idx]
            data = chip_groups[(fec, chip)]

            channels = list(range(128))
            means = [data.get(ch, (np.nan, np.nan))[0] for ch in channels]
            rms = [data.get(ch, (np.nan, np.nan))[1] for ch in channels]

            # 柱状图：Mean
            ax.bar(channels, means, width=1.0, color='#4C72B0', alpha=0.7, label='Mean')
            ax.set_ylim(10000, 20000)   #  修改 Mean 轴范围
            # 折线图：RMS（双Y轴）
            ax2 = ax.twinx()
            ax2.plot(channels, rms, color='#DD8452', linewidth=2, marker='o', markersize=2, label='RMS')

            ax.set_title(f'Chip {chip}', fontsize=12)
            ax.set_xlabel('Channel (0–127)')
            if idx % 3 == 0:
                ax.set_ylabel('Mean ADC', color='#4C72B0')
            ax2.set_ylabel('RMS ADC', color='#DD8452')
            ax2.set_ylim(0, 2000)  # RMS 上限可调

            ax.legend(loc='upper left', fontsize=9)
            ax2.legend(loc='upper right', fontsize=9)
            ax.grid(True, axis='y', ls='--', alpha=0.5)

        # 隐藏空的子图
        for idx in range(len(current_keys), 6):
            axes[idx].set_visible(False)

        plt.suptitle(f'Baseline Layer {layer + 1}/{num_layers} (6 Chips, 768 Channels)', 
                     fontsize=16, y=0.98)
        plt.tight_layout(rect=[0, 0, 1, 0.96])

        fig_path = Path(out_dir) / f"baseline_layer_{layer + 1}.png"
        plt.savefig(fig_path, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"第 {layer + 1} 层图已保存：{fig_path}")

    # --- 3. 全局热图：所有通道 Mean ---
    MAX_FEC, MAX_CHIP, CH = 8, 24, 128
    grid = np.full((MAX_FEC, MAX_CHIP, CH), np.nan)
    for gid, (m, _) in ms_map.items():
        fec = (gid >> 16) & 0xFF
        chip = (gid >> 8) & 0xFF
        ch = gid & 0xFF
        if fec < MAX_FEC and chip < MAX_CHIP and ch < CH:
            grid[fec, chip, ch] = m

    fig, axes = plt.subplots(2, 4, figsize=(24, 10), sharex=True, sharey=True)
    axes = axes.ravel()
    vmin, vmax = np.nanpercentile(grid, [2, 98])  # 去极值
    for i in range(MAX_FEC):
        im = axes[i].imshow(grid[i], cmap='plasma', vmin=vmin, vmax=vmax,
                            aspect='auto', interpolation='nearest')
        axes[i].set_title(f'FEC {i}', fontsize=14)
        axes[i].set_xlabel('Channel (0–127)')
        if i % 4 == 0:
            axes[i].set_ylabel('Chip (0–23)')

    for i in range(MAX_FEC, 8):
        axes[i].set_visible(False)

    cbar = fig.colorbar(im, ax=axes[:MAX_FEC].tolist(), shrink=0.8, pad=0.02)
    cbar.set_label('Mean ADC Value', rotation=270, labelpad=15)
    plt.suptitle('Baseline Mean Heatmap', fontsize=18, y=0.98)
    heat_path = Path(out_dir) / "baseline_mean_heatmap.png"
    plt.tight_layout(rect=[0, 0, 1, 0.95])
    plt.savefig(heat_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"热图已保存：{heat_path}")

    # --- 4. 自动打开所有图片 ---
    all_paths = [Path(out_dir) / f"baseline_layer_{i+1}.png" for i in range(num_layers)]
    all_paths.append(heat_path)
    for p in all_paths:
        if p.exists():
            webbrowser.open(f"file://{p.resolve()}")
    print(f"\n共生成 {num_layers} 张层图 + 1 张热图，已自动打开！")
def run_analysis(dat_file: str, output_dir: str = None):
    """
    供 GUI 调用的主函数
    - dat_file: .dat 文件路径
    - output_dir: 输出目录（可选）
    """
    dat_path = Path(dat_file)
    if not dat_path.exists():
        raise FileNotFoundError(f"文件不存在: {dat_file}")

    # 自动设置输出目录
    if output_dir is None:
        output_dir = dat_path.parent / f"{dat_path.stem}_result"
    else:
        output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 解析
    parser = DatParser()
    entries = parser.unpack_dat(str(dat_path))
    if not entries:
        raise ValueError("未解析到任何数据")

    # 统计
    ms_map = parser.ana_baseline(entries)

    # 绘图并保存
    plot_and_save(ms_map, str(output_dir))

    print(f"[GUI] 分析完成！结果保存至: {output_dir}")
    return str(output_dir)


# ==============================
# 7. 主函数（VSCode 运行入口）
# ==============================
if __name__ == "__main__":
    dat_path = Path(DAT_FILE)
    if not dat_path.exists():
        print(f"错误：文件不存在！\n{DAT_FILE}")
        input("按回车退出...")
        exit()

    parser = DatParser()
    entries = parser.unpack_dat(str(dat_path))
    if not entries:
        print("没有解析到数据！")
        input("按回车退出...")
        exit()

    ms_map = parser.ana_baseline(entries)
    print(f"统计完成：{len(ms_map)} 个通道")
    plot_and_save(ms_map, OUTPUT_DIR)
    print("全部完成！")
    input("\n按回车退出...")