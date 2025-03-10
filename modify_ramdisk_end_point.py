import argparse
import os
import re

def get_file_size(file_path):
    """
    获取文件大小（以字节为单位）。
    如果文件不存在，提示用户。
    """
    if not os.path.exists(file_path):
        print(f"文件 {file_path} 不存在。")
        return None

    size = os.path.getsize(file_path)
    return size

def modify_initrd_values(file_path, rootfs_size_decimal):
    """
    修改 DTS 文件中的 linux,initrd-end 的值。
    """
    # 正则表达式匹配 linux,initrd-start 和 linux,initrd-end
    initrd_pattern = r"linux,initrd-(start|end)\s*=\s*<\s*(0x[0-9a-fA-F]+)\s+(0x[0-9a-fA-F]+)\s*>;"

    with open(file_path, 'r') as file:
        content = file.read()

    # 查找 linux,initrd-start 的值
    match_start = re.search(initrd_pattern.replace("(start|end)", "start"), content, re.MULTILINE)
    if not match_start:
        raise ValueError("linux,initrd-start not found in the file.")

    # 调试输出匹配内容
    print("Matched start:", match_start.groups())

    # 提取高位和低位，并拼接为 64 位地址
    try:
        initrd_start_high = int(match_start.group(1), 16)
        initrd_start_low = int(match_start.group(2), 16)
    except (IndexError, ValueError) as e:
        raise ValueError("Failed to parse linux,initrd-start values: " + str(e))

    initrd_start_64 = (initrd_start_high << 32) | initrd_start_low

    # 计算新的 linux,initrd-end 的 64 位地址
    initrd_end_64 = initrd_start_64 + rootfs_size_decimal

    # 拆分成高位和低位
    initrd_end_high, initrd_end_low = split_64bit_to_high_low(initrd_end_64)

    # 替换 linux,initrd-end 的值
    def replace_end(match):
        return f"linux,initrd-end = < {initrd_end_high} {initrd_end_low} >;"

    content = re.sub(initrd_pattern.replace("(start|end)", "end"), replace_end, content, flags=re.MULTILINE)

    # 将修改后的内容写回文件
    with open(file_path, 'w') as file:
        file.write(content)


def split_64bit_to_high_low(value):
    """
    将一个 64 位整数拆分为高位和低位 32 位。

    参数：
        value (int): 64 位整数。

    返回：
        tuple: 高位和低位（十六进制字符串）。
    """
    high = (value >> 32) & 0xFFFFFFFF
    low = value & 0xFFFFFFFF
    return f"0x{high:08x}", f"0x{low:08x}"
def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description="修改 DTS 文件的 initrd 值。")
    parser.add_argument("ramdisk_path", type=str, help="用于获取大小的根文件系统镜像路径")
    parser.add_argument("dts_path", type=str, help="需要修改的 DTS 文件路径")
    args = parser.parse_args()

    # 获取文件大小
    try:
        ramdisk_size = get_file_size(args.ramdisk_path)
        if ramdisk_size is None:
            return
        print(f"文件 {args.ramdisk_path} 的大小为 {ramdisk_size:#x} 字节。")
    except FileNotFoundError as e:
        print(e)
        return

    # 修改 DTS 文件中的 initrd 值
    modify_initrd_values(args.dts_path, ramdisk_size)

if __name__ == "__main__":
    main()
