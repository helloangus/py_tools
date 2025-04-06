import os
import time
import pdfplumber
import re
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, A5
from reportlab.lib.units import mm
from decimal import Decimal
from pypdf import PdfReader, PdfWriter, Transformation

order_blk_id = 1
customer_blk_id = 4
item_blk_id = -1
last_third_id = -1
order_number = -1

def select_pdf_file():
    """
    获取并处理当前目录下的PDF文件。

    该函数会扫描当前目录下的所有PDF文件，并按修改时间从新到旧排序。用户可以选择最新的文件或从列表中选择其他文件。

    返回值:
        str: 用户选择的PDF文件路径。如果未找到PDF文件或用户取消选择，则返回None。
    """
    
    # 获取当前目录所有PDF文件（不区分大小写）
    pdf_files = [f for f in os.listdir() if f.lower().endswith('.pdf')]
    
    # 无PDF文件处理
    if not pdf_files:
        print("错误：当前目录下未找到PDF文件")
        return None
    
    # 按修改时间从新到旧排序（完整排序版）
    pdf_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    
    # 尝试选择最新文件
    latest_file = pdf_files[0]
    user_choice = input(f"发现最新PDF文件：{latest_file}\n直接按Enter选择该文件，输入n查看全部列表：").strip().lower()
    
    # 直接选择最新文件
    if user_choice in ('', 'y', 'yes'):
        return latest_file
    
    # 显示完整文件列表
    print("\n当前目录PDF文件列表：")
    for index, file in enumerate(pdf_files, 1):
        # 获取人类可读的修改时间
        mtime = os.path.getmtime(file)
        formatted_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
        print(f"[{index}] {file}（最后修改：{formatted_time}）")
    
    # 用户选择验证循环
    while True:
        selection = input(f"请输入文件编号（1-{len(pdf_files)}）：").strip()
        
        # 处理空输入（默认选择最新）
        if not selection:
            print("已自动选择最新文件")
            return latest_file
        
        # 有效性验证
        try:
            index = int(selection)
            if 1 <= index <= len(pdf_files):
                return pdf_files[index-1]
            print("错误：编号超出范围")
        except ValueError:
            print("错误：请输入有效数字")

def write_metadata_to_file(metadata, output_dir, file_name):
    """
    将元数据写入指定目录的文件中。

    参数:
    metadata (list of dict): 包含元数据的字典列表。每个字典应包含以下键：
        - 'ID': 元数据的唯一标识符。
        - '页码': 元数据所在的页码。
        - '文本': 元数据中的文本内容。
        - '字体': 文本使用的字体。
        - '字号': 文本的字号（以pt为单位）。
        - '位置': 文本在页面中的位置。
        - '宽度': 文本的宽度。
    output_dir (str): 输出文件的目录路径。如果目录不存在，将尝试创建。
    file_name (str): 输出文件的名称（不包含扩展名）。

    返回值:
    None: 如果成功写入文件，则返回None。如果遇到权限错误，则打印错误信息并返回None。
    """
    try:
        # 尝试创建输出目录，如果目录已存在则忽略
        os.makedirs(output_dir, exist_ok=True)
    except PermissionError:
        # 如果无权限创建目录，打印错误信息并返回
        print(f"错误：无权限创建目录 {output_dir}")
        return None

    # 构建输出文件的完整路径
    output_path = f"{output_dir}/{file_name}.txt"

    # 打开文件并写入元数据
    with open(output_path, 'w', encoding='utf-8') as f:
        for meta in metadata:
            # 将每个元数据项格式化为字符串并写入文件
            f.write(
                f"[ID: {meta['ID']}]\n"
                f"· 页码：{meta['页码']}\n"
                f"· 文本：{meta['文本']}\n"
                f"· 字体：{meta['字体']} ({meta['字号']}pt)\n"
                f"· 位置：{meta['位置']}\n"
                f"· 宽度：{meta['宽度']}\n\n"
            )

    # 打印成功信息
    print(f"新元数据文件已生成：{output_path}")
def extract_text_with_metadata(pdf_path):
    """
    从指定的PDF文件中提取文本及其元数据，并为每个文本对象生成唯一编号。

    参数:
    - pdf_path (str): PDF文件的路径。

    返回值:
    - list: 包含每个文本对象元数据的列表。每个元数据包括唯一ID、页码、文本内容、字体、字号、位置和宽度等信息。
           如果提取失败，则返回None。
    """

    try:
        # 使用pdfplumber打开PDF文件
        with pdfplumber.open(pdf_path) as pdf:      
            # 全局唯一ID计数器，初始化为-1，每次提取文本时自增
            text_id = -1
            # 用于存储所有文本对象的元数据
            metadata = []
            
            # 遍历PDF的每一页
            for page_num, page in enumerate(pdf.pages, 1):
                # 提取当前页的文本行对象
                text_objects = page.extract_text_lines(
                    x_tolerance=1,
                    y_tolerance=1,
                    keep_blank_chars=False,
                    use_text_flow=True,
                    split_at_punctuation=False
                )
                
                # 遍历当前页的每个文本对象
                for obj in text_objects:
                    text_id += 1  # 全局自增ID
                    
                    # 获取字符级属性，如字体、字号等
                    font_info = obj.get('chars', [{}])[0] if obj.get('chars') else {}
                    
                    # 构建当前文本对象的元数据字典
                    meta = {
                        "ID": text_id,  # 新增唯一ID
                        "页码": page_num,
                        "文本": obj.get('text', ''),
                        "字体": font_info.get('fontname', '未知字体'),
                        "字号": round(font_info.get('size', 0), 1),
                        "位置": (round(obj.get('x0', 0)), round(obj.get('top', 0), 1)),
                        "宽度": round(obj.get('width', 0), 1)
                    }
                    # 将元数据添加到列表中
                    metadata.append(meta)

        # 从元数据中提取订单号，并写入文件
        global order_number
        order_number = metadata[order_blk_id]['文本'].split('#')[-1] 
        output_dir = 'tmp'
        file_name = f"{order_number}_extracted"
        write_metadata_to_file(metadata, output_dir, file_name)
        print(f"写入元数据文件成功：{output_dir}/{file_name}.txt")
        return metadata

    except Exception as e:
        # 捕获并打印异常信息
        print(f"提取失败：{e.__class__.__name__}: {str(e)}")
    return None
def analyze_metadata(metadata):
    """
    元数据解析函数，用于从给定的元数据中提取订单信息并确定需要删除的内容范围。

    参数:
    metadata (list): 包含订单信息的元数据列表，每个元素为一个字典，包含'文本'和'位置'等键。

    返回值:
    list: 包含需要删除的元数据块ID的列表。如果无法确定删除范围，则返回None。

    异常:
    KeyError: 如果元数据中缺少必要的键，将抛出KeyError并提示缺失的键。
    """
    try:
        print("\n=== 订单信息提取 ===")
        # 提取订单编号和客户姓名
        print(f"[ID{order_blk_id}] 订单编号：{metadata[order_blk_id]['文本']}")
        print(f"[ID{customer_blk_id}] 客户姓名：{metadata[customer_blk_id]['文本']}")
        
        # 初始化待删除的块ID列表
        delete_ids = []
        
        # 查找ITEMS QUANTITY块作为关键锚点
        global item_blk_id
        for item_id, data in enumerate(metadata):
            if data.get('文本', '').strip() == 'ITEMS QUANTITY':
                item_blk_id = item_id
                break
        
        if item_blk_id:
            # 确定地址段结束位置，并添加地址段删除范围
            end_id = item_blk_id
            if end_id >=4:
                delete_ids.extend(list(range(5, end_id)))  # 包含end_id
            else:
                print(f"警告：ITEMS QUANTITY位置异常（ID{item_blk_id}），跳过地址段删除")
        else:
            print("警告：未找到ITEMS QUANTITY块，无法确定地址段")
            return None

        # 检测倒数第三块并添加到待删除列表
        global last_third_id
        if len(metadata) >=3:
            last_third_id = len(metadata) - 3
            delete_ids.append(last_third_id)
        else:
            print("警告：总块数不足，无法获取倒数第三块")
        
        # 打印待删除内容的信息
        print("\n=== 待删除内容 ===")
        for did in delete_ids:
            if metadata[did]:
                print(f"[ID{did}] {metadata[did]['文本'][:30]}... (位置：{metadata[did]['位置']})")
            else:
                print(f"警告：ID{did} 不存在")
        return delete_ids

    except KeyError as e:
        print(f"关键数据缺失：ID{e.args[0]} 不存在")
def create_new_metadata(metadata, delete_ids):
    """
    生成新元数据文件，新增区域和国家信息块，并删除指定ID的块。

    参数:
    - metadata (list): 包含多个字典的列表，每个字典代表一个元数据块。
    - delete_ids (list): 需要删除的元数据块的ID列表。

    返回值:
    - list: 处理后的元数据列表，包含新增的区域和国家信息块，并删除了指定ID的块。
    """
    # 提取关键信息
    order_block = metadata[order_blk_id]
    customer_block = metadata[customer_blk_id]
    global item_blk_id
    items_block = metadata[item_blk_id]
    country_block = metadata[item_blk_id-1]

    # 验证必要信息
    if not all([order_block, customer_block, country_block, items_block]):
        print("关键信息块缺失")
        return None

    # 格式化客户姓名
    def format_name(name):
        return ' '.join([n.capitalize() for n in re.split(r'(\W+)', name) if n.strip()])
    
    if customer_block:
        # 直接修改原始块的text字段
        customer_block['文本'] = format_name(customer_block['文本'])

    if items_block:
        # 直接修改原始块的text字段
        items_block['文本'] = 'ITEMS & QUANTITY'
        new_x = items_block['位置'][0] - 5
        new_position = (new_x, items_block['位置'][1])  # 创建新元组
        items_block['位置'] = new_position  # 替换整个元组

    # 创建新块
    # 区域信息块（ID3右侧）
    id3_block = metadata[3]
    region_block = {
        'ID': -1,  # 临时标识
        '页码': id3_block['页码'],
        '文本': 'Region or Country',
        '字体': id3_block['字体'],
        '字号': id3_block['字号'],
        '位置': (id3_block['位置'][0] + 200, id3_block['位置'][1]),
        '宽度': id3_block['宽度']
    }

    # 国家信息块（ID4样式）
    country_content = country_block['文本']
    country_new_block = {
        'ID': -2,  # 临时标识
        '页码': customer_block['页码'],
        '文本': country_content,
        '字体': customer_block['字体'],
        '字号': customer_block['字号'],
        '位置': (region_block['位置'][0], region_block['位置'][1] + 20),
        '宽度': customer_block['宽度']
    }

    # 插入新块到ITEMS QUANTITY前
    metadata.insert(item_blk_id, region_block)
    metadata.insert(item_blk_id + 1, country_new_block)

    # 初始化一个空列表，用于存储过滤后的元数据
    filtered_blocks = []

    # 遍历 metadata 列表中的每个元素
    for b in metadata:
        # 获取当前元素的 'ID' 值，如果不存在则返回默认值 -999
        block_id = b.get('ID', -999)
        # 检查 'ID' 是否不在 delete_ids 列表中
        if block_id not in delete_ids:
            # 如果条件满足，将当前元素添加到 filtered_blocks 列表中
            filtered_blocks.append(b)

    # 重新编号时重置所有ID
    for idx, meta in enumerate(filtered_blocks):
        meta['ID'] = idx  # 强制赋予新ID

    metadata = filtered_blocks
    # 查找ITEM QUANTITY块（关键锚点）
    for item_id, data in enumerate(metadata):
        if data.get('文本', '').strip() == 'ITEMS & QUANTITY':
            item_blk_id = item_id
            break

    # 生成新文件
    output_dir = 'tmp'
    file_name = f"{order_number}_modified"
    write_metadata_to_file(metadata, output_dir, file_name)
    print(f"写入元数据文件成功：{output_dir}/{file_name}.txt")

    return metadata
def adjust_metadata_positions(metadata):
    """
    调整内容块位置并保留完整元数据。

    该函数通过查找指定的锚点（如客户块和项目块），计算位置偏移量，并应用该偏移量来调整元数据中所有相关块的位置。
    最后，生成并保存调整后的元数据文件。

    参数:
    metadata (dict): 包含所有内容块及其位置信息的元数据字典。

    返回值:
    dict: 调整位置后的元数据字典。

    异常:
    ValueError: 如果关键锚点（如客户块或项目块）缺失，则抛出此异常。
    """
    # 查找定位锚点
    customer_block = metadata[customer_blk_id]
    items_block = metadata[item_blk_id]

    if not customer_block or not items_block:
        raise ValueError("关键锚点缺失，请检查Jasmine Perry和ITEMS QUANTITY是否存在")

    # 计算位置偏移量
    original_y = items_block['位置'][1]
    target_y = customer_block['位置'][1] + 40
    y_offset = target_y - original_y

    # 应用位置调整
    for block in metadata[item_blk_id:]:
        x, y = block['位置']
        block['位置'] = (x, y + y_offset)

    # 生成完整元数据文件
    output_dir = 'tmp'
    file_name = f"{order_number}_adjusted"
    write_metadata_to_file(metadata, output_dir, file_name)
    print(f"写入元数据文件成功：{output_dir}/{file_name}.txt")
    return metadata
def generate_new_pdf(metadata):
    """
    根据元数据生成A4尺寸的PDF文件。

    参数:
    metadata (list)

    返回值:
    str: 生成的PDF文件路径。如果生成过程中出现错误，则返回None。
    """
    # 设置A4页面尺寸（595.27pt × 841.89pt）
    page_width, page_height = A4

    # 创建输出目录（保持不变）
    output_dir = 'tmp'
    try:
        os.makedirs(output_dir, exist_ok=True)
    except PermissionError:
        print(f"错误：无权限创建目录 {output_dir}")
        return None

    pdf_name = f"{order_number}.pdf"
    pdf_path = os.path.join(output_dir, pdf_name)    
    
    # 创建PDF画布，并应用A4尺寸
    c = canvas.Canvas(pdf_path, pagesize=(page_width, page_height))  # 修改点2：应用A4尺寸
    
    # 字体映射表（保持不变）
    FONT_MAP = {
        'NotoSans-Regular': 'Helvetica',
        'NotoSans-Bold': 'Helvetica-Bold',
    }
    
    # 遍历元数据中的每个文本块，并在PDF上绘制
    for block in metadata:
        x, y = block.get('位置', (0, 0))
        
        font_key = block.get('字体', 'Helvetica').split('+')[-1]
        font_name = FONT_MAP.get(font_key, 'Helvetica')
        font_size = block.get('字号', 10)
        
        c.setFont(font_name, font_size)
        c.drawString(x, page_height - y, block.get('文本', ''))  # 保持Y轴转换逻辑
    
    # 保存PDF文件
    c.save()
    print(f"新PDF生成成功：{pdf_path}")
    return pdf_path

def split_a4_to_a5_vertical(input_pdf_path):
    """
    将A4尺寸的PDF文件垂直分割为两个A5尺寸的PDF文件，并保存到指定目录。

    参数:
    input_pdf_path (str): 输入的A4尺寸PDF文件路径。

    返回值:
    bool: 如果分割成功返回True，否则返回False。
    """
    output_dir = 'output'
    try:
        os.makedirs(output_dir, exist_ok=True)
    except PermissionError:
        print(f"错误：无权限创建目录 {output_dir}")
        return None

    pdf_name = f"{order_number}.pdf"
    output_pdf_path = os.path.join(output_dir, pdf_name)

    try:
        os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)
        
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        # 定义标准尺寸（单位：点）
        A4_WIDTH, A4_HEIGHT = A4
        A5_HEIGHT = A4_HEIGHT / 2  # 420.945pt

        for page in reader.pages:
            # ==================================================================
            # 处理上半部分 (关键修正点)
            # ==================================================================
            # 1. 创建新A5页面
            top_page = writer.add_blank_page(width=A4_WIDTH, height=A5_HEIGHT)
            
            # 2. 克隆原始页面
            top_clone = page.clone(pdf_dest=writer)
            
            # 3. 设置裁剪区域
            top_clone.cropbox.upper_left = (0, A4_HEIGHT)        # 原页面上半部分
            top_clone.cropbox.lower_right = (A4_WIDTH, A5_HEIGHT)
            
            # 4. 变换方向修正（关键！）
            transform = Transformation().translate(
                tx=0, 
                ty=-(A4_HEIGHT - A5_HEIGHT)  # 正确计算平移量
            )
            
            # 5. 应用变换并合并
            top_page.merge_page(top_clone)
            top_page.add_transformation(transform)  # 先合并后变换

            # ==================================================================
            # 处理下半部分 (已验证正确)
            # ==================================================================
            bottom_page = writer.add_blank_page(width=A4_WIDTH, height=A5_HEIGHT)
            bottom_clone = page.clone(pdf_dest=writer)
            bottom_clone.cropbox.upper_left = (0, A5_HEIGHT)
            bottom_clone.cropbox.lower_right = (A4_WIDTH, 0)
            bottom_page.merge_page(bottom_clone)

        # 保存输出文件
        with open(output_pdf_path, "wb") as f:
            writer.write(f)
            
        print(f"生成成功：{output_pdf_path}")
        return True

    except Exception as e:
        print(f"处理失败：{str(e)}")
        return False


if __name__ == "__main__":
    # 选择PDF文件，返回所选文件的路径
    selected_pdf = select_pdf_file()
    
    if selected_pdf is None:
        print("错误：未选择PDF文件")
        exit(1)

    # 从选定的PDF文件中提取文本及其元数据，返回包含文本和元数据的对象
    metadata = extract_text_with_metadata(selected_pdf)
    
    # 分析提取的元数据，确定需要删除的目标内容，返回需要删除的目标列表
    delete_targets = analyze_metadata(metadata)
    
    # 根据删除目标创建新的元数据，返回更新后的元数据对象
    new_meta = create_new_metadata(metadata, delete_targets)
    
    # 调整新元数据的位置信息，返回调整后的元数据对象
    adj_meta = adjust_metadata_positions(new_meta)
    
    # 打印提示信息，表示已准备好生成新PDF的元数据文件
    print("已准备好生成新PDF的元数据文件")
    
    # 根据调整后的元数据生成新的A4尺寸PDF文件，返回生成文件的路径
    A4_pdf_path = generate_new_pdf(adj_meta)
    
    # 将生成的A4尺寸PDF文件垂直分割为A5尺寸
    split_a4_to_a5_vertical(A4_pdf_path)