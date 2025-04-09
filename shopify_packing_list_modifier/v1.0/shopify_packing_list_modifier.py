import pdf_editor
import os
import re
from datetime import datetime
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader, PdfWriter, Transformation

def find_latest_pdf():
    pdf_files = [f for f in os.listdir() if f.lower().endswith('.pdf')]
    if not pdf_files:
        return None
    latest = max(pdf_files, key=lambda f: os.path.getctime(f))
    return latest

def select_pdf_interactive():
    latest = find_latest_pdf()
    if latest:
        confirm = input(f"检测到最新PDF文件: {latest} 是否使用？(Y/n): ")
        if confirm.strip().lower() in ('', 'y'):
            return latest
    
    pdf_list = [f for f in os.listdir() if f.lower().endswith('.pdf')]
    if not pdf_list:
        return None
    
    print("\n当前目录PDF文件：")
    for i, f in enumerate(pdf_list):
        print(f"[{i}] {f}")
    
    while True:
        try:
            choice = int(input("请选择文件编号: "))
            return pdf_list[choice]
        except (ValueError, IndexError):
            print("无效的编号，请重新输入")

def save_pdf_preview(editor, output_txt: str):
    try:
        preview_content = editor.show_text_preview()
        
        with open(output_txt, 'w', encoding='utf-8') as f:
            f.write('=== PDF文本块预览 ===\n')
            f.write(preview_content)
        
        print(f"预览内容已保存至: {os.path.abspath(output_txt)}")
        return True
    except Exception as e:
        print(f"处理失败: {str(e)}")
        return False

def extract_unique_name(text):
    seen = set()
    parts = text.split()
    unique_parts = []
    for part in parts:
        # 统一转为小写后比较去重，但保留原始大小写格式
        lower_part = part.lower()
        if lower_part not in seen:
            seen.add(lower_part)
            # 首字母大写处理
            unique_parts.append(part.lower().capitalize())
    return ' '.join(unique_parts)

def extract_order_number(blocks):
    """从文本块中提取订单号"""
    for block in blocks:
        match = re.search(r'Order #(\d+)', block['text'])
        if match:
            return match.group(1)
    return datetime.now().strftime('%Y%m%d%H%M')


def process_pdf_modifications(editor, output_file, blocks):
    modifications = []
    
    # 替换索引2为SHIP TO
    if len(blocks) > 2 and 'SHIP TO' in blocks[2]['text']:
        modifications.append({
            'type': 'replace',
            'page': blocks[2]['page'],
            'coordinates': blocks[2]['coordinates'],
            'new_text': 'SHIP TO',
            'offset': -10  # 新增垂直偏移量
        })
    
    # 处理客户姓名去重
    if len(blocks) > 3:
        unique_name = extract_unique_name(blocks[3]['text'])
        modifications.append({
            'type': 'replace',
            'page': blocks[3]['page'],
            'coordinates': blocks[3]['coordinates'],
            'new_text': unique_name,
            'offset': -10  # 新增更大偏移量
        })
    
    # 动态确定地址删除范围
    start_index = 4
    end_index = next((i for i, b in enumerate(blocks) if 'ITEMS QUANTITY' in b['text']), 7) - 1
    
    # 记录删除范围到preview
    with open(output_file, 'a', encoding='utf-8') as f:
        f.write(f'\n=== 删除地址块范围 ===\n{start_index}-{end_index}\n')
    
    # 删除地址区块
    for idx in range(start_index, end_index+1):
        if idx < len(blocks):
            modifications.append({
                'type': 'delete',
                'page': blocks[idx]['page'],
                'coordinates': blocks[idx]['coordinates']
            })
    
    # 删除倒数第三块（从preview.txt读取）
    try:
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            if '倒数第三块索引 ===' in content:
                index_section = content.split('倒数第三块索引 ===\n')
                last_third_index = int(index_section[-1].split('\n')[0].strip())
                if last_third_index < len(blocks):
                    modifications.append({
                        'type': 'delete',
                        'page': blocks[last_third_index]['page'],
                        'coordinates': blocks[last_third_index]['coordinates']
                    })
    except Exception as e:
        print(f'读取预览文件失败: {e}')
        return
        if last_third_index < len(blocks):
            modifications.append({
                'type': 'delete',
                'page': blocks[last_third_index]['page'],
                'coordinates': blocks[last_third_index]['coordinates']
            })
    
    # 创建输出目录并保存
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    order_num = extract_order_number(blocks)
    output_pdf = os.path.join(output_dir, f'{order_num}.pdf')
    
    editor.modify_pdf(modifications, output_pdf)
    print(f'\n修改后的PDF已保存至: {output_pdf}')

    return output_pdf

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

    # 使用输入的 PDF 文件名并添加 "_A5" 后缀
    base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
    pdf_name = f"{base_name}_A5.pdf"
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
    input_pdf = select_pdf_interactive()
    if not input_pdf:
        print("未找到可用的PDF文件")
        exit()
    
    editor = pdf_editor.PDFEditor(input_pdf)
    output_file = "preview.txt"
    print(f"即将生成预览文件: {output_file}")
    blocks = editor.extract_text_blocks()
    save_success = save_pdf_preview(editor, output_file)
    if not save_success:
        exit()
    
    # 提取并保存姓名信息
    if len(blocks) > 3:
        name_text = blocks[3]['text']
        unique_name = extract_unique_name(name_text)
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(f'\n\n=== 客户姓名 ===\n{unique_name}')
        # 提取并保存倒数第三个区块信息
        if len(blocks) >= 3:
            third_last_index = len(blocks) - 3
            with open(output_file, 'a', encoding='utf-8') as f:
                f.write(f'\n\n=== 倒数第三块索引 ===\n{third_last_index}')
    
    # 在main流程末尾添加处理
    A4_pdf = process_pdf_modifications(editor, output_file, blocks)

    split_a4_to_a5_vertical(A4_pdf)