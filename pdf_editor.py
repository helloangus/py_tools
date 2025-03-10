import os
import pdfplumber
import fitz
from typing import List, Dict
import re

class PDFEditor:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.doc = fitz.open(file_path)
        self.text_blocks = []

    def extract_text_blocks(self) -> List[Dict]:
        """提取带坐标的文本块并按页分组"""
        with pdfplumber.open(self.file_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                words = page.extract_words(keep_blank_chars=True)
                current_block = []
                for word in words:
                    if current_block and (word['top'] - current_block[-1]['top'] > 5):
                        self._add_block(page_num, current_block)
                        current_block = []
                    current_block.append(word)
                if current_block:
                    self._add_block(page_num, current_block)
        return self.text_blocks

    def _add_block(self, page_num: int, words: list):
        """构建文本块数据结构"""
        x0 = min(word['x0'] for word in words)
        top = min(word['top'] for word in words)
        x1 = max(word['x1'] for word in words)
        bottom = max(word['bottom'] for word in words)
        
        self.text_blocks.append({
            'page': page_num,
            'text': ' '.join(word['text'] for word in words),
            'coordinates': (x0, top, x1, bottom),
            'words': words
        })

    def modify_pdf(self, modifications: List[Dict], output_path: str):
        """执行PDF修改操作"""
        for mod in modifications:
            page = self.doc.load_page(mod['page'])
            
            # 添加白色覆盖层
            rect = mod['coordinates']
            page.draw_rect(rect, color=(1,1,1), fill=(1,1,1))
            
            # 如果是替换操作则添加新文本
            if mod['type'] == 'replace':
                page.insert_text(
                    point=(rect[0], rect[1] - mod.get('offset', 0)),
                    text=mod['new_text'],
                    fontsize=12,
                    color=(0,0,0)
                )
        
        if os.path.exists(output_path):
            os.remove(output_path)
        self.doc.save(output_path)
        self.doc.close()

    def show_text_preview(self):
        """生成带页码的文本预览"""
        preview = []
        current_page = -1
        for i, block in enumerate(self.text_blocks):
            if block['page'] != current_page:
                preview.append(f"\n=== Page {block['page']+1} ===")
                current_page = block['page']
            preview.append(f"[{i}] {block['text']}")
        return '\n'.join(preview)


def process_pdf(input_path: str, output_path: str):
    editor = PDFEditor(input_path)
    blocks = editor.extract_text_blocks()
    modifications = []
    
    while True:
        print("=== PDF文本块预览 ===")
        print(editor.show_text_preview())
        
        selections = input("请输入要修改的区块编号（多个用逗号分隔，输入q退出）: ").strip()
        if selections.lower() == 'q':
            break
        
        selections = selections.split(',')
        
        for idx in selections:
            idx = idx.strip()
            if not idx:
                continue
            try:
                block = blocks[int(idx)]
            except (ValueError, IndexError):
                print(f"无效的区块编号: {idx}")
                continue
            
            print(f"\n当前选择区块: {block['text']}")
            action = input("操作类型（delete/replace）: ").lower()
            
            if action == 'replace':
                new_text = input("输入替换文本: ")
                modifications.append({
                    'type': 'replace',
                    'page': block['page'],
                    'coordinates': block['coordinates'],
                    'new_text': new_text
                })
            elif action == 'delete':
                modifications.append({
                    'type': 'delete',
                    'page': block['page'],
                    'coordinates': block['coordinates']
                })
            else:
                print("无效的操作类型，已跳过。")
        
        continue_edit = input("是否继续修改其他区域？(y/n): ").lower()
        if continue_edit != 'y':
            break
    
    if modifications:
        editor.modify_pdf(modifications, output_path)
        print(f"\n修改后的PDF已保存至: {output_path}")
    else:
        print("\n未做任何修改。")


if __name__ == "__main__":
    input_pdf = input("输入PDF文件路径: ")
    output_pdf = input("输入保存路径: ")
    process_pdf(input_pdf, output_pdf)