# Python工具集说明

## 1. modify_ramdisk_end_point.py

### 功能说明
用于修改DTS文件中的linux,initrd-end值，自动计算根文件系统镜像大小并更新相关配置

### 使用方式
```bash
python modify_ramdisk_end_point.py <ramdisk镜像路径> <dts文件路径>
```

### 参数说明
- ramdisk_path: 根文件系统镜像路径
- dts_path: 需要修改的DTS文件路径

---

## 2. pdf_editor.py

### 功能说明
PDF文档编辑器，提供以下功能：
1. 提取带坐标的文本块
2. 支持文本替换和删除操作
3. 生成修改后的PDF文档

### 依赖
```bash
pip install pymupdf pdfplumber
```

### 使用方式
```bash
python pdf_editor.py
```

---

## 3. shopify_packing_list_modify.py

### 功能说明
Shopify发货单自动化处理工具，支持：
1. 自动检测最新PDF文件
2. 交互式文件选择
3. SHIP TO信息替换
4. 客户姓名去重处理
5. 地址区块自动删除

### 依赖关系
```bash
# 依赖pdf_editor.py模块
# 无需单独安装依赖
```

### 使用方式
```bash
python shopify_packing_list_modifier.py
```

> 另：可使用`pyinstaller --onefile xxx.py`的方式将其转化为一个exe文件
