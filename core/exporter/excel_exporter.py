"""
Excel导出器
支持将提取的数据导出为不同ERP格式的Excel文件
"""
import pandas as pd
import os
import logging
from typing import List, Dict
from io import BytesIO

from .erp_templates import get_template

logger = logging.getLogger(__name__)


class ExcelExporter:
    """Excel导出器"""

    # 源数据字段到通用语义的映射
    FIELD_MAPPING = {
        "条码": "条码",
        "品名": "品名",
        "规格": "规格",
        "单位": "单位",
        "进价": "进价",
        "零售价": "零售价"
    }

    # 通用语义到各模板字段的映射规则（关键词匹配）
    SEMANTIC_KEYWORDS = {
        "条码": ["货号", "条码", "条形码", "商品码", "编码", "自编码", "code", "barcode"],
        "品名": ["品名", "名称", "商品名", "商品名称", "产品名", "name", "product"],
        "规格": ["规格", "规格型号", "型号", "容量", "净含量", "spec"],
        "单位": ["单位", "计量单位", "包装", "unit"],
        "进价": ["进价", "进货价", "成本价", "采购价", "批发价", "进货单价", "cost", "price"],
        "零售价": ["零售价", "售价", "销售价", "销售单价", "零售金额", "retail", "sell"],
        "品牌": ["品牌", "品牌编码", "brand"],
        "产地": ["产地", "生产地", "origin"],
        "供应商编码": ["供应商编码", "主供应商编码", "supplier"],
        "类别编码": ["类别编码", "分类编码", "category"],
    }

    def export(self, data: List[Dict], manual_codes: Dict = None,
               template_name: str = 'default') -> str:
        """
        导出为Excel文件

        Args:
            data: 识别的数据列表
            manual_codes: 手工编码 {'supplier_code': '', 'category_code': ''}
            template_name: ERP模板名称

        Returns:
            导出文件的路径
        """
        if manual_codes is None:
            manual_codes = {}

        # 获取模板配置
        template = get_template(template_name)

        # 合并数据和编码
        merged_data = self._merge_data(data, manual_codes, template)

        # 按字段顺序导出
        df = pd.DataFrame(merged_data)
        # 确保所有模板字段都存在，不存在的填充空值
        df = df.reindex(columns=template['field_order'], fill_value='')

        # 保存到文件
        output_dir = "data/uploads"
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f"export_{template_name}_{pd.Timestamp.now().timestamp():.0f}.xlsx")

        df.to_excel(output_file, index=False, engine='openpyxl')

        logger.info(f"导出成功: {output_file}, 共 {len(data)} 行")
        return output_file

    def export_to_bytes(self, data: List[Dict], manual_codes: Dict = None,
                        template_name: str = 'default') -> BytesIO:
        """
        导出为Excel字节流

        Args:
            data: 识别的数据列表
            manual_codes: 手工编码 {'supplier_code': '', 'category_code': ''}
            template_name: ERP模板名称

        Returns:
            Excel文件的字节流
        """
        if manual_codes is None:
            manual_codes = {}

        # 获取模板配置
        template = get_template(template_name)

        # 合并数据和编码
        merged_data = self._merge_data(data, manual_codes, template)

        # 按字段顺序导出
        df = pd.DataFrame(merged_data)
        # 确保所有模板字段都存在，不存在的填充空值
        df = df.reindex(columns=template['field_order'], fill_value='')

        # 保存到字节流
        output = BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        logger.info(f"导出字节流成功, 共 {len(data)} 行")
        return output

    def _merge_data(self, data: List[Dict], codes: Dict, template: Dict) -> List[Dict]:
        """
        合并数据和编码

        Args:
            data: 识别的数据列表
            codes: 手工编码
            template: ERP模板配置

        Returns:
            合并后的数据列表
        """
        merged_data = []

        for item in data:
            merged_item = {}

            # 如果没有字段映射，直接使用原字段名
            if not template['field_mapping']:
                merged_item = item.copy()
            else:
                # 映射字段名
                for src_field, dst_field in template['field_mapping'].items():
                    if src_field in item:
                        merged_item[dst_field] = item[src_field]

            # 添加手工编码
            for field in template.get('manual_fields', []):
                if field == '供应商编码':
                    merged_item[field] = codes.get('supplier_code', '')
                elif field == '类别编码':
                    merged_item[field] = codes.get('category_code', '')

            merged_data.append(merged_item)

        return merged_data

    def _detect_template_header_row(self, template_file: str) -> int:
        """
        自动检测模板文件的表头行
        跳过说明行、合并单元格等非数据行，找到包含实际字段名的行

        Args:
            template_file: 模板文件路径

        Returns:
            表头行索引（从0开始）
        """
        # 常见的字段名关键词，用于识别表头行
        header_keywords = [
            '货号', '品名', '条码', '名称', '规格', '单位',
            '进价', '售价', '零售价', '编码', '类别', '品牌'
        ]

        # 读取前10行，查找表头行
        for header_row in range(min(10, 20)):
            try:
                df = pd.read_excel(template_file, header=header_row, nrows=1)
                cols = [str(c).strip() for c in df.columns.tolist()]

                # 过滤无效列名（Unnamed、空值、纯数字）
                valid_cols = [
                    c for c in cols
                    if c and not c.startswith('Unnamed') and not c.replace('.', '').isdigit()
                ]

                # 检查有效列名中是否包含至少2个已知关键词
                match_count = sum(
                    1 for kw in header_keywords
                    if any(kw in col for col in valid_cols)
                )

                if match_count >= 2 and len(valid_cols) >= 3:
                    logger.info(f"检测到模板表头行: 第{header_row + 1}行, "
                                f"有效字段数: {len(valid_cols)}, "
                                f"关键词匹配数: {match_count}")
                    return header_row
            except Exception:
                continue

        # 默认使用第一行
        logger.warning(f"未检测到有效表头行，使用默认第1行")
        return 0

    def export_by_template_file(self, data: List[Dict], manual_codes: Dict = None,
                                template_file: str = None) -> str:
        """
        根据Excel模板文件导出数据

        Args:
            data: 识别的数据列表
            manual_codes: 手工编码 {'supplier_code': '', 'category_code': ''}
            template_file: Excel模板文件路径

        Returns:
            导出文件的路径
        """
        if manual_codes is None:
            manual_codes = {}

        if not template_file or not os.path.exists(template_file):
            raise ValueError(f"模板文件不存在: {template_file}")

        # 自动检测表头行
        header_row = self._detect_template_header_row(template_file)

        # 读取模板，获取字段顺序
        template_df = pd.read_excel(template_file, header=header_row)
        field_order = template_df.columns.tolist()

        # 过滤无效列名
        field_order = [
            str(f).strip() for f in field_order
            if str(f).strip() and not str(f).startswith('Unnamed')
        ]

        logger.info(f"读取模板文件: {template_file}, 表头行: {header_row + 1}, 字段数: {len(field_order)}")
        logger.info(f"模板字段: {field_order}")

        # 合并数据
        merged_data = self._merge_data_by_template(data, manual_codes, field_order)

        # 按模板顺序输出
        df = pd.DataFrame(merged_data)
        # 确保所有模板字段都存在，不存在的填充空值
        df = df.reindex(columns=field_order, fill_value='')

        # 保存到文件
        output_dir = "data/uploads"
        os.makedirs(output_dir, exist_ok=True)
        timestamp = pd.Timestamp.now().timestamp()
        output_file = os.path.join(output_dir, f"export_template_{int(timestamp)}.xlsx")

        df.to_excel(output_file, index=False, engine='openpyxl')

        logger.info(f"导出成功: {output_file}, 共 {len(data)} 行")
        return output_file

    def _match_field_to_template(self, semantic_field: str, template_fields: List[str]) -> str:
        """
        将语义字段匹配到模板中的实际字段名

        Args:
            semantic_field: 语义字段名（如 "条码"）
            template_fields: 模板字段名列表

        Returns:
            匹配到的模板字段名，未匹配到则返回语义字段名本身
        """
        keywords = self.SEMANTIC_KEYWORDS.get(semantic_field, [])
        if not keywords:
            return semantic_field

        for template_field in template_fields:
            template_lower = str(template_field).strip().lower()
            for keyword in keywords:
                if keyword.lower() == template_lower or keyword.lower() in template_lower:
                    return str(template_field).strip()

        return semantic_field

    def _merge_data_by_template(self, data: List[Dict], codes: Dict,
                                field_order: List[str]) -> List[Dict]:
        """
        根据模板字段智能合并数据
        基于关键词匹配，将源数据字段映射到模板实际字段名

        Args:
            data: 识别的数据列表
            codes: 手工编码
            field_order: 模板字段顺序（实际字段名列表）

        Returns:
            合并后的数据列表
        """
        merged_data = []

        # 预计算：每个语义字段匹配到哪个模板字段
        semantic_to_template = {}
        for src_field in self.FIELD_MAPPING.keys():
            matched = self._match_field_to_template(src_field, field_order)
            semantic_to_template[src_field] = matched

        logger.info(f"字段映射结果: {semantic_to_template}")

        for item in data:
            merged_item = {}

            # 按语义字段映射到模板实际字段名
            for src_field, template_field in semantic_to_template.items():
                if src_field in item and item[src_field]:
                    merged_item[template_field] = item[src_field]

            # 添加手工编码 - 智能匹配模板中的字段名
            supplier_field = self._match_field_to_template("供应商编码", field_order)
            category_field = self._match_field_to_template("类别编码", field_order)
            merged_item[supplier_field] = codes.get("supplier_code", "")
            merged_item[category_field] = codes.get("category_code", "")

            # 直接复制源数据中已存在的字段（如果模板也有该字段）
            for src_field, value in item.items():
                if src_field in self.FIELD_MAPPING:
                    continue
                # 如果源字段名恰好与模板字段名一致，直接使用
                if src_field in field_order and src_field not in merged_item:
                    merged_item[src_field] = value

            merged_data.append(merged_item)

        return merged_data