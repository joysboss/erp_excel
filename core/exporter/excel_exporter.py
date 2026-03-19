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

    # 字段映射关系
    FIELD_MAPPING = {
        "条码": "条码",
        "品名": "名称",
        "规格": "规格",
        "单位": "单位",
        "进价": "进价",
        "零售价": "零售价"
    }

    # 默认值
    DEFAULT_VALUES = {
        "自编码": "0",
        "批发价": "",
        "会员价": "",
        "配送价": "",
        "联营扣率": "",
        "进货规格": "",
        "产地": "",
        "计价方式": "",
        "是否积分": "",
        "前台议价": "",
        "前台折扣": "",
        "门店变价": "",
        "助记码": "",
        "经营方式": "",
        "品牌编码": "",
        "保质期": "",
        "课组": ""
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

        # 读取模板，获取字段顺序
        template_df = pd.read_excel(template_file)
        field_order = template_df.columns.tolist()

        logger.info(f"读取模板文件: {template_file}, 字段数: {len(field_order)}")

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

    def _merge_data_by_template(self, data: List[Dict], codes: Dict,
                                field_order: List[str]) -> List[Dict]:
        """
        根据模板字段合并数据

        Args:
            data: 识别的数据列表
            codes: 手工编码
            field_order: 模板字段顺序

        Returns:
            合并后的数据列表
        """
        merged_data = []

        for item in data:
            merged_item = {}

            # 映射字段
            for src_field, dst_field in self.FIELD_MAPPING.items():
                if src_field in item:
                    merged_item[dst_field] = item[src_field]

            # 添加手工编码
            merged_item["类别编码"] = codes.get("category_code", "")
            merged_item["主供应商编码"] = codes.get("supplier_code", "")

            # 添加默认值
            for field, default_value in self.DEFAULT_VALUES.items():
                if field not in merged_item:
                    merged_item[field] = default_value

            merged_data.append(merged_item)

        return merged_data