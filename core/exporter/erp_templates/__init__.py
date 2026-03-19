"""
ERP模板配置
定义各ERP系统的字段映射和导出规则
"""
from typing import Dict, List

# ERP模板定义
ERP_TEMPLATES = {
    'default': {
        'name': '默认模板',
        'field_order': ['条码', '品名', '规格', '单位', '进价', '零售价'],
        'field_mapping': {}
    },
    'sixun_shangyun_x': {
        'name': '思迅商云X',
        'field_order': [
            '商品编码', '商品名称', '规格型号', '计量单位',
            '进货单价', '销售单价', '供应商编码', '类别编码'
        ],
        'field_mapping': {
            '条码': '商品编码',
            '品名': '商品名称',
            '规格': '规格型号',
            '单位': '计量单位',
            '进价': '进货单价',
            '零售价': '销售单价'
        },
        'manual_fields': ['供应商编码', '类别编码']
    }
}

def get_template(template_name: str) -> Dict:
    """获取ERP模板配置"""
    return ERP_TEMPLATES.get(template_name, ERP_TEMPLATES['default'])

def get_supported_templates() -> Dict[str, str]:
    """获取支持的ERP模板列表"""
    return {k: v['name'] for k, v in ERP_TEMPLATES.items()}