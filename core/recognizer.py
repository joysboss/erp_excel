"""
列结构识别引擎
基于Excel列结构进行智能识别，准确率接近100%
"""
import pandas as pd
import re
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, asdict
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@dataclass
class ColumnMapping:
    """列映射信息"""
    column_index: int
    column_name: str
    field_type: str
    confidence: float
    
    def to_dict(self):
        return asdict(self)


@dataclass
class ExtractionResult:
    """提取结果"""
    条码: Optional[str] = None
    品名: Optional[str] = None
    进价: Optional[float] = None
    零售价: Optional[float] = None
    单位: Optional[str] = None
    规格: Optional[str] = None
    row_index: Optional[int] = None
    confidence: float = 1.0
    
    def to_dict(self):
        return {k: v for k, v in asdict(self).items() if k != 'row_index'}


class SmartRecognizer:
    """兼容旧接口的轻量包装器，委托给 ColumnBasedRecognizer 实现。"""

    def __init__(self):
        from .column_based_recognizer import ColumnBasedRecognizer
        self.engine = ColumnBasedRecognizer()

    def process(self, df: pd.DataFrame, skip_rows: int = 0) -> Dict[str, Any]:
        """调用底层 ColumnBasedRecognizer 并将结果适配为原有格式。"""
        try:
            res = self.engine.process_dataframe(df if skip_rows == 0 else df.iloc[skip_rows:].reset_index(drop=True))

            mappings = res.get('mappings', {})
            results = res.get('results', [])
            header_row = res.get('stats', {}).get('header_row', 0)

            # 计算额外统计信息以兼容旧接口
            if header_row is not None:
                data_start_row = header_row + 1
                total_data_rows = max(0, len(df) - data_start_row)
            else:
                # 如果无法检测到表头行，使用默认值
                data_start_row = 0
                total_data_rows = len(df)

            extracted_rows = len(results)
            missing_fields = [f for f in ("条码", "品名") if f not in mappings]

            stats = {
                'total_rows': len(df),
                'header_row': header_row,
                'data_rows': total_data_rows,
                'extracted_rows': extracted_rows,
                'success_rate': round(extracted_rows / max(1, total_data_rows) * 100, 2),
                'mapped_fields': list(mappings.keys()),
                'missing_fields': missing_fields
            }

            return {
                'success': True,
                'header_row': header_row,
                'mappings': mappings,  # 新格式：{field_name: {column_index, column_name, confidence}}
                'results': results,
                'stats': stats,
                'error': None
            }

        except Exception as e:
            logger.error(f"处理失败: {e}", exc_info=True)
            return {'success': False, 'header_row': None, 'mappings': {}, 'results': [], 'stats': {}, 'error': str(e)}