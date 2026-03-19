"""
基于列结构的智能识别器
通过分析Excel列结构，将特定列映射到商品资料字段，实现精确识别
"""

import re
import json
import os
from typing import Dict, List, Any, Optional
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

class ColumnBasedRecognizer:
    """基于列结构的商品资料智能识别器"""
    
    @staticmethod
    def _load_config() -> tuple[Dict[str, List[str]], Dict[str, str]]:
        """
        从JSON文件加载配置，如果失败则返回默认配置（降级方案）
        
        Returns:
            (field_keywords, unit_map) 元组
        """
        # 默认配置（降级方案）
        default_field_keywords = {
            '品名': ['名称', '产品', '品名', '商品', '货品', '物品', '名字', '标题'],
            '规格': ['规格', '尺寸', '参数', '特性', '描述', '说明'],
            '型号': ['型号', 'model', 'type', 'type_no', '型号编号'],
            '单位': ['单位', '计量单位', '包装单位', '计量', '单位名称'],
            '数量': ['数量', '库存', '存量', '余量', '总量', 'qty'],
            '进价': ['进价', '采购价', '进货价', '买入价', '成本价', '单价', '价格'],
            '零售价': ['零售价', '售价', '销售价', '定价', '零售金额'],
            '条码': ['条形码', '条码', 'Barcode', 'barcode', '货号'],
            '分类': ['分类', '类别', '种类', '类型', '品类', '分组'],
            '品牌': ['品牌', '商标', '牌子', '厂牌', '制造商'],
            '供应商': ['供应商', '供货商', '厂商', '厂家', '提供商'],
        }
        
        default_unit_map = {
            '米': '米', '厘米': '厘米', '毫米': '毫米', '千克': '千克', '克': '克',
            '个': '个', '件': '件', '套': '套', '台': '台', '盒': '盒',
            '升': '升', '毫升': '毫升',
        }
        
        # 尝试从JSON文件加载配置
        try:
            # 获取配置文件路径
            current_file = Path(__file__)
            config_path = current_file.parent / "field_mapping_config.json"
            
            if not config_path.exists():
                logger.warning(f"配置文件不存在: {config_path}，使用默认配置")
                return default_field_keywords, default_unit_map
            
            # 读取并解析JSON文件
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 验证配置结构
            if 'field_keywords' not in config or 'unit_map' not in config:
                logger.warning(f"配置文件结构无效: {config_path}，使用默认配置")
                return default_field_keywords, default_unit_map
            
            # 验证数据类型
            if not isinstance(config['field_keywords'], dict) or not isinstance(config['unit_map'], dict):
                logger.warning(f"配置文件数据类型无效: {config_path}，使用默认配置")
                return default_field_keywords, default_unit_map
            
            logger.info(f"成功加载配置文件: {config_path}")
            return config['field_keywords'], config['unit_map']
            
        except json.JSONDecodeError as e:
            logger.error(f"配置文件JSON解析失败: {e}，使用默认配置")
            return default_field_keywords, default_unit_map
        except Exception as e:
            logger.error(f"加载配置文件失败: {e}，使用默认配置")
            return default_field_keywords, default_unit_map
    
    def __init__(self):
        # 从配置文件加载字段关键词映射（支持降级方案）
        self.field_keywords, self.unit_map = self._load_config()

    def detect_header_row(self, df) -> Optional[int]:
        """
        检测表头所在的行号
        通过分析每行的内容特征来判断哪一行最可能是表头
        
        策略:
        1. 扫描前50行（支持复杂的表单格式）
        2. 统计每行的非空文本单元格数量
        3. 非空文本单元格比例最高的行即为表头
        """
        best_row = 0
        best_score = 0
        
        for idx in range(min(50, len(df))):  # 检查前50行
            row_values = df.iloc[idx].astype(str).values
            
            # 统计非空文本单元格数量
            text_count = 0
            non_empty_count = 0
            
            for val in row_values:
                if val and val != 'nan':
                    non_empty_count += 1
                    if self._is_text_like(val):
                        text_count += 1
            
            # 计算分数：文本比例 * 非空比例
            if non_empty_count > 0:
                score = (text_count / non_empty_count) * (non_empty_count / len(row_values))
            else:
                score = 0
            
            # 优先选择非空文本单元格较多的行
            if text_count > 0 and score > best_score:
                best_score = score
                best_row = idx
        
        return best_row if best_score > 0.3 else 0

    def _is_text_like(self, value: str) -> bool:
        """判断值是否像文本（而非数值）"""
        value = str(value).strip().lower()
        if value == '' or value == 'nan' or value == 'none':
            return False
        # 检查是否主要是数字或数字+单位
        digit_ratio = sum(1 for c in value if c.isdigit()) / max(len(value), 1)
        if digit_ratio > 0.7:  # 如果数字占比超过70%，则认为不是文本
            return False
        return True

    def map_columns(self, header_row: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        将表头列映射到标准字段
        返回字段名到映射信息的映射

        每个字段只映射到置信度最高的列
        使用贪心算法解决冲突：优先处理最精确的字段（置信度、关键词长度、匹配数量）
        """
        # 存储所有可能的匹配: {field_name: [(col_idx, col_name, confidence, best_keyword_len, match_count), ...]}
        all_matches = {}

        # 第一步：收集所有可能的匹配
        for col_idx, col_name in enumerate(header_row):
            col_name_clean = str(col_name).strip().replace(' ', '')  # 去除所有空格，处理"规  格"等
            if not col_name_clean or col_name_clean.lower() == 'nan':
                continue

            # 对每个字段类型检查关键词匹配
            for field_name, keywords in self.field_keywords.items():
                is_match, confidence, match_count, keyword_priority = self._match_field_keyword(col_name_clean, keywords)
                if is_match:
                    # 找到匹配的关键词长度
                    best_keyword_len = 0
                    for keyword in keywords:
                        keyword_clean = keyword.replace(' ', '')  # 去除关键词中的空格
                        if keyword_clean.lower() in col_name_clean.lower():
                            if len(keyword_clean) > best_keyword_len:
                                best_keyword_len = len(keyword_clean)

                    if field_name not in all_matches:
                        all_matches[field_name] = []
                    all_matches[field_name].append({
                        'column_index': col_idx,
                        'column_name': col_name_clean,
                        'confidence': confidence,
                        'keyword_len': best_keyword_len,
                        'match_count': match_count,  # 添加匹配的关键词数量
                        'keyword_priority': keyword_priority  # 添加关键词优先级
                    })

        # 第二步：为每个字段选择最佳匹配列
        field_best_match = {}
        for field_name, matches in all_matches.items():
            if matches:
                # 选择该字段的最佳匹配列（置信度、关键词优先级、关键词长度、匹配数量）
                # 优先使用关键词优先级（keyword_priority），然后才是置信度和关键词长度
                best_match = max(matches, key=lambda x: (
                    x['keyword_priority'],    # 关键词优先级最高
                    x['confidence'],           # 置信度
                    x['keyword_len'],          # 关键词长度
                    x['match_count']           # 匹配数量
                ))
                field_best_match[field_name] = best_match

        # 第三步：使用贪心算法解决冲突
        # 按匹配质量排序（字段优先级、置信度、关键词长度、匹配数量）
        def _get_field_priority(field_name):
            """获取字段优先级（核心字段优先级更高）"""
            priority_map = {
                # 核心字段（最高优先级）
                '条码': 100,
                '品名': 95,
                '进价': 90,
                '零售价': 85,
                '规格': 80,
                '单位': 75,
                '数量': 70,
                # 常见字段
                '分类': 60,
                '品牌': 55,
                '供应商': 50,
                '型号': 45,
                '产地': 40,
                # 扩展字段（低优先级）
                '简称': 30,
                '批发价': 25,
                '会员价': 20,
                '配送价': 15,
                '保质期': 10,
                '生鲜标志': 5,
            }
            return priority_map.get(field_name, 0)

        sorted_fields = sorted(
            field_best_match.items(),
            key=lambda x: (
                _get_field_priority(x[0]),        # 字段优先级
                x[1]['confidence'],              # 置信度
                x[1]['keyword_len'],             # 关键词长度
                x[1]['match_count']              # 匹配数量
            ),
            reverse=True
        )

        column_mapping = {}
        used_columns = set()

        for field_name, match_info in sorted_fields:
            col_idx = match_info['column_index']
            if col_idx not in used_columns:
                # 该列未被占用，可以使用
                column_mapping[field_name] = match_info
                used_columns.add(col_idx)
            else:
                # 该列已被占用，跳过该字段
                pass

        return column_mapping

    def _match_field_keyword(self, col_name: str, keywords: List[str]) -> tuple[bool, float, int, float]:
        """
        检查列名是否匹配关键词

        返回: (是否匹配, 置信度, 匹配的关键词数量, 关键词优先级)
        置信度范围: 0.0-1.0，精确匹配为1.0，包含匹配根据匹配程度计算
        关键词优先级: 0-100，数字越大优先级越高
        """
        # 去除所有空格，处理"规  格"等变体
        col_name_clean = col_name.lower().replace(' ', '')

        # 定义条码主要关键词的优先级（数字越大优先级越高）
        barcode_primary_priority = {
            '条形码': 100,
            '商品条码': 95,
            '产品条码': 95,
            '条码': 90,
            'barcode': 80,
        }

        # 1. 精确匹配（置信度最高）
        for keyword in keywords:
            keyword_clean = keyword.lower().replace(' ', '')
            if keyword_clean == col_name_clean:
                # 如果匹配到条码主要关键词，根据优先级调整置信度
                if keyword_clean in barcode_primary_priority:
                    priority = barcode_primary_priority[keyword_clean]
                    # 优先级越高，置信度越接近1.0
                    confidence = 0.95 + (priority / 1000.0)  # 0.95 ~ 1.05
                    confidence = min(confidence, 1.0)
                    return True, confidence, 1, priority
                else:
                    # 其他关键词使用默认置信度和优先级
                    return True, 0.90, 1, 0

        # 2. 包含匹配（根据匹配长度计算置信度）
        best_match = False
        best_confidence = 0.0
        match_count = 0
        best_priority = 0

        for keyword in keywords:
            keyword_clean = keyword.lower().replace(' ', '')
            if keyword_clean in col_name_clean:
                # 计算置信度：关键词长度 / 列名长度
                confidence = len(keyword_clean) / max(len(col_name_clean), 1)
                # 如果匹配到条码主要关键词，根据优先级提高置信度
                if keyword_clean in barcode_primary_priority:
                    priority = barcode_primary_priority[keyword_clean]
                    confidence = min(confidence * (1 + priority/1000.0), 0.98)
                    if confidence > best_confidence:
                        best_priority = priority
                if confidence > best_confidence:
                    best_confidence = confidence
                    best_match = True
                match_count += 1

        return best_match, best_confidence, match_count, best_priority

    def extract_data(self, df, column_mapping: Dict[str, Dict[str, Any]], header_row_idx: int) -> List[Dict[str, Any]]:
        """
        根据列映射提取数据
        """
        extracted_data = []

        # 从表头行之后开始提取数据
        for row_idx in range(header_row_idx + 1, len(df)):
            row_data = df.iloc[row_idx]
            item = {}

            for field_name, mapping_info in column_mapping.items():
                col_idx = mapping_info['column_index']
                if col_idx < len(row_data):
                    value = row_data.iloc[col_idx] if hasattr(row_data, 'iloc') else row_data[col_idx]
                    value = self._clean_value(value)

                    if field_name == '单位':
                        # 特殊处理单位字段，进行标准化
                        value = self._standardize_unit(value)

                    item[field_name] = value

            # 只有当有有效数据时才添加到结果中
            if item:
                # 验证数据有效性
                if self._is_valid_row(item):
                    extracted_data.append(item)

        return extracted_data

    def _is_valid_row(self, item: Dict[str, Any]) -> bool:
        """
        检查数据行是否有效（宽松版本，适应手工录入表格）

        规则：
        1. 只要有品名且有效，就认为该行有效
        2. 品名不能包含"品名"、"商品名称"等关键词（避免误识别表头）
        3. 排除合计、总计、汇总等统计行
        4. 排除只包含数字或金额的行
        5. 其他字段可以为空
        """
        # 检查品名是否存在且有效
        product_name = item.get('品名', '')
        if product_name and product_name.strip():
            product_name_clean = product_name.strip()
            
            # 排除字段关键词本身
            if product_name_clean in ['品名', '商品名称', '名称', '产品', '商品', '货品']:
                return False
            
            # 排除合计、总计、汇总等统计行
            invalid_keywords = ['合计', '总计', '汇总', '总计金额', '合计金额', '小计', '累计', '总计数量']
            for keyword in invalid_keywords:
                if keyword in product_name_clean:
                    return False
            
            # 检查是否为纯数字（可能是条码列被误识别为品名）
            if product_name_clean.isdigit() and len(product_name_clean) >= 8:
                return False
            
            # 检查是否为纯小数（可能是金额被误识别为品名）
            try:
                float(product_name_clean)
                # 如果能转换为数字，很可能是统计行
                return False
            except ValueError:
                pass
            
            # 检查是否为统计关键词（如"优惠金额"、"欠款金额"等）
            invalid_keywords = ['优惠金额', '欠款金额', '应收金额', '应付金额', '实收金额', '实付金额']
            for keyword in invalid_keywords:
                if keyword in product_name_clean:
                    return False
            
            return True

        return False

    def _clean_value(self, value: Any) -> str:
        """清理单元格值"""
        if value is None or str(value).lower() in ['nan', 'none', '']:
            return ""
        return str(value).strip()

    def _standardize_unit(self, unit_value: str) -> str:
        """标准化单位值"""
        if not unit_value:
            return unit_value
            
        unit_value = str(unit_value).strip()
        # 尝试在单位映射中查找
        for original, standardized in self.unit_map.items():
            if original.lower() in unit_value.lower() or original in unit_value:
                return standardized
        
        # 如果没有找到匹配项，返回原值
        return unit_value

    def process_dataframe(self, df):
        """
        处理DataFrame并返回识别结果
        """
        try:
            # 检测表头行
            header_row_idx = self.detect_header_row(df)

            # 确保header_row_idx不为None
            if header_row_idx is None:
                header_row_idx = 0

            # 获取表头
            if header_row_idx < len(df):
                header_row = df.iloc[header_row_idx].astype(str).tolist()
            else:
                header_row = []

            # 映射列到字段
            column_mapping = self.map_columns(header_row)

            # 提取数据
            extracted_data = self.extract_data(df, column_mapping, header_row_idx)

            # 构建结果
            result = {
                "results": extracted_data,
                "mappings": column_mapping,  # 直接使用映射信息（已包含置信度）
                "stats": {
                    "data_rows": len(df),
                    "header_row": header_row_idx,
                    "columns_detected": len(column_mapping),
                    "extracted_rows": len(extracted_data),
                    "success_rate": 100.0 if len(df) > 0 else 0.0,
                }
            }

            return result

        except Exception as e:
            logger.error(f"处理DataFrame失败: {e}", exc_info=True)
            # 返回基本结构以避免错误
            return {
                "results": [],
                "mappings": {},
                "stats": {
                    "data_rows": len(df) if len(df) > 0 else 0,
                    "header_row": -1,
                    "columns_detected": 0,
                    "extracted_rows": 0,
                    "success_rate": 0.0,
                }
            }