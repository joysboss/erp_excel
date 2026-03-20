"""测试天店模板导出"""
import sys
sys.path.insert(0, '/app')

from core.exporter.excel_exporter import ExcelExporter
import pandas as pd

# 模拟识别数据
test_data = [
    {'条码': '6901028137492', '品名': '七匹狼(红)条', '规格': '1*10', '单位': '条', '进价': 90, '零售价': 120},
    {'条码': '6901028137508', '品名': '七匹狼(红)', '规格': '20支', '单位': '盒', '进价': 9, '零售价': 15},
]
manual_codes = {'supplier_code': 'SUP001', 'category_code': '烟酒'}
exporter = ExcelExporter()

# 测试1: 表头检测
print('=== 测试1: 表头行检测 ===')
header_row = exporter._detect_template_header_row('/app/moban/天店商品导入模板.xlsx')
print(f'检测到表头行: 第{header_row + 1}行')
assert header_row == 1, f'期望表头行=1, 实际={header_row}'
print('PASS')

# 测试2: 字段匹配
print('\n=== 测试2: 字段匹配 ===')
template_fields = ['货号', '品名', '自编码', '规格', '单位', '类别编码', '类别', '品牌',
                   '供应商编码', '供应商', '进货价', '零售价', '会员价', '最低售价',
                   '批发价', '商品状态', '计价方式', '产地', '进货规格', '联营扣率',
                   '配送价', '商品类型', '允许积分', '允许折扣']
for semantic in ['条码', '品名', '规格', '单位', '进价', '零售价', '供应商编码', '类别编码']:
    matched = exporter._match_field_to_template(semantic, template_fields)
    print(f'  {semantic} -> {matched}')

# 测试3: 完整导出
print('\n=== 测试3: 完整导出 ===')
output = exporter.export_by_template_file(test_data, manual_codes, '/app/moban/天店商品导入模板.xlsx')
print(f'导出文件: {output}')

df = pd.read_excel(output)
print(f'导出字段: {df.columns.tolist()}')
print(f'数据行数: {len(df)}')
print(f'\n第1行数据:')
for col in df.columns:
    val = df.iloc[0][col]
    if pd.notna(val) and str(val).strip():
        print(f'  {col}: {val}')

# 验证
assert '货号' in df.columns, '缺少货号列'
assert '进货价' in df.columns, '缺少进货价列'
assert str(df.iloc[0]['货号']) == '6901028137492', '货号数据错误'
assert str(df.iloc[0]['品名']) == '七匹狼(红)条', '品名数据错误'
assert str(df.iloc[0]['进货价']) == '90', '进价数据错误'
assert str(df.iloc[0]['零售价']) == '120', '零售价数据错误'
print('\nALL TESTS PASSED')
