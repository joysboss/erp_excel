# ERP 模板使用指南

## 概述

Smart Excel Extractor 支持多种 ERP 系统的导出模板，自动将识别的数据转换为各 ERP 系统所需的格式。

---

## 支持的 ERP 模板

### 1. 默认模板

**适用场景**: 通用数据导出

**字段列表**:
```
条码、品名、规格、单位、进价、零售价
```

**特点**:
- 最简单的导出格式
- 包含基本的商品信息
- 适用于数据查看和初步处理

---

### 2. 思迅天店

**适用场景**: 思迅天店零售管理系统

**字段列表**:
```
商品编码、商品名称、规格型号、计量单位、
进货价、销售价、供应商编码、商品类别、
品牌、产地、保质期(天)
```

**字段映射**:
| 识别字段 | 导出字段 | 说明 |
|---------|---------|------|
| 条码 | 商品编码 | 商品唯一标识 |
| 品名 | 商品名称 | 商品名称 |
| 规格 | 规格型号 | 规格描述 |
| 单位 | 计量单位 | 计量单位 |
| 进价 | 进货价 | 采购价格 |
| 零售价 | 销售价 | 销售价格 |
| 品牌 | 品牌 | 品牌名称（可选）|
| 产地 | 产地 | 产地信息（可选）|
| 保质期 | 保质期(天) | 保质期天数（可选）|

**手工填写字段**:
- **供应商编码**: 必填，在导出界面手动输入
- **商品类别**: 必填，在导出界面手动输入

**默认值**:
- 品牌: 空
- 产地: 空
- 保质期(天): 空

**使用步骤**:
1. 上传 Excel 文件并识别
2. 在导出配置中选择"思迅天店"模板
3. 填写供应商编码和商品类别
4. 点击"导出 Excel"
5. 导入到思迅天店系统

---

### 3. 思迅商云 X

**适用场景**: 思迅商云 X 企业管理系统

**字段列表**:
```
商品编码、商品名称、规格型号、计量单位、
进货单价、销售单价、供应商编码、类别编码
```

**字段映射**:
| 识别字段 | 导出字段 | 说明 |
|---------|---------|------|
| 条码 | 商品编码 | 商品唯一标识 |
| 品名 | 商品名称 | 商品名称 |
| 规格 | 规格型号 | 规格描述 |
| 单位 | 计量单位 | 计量单位 |
| 进价 | 进货单价 | 采购价格 |
| 零售价 | 销售单价 | 销售价格 |

**手工填写字段**:
- **供应商编码**: 必填
- **类别编码**: 必填

**模板文件**: `moban/商品档案模版.xls`

**使用步骤**:
1. 上传 Excel 文件并识别
2. 在导出配置中选择"思迅商云X（商品档案）"
3. 填写供应商编码和类别编码
4. 点击"导出 Excel"
5. 导入到思迅商云 X 系统

---

## 模板对比

| 特性 | 默认模板 | 思迅天店 | 思迅商云X |
|------|---------|---------|-----------|
| 字段数量 | 6 | 11 | 8 |
| 手工字段 | 无 | 2 个 | 2 个 |
| 支持品牌 | ❌ | ✅ | ❌ |
| 支持产地 | ❌ | ✅ | ❌ |
| 支持保质期 | ❌ | ✅ | ❌ |
| 适用系统 | 通用 | 思迅天店 | 思迅商云X |

---

## 使用指南

### Web 界面使用

**步骤 1: 上传文件**
```
访问 http://localhost:8000
拖拽或点击上传 Excel 文件
```

**步骤 2: 查看识别结果**
```
系统自动识别并显示:
- 列映射关系
- 提取的数据预览
- 统计信息（行数、成功率、速度）
```

**步骤 3: 手动指正（可选）**
```
如果自动识别不正确:
1. 在"手动指正配置"区域选择正确的列
2. 点击"应用指正"
3. 系统重新提取数据
```

**步骤 4: 配置导出**
```
在"导出配置"区域:
1. 选择 ERP 模板（默认/思迅天店/思迅商云X）
2. 填写供应商编码（思迅天店/思迅商云X必填）
3. 填写类别编码（思迅天店/思迅商云X必填）
```

**步骤 5: 导出数据**
```
点击"导出 Excel"按钮
下载生成的 Excel 文件
```

---

### API 调用使用

#### 示例 1: 使用默认模板

```python
import requests

# 上传文件
url = "http://localhost:8000/api/upload"
files = {"file": open("data.xlsx", "rb")}
response = requests.post(url, files=files)
result = response.json()

# 导出（默认模板）
export_url = "http://localhost:8000/api/export"
export_data = {
    "results": result['results'],
    "erp_template": "default"
}
export_response = requests.post(export_url, json=export_data)

# 保存文件
with open("output.xlsx", "wb") as f:
    f.write(export_response.content)
```

#### 示例 2: 使用思迅天店模板

```python
import requests

# 上传文件
url = "http://localhost:8000/api/upload"
files = {"file": open("data.xlsx", "rb")}
response = requests.post(url, files=files)
result = response.json()

# 导出（思迅天店模板）
export_url = "http://localhost:8000/api/export"
export_data = {
    "results": result['results'],
    "manual_codes": {
        "supplier_code": "SUP001",
        "category_code": "CAT001"
    },
    "erp_template": "sixun_tiandian"
}
export_response = requests.post(export_url, json=export_data)

# 保存文件
with open("tiandian_output.xlsx", "wb") as f:
    f.write(export_response.content)
```

#### 示例 3: 使用 Excel 模板文件

```python
import requests

# 上传文件
url = "http://localhost:8000/api/upload"
files = {"file": open("data.xlsx", "rb")}
response = requests.post(url, files=files)
result = response.json()

# 使用自定义模板文件导出
export_url = "http://localhost:8000/api/export"
export_data = {
    "results": result['results'],
    "manual_codes": {
        "supplier_code": "SUP001",
        "category_code": "CAT001"
    },
    "template_file": "moban/商品档案模版.xls"
}
export_response = requests.post(export_url, json=export_data)

# 保存文件
with open("custom_output.xlsx", "wb") as f:
    f.write(export_response.content)
```

---

## 添加自定义模板

### 方法一：修改配置文件

编辑 `core/exporter/erp_templates/__init__.py`:

```python
ERP_TEMPLATES = {
    # ... 现有模板 ...
    
    'my_custom_erp': {
        'name': '我的自定义ERP',
        'field_order': [
            '商品编码', '商品名称', '规格', '单位',
            '进价', '售价', '供应商', '类别'
        ],
        'field_mapping': {
            '条码': '商品编码',
            '品名': '商品名称',
            '规格': '规格',
            '单位': '单位',
            '进价': '进价',
            '零售价': '售价'
        },
        'manual_fields': ['供应商', '类别'],
        'defaults': {
            '供应商': '',
            '类别': ''
        }
    }
}
```

### 方法二：使用 Excel 模板文件

1. **准备模板文件**
   - 创建 Excel 文件，第一行为字段名
   - 保存到 `moban/` 目录

2. **使用模板文件导出**
   ```python
   export_data = {
       "results": result['results'],
       "template_file": "moban/我的模板.xls"
   }
   ```

### 方法三：前端添加选项

编辑 `templates/index.html`，在下拉框中添加选项:

```html
<select x-model="selectedTemplate" ...>
    <option value="default">默认模板</option>
    <option value="sixun_tiandian">思迅天店</option>
    <option value="moban/商品档案模版.xls">思迅商云X（商品档案）</option>
    <option value="my_custom_erp">我的自定义ERP</option>
</select>
```

---

## 字段映射规则

### 自动映射

系统根据关键词自动识别并映射字段：

| 识别字段 | 关键词 |
|---------|--------|
| 条码 | 条码、条形码、商品码、货号、编码、code、barcode |
| 品名 | 品名、商品名、商品全名、产品名、名称、商品 |
| 进价 | 进价、成本价、批发价、采购价、单价、cost |
| 零售价 | 零售价、售价、零售金额、销售价、价格、retail |
| 单位 | 单位、包装、计量单位 |
| 规格 | 规格、型号、容量、净含量 |

### 手动指正

如果自动映射不准确，可以在前端手动选择正确的列。

### 默认值填充

对于模板中的可选字段，系统自动填充默认值：

```python
'defaults': {
    '品牌': '',
    '产地': '',
    '保质期(天)': ''
}
```

---

## 注意事项

### 1. 必填字段

- **思迅天店**: 供应商编码、商品类别
- **思迅商云X**: 供应商编码、类别编码
- **默认模板**: 无必填字段

### 2. 数据验证

系统会自动验证：
- 条码格式（13位/69开头）
- 价格合理性（进价 ≤ 零售价）
- 必填字段完整性

### 3. 导出格式

- 文件格式: Excel (.xlsx)
- 编码: UTF-8
- 第一行: 字段名
- 后续行: 数据

### 4. 性能建议

- 单次导出建议不超过 10 万行
- 大文件建议分批处理
- 使用 Docker 部署可获得更好性能

---

## 故障排查

### 问题 1: 导出字段缺失

**原因**: 原始 Excel 中没有对应的列

**解决**:
- 使用手动指正功能选择正确的列
- 或在原始数据中添加缺失的列

### 问题 2: 手工编码未生效

**原因**: 未在导出配置中填写

**解决**:
- 确保在导出前填写供应商编码和类别编码
- 检查 API 请求中的 `manual_codes` 参数

### 问题 3: 模板选择无效

**原因**: 前端选择与后端配置不匹配

**解决**:
- 检查 `erp_templates/__init__.py` 中的模板名称
- 确保前端 value 与后端 key 一致

---

## 更新日志

### v1.0.0 (2026-02-06)

- ✅ 添加默认模板
- ✅ 添加思迅商云 X 模板
- ✅ 添加思迅天店模板
- ✅ 支持手工编码输入
- ✅ 支持自定义 Excel 模板文件

---

## 联系支持

- **项目地址**: https://github.com/joysboss/erp_excel
- **问题反馈**: https://github.com/joysboss/erp_excel/issues
- **文档更新**: 查看 `ERP_TEMPLATE_GUIDE.md`
