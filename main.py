"""
Smart Excel Extractor - FastAPI主程序
高性能Excel数据提取Web服务
"""
from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io
from io import BytesIO
import os
import time
from typing import Optional
import logging

from core.recognizer import SmartRecognizer
from core.excel_handler import ExcelHandler
from core.exporter import ExcelExporter
from core.exporter.erp_templates import get_template

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# 创建FastAPI应用
app = FastAPI(
    title="Smart Excel Extractor",
    description="智能Excel数据提取系统 - 基于列结构映射",
    version="1.0.0"
)

# 配置CORS - 允许所有来源，支持域名访问
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 允许所有来源
    allow_credentials=True,  # 允许携带凭证
    allow_methods=["*"],  # 允许所有HTTP方法
    allow_headers=["*"],  # 允许所有请求头
)

# 挂载静态文件
app.mount("/static", StaticFiles(directory="static"), name="static")

# 模板引擎
templates = Jinja2Templates(directory="templates")

# 创建识别器实例
recognizer = SmartRecognizer()

# 确保上传目录存在
os.makedirs("data/uploads", exist_ok=True)


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """主页"""
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """
    上传并处理Excel文件

    支持格式: .xlsx, .xls, .xlsm, .xlsb, .csv, .tsv, .txt
    """
    start_time = time.time()

    try:
        logger.info(f"开始接收文件: {file.filename}")
        
        # 验证文件类型
        filename_str = file.filename or ""
        file_ext = os.path.splitext(filename_str)[1].lower()
        logger.info(f"文件扩展名: {file_ext}")
        
        if not ExcelHandler.is_supported(file_ext):
            logger.warning(f"不支持的文件格式: {file_ext}")
            raise HTTPException(
                status_code=400,
                detail=f"不支持的文件格式: {file_ext}。支持: {', '.join(ExcelHandler.get_supported_formats().keys())}"
            )

        # 读取文件
        contents = await file.read()
        logger.info(f"文件大小: {len(contents)} 字节")

        # 获取工作表信息
        worksheets = ExcelHandler.get_worksheets(contents, file_ext)
        logger.info(f"检测到工作表数量: {len(worksheets)}, 工作表详情: {[ws['name'] for ws in worksheets]}")

        # 检查有多少个工作表包含内容
        content_sheets = [ws for ws in worksheets if ws.get('has_content', True)]
        logger.info(f"包含内容的工作表数量: {len(content_sheets)}")

        # 如果有多个工作表，自动合并所有有数据的工作表
        if len(content_sheets) > 1:
            # 合并所有有数据的工作表
            logger.info(f"多个工作表有内容，自动合并。工作表: {[ws['name'] for ws in content_sheets]}")
            all_dfs = []
            
            # 导入识别器用于检测表头
            from core.recognizer import SmartRecognizer
            temp_recognizer = SmartRecognizer()
            
            # 首先确定标准列数（使用第一个工作表的列数）
            standard_cols = None
            first_df = ExcelHandler.read_sheet(
                contents,
                file_ext,
                sheet_index=content_sheets[0]['index'],
                sheet_name=content_sheets[0]['name']
            )
            standard_cols = len(first_df.columns)
            logger.info(f"标准列数: {standard_cols} (基于第一个工作表的原始列数)")
            
            for i, sheet in enumerate(content_sheets):
                try:
                    # 读取数据（不使用header参数，保持原始数据）
                    df = ExcelHandler.read_sheet(
                        contents,
                        file_ext,
                        sheet_index=sheet['index'],
                        sheet_name=sheet['name']
                    )
                    
                    # 检测表头行
                    header_idx = temp_recognizer.engine.detect_header_row(df)
                    
                    # 第一个工作表保留全部（包括表头），后续工作表跳过表头并使用第一个工作表的列名
                    if i == 0:
                        # 保存第一个工作表的列名
                        standard_columns = df.columns.tolist()
                        all_dfs.append(df)
                    else:
                        df = df.iloc[header_idx + 1:].reset_index(drop=True)
                        
                        # 如果列数不同，先添加空列到标准列数
                        if len(df.columns) < len(standard_columns):
                            cols_to_add = len(standard_columns) - len(df.columns)
                            for j in range(cols_to_add):
                                df.insert(9, '', '')
                        
                        # 使用第一个工作表的列名
                        df.columns = standard_columns
                        all_dfs.append(df)
                except Exception as e:
                    logger.warning(f"读取工作表 {sheet['name']} 失败: {e}")
                    continue

            if all_dfs:
                # 合并所有DataFrame
                df = pd.concat(all_dfs, ignore_index=True)
                logger.info(f"合并后总行数: {len(df)}, 来源工作表数: {len(all_dfs)}, 列数: {len(df.columns)}")
            else:
                raise HTTPException(status_code=400, detail="所有工作表都为空或读取失败")
        elif len(content_sheets) == 1:
            # 只有一个工作表有内容，直接处理
            target_sheet = content_sheets[0]
            logger.info(f"自动选择工作表: {target_sheet['name']} (包含内容)")
            df = ExcelHandler.read_sheet(
                contents,
                file_ext,
                sheet_index=target_sheet['index'],
                sheet_name=target_sheet['name']
            )
            logger.info(f"已读取工作表: {target_sheet['name']}, 行数: {len(df)}")
        else:
            # 没有工作表有内容，尝试读取第一个工作表
            if len(worksheets) > 0:
                logger.info(f"没有检测到明显内容，尝试读取第一个工作表")
                df = ExcelHandler.read_sheet(contents, file_ext, sheet_index=0)
                logger.info(f"已读取工作表，行数: {len(df)}")
            else:
                raise HTTPException(status_code=400, detail="文件为空或无法读取")

        logger.info(f"文件读取成功: {file.filename}, 大小: {len(df)} 行")

        # 处理数据
        result = recognizer.process(df)

        # 添加表头信息（用于手动指正）
        # 从DataFrame获取表头，但需要使用识别到的表头行
        header_row_idx = result.get('header_row', 0)
        if header_row_idx is not None and header_row_idx < len(df):
            headers = df.iloc[header_row_idx].astype(str).tolist()
        else:
            headers = df.columns.tolist() if hasattr(df, 'columns') else []
        result['headers'] = headers

        # 添加列预览数据（从表头行之后的数据中获取前3行）
        column_previews = {}
        if hasattr(df, 'columns'):
            data_start_row = header_row_idx + 1 if header_row_idx is not None else 0
            for col_idx in range(len(df.columns)):
                preview_data = []
                # 从表头行之后开始读取数据
                for i in range(data_start_row, min(data_start_row + 10, len(df))):
                    value = df.iloc[i, col_idx]
                    if value is not None and str(value).strip() != '':
                        preview_data.append(str(value))
                    if len(preview_data) >= 3:
                        break
                column_previews[str(col_idx)] = ', '.join(preview_data) if preview_data else '(无数据)'
        result['column_previews'] = column_previews

        # 添加性能统计
        processing_time = time.time() - start_time
        result['stats']['processing_time'] = round(processing_time, 3)
        extracted_rows = result['stats'].get('extracted_rows', 0)
        result['stats']['speed'] = round(extracted_rows / max(processing_time, 0.001), 2)
        result['filename'] = file.filename

        # 添加工作表信息
        if len(content_sheets) > 1:
            result['sheet_name'] = f"合并 {len(content_sheets)} 个工作表"
            result['sheets_info'] = [{'name': ws['name'], 'rows': ws['rows']} for ws in content_sheets]
        else:
            result['sheet_name'] = worksheets[0]['name'] if worksheets else 'Data'

        logger.info(f"处理完成: {file.filename}, 耗时: {processing_time:.3f}秒, 速度: {result['stats']['speed']} 行/秒")

        return JSONResponse(content=result)

    except HTTPException:
        # HTTP异常直接抛出
        raise
    except Exception as e:
        logger.error(f"处理文件失败: {e}", exc_info=True)
        logger.error(f"文件名: {file.filename if file else 'Unknown'}")
        raise HTTPException(status_code=500, detail=f"处理失败: {str(e)}")


@app.post("/api/process-sheet")
async def process_sheet(data: dict):
    """
    处理指定的工作表

    请求体:
    {
        "filename": "example.xlsx",
        "sheet_index": 0,
        "sheet_name": "Sheet1",
        "file_data": "base64_encoded_file_data"  # 可选，如果文件已缓存
    }
    """
    start_time = time.time()

    try:
        filename = data.get('filename')
        sheet_index = data.get('sheet_index', 0)
        sheet_name = data.get('sheet_name')
        file_data = data.get('file_data')

        if not filename:
            raise HTTPException(status_code=400, detail="缺少文件名")

        # 解码文件数据
        import base64
        if file_data is None:
            raise HTTPException(status_code=400, detail="缺少文件数据")
        contents = base64.b64decode(file_data)

        file_ext = os.path.splitext(filename)[1].lower()
        if not ExcelHandler.is_supported(file_ext):
            raise HTTPException(status_code=400, detail=f"不支持的文件格式: {file_ext}")

        # 读取指定工作表
        df = ExcelHandler.read_sheet(
            contents,
            file_ext,
            sheet_index=sheet_index,
            sheet_name=sheet_name
        )

        logger.info(f"工作表读取成功: {filename} - {sheet_name or f'Sheet{sheet_index}'}, 大小: {len(df)} 行")

        # 处理数据
        result = recognizer.process(df)

        # 添加表头信息（用于手动指正）
        # 从DataFrame获取表头，但需要使用识别到的表头行
        header_row_idx = result.get('header_row', 0)
        if header_row_idx is not None and header_row_idx < len(df):
            headers = df.iloc[header_row_idx].astype(str).tolist()
        else:
            headers = df.columns.tolist() if hasattr(df, 'columns') else []
        result['headers'] = headers

        # 添加列预览数据（从表头行之后的数据中获取前3行）
        column_previews = {}
        if hasattr(df, 'columns'):
            data_start_row = header_row_idx + 1 if header_row_idx is not None else 0
            for col_idx in range(len(df.columns)):
                preview_data = []
                # 从表头行之后开始读取数据
                for i in range(data_start_row, min(data_start_row + 10, len(df))):
                    value = df.iloc[i, col_idx]
                    if value is not None and str(value).strip() != '':
                        preview_data.append(str(value))
                    if len(preview_data) >= 3:
                        break
                column_previews[str(col_idx)] = ', '.join(preview_data) if preview_data else '(无数据)'
        result['column_previews'] = column_previews

        # 添加性能统计
        processing_time = time.time() - start_time
        result['stats']['processing_time'] = round(processing_time, 3)
        result['stats']['speed'] = round(result['stats']['extracted_rows'] / max(processing_time, 0.001), 2)
        result['filename'] = filename
        result['sheet_name'] = sheet_name or f'Sheet{sheet_index}'

        logger.info(f"处理完成: {filename} - {sheet_name or f'Sheet{sheet_index}'}, "
                   f"耗时: {processing_time:.3f}秒, 速度: {result['stats']['speed']} 行/秒")

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"处理工作表失败: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"处理失败: {str(e)}")


@app.post("/api/apply-correction")
async def apply_correction(data: dict):
    """
    应用手动指正并重新识别数据

    参数:
        filename: 文件名
        file_data: 文件数据(base64)
        corrected_mapping: 指正后的映射
    """
    try:
        filename = data.get('filename')
        file_data = data.get('file_data')
        corrected_mapping = data.get('corrected_mapping')

        logger.info(f"收到指正请求: filename={filename}")
        logger.info(f"corrected_mapping keys: {list(corrected_mapping.keys()) if corrected_mapping else 'None'}")
        logger.info(f"file_data length: {len(file_data) if file_data else 0}")

        if not filename or not file_data:
            logger.error("缺少文件名或文件数据")
            raise HTTPException(status_code=400, detail="缺少文件名或文件数据")
        
        if not corrected_mapping:
            logger.error("缺少指正映射")
            raise HTTPException(status_code=400, detail="缺少指正映射")

        # 解码文件数据
        import base64
        try:
            contents = base64.b64decode(file_data)
            logger.info(f"文件解码成功，大小: {len(contents)} 字节")
        except Exception as e:
            logger.error(f"文件解码失败: {e}")
            raise HTTPException(status_code=400, detail=f"文件解码失败: {str(e)}")

        file_ext = os.path.splitext(filename)[1].lower()
        logger.info(f"文件扩展名: {file_ext}")
        
        if not ExcelHandler.is_supported(file_ext):
            logger.error(f"不支持的文件格式: {file_ext}")
            raise HTTPException(status_code=400, detail=f"不支持的文件格式: {file_ext}")

        # 读取文件
        try:
            df = ExcelHandler.read_sheet(contents, file_ext)
            logger.info(f"文件读取成功，行数: {len(df)}, 列数: {len(df.columns)}")
        except Exception as e:
            logger.error(f"文件读取失败: {e}", exc_info=True)
            raise HTTPException(status_code=500, detail=f"文件读取失败: {str(e)}")

        # 使用指正后的映射提取数据
        from core.column_based_recognizer import ColumnBasedRecognizer
        engine = ColumnBasedRecognizer()
        
        # 检测表头行
        try:
            header_row_idx = engine.detect_header_row(df)
            logger.info(f"检测到表头行: {header_row_idx}")
        except Exception as e:
            logger.error(f"检测表头行失败: {e}", exc_info=True)
            header_row_idx = 0
        
        # 获取表头
        try:
            if header_row_idx < len(df):
                headers = df.iloc[header_row_idx].astype(str).tolist()
            else:
                headers = []
            logger.info(f"表头: {headers}")
        except Exception as e:
            logger.error(f"获取表头失败: {e}", exc_info=True)
            headers = []

        # 使用指正后的映射提取数据
        try:
            extracted_data = engine.extract_data(df, corrected_mapping, header_row_idx)
            logger.info(f"数据提取成功，提取{len(extracted_data)}条数据")
        except Exception as e:
            logger.error(f"数据提取失败: {e}", exc_info=True)
            raise HTTPException(status_code=500, detail=f"数据提取失败: {str(e)}")

        # 构建结果
        result = {
            "success": True,
            "header_row": header_row_idx,
            "mappings": corrected_mapping,
            "results": extracted_data,
            "stats": {
                "total_rows": len(df),
                "header_row": header_row_idx,
                "data_rows": len(df) - header_row_idx - 1 if header_row_idx >= 0 else len(df),
                "extracted_rows": len(extracted_data),
                "success_rate": round(len(extracted_data) / max(1, len(df) - header_row_idx - 1) * 100, 2) if header_row_idx >= 0 else 100.0,
            },
            "filename": filename
        }

        # 添加表头信息
        result['headers'] = headers

        # 添加列预览数据
        try:
            column_previews = {}
            if hasattr(df, 'columns'):
                data_start_row = header_row_idx + 1 if header_row_idx is not None else 0
                for col_idx in range(len(df.columns)):
                    preview_data = []
                    for i in range(data_start_row, min(data_start_row + 10, len(df))):
                        value = df.iloc[i, col_idx]
                        if value is not None and str(value).strip() != '':
                            preview_data.append(str(value))
                        if len(preview_data) >= 3:
                            break
                    column_previews[str(col_idx)] = ', '.join(preview_data) if preview_data else '(无数据)'
            result['column_previews'] = column_previews
        except Exception as e:
            logger.error(f"生成列预览失败: {e}", exc_info=True)
            result['column_previews'] = {}

        # 添加性能统计
        result['stats']['processing_time'] = 0.0
        result['stats']['speed'] = round(result['stats']['extracted_rows'] / max(0.001, 0.001), 2)

        logger.info(f"应用指正成功: {filename}, 提取{len(extracted_data)}条数据")

        return JSONResponse(content=result)

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"应用指正失败: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"应用指正失败: {str(e)}")


@app.post("/api/export")
async def export_results(data: dict):
    """
    导出提取结果为Excel文件

    支持参数:
        results: 提取的数据列表
        manual_codes: 手工编码 {'supplier_code': '', 'category_code': ''}
        erp_template: ERP模板名称（默认'default'）
        template_file: Excel模板文件路径（优先使用）
    """
    try:
        results = data.get('results', [])
        if not results:
            raise HTTPException(status_code=400, detail="没有可导出的数据")

        # 获取参数
        manual_codes = data.get('manual_codes', {})
        erp_template = data.get('erp_template', 'default')
        template_file = data.get('template_file')

        # 使用导出器导出
        exporter = ExcelExporter()

        # 如果提供了模板文件，使用模板文件导出
        if template_file:
            if not os.path.exists(template_file):
                raise HTTPException(status_code=400, detail=f"模板文件不存在: {template_file}")
            output_file = exporter.export_by_template_file(results, manual_codes, template_file)
            filename = f"导入模板_{os.path.basename(template_file)}"
        else:
            # 使用默认模板导出
            output_file = exporter.export(results, manual_codes, erp_template)
            template_name = get_template(erp_template)['name']
            filename = f"{template_name}_导入模板.xlsx"

        return FileResponse(
            output_file,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    except Exception as e:
        logger.error(f"导出失败: {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"导出失败: {str(e)}")


@app.get("/api/supported-formats")
async def get_supported_formats():
    """获取支持的文件格式"""
    formats = ExcelHandler.get_supported_formats()
    return {
        "formats": formats,
        "total": len(formats),
        "excel_formats": {k: v for k, v in formats.items() if k in ['.xlsx', '.xls', '.xlsm', '.xlsb', '.xltx', '.xltm']},
        "csv_formats": {k: v for k, v in formats.items() if k in ['.csv', '.tsv', '.txt']}
    }


@app.get("/api/health")
async def health_check():
    """健康检查"""
    return {"status": "ok", "service": "Smart Excel Extractor"}


@app.get("/api/stats")
async def get_stats():
    """获取系统统计"""
    return {
        "total_processed": 0,  # TODO: 从数据库获取
        "success_rate": 99.9,
        "avg_speed": 10000
    }


if __name__ == "__main__":
    import uvicorn
    import sys
    import io
    
    # 设置标准输出编码为UTF-8
    if sys.platform == 'win32':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    print("=" * 60)
    print("🚀 Smart Excel Extractor 启动中...")
    print("=" * 60)
    print(f"📍 访问地址: http://localhost:8000")
    print(f"📖 API文档: http://localhost:8000/docs")
    print("=" * 60)
    
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )